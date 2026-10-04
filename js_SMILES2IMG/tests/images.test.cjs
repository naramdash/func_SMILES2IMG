const assert = require("node:assert/strict");
const { readFileSync } = require("node:fs");
const path = require("node:path");
const { test } = require("node:test");
const vm = require("node:vm");
const ts = require("typescript");
const { generateCustomFunctionsMetadata } = require("custom-functions-metadata");

function loadRuntime({ supported = true, preview = true, decodeFailures = 0, drawingFailure = false, encodingFailure = false } = {}) {
  const records = { drawn: 0, parsed: [], decoded: 0, rasterized: 0 };
  class FunctionError extends Error {
    constructor(code, message) { super(message); this.code = code; }
  }
  const smilesDrawer = {
    Parser: {
      parse(input) {
        records.parsed.push(input);
        if (input === "invalid") throw new Error("Parsing failed");
        return { smiles: input };
      },
    },
    SvgDrawer: class {
      constructor(options) {
        assert.equal(options.width, 300);
        assert.equal(options.height, 200);
      }
      draw(tree, target, theme) {
        assert.equal(target, null);
        assert.equal(theme, "light");
        if (drawingFailure) throw new Error("Drawing failed");
        records.drawn++;
        return { outerHTML: "<svg><path /></svg>" };
      }
    },
  };
  const sandbox = {
    Office: { context: { requirements: { isSetSupported: () => supported } } },
    Excel: { CellValueType: preview ? { localImage: "LocalImage" } : {} },
    CustomFunctions: { Error: FunctionError, ErrorCode: { invalidValue: "#VALUE!", notAvailable: "#N/A" } },
    fetch: () => { throw new Error("Image rendering must not make API requests"); },
    Image: class {
      async decode() {
        assert.match(this.src, /^data:image\/svg\+xml/);
        records.decoded++;
        if (records.decoded <= decodeFailures) throw new Error("Decoding failed");
        await Promise.resolve();
        this.ready = true;
      }
    },
    document: {
      createElement(tag) {
        assert.equal(tag, "canvas", "Must not load an external rendering script");
        return {
          getContext: () => ({
            fillRect() {},
            drawImage(bitmap, x, y, width, height) {
              assert.equal(bitmap.ready, true, "Image must be decoded before rasterization");
              assert.equal(width, 300);
              assert.equal(height, 200);
              records.rasterized++;
            },
          }),
          toDataURL: () => encodingFailure ? "data:," : "data:image/png;base64,iVBORw0KGgo=",
        };
      },
    },
  };
  const context = vm.createContext(sandbox);
  const modules = new Map();
  function load(file) {
    file = path.resolve(file);
    if (modules.has(file)) return modules.get(file);
    const exports = {};
    modules.set(file, exports);
    const js = ts.transpileModule(readFileSync(file, "utf8"), {
      compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2025 },
    }).outputText;
    vm.runInContext("(function(exports, require) {" + js + "\n})", context)(exports, (id) =>
      id === "smiles-drawer" ? { __esModule: true, default: smilesDrawer } : load(path.resolve(path.dirname(file), id + ".ts")));
    return exports;
  }
  return {
    ...load(path.join(__dirname, "../src/functions/functions.ts")),
    ...load(path.join(__dirname, "../src/rendering/molecule.ts")),
    records,
  };
}

test("IMG returns LocalImage with raw PNG Base64 without fetching scripts or images", async () => {
  const { img, records } = loadRuntime();
  const result = await img(" [NH4+].[Cl-] ");
  assert.equal(result.type, "LocalImage");
  assert.equal(result.image.type, "PNG");
  assert.equal(result.image.data, "iVBORw0KGgo=");
  assert.equal("address" in result, false);
  assert.match(result.altText, /\[NH4\+\]/);
  assert.equal(records.rasterized, 1);
});

test("concurrent cells share rendering and cached images avoid recomputation", async () => {
  const { renderMolecule, records } = loadRuntime();
  const [first, second] = await Promise.all([renderMolecule("CCO"), renderMolecule("CCO")]);
  assert.strictEqual(first, second);
  assert.strictEqual(await renderMolecule(" CCO "), first);
  await renderMolecule("CCN");
  assert.equal(records.drawn, 2);
  assert.equal(records.rasterized, 2);
});

test("invalid input and SMILES become #VALUE! and do not poison future calls", async () => {
  const { img } = loadRuntime();
  for (const input of [null, 123, "", " ", "C".repeat(2001), "invalid"]) {
    await assert.rejects(img(input), (error) => error.code === "#VALUE!");
  }
  assert.equal((await img("CCO")).type, "LocalImage");
});

test("unsupported runtime or missing Preview library becomes #N/A before rendering", async () => {
  for (const options of [{ supported: false }, { preview: false }]) {
    const { img, records } = loadRuntime(options);
    await assert.rejects(img("CCO"), (error) => error.code === "#N/A");
    assert.equal(records.drawn, 0);
  }
});

test("image decoding failure can be retried without caching a failed result", async () => {
  const { img, records } = loadRuntime({ decodeFailures: 1 });
  await assert.rejects(img("CCO"), (error) => error.code === "#N/A");
  assert.equal((await img("CCO")).type, "LocalImage");
  assert.equal(records.drawn, 2);
  assert.equal(records.rasterized, 1);
});

test("drawing and encoding failures become #N/A", async () => {
  for (const options of [{ drawingFailure: true }, { encodingFailure: true }]) {
    await assert.rejects(loadRuntime(options).img("CCO"), (error) => error.code === "#N/A");
  }
});

test("image cache keeps recently accessed results and evicts older entries", async () => {
  const { renderMolecule, records } = loadRuntime();
  for (let i = 1; i <= 128; i++) await renderMolecule("C".repeat(i));
  await renderMolecule("C");
  await renderMolecule("C".repeat(129));
  await renderMolecule("C");
  assert.equal(records.drawn, 129);
  await renderMolecule("CC");
  assert.equal(records.drawn, 130);
});

test("real SmilesDrawer parses charged, aromatic, and stereochemical SMILES and rejects malformed input", async () => {
  const { default: smilesDrawer } = await import("smiles-drawer");
  for (const input of ["CCO", "c1ccccc1", "C[C@H](O)C(=O)O", "[NH4+].[Cl-]"]) {
    assert.ok(smilesDrawer.Parser.parse(input));
  }
  for (const input of ["not-a-smiles", "C(C"]) {
    assert.throws(() => smilesDrawer.Parser.parse(input));
  }
});

test("metadata registers only IMG, enables rich values, and uses the SMILES namespace", async () => {
  const generated = await generateCustomFunctionsMetadata(path.join(__dirname, "../src/functions/functions.ts"));
  assert.deepEqual(generated.errors, []);
  const metadata = JSON.parse(generated.metadataJson);
  assert.equal(metadata.allowCustomDataForDataTypeAny, true);
  assert.deepEqual(metadata.functions.map((fn) => fn.id), ["IMG"]);
  assert.equal(metadata.functions[0].result.type || "any", "any");
  const manifest = readFileSync(path.join(__dirname, "../manifest.xml"), "utf8");
  assert.match(manifest, /id="Functions.Namespace" DefaultValue="SMILES"/);
});
