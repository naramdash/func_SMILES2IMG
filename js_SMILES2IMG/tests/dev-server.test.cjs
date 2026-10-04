const assert = require("node:assert/strict");
const { test } = require("node:test");
const vm = require("node:vm");
const webpack = require("webpack");
const WebpackDevServer = require("webpack-dev-server");
const configure = require("../webpack.config.js");

test("Webpack bundles SmilesDrawer and serves IMG without rendering assets or an image API", { timeout: 20000 }, async () => {
  // Test routing over loopback HTTP without installing a development certificate.
  const config = await configure({ WEBPACK_BUILD: true }, { mode: "development" });
  config.mode = "development";
  config.infrastructureLogging = { level: "none" };
  config.stats = "none";
  const server = new WebpackDevServer({
    ...config.devServer,
    server: "http", host: "127.0.0.1", port: 0,
    open: false, client: false, hot: false,
  }, webpack(config));
  try {
    await server.start();
    await new Promise((resolve) => server.middleware.waitUntilValid(resolve));
    const origin = `http://127.0.0.1:${server.server.address().port}`;
    const metadata = await (await fetch(`${origin}/functions.json`)).json();
    assert.equal(metadata.functions[0].name, "IMG");
    assert.equal(metadata.allowCustomDataForDataTypeAny, true);
    const functions = await (await fetch(`${origin}/functions.js`)).text();
    const registered = [];
    vm.runInNewContext(functions, { CustomFunctions: { associate: (id, fn) => registered.push({ id, fn }) } });
    assert.deepEqual(registered.map(({ id }) => id), ["IMG"]);
    assert.equal(typeof registered[0].fn, "function");
    const taskpane = await (await fetch(`${origin}/taskpane.html`)).text();
    assert.match(taskpane, /=SMILES\.IMG\(A2\)/);
    assert.match(taskpane, /functions\.js/);
    assert.match(taskpane, /\/beta\/office\.js/);
    assert.match(functions, /smiles-drawer/);
    assert.equal((await fetch(`${origin}/rdkit/RDKit_minimal.js`)).status, 404);
    assert.equal((await fetch(`${origin}/rdkit/RDKit_minimal.wasm`)).status, 404);
    assert.equal((await fetch(`${origin}/api/smiles.png?smiles=CCO`)).status, 404);
  } finally {
    await server.stop();
  }
});
