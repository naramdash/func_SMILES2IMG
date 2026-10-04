import SmilesDrawer from "smiles-drawer";

export interface MoleculeImage {
  pngBase64: string;
  width: number;
  height: number;
}

export class InvalidSmilesError extends Error {}

export function normalizeSmiles(smiles: string): string {
  if (typeof smiles !== "string" || !smiles.trim() || smiles.trim().length > 2000) {
    throw new InvalidSmilesError("Enter a SMILES string containing 1 to 2000 characters.");
  }
  return smiles.trim();
}

const cache = new Map<string, MoleculeImage>();
const pending = new Map<string, Promise<MoleculeImage>>();
const cacheLimit = 128;

export async function renderMolecule(smiles: string): Promise<MoleculeImage> {
  const input = normalizeSmiles(smiles);
  const cached = cache.get(input);
  if (cached) {
    cache.delete(input);
    cache.set(input, cached);
    return cached;
  }
  const running = pending.get(input);
  if (running) return running;

  const rendering = (async () => {
    let tree: ReturnType<typeof SmilesDrawer.Parser.parse>;
    try {
      tree = SmilesDrawer.Parser.parse(input);
    } catch {
      throw new InvalidSmilesError("Invalid SMILES string.");
    }

    // SmilesDrawer 2.x rasterizes its vector drawing asynchronously.
    // Decode the local image before encoding PNG to avoid returning a blank canvas.
    const width = 300;
    const height = 200;
    const drawer = new SmilesDrawer.SvgDrawer({ width, height });
    const drawing = drawer.draw(tree, null, "light");
    const bitmap = new Image();
    bitmap.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(drawing.outerHTML);
    await bitmap.decode();

    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    if (!context) throw new Error("Canvas 2D rendering is unavailable.");
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, width, height);
    context.drawImage(bitmap, 0, 0, width, height);
    const dataUrl = canvas.toDataURL("image/png");
    const prefix = "data:image/png;base64,";
    if (!dataUrl.startsWith(prefix) || dataUrl.length <= prefix.length) {
      throw new Error("Canvas could not encode a PNG image.");
    }
    const image: MoleculeImage = {
      pngBase64: dataUrl.slice(prefix.length),
      width,
      height,
    };
    cache.set(input, image);
    if (cache.size > cacheLimit) cache.delete(cache.keys().next().value!);
    return image;
  })();
  pending.set(input, rendering);
  try {
    return await rendering;
  } finally {
    pending.delete(input);
  }
}
