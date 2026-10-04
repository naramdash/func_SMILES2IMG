import { InvalidSmilesError, normalizeSmiles, renderMolecule } from "../rendering/molecule";
import { hasLocalImagePrerequisites } from "../runtime";

/**
 * Converts a SMILES string into a locally generated molecular image (Preview).
 * @customfunction IMG
 * @param smiles SMILES string, for example CCO, or a cell containing one.
 * @returns The molecular structure as a local PNG image.
 */
export async function img(smiles: string): Promise<Excel.LocalImageCellValue> {
  let input: string;
  try {
    input = normalizeSmiles(smiles);
  } catch {
    throw new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "Enter a SMILES string containing 1 to 2000 characters.");
  }
  if (!hasLocalImagePrerequisites()) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      "SMILES.IMG requires the Office.js Preview library and Microsoft 365 Excel with CustomFunctionsRuntime 1.4."
    );
  }
  try {
    const image = await renderMolecule(input);
    return {
      type: "LocalImage",
      image: { type: "PNG", data: image.pngBase64 },
      altText: `Molecular structure: ${input}`,
    };
  } catch (error) {
    throw new CustomFunctions.Error(
      error instanceof InvalidSmilesError ? CustomFunctions.ErrorCode.invalidValue : CustomFunctions.ErrorCode.notAvailable,
      error instanceof InvalidSmilesError ? "Invalid SMILES string." : "Cannot generate the molecular image with SmilesDrawer."
    );
  }
}
