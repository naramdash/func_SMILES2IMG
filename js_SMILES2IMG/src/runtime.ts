// These checks confirm the library and stable prerequisites, not native BETA support.
export function hasLocalImagePrerequisites(): boolean {
  return (
    Office.context.requirements.isSetSupported("CustomFunctionsRuntime", "1.4") &&
    typeof Excel !== "undefined" &&
    Excel.CellValueType?.localImage === "LocalImage"
  );
}
