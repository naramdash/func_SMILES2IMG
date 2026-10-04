import { hasLocalImagePrerequisites } from "../runtime";
import { renderMolecule } from "../rendering/molecule";

document.getElementById("render")!.onclick = async () => {
  const button = document.getElementById("render") as HTMLButtonElement;
  const status = document.getElementById("preview-status")!;
  button.disabled = true;
  status.textContent = "분자 이미지를 생성하고 있습니다.";
  document.getElementById("preview-result")!.hidden = true;
  try {
    const smiles = (document.getElementById("smiles") as HTMLInputElement).value;
    const image = await renderMolecule(smiles);
    const pngUrl = `data:image/png;base64,${image.pngBase64}`;
    const preview = document.getElementById("preview") as HTMLImageElement;
    preview.src = pngUrl;
    preview.alt = `분자 구조: ${smiles.trim()}`;
    (document.getElementById("download-png") as HTMLAnchorElement).href = pngUrl;
    document.getElementById("preview-result")!.hidden = false;
    status.textContent = "이미지가 생성되었습니다.";
  } catch (error) {
    status.textContent = error instanceof Error ? error.message : "이미지를 생성하지 못했습니다.";
  } finally {
    button.disabled = false;
  }
};

Office.onReady(({ host }) => {
  if (host !== Office.HostType.Excel) return;
  document.getElementById("sideload-msg")!.hidden = true;
  document.getElementById("app-body")!.hidden = false;
  document.getElementById("status")!.textContent = hasLocalImagePrerequisites()
    ? "기본 요구사항을 충족합니다. 셀 이미지는 LocalImage Preview를 지원하는 Excel 빌드가 필요합니다."
    : "Office.js Preview와 CustomFunctionsRuntime 1.4를 지원하는 Microsoft 365 Excel이 필요합니다.";
});
