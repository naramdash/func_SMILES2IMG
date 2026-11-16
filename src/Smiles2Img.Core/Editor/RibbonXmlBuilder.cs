namespace Smiles2Img.Editor;

/// <summary>
/// Builds shared Fluent Ribbon XML and context-menu fragments across Excel, Word, and PowerPoint add-ins.
/// All imageMso attributes reference verified icons listed in AGENTS.md and tests/Program.cs.
/// </summary>
public static class RibbonXmlBuilder
{
    public static string BuildRibbonMoleculeInfoMenu(string appTag)
    {
        return $@"          <menu id=""menuRibbon{appTag}MoleculeInfo""
                label=""Molecule Info""
                screentip=""Copy Molecule Info &amp; Substance Properties""
                supertip=""Directly copy substance name, formula, MW, InChI, or open the detailed molecule properties window.""
                imageMso=""FileProperties""
                size=""large"">
            <button id=""btnRibbon{appTag}CopyName""
                    label=""Copy Substance Name""
                    imageMso=""FontSize""
                    onAction=""OnCopySubstanceName"" />
            <button id=""btnRibbon{appTag}CopyFormula""
                    label=""Copy Chemical Formula""
                    imageMso=""Subscript""
                    onAction=""OnCopyFormula"" />
            <button id=""btnRibbon{appTag}CopyMW""
                    label=""Copy Molecular Weight (MW)""
                    imageMso=""Calculator""
                    onAction=""OnCopyMolecularWeight"" />
            <button id=""btnRibbon{appTag}SubCopySmiles""
                    label=""Copy SMILES""
                    imageMso=""Copy""
                    onAction=""OnCopySmiles"" />
            <button id=""btnRibbon{appTag}SubCopyInchi""
                    label=""Copy InChI""
                    imageMso=""SignatureLineInsert""
                    onAction=""OnCopyInchi"" />
            <button id=""btnRibbon{appTag}SubCopyInchiKey""
                    label=""Copy InChIKey""
                    imageMso=""FilePermission""
                    onAction=""OnCopyInchiKey"" />
            <button id=""btnRibbon{appTag}SubCopySvg""
                    label=""Copy Vector SVG""
                    imageMso=""CopyPicture""
                    onAction=""OnCopySvg"" />
            <button id=""btnRibbon{appTag}CopyAllSummary""
                    label=""Copy All Summary (Text)""
                    imageMso=""FileDocumentInspect""
                    onAction=""OnCopyAllSummary"" />
            <button id=""btnRibbon{appTag}CopyTsvRow""
                    label=""Copy as Excel Row (TSV Columns)""
                    imageMso=""TableInsert""
                    onAction=""OnCopyTsvRow"" />
            <button id=""btnRibbon{appTag}CopyMultiLine""
                    label=""Copy as Multi-Line List""
                    imageMso=""Bullets""
                    onAction=""OnCopyMultiLine"" />
            <button id=""btnRibbon{appTag}CopyMarkdown""
                    label=""Copy as Markdown Table""
                    imageMso=""TableOfContentsGallery""
                    onAction=""OnCopyMarkdown"" />
            <menuSeparator id=""sepRibbon{appTag}SubDialog"" />
            <button id=""btnRibbon{appTag}ShowDialog""
                    label=""Detailed Molecule Info Window...""
                    imageMso=""FileProperties""
                    onAction=""OnCopyMoleculeInfo"" />
          </menu>";
    }

    public static string BuildRibbonBatchStyleMenu(
        string menuId,
        string screentip,
        string supertip,
        string colorBtnId,
        string bwBtnId,
        string whiteBgBtnId,
        string darkBtnId)
    {
        return $@"          <menu id=""{menuId}""
                label=""Batch Style""
                screentip=""{screentip}""
                supertip=""{supertip}""
                imageMso=""PictureRecolorGallery""
                size=""large"">
            <button id=""{colorBtnId}""
                    label=""Standard Color (Transparent)""
                    imageMso=""ColorMenu""
                    onAction=""OnBatchStyleColorTrans"" />
            <button id=""{bwBtnId}""
                    label=""Publication B&amp;W (ACS Monochrome)""
                    imageMso=""BlackAndWhiteGrayscale""
                    onAction=""OnBatchStyleBw"" />
            <button id=""{whiteBgBtnId}""
                    label=""White Background (Color)""
                    imageMso=""BlackAndWhiteLightGrayscale""
                    onAction=""OnBatchStyleWhiteBg"" />
            <button id=""{darkBtnId}""
                    label=""Dark Mode (White Ink)""
                    imageMso=""BlackAndWhiteBlack""
                    onAction=""OnBatchStyleDark"" />
          </menu>";
    }

    public static string BuildContextMoleculeInfoMenu(string menuId, string prefix, bool insertBeforeCut = false)
    {
        var anchorAttr = insertBeforeCut ? "\n            insertBeforeMso=\"Cut\"" : string.Empty;
        return $@"      <menu id=""{menuId}""
            label=""Copy Molecule Info""
            imageMso=""GroupTaskProperties""{anchorAttr}>
        <button id=""btnContext{prefix}CopyName""
                label=""Copy Substance Name""
                imageMso=""PropertySheet""
                onAction=""OnCopySubstanceName"" />
        <button id=""btnContext{prefix}CopyFormula""
                label=""Copy Chemical Formula""
                imageMso=""EquationInsertNew""
                onAction=""OnCopyFormula"" />
        <button id=""btnContext{prefix}CopyMW""
                label=""Copy Molecular Weight (MW)""
                imageMso=""Calculator""
                onAction=""OnCopyMolecularWeight"" />
        <button id=""btnContext{prefix}SubCopySmiles""
                label=""Copy SMILES""
                imageMso=""Copy""
                onAction=""OnCopySmiles"" />
        <button id=""btnContext{prefix}SubCopyInchi""
                label=""Copy InChI""
                imageMso=""Copy""
                onAction=""OnCopyInchi"" />
        <button id=""btnContext{prefix}SubCopyInchiKey""
                label=""Copy InChIKey""
                imageMso=""SignatureLineInsert""
                onAction=""OnCopyInchiKey"" />
        <button id=""btnContext{prefix}CopyAllSummary""
                label=""Copy All Summary (Text)""
                imageMso=""FileDocumentInspect""
                onAction=""OnCopyAllSummary"" />
        <button id=""btnContext{prefix}CopyTsvRow""
                label=""Copy as Excel Row (TSV Columns)""
                imageMso=""TableInsert""
                onAction=""OnCopyTsvRow"" />
        <button id=""btnContext{prefix}CopyMultiLine""
                label=""Copy as Multi-Line List""
                imageMso=""Bullets""
                onAction=""OnCopyMultiLine"" />
        <button id=""btnContext{prefix}CopyMarkdown""
                label=""Copy as Markdown Table""
                imageMso=""TableOfContentsGallery""
                onAction=""OnCopyMarkdown"" />
        <menuSeparator id=""sepContext{prefix}SubDialog"" />
        <button id=""btnContext{prefix}ShowDialog""
                label=""Detailed Molecule Info Window...""
                imageMso=""GroupTaskProperties""
                onAction=""OnCopyMoleculeInfo"" />
      </menu>";
    }

    public static string BuildContextStructureCopyButtons(
        string smilesBtnId,
        string inchiBtnId,
        string ketBtnId,
        string svgBtnId,
        bool insertBeforeCut = false)
    {
        var anchorAttr = insertBeforeCut ? "\n              insertBeforeMso=\"Cut\"" : string.Empty;
        return $@"      <button id=""{smilesBtnId}""
              label=""Copy SMILES""
              imageMso=""Copy""
              onAction=""OnCopySmiles""{anchorAttr} />
      <button id=""{inchiBtnId}""
              label=""Copy InChI""
              imageMso=""Copy""
              onAction=""OnCopyInchi""{anchorAttr} />
      <button id=""{ketBtnId}""
              label=""Copy Canvas Data (KET)""
              imageMso=""Copy""
              onAction=""OnCopyKet""{anchorAttr} />
      <button id=""{svgBtnId}""
              label=""Copy Vector SVG""
              imageMso=""CopyPicture""
              onAction=""OnCopySvg""{anchorAttr} />";
    }

    public static string BuildObjectContextMenu(
        string contextMenuIdMso,
        string btnPrefix,
        bool editSuffixMolecule,
        string infoMenuId,
        string infoPrefix,
        string separatorId)
    {
        var editBtnId = editSuffixMolecule ? $"{btnPrefix}EditMolecule" : $"{btnPrefix}Edit";
        var copyButtons = BuildContextStructureCopyButtons(
            $"{btnPrefix}CopySmiles",
            $"{btnPrefix}CopyInchi",
            $"{btnPrefix}CopyKet",
            $"{btnPrefix}CopySvg",
            insertBeforeCut: true);
        var infoMenu = BuildContextMoleculeInfoMenu(infoMenuId, infoPrefix, insertBeforeCut: true);

        return $@"    <contextMenu idMso=""{contextMenuIdMso}"">
      <button id=""{editBtnId}""
              label=""Edit Structure (Ketcher)...""
              imageMso=""DesignMode""
              onAction=""OnEditMolecule""
              insertBeforeMso=""Cut"" />
      <button id=""{btnPrefix}EditMolstar""
              label=""Edit 3D Structure (Mol*)...""
              imageMso=""Insert3DModelFromFile""
              onAction=""OnEditMolstar""
              insertBeforeMso=""Cut"" />
{copyButtons}
{infoMenu}
      <menuSeparator id=""{separatorId}"" insertBeforeMso=""Cut"" />
    </contextMenu>";
    }

    public static string BuildTextOrTableContextMenu(
        string contextMenuIdMso,
        string menuId,
        string btnPrefix,
        string sepPrefix,
        string infoMenuId,
        string infoPrefix,
        string trailingSeparatorId)
    {
        var copyButtons = BuildContextStructureCopyButtons(
            $"{btnPrefix}CopySmiles",
            $"{btnPrefix}CopyInchi",
            $"{btnPrefix}CopyKet",
            $"{btnPrefix}CopySvg",
            insertBeforeCut: false);
        var infoMenu = BuildContextMoleculeInfoMenu(infoMenuId, infoPrefix, insertBeforeCut: false);

        return $@"    <contextMenu idMso=""{contextMenuIdMso}"">
      <menu id=""{menuId}""
            label=""Chemical Structure""
            imageMso=""EquationInsertNew""
            insertAfterMso=""PasteGalleryMini"">
        <button id=""{btnPrefix}Insert""
                label=""Insert / Edit Structure...""
                imageMso=""EquationInsertNew""
                onAction=""OnInsertMolecule"" />
        <button id=""{btnPrefix}InsertMolstar""
                label=""3D Structure (Mol*)...""
                imageMso=""Insert3DModelFromFile""
                onAction=""OnInsertMolstar"" />
        <button id=""{btnPrefix}QuickInsert""
                label=""Paste Structure""
                imageMso=""Paste""
                onAction=""OnQuickInsert"" />
        <menuSeparator id=""{sepPrefix}SubCopy"" />
{copyButtons}
        <menuSeparator id=""{sepPrefix}SubInfo"" />
{infoMenu}
      </menu>
      <menuSeparator id=""{trailingSeparatorId}"" insertAfterMso=""PasteGalleryMini"" />
    </contextMenu>";
    }

    public static string BuildExcelRangeContextMenu(
        string contextMenuIdMso,
        string preSeparatorId,
        string menuId,
        string insertBtnPrefix,
        string insertMolstarBtnId,
        string quickInsertBtnId,
        string sepPrefix,
        string copyBtnPrefix,
        string infoMenuId,
        string infoPrefix,
        string trailingSeparatorId)
    {
        var copyButtons = BuildContextStructureCopyButtons(
            $"{copyBtnPrefix}CopySmilesCell",
            $"{copyBtnPrefix}CopyInchiCell",
            $"{copyBtnPrefix}CopyKetCell",
            $"{copyBtnPrefix}CopySvgCell",
            insertBeforeCut: false);
        var infoMenu = BuildContextMoleculeInfoMenu(infoMenuId, infoPrefix, insertBeforeCut: false);

        return $@"    <contextMenu idMso=""{contextMenuIdMso}"">
      <menuSeparator id=""{preSeparatorId}"" insertAfterMso=""PasteGalleryMini"" />
      <menu id=""{menuId}""
            label=""Chemical Structure""
            imageMso=""EquationInsertNew""
            insertAfterMso=""PasteGalleryMini"">
        <button id=""{insertBtnPrefix}InsertMolecule""
                label=""Insert / Edit Structure (SMILES2IMG)...""
                imageMso=""EquationInsertNew""
                onAction=""OnInsertMolecule"" />
        <button id=""{insertBtnPrefix}InsertInchiMolecule""
                label=""Insert / Edit Structure (INCHI2IMG)...""
                imageMso=""FunctionWizard""
                onAction=""OnInsertInchiMolecule"" />
        <button id=""{insertBtnPrefix}InsertKetMolecule""
                label=""Insert / Edit Structure (KET2IMG)...""
                imageMso=""SmartArtInsert""
                onAction=""OnInsertKetMolecule"" />
        <button id=""{insertMolstarBtnId}""
                label=""3D Structure (Mol*)...""
                imageMso=""Insert3DModelFromFile""
                onAction=""OnInsertMolstar"" />
        <button id=""{quickInsertBtnId}""
                label=""Paste Structure""
                imageMso=""Paste""
                onAction=""OnQuickInsert"" />
        <menuSeparator id=""{sepPrefix}SubCopy"" />
{copyButtons}
        <menuSeparator id=""{sepPrefix}SubInfo"" />
{infoMenu}
      </menu>
      <menuSeparator id=""{trailingSeparatorId}"" insertAfterMso=""PasteGalleryMini"" />
    </contextMenu>";
    }
}
