# Reaction Schemes

Draw multi-step organic synthesis schemes with reaction arrows, `+` signs, and catalyst/condition labels directly inside Microsoft Word—while keeping full atomic coordinates editable.

<FeatureMedia id="word-ch4-reaction-scheme" />

## How Lossless KET Persistence Works in Word

1. Click **`[SMILES]` ➔ `[Insert / Edit Structure]`** (<kbd>Ctrl</kbd> + <kbd>Shift</kbd> + <kbd>K</kbd>).
2. Sketch reactants, reaction arrows (`➔`), `+` signs, catalysts, temperatures, and yields in **Ketcher 2D**.
3. Click **`Apply to Document`** (<kbd>Ctrl</kbd> + <kbd>Enter</kbd>).
4. Full KET JSON coordinates are embedded inside the picture's `AlternativeText` and Word `CustomXMLParts`. Even after saving, closing, and sharing the `.docx` with co-authors, **double-clicking** the scheme on a PC with `func_SMILES2IMG` restores the exact arrows and layout.
