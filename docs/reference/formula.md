# Excel Function & Parameter Specification

Complete reference for Excel worksheet functions provided by **SMILES for Microsoft Office Suite**.

---

## 📐 Function Signatures

### 1. `SMILES2IMG`
```text
=SMILES2IMG(smiles_or_ket, [background], [options])
```
* **Purpose:** Renders chemical structure as a **native in-cell image** (Microsoft 365 / Excel 2024+).
* **Return Value:** In-cell image entity that auto-scales with cell boundaries.

### 2. `SMILES2IMG.FLOAT`
```text
=SMILES2IMG.FLOAT(smiles_or_ket, [background], [options])
```
* **Purpose:** Inserts a floating shape locked to cell bounds (`xlMoveAndSize`). Compatible with Excel 2016, 2019, and 2021.
* **Return Value:** In-cell text string while managing the associated overlay shape.

### 3. `KET2IMG` & `KET2IMG.FLOAT`
```text
=KET2IMG(ket_json_string, [background], [options])
=KET2IMG.FLOAT(ket_json_string, [background], [options])
```
* **Purpose:** Renders multi-component organic chemical reactions, reaction arrows, plus signs, and text annotations directly from Ketcher KET JSON representations.

---

## ⚙️ Parameter Reference

| Parameter | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `smiles_or_ket` | `string` | *(Required)* | Valid SMILES string (e.g., `CN1C=NC2=C1C(=O)N(C(=O)N2C)C`) or Ketcher JSON string (`{"root": ...}`). |
| `background` | `string` | `""` (Auto) | Background color. Empty string defaults to transparent in Ketcher and transparent/white in Excel. Accepts hex codes (`"#FFFFFF"`, `"#1A1A1A"`, `"#00000000"`), named colors (`"white"`, `"transparent"`, `"black"`). |
| `options` | `string` | `""` | Space-, comma-, semicolon-, or pipe-separated styling and rendering flags. |

---

## 🎨 Option Tokens Specification

| Token | Category | Description | Example |
| :--- | :--- | :--- | :--- |
| `bw` / `monochrome` | Mode | Black-and-white line art (ACS / patent specification) | `"bw"` |
| `colored` | Mode | Standard IUPAC CPK element color scheme | `"colored"` |
| `num` / `numbers` | Label | Display 1-based atom indices for NMR/MS assignments | `"num"` |
| `stereo` | Label | Force CIP stereochemistry annotations (*R/S*, *E/Z*) | `"stereo"` |
| `d=<float>` | Quality | Supersampling resolution multiplier (e.g. `d=1.5`, `d=2.0`) | `"d=2.0"` |
| `flip` | Transform | Flips molecular coordinates horizontally | `"flip"` |
| `flipv` | Transform | Flips molecular coordinates vertically | `"flipv"` |
| `rot90` | Transform | Rotates molecule 90 degrees clockwise | `"rot90"` |
| `rot180` | Transform | Rotates molecule 180 degrees | `"rot180"` |
| `rot270` | Transform | Rotates molecule 270 degrees clockwise | `"rot270"` |
| `scale=<float>` | Sizing | Scaling factor for bond lengths | `"scale=1.2"` |
| `w=<int>` | Sizing | Custom width in pixels (Floating images) | `"w=240"` |
| `h=<int>` | Sizing | Custom height in pixels (Floating images) | `"h=180"` |

---

## 🔀 Flexible Delimiter Syntax

Option tokens can be formatted with any standard delimiter:

```text
=SMILES2IMG(A2, "", "bw d=2.0 rot90")
=SMILES2IMG(A2, "", "bw, d=2.0, rot90")
=SMILES2IMG(A2, "", "bw; d=2.0; rot90")
=SMILES2IMG(A2, "", "bw | d=2.0 | rot90")
```

---

## 💡 Practical Examples

* **Monochrome Patent Line Art:**
  ```text
  =SMILES2IMG(A2, "", "bw")
  ```
* **Dark Mode with White Bonds:**
  ```text
  =SMILES2IMG(A2, "#1e1e1e", "white d=2.0")
  ```
* **NMR Peak Mapping with Stereochemistry:**
  ```text
  =SMILES2IMG(A2, "", "num stereo d=2.0")
  ```
* **Full Multi-Step Organic Reaction:**
  ```text
  =KET2IMG(B2)
  ```
