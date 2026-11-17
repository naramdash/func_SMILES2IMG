# Excel Cookbook Part 2: MW, Lipinski Ro5, SAR Screening & SDF Files

> **Excel Cookbook Part 2:** Ready-to-copy recipes for zero-latency physicochemical properties, 3D complexity metrics, substructure/similarity SAR screening, cross-notation conversions, and 12-column `.sdf` file import/export—each backed by live Excel captures.

<div class="s2i-card-grid" style="margin-top: 12px;">
  <a href="/func_SMILES2IMG/sample_compounds.sdf" download="sample_compounds.sdf" class="s2i-card">
    <div class="s2i-card-title">📥 Download Sample Chemical DB (.sdf)</div>
    <p class="s2i-card-desc">Test [SMILES] ➔ [SDF File] ➔ [Import SDF] in Excel to generate a 12-column SAR &amp; 3D complexity table in 1 click</p>
  </a>
  <a href="/func_SMILES2IMG/SMILES2IMG_Comprehensive_Showcase.xlsx" download="SMILES2IMG_Comprehensive_Showcase.xlsx" class="s2i-card">
    <div class="s2i-card-title">📥 Download 8-Tab Master Excel Showcase (.xlsx)</div>
    <p class="s2i-card-desc">0.Quick Guide · 1.Structures · 2.Properties/Lipinski/3D · 3.Substructure · 4.Styles Matrix · 5.Reactions · 6.Legacy .FLOAT · 7.Ribbon Tools</p>
  </a>
</div>

---

## Recipe 6 & 7: Physicochemical Properties, Lipinski Rule of 5 & 3D Complexity

**Goal:** Compute Molecular Weight, HRMS Exact Mass, LogP, TPSA, HBD/HBA/RotBonds, Rings, Aromatic Rings, Fraction Csp3 (`Fsp3`), Chiral Centers, and Lipinski Pass/Fail directly beside your structure column in 0ms offline.

```text
=SMILES2MW(A2)
=SMILES2EXACTMASS(A2)
=SMILES2LIPINSKI(A2)
=SMILES2FSP3(A2)
```

<FeatureMedia id="excel-cheminformatics" />

| Descriptor | Formula | Lipinski / MedChem Rule | Aspirin (`A2`) | Notes |
| :--- | :--- | :---: | :---: | :--- |
| **Chemical Formula** | `=SMILES2FORMULA(A2)` | — | `C9H8O4` | Standard Hill notation |
| **Molecular Weight** | `=SMILES2MW(A2)` | **≤ 500** g/mol | `180.16` | Numeric double for stoichiometry |
| **Exact Mass (HRMS)** | `=SMILES2EXACTMASS(A2)` | — | `180.0423` | `[M+H]+` = `+1.0073`, `[M+Na]+` = `+22.9892` |
| **LogP / TPSA** | `=SMILES2LOGP(A2)` / `=SMILES2TPSA(A2)` | **≤ 5.0** / **≤ 140 Å²** | `1.31` / `63.60` | CNS BBB ≤ 90 Å² |
| **HBD / HBA / RotB** | `=SMILES2HBD` / `HBA` / `ROTBONDS` | **≤ 5** / **≤ 10** / **≤ 10** | `1` / `4` / `3` | Oral bioavailability descriptors |
| **Lipinski Pass/Fail** | `=SMILES2LIPINSKI(A2)` | — | `PASS` | Combined 4-rule evaluation |
| **Heavy / Rings / Arom** | `=SMILES2HEAVY` / `RINGS` / `AROMATIC` | **Arom ≤ 3** | `13` / `1` / `1` | Developability & scaffold tracking |
| **3D Fsp³ / Chiral** | `=SMILES2FSP3(A2)` / `=SMILES2CHIRAL(A2)` | **Fsp³ ≥ 0.25** | `0.11` / `0` | "Escape from Flatland" metrics |

> 🔄 Using InChI or KET? Replace `SMILES2` with `INCHI2` or `KET2` (e.g., `=INCHI2MW(A2)`, `=KET2EXACTMASS(A2)`).

---

## Recipe 8: Substructure Filtering (`SMILES.MATCH`), Highlighting & Tanimoto Similarity (`SAR`)

**Goal:** Filter libraries containing a target scaffold, visually highlight the matching atoms in color, and rank analogs by 2D fingerprint similarity (`0.0–1.0`).

```text
=SMILES.MATCH(A2, "c1ccccc1")
=SMILES.SIMILARITY(A2, $B$2)
=SMILES2IMG(A2, "", "hl:c1ccccc1")
```

<FeatureMedia id="excel-sar-screening" />

### 12 Ready-to-Copy SMARTS Functional Group Patterns
| Functional Group | Pattern | Match Formula | Highlight Formula |
| :--- | :--- | :--- | :--- |
| **Benzene Ring** | `c1ccccc1` | `=SMILES.MATCH(A2, "c1ccccc1")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1")` |
| **Pyridine Ring** | `c1ccncc1` | `=SMILES.MATCH(A2, "c1ccncc1")` | `=SMILES2IMG(A2, "", "hl:c1ccncc1")` |
| **Carboxylic Acid** | `C(=O)[OH]` | `=SMILES.MATCH(A2, "C(=O)[OH]")` | `=SMILES2IMG(A2, "", "hl:C(=O)[OH]")` |
| **Ester** | `C(=O)O[#6]` | `=SMILES.MATCH(A2, "C(=O)O[#6]")` | `=SMILES2IMG(A2, "", "hl:C(=O)O[#6]")` |
| **Amide** | `C(=O)N` | `=SMILES.MATCH(A2, "C(=O)N")` | `=SMILES2IMG(A2, "", "hl:C(=O)N")` |
| **Sulfonamide** | `S(=O)(=O)N` | `=SMILES.MATCH(A2, "S(=O)(=O)N")` | `=SMILES2IMG(A2, "", "hl:S(=O)(=O)N")` |
| **Ketone** | `[#6]C(=O)[#6]` | `=SMILES.MATCH(A2, "[#6]C(=O)[#6]")` | `=SMILES2IMG(A2, "", "hl:[#6]C(=O)[#6]")` |
| **Aldehyde** | `[CX3H1](=O)` | `=SMILES.MATCH(A2, "[CX3H1](=O)")` | `=SMILES2IMG(A2, "", "hl:[CX3H1](=O)")` |
| **Phenol** | `c1ccccc1[OH]` | `=SMILES.MATCH(A2, "c1ccccc1[OH]")` | `=SMILES2IMG(A2, "", "hl:c1ccccc1[OH]")` |
| **Nitro** | `[N+](=O)[O-]` | `=SMILES.MATCH(A2, "[N+](=O)[O-]")` | `=SMILES2IMG(A2, "", "hl:[N+](=O)[O-]")` |
| **Halogen** | `[F,Cl,Br,I]` | `=SMILES.MATCH(A2, "[F,Cl,Br,I]")` | `=SMILES2IMG(A2, "", "hl:[F,Cl,Br,I]")` |
| **Trifluoromethyl** | `C(F)(F)F` | `=SMILES.MATCH(A2, "C(F)(F)F")` | `=SMILES2IMG(A2, "", "hl:C(F)(F)F")` |

> Similarity `[metric]` supports `"tanimoto"` (default), `"tversky"`, and `"euclid-sub"`.

---

## Recipe 9: Convert Between Name, CAS, SMILES, InChI & InChIKey

**Goal:** Standardize messy compound lists into canonical SMILES, IUPAC InChI, and 27-character InChIKeys for `XLOOKUP`/`VLOOKUP` deduplication, or render directly from InChI.

```text
=NAME2SMILES("50-78-2")
=SMILES2INCHI(B2)
=SMILES2INCHIKEY(B2)
=INCHI2IMG(C2)
```

<FeatureMedia id="excel-conversions" />

---

## Recipe 10: Import & Export 12-Column `.SDF` / `.MOL` Chemical Databases

**Goal:** Open multi-compound `.sdf` libraries directly into a formatted 12-column live-formula Excel sheet (`Name`, `Structure`, `SMILES`, `Formula`, `MW`, `Exact Mass`, `LogP`, `Rings`, `Aromatic Rings`, `Fsp3`, `Chiral Centers`, `Lipinski`) or export worksheet rows back to V2000 `.sdf`.

<FeatureMedia id="excel-sdf-workflow" />

1. Click **`[SMILES] ➔ [SDF File] ➔ [Import SDF File (.sdf)...]`** and select an `.sdf` or `.mol` file.
2. All 12 columns and structure images populate automatically with adjusted row heights and column widths.
3. Click **`[SMILES] ➔ [SDF File] ➔ [Export Sheet to SDF (.sdf)...]`** anytime to save compounds back to `.sdf`.

---

## Combining with Excel `IF` / `AND` & Conditional Formatting

| Screening Rule | Copy-Paste Formula (`A2` = SMILES) |
| :--- | :--- |
| **Priority Pyridine Hit (`Ro5 PASS`)** | `=IF(AND(SMILES2LIPINSKI(A2)="PASS", SMILES.MATCH(A2, "c1ccncc1")), "✅ Priority Hit", "Exclude")` |
| **CNS BBB Candidate (`TPSA ≤ 90`, `MW ≤ 450`)** | `=IF(AND(SMILES2TPSA(A2)<=90, SMILES2MW(A2)<=450), "🧠 CNS Candidate", "Standard")` |
| **3D-Rich Lead (`Fsp3 ≥ 0.25`, `Arom ≤ 3`)** | `=IF(AND(SMILES2FSP3(A2)>=0.25, SMILES2AROMATIC(A2)<=3), "✨ Escape Flatland", "Flat")` |

---

## 🚀 Continue Exploring

<div class="s2i-card-grid">
  <a href="../reference/formula" class="s2i-card">
    <div class="s2i-card-title">📑 40 Formula &amp; Style Reference</div>
    <p class="s2i-card-desc">Browse all 40 chemical UDFs, return data types, and error code specifications</p>
  </a>
  <a href="../tools/formula-builder" class="s2i-card">
    <div class="s2i-card-title">⚡ Interactive Formula Builder</div>
    <p class="s2i-card-desc">Click styling and highlight options to build formulas visually</p>
  </a>
</div>
