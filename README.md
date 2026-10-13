# SmartBridge — SolidWorks API Footbridge Generator

[![License: All Rights Reserved](https://img.shields.io/badge/License-All%20Rights%20Reserved-red.svg)](LICENSE.txt)
[![Personal use: free](https://img.shields.io/badge/Personal%20use-Free-brightgreen.svg)](LICENSE.txt)
[![Academic submission: prohibited](https://img.shields.io/badge/Academic%20submission-Prohibited-critical.svg)](LICENSE.txt)
[![Commercial use: by agreement](https://img.shields.io/badge/Commercial%20use-By%20agreement-orange.svg)](LICENSE.txt)
[![SolidWorks 2024–2025](https://img.shields.io/badge/SolidWorks-2024--2025-e2231a.svg)](#requirements)
[![VBA macro](https://img.shields.io/badge/Language-VBA-867db1.svg)](#structure)
[![No add-ins](https://img.shields.io/badge/Add--ins-None-blue.svg)](#requirements)

A VBA automation tool for SolidWorks 2024–2025 that designs and models a
complete steel footbridge from minimal input.

Enter a span, width, load class, material and finish into one UserForm. The
program validates the design against simply-supported beam theory, selects
the optimum standard section automatically, and generates the full design
package with no manual modelling.

<img width="1082" height="714" alt="SmartBridge UserForm and generated footbridge" src="https://github.com/user-attachments/assets/4f3ac37b-2894-41ac-b588-0a2af348e9c1" />

## What it does

- **Validates before it builds.** Every catalogued beam section is checked
  against five structural criteria: bending, shear, von Mises stress,
  deflection and natural frequency (2.5 Hz minimum). The Create button stays
  disabled until a design passes, so a failing bridge can never reach
  geometry.
- **Selects the beam for you.** The section loop picks the lightest or
  cheapest section that passes all five checks, and reports the nearest miss
  when nothing does.
- **Generates everything.** One validated run produces about 30 files: 7
  parametric part files, the assembly, a dimensioned GA drawing plus 7 part
  drawings with a BOM table, PDF/DWG exports, a CSV bill of materials and a
  text design report, all stamped with custom properties.
- **Reads its data from Excel.** Materials, finishes and the beam section
  catalogue live in `SmartBridgeData.xlsx`, so the catalogue can be extended
  without touching the code. Header matching is tolerant, and built-in data
  is used as a fallback if the workbook is missing or malformed.

## Requirements

- SolidWorks 2024 or 2025 (Windows) with default document templates
  configured
- Microsoft Excel (for the external data workbook)
- No add-ins or type-library references — all COM objects are late-bound

## Usage

1. Open SolidWorks.
2. Run the macro and open the `frmSmartBridge` UserForm.
3. Set the data workbook path and output folder, then press **Load workbook**.
4. Enter the design parameters and press **Validate Design**.
5. When the status shows PASS, press **Create** and wait for the staged
   build to finish. The output folder opens on completion.

## Input parameters

| Parameter | Meaning |
|---|---|
| Span | Clear span of the footbridge |
| Width | Deck width |
| Load class | Live load and deck load applied to the deck |
| Material | Beam material: Young's modulus, yield strength, density |
| Finish | Surface finish, which sets the steelwork cost multiplier |

## Data workbook format

`SmartBridgeData.xlsx` needs three sheets with these headers. Unit variants
such as "m2"/"m²" and "£"/"GBP" are accepted.

| Sheet | Columns |
|---|---|
| Materials | Material Name, E (Pa), Fy (Pa), Density (kg/m³) |
| Finishes | Finish Name, Multiplier, RGB Color (optional) |
| Sections | Section Name, Depth (m), Width (m), Area (m²), Ixx (m⁴), Zxx (m³), Mass (kg/m), Cost (£/m) |

All section properties are base SI: catalogue values in cm⁴ and cm³ must be
converted to m⁴ and m³ before entry.

## Structure

| Part | Role |
|---|---|
| `frmSmartBridge` UserForm | All calculation and validation logic |
| Ten modules | Parts, assembly, drawings, reports, SolidWorks core utilities, file utilities, shared types |
| `BridgeDesign` structure | The only data passed from the form to the model builder |

The two sides communicate only through read-only properties and the
`BridgeDesign` structure, so the model builder never receives unvalidated
data. Build failures are reported by a staged error handler that names the
failing phase, file and SolidWorks error code.

The calculations flow through three data structures before being bundled
into `BridgeDesign`:

| Structure | Holds |
|---|---|
| `BridgeInputs` | Raw user parameters: span, width, live and deck loads, material properties ($E$, $F_y$, $\rho$), safety factors |
| `SectionCalc` | Structural results for one beam profile ($M$, $V$, $\sigma$, $\tau$, $\delta$, $f_1$) and its utilisations |
| `LayoutCalc` | Geometric layout and commercial estimates: member counts, actual spacings, total mass, total cost |

---

# Equations

The model assumes two simply supported main beams under a uniformly
distributed load. All quantities are SI. These formulas correspond directly
to the generated `Report.txt`.

## 1. Fixed geometry

Secondary members use fixed dimensions, defined in VBA as `SB` constants.

| Component | Dimension | Value (m) |
|---|---|---|
| Cross member | Profile ($X \times Z$) | $0.08 \times 0.08$ |
| Deck | Thickness | $0.04$ |
| Guard post | Profile size | $0.06$ |
| Handrail | Profile size | $0.05$ |
| Kickplate | Height $\times$ thickness | $0.15 \times 0.012$ |
| Endplate | Thickness | $0.012$ |
| Endplate allowances | Width / height | $0.20$ / $0.15$ |
| Endplate holes | Diameter / edge distance | $0.014$ / $0.06$ |

## 2. Loads

**Rail and cross-member allowance** (the rail allowance scales with guardrail height)

$$q_{\text{rc}} = q_{\text{crossAllow}} + q_{\text{railAllow}} \left( \frac{h_{\text{rail}}}{1100} \right)$$

**Total line load on the deck** (live load factored for dynamic effects)

$$q_{\text{deckTotal}} = \left( q_{\text{live}} \cdot f_{\text{dyn}} + q_{\text{deck}} \right) W_{\text{bridge}} + q_{\text{rc}}$$

**Line load per main beam** (shared between two beams, plus beam self-weight; `qBeam` in `SectionCalc`)

$$w = \frac{q_{\text{deckTotal}}}{2} + m_{\text{beam}} \cdot g$$

## 3. Structural analysis

**Bending moment and shear force**

$$M_{\max} = \frac{w L^2}{8}, \qquad V_{\max} = \frac{w L}{2}$$

**Bending, shear and von Mises stress** (parabolic shear distribution)

$$\sigma = \frac{M}{Z_{xx}}, \qquad \tau = \frac{1.5\,V}{A}, \qquad \sigma_{VM} = \sqrt{\sigma^2 + 3\tau^2}$$

**Deflection**

$$\delta = \frac{5\,w\,L^4}{384\,E\,I_{xx}}$$

**First natural frequency** (checked against resonant pedestrian loading)

$$f_1 = \frac{\pi}{2} \sqrt{\frac{E\,I_{xx}}{m_{\text{line}}\,L^4}}, \qquad m_{\text{line}} = \frac{w}{g}$$

| Symbol | Meaning |
|---|---|
| $L$ | Span length |
| $Z_{xx}$, $I_{xx}$, $A$ | Section modulus, second moment of area and area of the selected beam |
| $E$ | Young's modulus |
| $m_{\text{line}}$ | Mass per unit length of the loaded beam |

## 4. Utilisation checks

**Allowable stresses** ($F_y$ = yield strength, $FoS$ = factor of safety)

$$\sigma_{\text{allow}} = \frac{F_y}{FoS}, \qquad \tau_{\text{allow}} = \frac{F_y / \sqrt{3}}{FoS}$$

**Utilisation ratios.** A value of 1.0 or below is a pass.

| Check | Utilisation |
|---|---|
| Bending | $u_B = \dfrac{\sigma}{\sigma_{\text{allow}}}$ |
| Shear | $u_S = \dfrac{\tau}{\tau_{\text{allow}}}$ |
| Von Mises | $u_{VM} = \dfrac{\sigma_{VM}}{\sigma_{\text{allow}}}$ |
| Deflection | $u_D = \dfrac{\delta}{L / \text{Ratio}_{\text{defl}}}$ |
| Frequency | $u_F = \dfrac{f_{\min}}{f_1}$, with $f_{\min}$ = 2.5 Hz |

**Governing utilisation**

$$u_{\max} = \max(u_B,\ u_S,\ u_{VM},\ u_D,\ u_F)$$

## 5. Geometric layout

**Cross members** (minimum 2)

$$n_{\text{cross}} = \left\lceil \frac{L}{S_{\text{cross}}} \right\rceil + 1, \qquad S_{\text{actual}} = \frac{L}{n_{\text{cross}} - 1}$$

**Guard posts** (per side, then doubled)

$$n_{\text{posts}} = 2 \left( \left\lceil \frac{L}{S_{\text{post}}} \right\rceil + 1 \right)$$

**Deck panels** (minimum 1)

$$n_{\text{panels}} = \left\lceil \frac{L}{L_{\text{panel}}} \right\rceil$$

## 6. Cost and mass

| Component | Cost |
|---|---|
| Main beams | $C_{\text{beam}} = 2\,L \cdot Cost_{\text{beam}}$ |
| Cross members | $C_{\text{cross}} = n_{\text{cross}} \cdot W_{\text{bridge}} \cdot Cost_{\text{cross}}$ |
| Posts | $C_{\text{post}} = n_{\text{posts}} \cdot Cost_{\text{postEa}}$ |
| Rails and kickplates | $C_{\text{rail}} = 4\,L \cdot Cost_{\text{rail}} + 2\,L \cdot Cost_{\text{kickplate}}$ |
| Decking | $C_{\text{deck}} = L \cdot W_{\text{bridge}} \cdot Cost_{\text{deckType}}$ |

**Total cost** (steelwork is scaled by the finish multiplier, decking is not)

$$\text{Total Cost} = \left( C_{\text{beam}} + C_{\text{cross}} + C_{\text{post}} + C_{\text{rail}} \right) M_{\text{finish}} + C_{\text{deck}}$$

**Total mass**

$$M_{\text{total}} = M_{\text{beams}} + M_{\text{cross}} + M_{\text{deck}} + M_{\text{posts}} + M_{\text{rails}}$$

---

# Limitations

- Simply-supported beam theory with a uniformly distributed load only; no
  point loads, lateral-torsional buckling, connection or fatigue checks.
- Shear stress assumes a parabolic distribution over the whole section area.
- Costs are estimates from the unit rates in the data workbook.
- The geometry is a conceptual final design, not a fabrication package.

(disclaimer: this description was created with the help of AI)

## Author

Sajjad Mohammed — 2026.

## License

[![License: All Rights Reserved](https://img.shields.io/badge/License-All%20Rights%20Reserved-red.svg?style=for-the-badge)](LICENSE.txt)

Copyright (c) 2026 Sajjad Mohammed. All rights reserved. This is **not** an
open-source project. The full terms are in [`LICENSE.txt`](LICENSE.txt).

| | You may | Condition |
|:-:|---|---|
| ✅ | View, download and run the macro | Personal, educational, non-commercial use only |
| ✅ | Modify it | For personal, educational, non-commercial purposes only |
| ✅ | Keep a copy | The licence file and copyright notice must stay with it |

| | You may not | Consequence |
|:-:|---|---|
| ❌ | Submit it, or a derivative, for academic assessment | Breach of licence |
| ❌ | Present it as your own work | Breach of licence |
| ❌ | Redistribute it, free or paid | Needs the author's written permission |
| ❌ | Use it commercially | Needs a written agreement with the author first |
| ❌ | Design or build a real footbridge from it | Prohibited without the author's approval and sign-off by a qualified engineer |

> [!WARNING]
> **Students:** submitting this macro, or any derivative of it, for
> coursework is a breach of the licence.

> [!NOTE]
> **No warranty, no liability.** The software is provided as is. It performs
> illustrative structural calculations for an academic project and must not
> be relied upon for the design of any real structure.

Commercial or licensing enquiries: [github.com/BleoSMV](https://github.com/BleoSMV)
