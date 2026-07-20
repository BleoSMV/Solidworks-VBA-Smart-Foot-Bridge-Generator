# SmartBridge — SolidWorks API Footbridge Generator

A VBA automation tool for SolidWorks 2024-5 that designs and models a complete
steel footbridge from minimal input. 

Enter a span, width, load class, material and finish into one UserForm. The
program validates the design against simply-supported beam theory, selects the
optimum standard section automatically, and generates the full design package
with no manual modelling.


<img width="1082" height="714" alt="image" src="https://github.com/user-attachments/assets/4f3ac37b-2894-41ac-b588-0a2af348e9c1" />

## What it does

- **Validates before it builds.** Every catalogued beam section is checked
  against five structural criteria: bending, shear, von Mises stress,
  deflection and natural frequency (2.5 Hz minimum). The Create button stays
  disabled until a design passes, so a failing bridge can never reach geometry.
- **Selects the beam for you.** The section loop picks the lightest or
  cheapest section that passes all five checks, and reports the nearest miss
  when nothing does.
- **Generates everything.** One validated run produces ~30 files: 7 parametric
  part files, the assembly, a dimensioned GA drawing plus 7 part drawings with
  a BOM table, PDF/DWG exports, a CSV bill of materials and a text design
  report, all stamped with custom properties.
- **Reads its data from Excel.** Materials, finishes and the beam section
  catalogue live in `SmartBridgeData.xlsx` (sheets: Materials, Finishes,
  Sections), so the catalogue can be extended without touching the code.
  Header matching is tolerant, and built-in data is used as a fallback if the
  workbook is missing or malformed.

## Requirements

- SolidWorks 2024 or 2025 (Windows) with default document templates configured
- Microsoft Excel (for the external data workbook)
- No add-ins or type-library references required — all COM objects are
  late-bound

## Usage

1. Open SolidWorks.
2. Run the macro and open the `frmSmartBridge` UserForm.
3. Set the data workbook path and output folder, then press **Load workbook**.
4. Enter the design parameters and press **Validate Design**.
5. When the status shows PASS, press **Create** and wait for the staged build
   to finish. The output folder opens on completion.

## Data workbook format

`SmartBridgeData.xlsx` needs three sheets with these headers (unit variants
such as "m2"/"m²" and "£"/"GBP" are accepted):

| Sheet     | Columns                                                                 |
|-----------|-------------------------------------------------------------------------|
| Materials | Material Name, E (Pa), Fy (Pa), Density (kg/m³)                          |
| Finishes  | Finish Name, Multiplier, RGB Color (optional)                            |
| Sections  | Section Name, Depth (m), Width(m), Area(m²), Ixx (m⁴), Zxx (m³), Mass (kg/m), Cost (£/m)|

All section properties are base SI: catalogue values in cm⁴/cm³ must be
converted to m⁴/m³ before entry.

## Structure

One UserForm (`frmSmartBridge`) containing all calculation and validation
logic, plus ten modules handling parts, assembly, drawings, reports, SolidWorks
core utilities, file utilities and shared types. The two sides communicate only
through read-only properties and a `BridgeDesign` structure, so the model
builder never receives unvalidated data. Build failures are reported by a
staged error handler that names the failing phase, file and SolidWorks error
code.

## Equations
# Structural & Layout Calculations

This section details the mathematical models, fixed constants, and data structures used by the SmartBridge API to validate footbridge designs, calculate utilisations, generate CAD models, and estimate project costs.

## 1. System Constants & Fixed Geometry

The API relies on standard hardcoded geometric constants for secondary members (defined in VBA as `SB_...` constants):

| Component | Dimension | Value (m) |
| :--- | :--- | :--- |
| **Cross Member** | Profile ($X \times Z$) | $0.08 \times 0.08$ |
| **Deck** | Thickness | $0.04$ |
| **Guard Post** | Profile Size | $0.06$ |
| **Handrail** | Profile Size | $0.05$ |
| **Kickplate** | Height $\times$ Thickness | $0.15 \times 0.012$ |
| **Endplate** | Thickness | $0.012$ |
| **Endplate Allowances** | Width / Height | $0.20$ / $0.15$ |
| **Endplate Holes** | Diameter / Edge Dist | $0.014$ / $0.06$ |

---

## 2. Core Data Structures (VBA Types)

The calculations flow through three primary data structures before being bundled into the final `BridgeDesign` record for SolidWorks generation and reporting:

1. **`BridgeInputs`**: Stores the raw user parameters (Span, Width, live/deck loads, material properties ($E, F_y, \rho$), and safety factors).
2. **`SectionCalc`**: Stores the output of the structural analysis for a specific beam profile ($M, V, \sigma, \tau, \delta, f_1$).
3. **`LayoutCalc`**: Stores the automated geometric layout and commercial estimates (member counts, actual spacings, total mass, total cost).

---

## 3. Load Calculations

Before structural validation can occur, the raw inputs are converted into a uniform line load applied to the main beams.

**Rail and Cross-Member Allowance**
Accounts for the base weight of the cross members and scales the rail allowance based on the specified guardrail height.
$$q_{\text{rc}} = q_{\text{cross\\_allow}} + q_{\text{rail\\_allow}} \left( \frac{h_{\text{rail}}}{1100} \right)$$

**Total Area Load (excluding main beams)**
Combines live loads (factored for dynamic effects) and deck loads across the width of the bridge, adding the allowances.
$$q_{\text{deck\\_total}} = \left( q_{\text{live}} \cdot f_{\text{dyn}} + q_{\text{deck}} \right) W_{\text{bridge}} + q_{\text{rc}}$$

**Line Load Per Main Beam (`qBeam` or $w$)**
Distributes the total deck load to the two main structural beams and adds the self-weight of the selected beam section.
$$w = \frac{q_{\text{deck\\_total}}}{2} + (m_{\text{beam}} \cdot g)$$

*(Note: $w$ corresponds to `qBeam` in the `SectionCalc` structure).*

*Where:*
*   $q_{\text{cross\\_allow}}$ = Cross member allowance (**0.2 kN/m**)
*   $q_{\text{rail\\_allow}}$ = Base rail allowance (**0.35 kN/m**)
*   $h_{\text{rail}}$ = Guard rail height ($mm$)
*   $q_{\text{live}}$ = Live load class ($kN/m^2$)
*   $f_{\text{dyn}}$ = Dynamic factor (**1.15** if dynamic, **1.0** if static)
*   $q_{\text{deck}}$ = Deck load ($kN/m^2$)
*   $W_{\text{bridge}}$ = Bridge width ($m$)
*   $m_{\text{beam}}$ = Mass of selected beam section ($kg/m$)
*   $g$ = Gravity (**9.81 m/s²**)

---

## 4. Structural Analysis Equations

The model assumes a simply supported beam subject to a uniformly distributed load (UDL). These formulas correspond directly to the generated `Report.txt`.

**Maximum Bending Moment (`M`)**
$$M_{max} = \frac{w \cdot L^2}{8}$$

**Maximum Shear Force (`V`)**
$$V_{max} = \frac{w \cdot L}{2}$$

**Bending Stress (`sigma` / $\sigma$)**
$$\sigma = \frac{M}{Z_{xx}}$$

**Maximum Shear Stress (`tau` / $\tau$)**
Assumes a standard parabolic shear stress distribution across a generic web cross-section.
$$\tau = \frac{1.5 \cdot V}{A}$$

**Von Mises Stress (`vm` / $\sigma_{VM}$)**
Combined stress state calculation.
$$\sigma_{VM} = \sqrt{\sigma^2 + 3\tau^2}$$

**Maximum Deflection (`delta` / $\delta$)**
$$\delta = \frac{5 \cdot w \cdot L^4}{384 \cdot E \cdot I_{xx}}$$

**Natural Frequency (First Mode) (`f1` / $f_1$)**
Calculates the lowest natural frequency of the beam to check against resonant pedestrian loading.
$$f_1 = \frac{\pi}{2} \sqrt{\frac{E \cdot I_{xx}}{m_{\text{line}} \cdot L^4}}$$

*Where:*
*   $L$ = Span length (`SpanM`)
*   $Z_{xx}$, $I_{xx}$, $A$ = Section properties of the selected beam
*   $E$ = Young's Modulus (`YoungsE`)
*   $m_{\text{line}}$ = Mass per unit length of the loaded beam ($\frac{w}{g}$)

---

## 5. Utilisation Checks

The design is validated by checking the calculated values against allowable limits dictated by the material yield strength ($F_y$) and the Factor of Safety ($FoS$). A utilisation score $\le 1.0$ is a pass.

**Allowable Stress Limits**
$$\sigma_{\text{allow}} = \frac{F_y}{FoS}$$
$$\tau_{\text{allow}} = \frac{F_y / \sqrt{3}}{FoS}$$

**Utilisation Ratios (`SectionCalc`)**
*   **Bending (`uB`):** $u_B = \frac{\sigma}{\sigma_{\text{allow}}}$
*   **Shear (`uS`):** $u_S = \frac{\tau}{\tau_{\text{allow}}}$
*   **Von Mises (`uVM`):** $u_{VM} = \frac{\sigma_{vm}}{\sigma_{\text{allow}}}$
*   **Deflection (`uD`):** $u_D = \frac{\delta}{L / \text{Ratio}_{\text{defl}}}$
*   **Frequency (`uF`):** $u_F = \frac{f_{\text{min}}}{f_1}$ *(where $f_{\text{min}}$ is **2.5 Hz**)*

**Overall Validation (`uMax`)**
The maximum utilisation governs the pass/fail state.
$$u_{\text{max}} = \max(u_B, u_S, u_{VM}, u_D, u_F)$$

---

## 6. Geometric Layout

Calculates the exact quantity of required secondary members based on span length and user-defined maximum spacings (stored in `LayoutCalc`).

**Cross Members (`nCross`, `actualCrossSpacing`)**
$$n_{\text{cross}} = \left\lceil \frac{L}{S_{\text{cross}}} \right\rceil + 1 \quad (\text{Minimum: 2})$$
$$S_{\text{actual}} = \frac{L}{n_{\text{cross}} - 1}$$

**Guard Posts (`nPosts`)**
Calculated per side, then doubled.
$$n_{\text{posts}} = 2 \cdot \left( \left\lceil \frac{L}{S_{\text{post}}} \right\rceil + 1 \right)$$

**Deck Panels (`nPanels`)**
$$n_{\text{panels}} = \left\lceil \frac{L}{L_{\text{panel}}} \right\rceil \quad (\text{Minimum: 1})$$

---

## 7. Cost and Mass Estimation

Calculates total material mass and applies financial unit costs depending on materials and selected aesthetic finish (`finishMult`). 

**Component Costs**
*   **Main Beams:** $C_{\text{beam}} = 2 \cdot L \cdot Cost_{\text{beam}}$
*   **Cross Members:** $C_{\text{cross}} = n_{\text{cross}} \cdot W_{\text{bridge}} \cdot Cost_{\text{cross}}$
*   **Posts:** $C_{\text{post}} = n_{\text{posts}} \cdot Cost_{\text{post\\_ea}}$
*   **Rails & Kickplates:** $C_{\text{rail}} = (4 \cdot L \cdot Cost_{\text{rail}}) + (2 \cdot L \cdot Cost_{\text{kickplate}})$
*   **Decking:** $C_{\text{deck}} = L \cdot W_{\text{bridge}} \cdot Cost_{\text{deck\\_type}}$

**Total Project Cost (`totalCost`)**
Steelwork costs are scaled by a multiplier derived from the chosen finish.
$$\text{Total Cost} = \left[ (C_{\text{beam}} + C_{\text{cross}} + C_{\text{post}} + C_{\text{rail}}) \cdot M_{\text{finish}} \right] + C_{\text{deck}}$$

**Total Mass (`totalMass`)**
Sums the physical mass of all components to output the final expected weight of the structure.
$$M_{\text{total}} = M_{\text{beams}} + M_{\text{cross}} + M_{\text{deck}} + M_{\text{posts}} + M_{\text{rails}}$$

## Author

Sajjad Mohammed - 2026.

License

Copyright (c) 2026 Sajjad Mohammed. All rights reserved. This is not an
open-source project — see LICENSE.txt for the full terms. In
short:


Personal and educational use is free. Individuals, students and
educators may view, download, run and modify the code for non-commercial
purposes.
You may not pass this work off as your own. No submitting it (or a
derivative) for academic assessment, republishing it under your own name,
or redistributing it without written permission.
Commercial use requires a written agreement with the author. Any
business or commercial use must be discussed with author and terms agreed
before use begins — contact me via [https://github.com/BleoSMV].
No warranty, no liability. The software is provided as is. It performs
illustrative structural calculations for an academic project and must not
be relied upon for the design of any real structure.
