# SmartBridge — SolidWorks API Footbridge Generator

A VBA automation tool for SolidWorks 2024-5 that designs and models a complete
steel footbridge from minimal input. Built for module MP3604 Advanced Computer
Aided Design (UCLan).

Enter a span, width, load class, material and finish into one UserForm. The
program validates the design against simply-supported beam theory, selects the
optimum standard section automatically, and generates the full design package
with no manual modelling.

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
| Sections  | Section Name,Depth(m),Width(m),Area(m²),Ixx(m⁴),Zxx(m³),Mass(kg/m),Cost(£/m)|

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

## Limitations

Parts are rectangular extrusions (an I-beam is modelled as its bounding
envelope); drawing dimensions are notes and guide lines rather than parametric
model items; assembly components are placed by coordinates without mates.
Single-beam physics limits practical spans to roughly 30 m at standard
pedestrian loading, governed by the natural-frequency check. See the project
report for the full reflection and planned improvements.

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
business or commercial use must be discussed with me and terms agreed
before use begins — contact me via [INSERT email or GitHub profile].
No warranty, no liability. The software is provided as is. It performs
illustrative structural calculations for an academic project and must not
be relied upon for the design of any real structure.
