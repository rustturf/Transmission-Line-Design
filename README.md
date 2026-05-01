# Copyright (c) 2026 rustturf

# All rights reserved.




# Transmission Line Design Tool



This application is for automated overhead transmission line design. Given power transfer requirements, it allows for voltage selection, technical analysis, air clearance, insulator sizing, conductor selection, voltage regulation, corona check, and tower design  outputting a formatted text report and an Excel tower design workbook.

---

## File Structure
The file structure is as follows:
```
project-root/
│
├── gui.py                          # Main application entry point (tkinter UI)
├── calculations.py                 # All engineering calculation logic
├── report_writer.py                # Text report generator
│
├── Tables/
│   ├── withstand voltage capability for different system voltage.csv
│   │                               # Maps max system voltage → dry/wet/impulse withstand (kV)
│   └── Flashover voltage for 254by154mm disc insulators.csv
│                                   # Maps number of discs → FOV and impulse withstand (kV)
│
└── matlabandexcelfiles/
    ├── british.xlsx                # ACSR conductor catalogue
    │                               #   Columns used: ACSR name, Current Rating Ampacity,
    │                               #   Resistance at 20°C, Overall Diameter, Overall Area,
    │                               #   Lactual (thermal expansion coeff), Weight, UTS, YME
    └── towerheightcalculation.xlsx # Conductor mechanical properties for tower sag/tension
                                    #   Same columns as british.xlsx (used separately)
```
All the excel and csv files contains the ACSR conductor data based on the british standard. I have included the actual picture of the conductor sizing available. Please note that the code must be ran manually and must have Excel installed in your PC to make it work. 

### Output Files (auto-saved to Desktop)

| File | Content |
|---|---|
| `TL_Design_Report.txt` | Full formatted text report of all design stages |
| `TL_Tower_Design.xlsx` | Tower design workbook: Sag_Results, Bending_Moment, Tower (cost) |
---

## How to Run

### Prerequisites

```bash
pip install pandas numpy openpyxl
```

Python 3.8+ with `tkinter` (included in standard CPython distributions) is required.

### Launch

```bash
python gui.py
```

The GUI window will open. Fill in the input fields and click **▶ RUN DESIGN**.

### Input Fields

| Field | Description | Example |
|---|---|---|
| Power (MW) | Total power to be transmitted | `300` |
| Length (km) | Line length | `200` |
| Power Factor | System power factor (0–1) | `0.95` |
| Weight of Ice (kg/m) | Ice load per unit length; enter `0` for no ice | `0` |
| Output File | Path for the text report | auto-set to Desktop, and you can change it manually too|

> **Manual GMD Dialog:** If the selected configuration has `nc > 2` or `nb > 1 or nc=1 and nb>=2` (where auto-calculation isn't implemented), a popup will appear asking you to manually enter GMD, GMRl, and GMRc in metres. I haven't had the time to do it for different circuit configurations. Hence, if the spacing of design of conductor spacing is not as in the formula, you may have different results and mismatch in calculations. Although, that case has not been included. The structure is as shown 
For single circuit 
               O
> 
            O
            
                O     

For double circuit
 O     O

 O     O

 O     O
---
## Calculation Stages — Assumptions & Detail

### Stage 1 · Most Economical Voltage

**Formula (Economical Voltage formula):**

```
V = 5.5 × √( L/1.6 + P×1000 / (150 × nc × pf) )
```

Where `L` = length (km), `P` = power (MW), `pf` = power factor, `nc` = number of circuits (1 or 2).

**Assumptions:**
- Formula is evaluated for both `nc = 1` and `nc = 2`.
- The coefficient `5.5`, divisors `1.6` and `150` are empirical constants from Still's formula.
- Result is in kV.

**Standard voltage snap:** Calculated voltage is snapped to the nearest standard level from:
`[66, 132, 220, 400, 500, 765]` kV.  
If the technical analysis requires stepping up, the *next higher* standard voltage is selected instead.

---

### Stage 2 · Technical Analysis & nc Selection

**Interpolated mf limit** from a lookup table:

| Length (km) | mf Limit |
|---|---|
| 80  | 2.75 |
| 160 | 2.25 |
| 240 | 1.75 |
| 320 | 1.35 |
| 480 | 1.00 |
| 640 | 0.75 |

Linear interpolation is used between table entries. Extrapolation is applied for lengths outside the table range.

**Surge Impedance Loading (SIL):**

```
SIL = V² / Z_SIL
```

**Assumed SIL impedances:**
- `nc = 1`: Z = 400 Ω
- `nc = 2`: Z = 200 Ω

**Loading factor:**
```
mf = P / SIL
```

**Margin:**
```
mfmargin = |mf_limit − mf|
```

**nc selection rule:**
- If both `nc = 1` and `nc = 2` give `mfmargin < mf_limit`: choose the one with the smaller margin.
- If only one satisfies the condition: use that circuit count.
- If neither satisfies: step up to next higher standard voltage and repeat (up to 10 attempts, max 765 kV).

---

### Stage 3 · Air Clearance Calculation

All dimensions calculated in **cm**.

**Assumed shielding angle:** 45°

| Parameter | Formula |
|---|---|
| Min. air clearance `a` | `(V/√3 × 1.1 × √2) + 20` |
| Cross-arm length `cl` | `2 × a` |
| Tower width `b` | `2 × a` |
| Insulator string length `l` | `√2 × a` (45° shielding) |
| Horizontal phase separation `c` | `b + 2 × cl` |
| Vertical phase separation `y` | `(l + a) / √(1 − (ratio)² × (1/3)²)` |
| Earth wire height `d` | `√3 × cl` (nc=1)  or  `√3 × (b/2 + a)` (nc=2) |

The factor `1.1` accounts for the 10% overvoltage allowance; `√2` converts RMS to peak.

---

### Stage 4 · Insulator Selection

**Correction factors (all fixed/assumed):**

| Factor | Symbol | Value |
|---|---|---|
| Flashover-withstand ratio | FWR | 1.15 |
| Non-standard atmosphere correction | NAC | 1.10 |
| Safety factor | FS | 1.20 |
| Switching surge ratio | SSR | 2.80 |
| Surge impedance ratio | SIR | 1.20 |
| Earthing Factor factor | EF | 0.80 |

**Equivalent withstand levels:**

| Criterion | Formula |
|---|---|
| 1-min dry withstand | `Dry_table × FWR × NAC × FS` |
| 1-min wet withstand | `Wet_table × FWR × NAC × FS` |
| Temporary OV withstand | `EF × √2 × V_max × FWR × NAC × FS` |
| Lightning OV withstand | `Impulse_table × FWR × NAC × FS` |
| Switching OV withstand | `(V/√3) × 1.1 × √2 × SSR × SIR × FWR × NAC × FS` |

Disc count is looked up from the CSV for each criterion. The **maximum** disc count across all five criteria governs.

**Disc type assumed:** 254 × 154 mm standard disc insulators.

---

### Stage 5 · Conductor Selection

Conductor database: `british.xlsx` (ACSR conductors).

**Temperature assumption:**
- Operating temperature `θ₂` = **65°C**
- Reference temperature `θ₁` = **20°C**
- Resistivity temperature coefficient = **0.004 /°C**

```
r65 = r20 × (1 + 0.004 × (65 − 20))
```

**Selection algorithm:**
1. Start with `nb = 1` bundle conductor.
2. Calculate operating current: `I = P×10⁶ / (√3 × V×10³ × pf × nc × nb)`
3. Find the smallest ACSR conductor with current rating ≥ `I`.
4. Compute transmission efficiency:
   ```
   Loss = 3 × I² × r65 × L × nc × nb
   η = (1 − Loss/P) × 100%
   ```
5. If `η < 94%`, step to the next larger conductor. If the table is exhausted, increment `nb` and repeat.
6. If forced `nb` is provided (from VR or corona loop), that value is used directly.

**Efficiency threshold assumed: 94%**

---

### Stage 6 · Voltage Regulation

**Circuit model:** Nominal-π  model is selected, (medium line model). For long transmission lines, you can change the parameters to that of the nominal T model digging through the code. 

```


Where `Z = R + jXL` and `Y = jBc`.

**GMD / GMR calculation (auto, for standard configs):**
The diameter is taken of the conductor itself to calculate the GMD And GMR
| Config | GMD Formula | GMRl | GMRc |
|---|---|---|---|
| nc=1, nb=1 | `(( c/2)² + y²) × c)^(1/3)` | `0.7788 × r` | `r` |
| nc=2, nb=1 | Full phase-to-phase geometric mean across R/Y/B phases | Geometric mean of per-phase GMRl | Geometric mean of per-phase GMRc |
| All other | **User must supply manually** via the GMD dialog | — | — |

Where `c` = horizontal phase separation (m), `y` = vertical phase separation (m), `r` = conductor radius (m).

**Frequency assumed: 50 Hz**

**VR limit: 12%** — if exceeded, `nb` is incremented (up to `nb = 6`) and the conductor selection / VR calculation is repeated.

---

### Stage 7 · Corona Check

**Peek's formula (adapted):**

```
Vcr = √3 × (30/√2) × 100 × GMRc × ln(GMD/GMRc) × 0.95 × 0.95
```

**Assumed factors:**
- `30/√2` kV/cm — critical gradient (Peek's constant for standard air)
- First `0.95` — surface irregularity factor (stranded conductor)
- Second `0.95` — weather/altitude factor

If `Vcr < V_system`, bundles are incremented and the check repeats (max `nb = 6`).

---

### Stage 8 · Tower Design

Spans evaluated: The ruling spans taken are : **200 m, 250 m, 300 m, 350 m** for every ACSR conductor in the database.

**Fixed assumptions:**

| Parameter | Value |
|---|---|
| Wind pressure `wp` | 100 N/m² |
| Safety factor `fs` | 2 |
| Earth wire diameter `dew` | 14.60 mm |
| Earth wire UTS `utsew` | 6664 N |
| Number of earth wires `ne` | 2 |
| Temperature states | θ₁=0°C, θ₂=27°C, θ₃=65°C | Temperature states are the ones at the toughest condition, stringing condition and the easiest condition
| Steel cost | 150,000 (currency/tonne) |

**Temperature load angles (for tension change calculation):**
The angle of deviation is selected on the basis of class of towers for type a towers assumeed to have a maximum deviation of 2 degree. 15 degree and 30 degree. Note that alpha is calculated by dividing this angle of deviation by 2.
- α₁ = 1°, α₂ = 7.5°, α₃ = 15° (converted to radians internally)
- Sin weighting: `0.8 sin(α₁) + 0.15 sin(α₂) + 0.05 sin(α₃)`

**Wind load on conductor:**
```
Ww = wp × (2/3) × dc
```

**Combined mechanical load:**
```
w1 = √(Ww² + (wc + wice)²)
```

**Sag:**
```
sagmax = (wc/1000) × span² / (8 × T3)
```

**Minimum ground clearance (ho):**
```
ho = (((V × 1.1 − 33) / 33) + 17) × 0.305   [metres]
```

**Tower heights:**
- `h1 = ho + sagmax` (lowest conductor)
- `h2 = h1 + y` (middle conductor; `h1 + y/2` if nc=1)
- `h3 = h2 + y` (top conductor)
- `ht = h3 + d` (earth wire attachment)

**Tower weight:**
Calculated using the formula shown below and the answer is in tonnes.
```
TW = 0.000631 × ht × √(TBM × fs)
```

**Output Excel sheets:**
- **Sag_Results** — per conductor per span: k1, k2, k11, T2, T3, sagmax, h1, h2, h3, ht
- **Bending_Moment** — wind/tension loads, total bending moment, tower weight
- **Tower** — number of towers, cost per tower, cost per unit length

---

## Known Limitations & Manual Inputs

- **GMD auto-calculation** is only implemented for `nc ∈ {1,2}` and `nb = 1`. Any other combination requires the user to supply GMD, GMRl, and GMRc manually.
- **Maximum bundles** cap is 6. If VR or corona still fails at `nb = 6`, the tool recommends moving to a higher voltage level.
- **Wind pressure** is hardcoded at 100 N/m² — not an input field.
- **Tower design** runs all conductors in the database, not just the selected one. Filter results by the conductor name identified in Stage 5.
- The tool uses `os.startfile()` to open outputs, which is **Windows-only**. On Linux/macOS, open the files manually from the Desktop.

The program dosent yet do select the most economical span based on the financial analysis due to time limitations. The sample calculation has been included in the report provided, please have a look there if you want to proceed further.
---

##  Report Structure (TL_Design_Report.txt)

```
Section 1  — Input Parameters
Section 2  — Voltage Selection
Section 3  — Technical Analysis
Section 4  — Air Clearance Calculation
Section 5  — Insulator Selection
Section 6  — Conductor Selection
Section 7  — Voltage Regulation
Section 8  — Corona Check
Section 9  — Tower Design (reference to Excel)
Section 10 — Notes & Warnings (if any)
```
Any feedback or suggestions to improve the code are welcome, and users may also submit a pull request to modify or enhance it further. However, I will not be actively working on this project for now and may only revisit it occasionally.
