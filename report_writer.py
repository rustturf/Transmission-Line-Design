
# Copyright (c) 2026 rustturf
# All rights reserved.


"""
report_writer.py
Writes all design results to a formatted text report file.
"""

from datetime import datetime
import os


# ─────────────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _line(char="─", width=70):
    return char * width

def _section(title):
    lines = [
        "",
        _line("═"),
        f"  {title.upper()}",
        _line("═"),
    ]
    return "\n".join(lines)

def _subsection(title):
    return f"\n  ── {title} ──"

def _row(label, value, unit="", indent=4):
    pad = " " * indent
    label_str = f"{pad}{label}:".ljust(52)
    return f"{label_str}{value} {unit}".rstrip()

def _list_row(label, values, unit="", indent=4):
    pad   = " " * indent
    label_str = f"{pad}{label}:".ljust(52)
    vals  = "  |  ".join(f"nc={i+1}: {v} {unit}".strip() for i, v in enumerate(values))
    return f"{label_str}{vals}"


# ─────────────────────────────────────────────────────────────────────────────
#  Public API
# ─────────────────────────────────────────────────────────────────────────────

def create_report(results: dict, output_path: str = None) -> str:
    """
    Compile all results into a formatted report and save to a .txt file.

    Parameters
    ----------
    results : dict  — keys defined below, filled by the GUI run-logic.
    output_path : str — optional explicit path; auto-generated if None.

    Returns
    -------
    str : path to the saved report file.
    """
    if output_path is None:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"TL_Design_Report_{ts}.txt"

    lines = []

    # ── Cover ────────────────────────────────────────────────────────────────
    lines += [
        _line("═"),
        "  TRANSMISSION LINE DESIGN REPORT",
        f"  Generated : {datetime.now().strftime('%d %B %Y  %H:%M:%S')}",
        _line("═"),
    ]

    # ── 1. Input Parameters ───────────────────────────────────────────────────
    inp = results.get("inputs", {})
    lines += [
        _section("1. Input Parameters"),
        _row("Power to be transferred",      inp.get("power",       "—"), "MW"),
        _row("Length of transmission line",  inp.get("length",      "—"), "km"),
        _row("Power factor",                 inp.get("powerfactor", "—")),
        _row("Wind pressure wp (fixed)",     100,                          "N/m²"),
        _row("Weight of ice (wice)",          inp.get("wice",        0),    "N/m"),
    ]

    # ── 2. Voltage Selection ──────────────────────────────────────────────────
    vs = results.get("voltage_selection", {})
    lines += [
        _section("2. Voltage Selection"),
        _subsection("Most Economical Voltage (calculated)"),
        _list_row("Economical voltage", [f"{v:.2f}" for v in vs.get("veco_original", [])], "kV"),
        _subsection("Standard Voltage Selected (nearest standard)"),
        _list_row("Standard voltage", vs.get("veconomical", []), "kV"),
    ]

    # ── 3. Technical Analysis ─────────────────────────────────────────────────
    ta = results.get("technical_analysis", {})
    lines += [
        _section("3. Technical Analysis"),
        _row("mf Limit (from length)",       f"{ta.get('mflimit', '—'):.4f}"),
        _list_row("Surge Impedance Loading (SIL)", [f"{s:.2f}" for s in ta.get("sil", [])], "MW"),
        _list_row("Loading factor mf",        [f"{m:.4f}" for m in ta.get("mf",  [])]),
        _list_row("mf Margin",                [f"{m:.4f}" for m in ta.get("mfmargin", [])]),
        "",
        _row("Selected number of circuits (nc)", ta.get("nc", "—")),
        _row("Selected voltage level",        f"{ta.get('selected_voltage', '—')}", "kV"),
    ]

    # ── 4. Air Clearance ──────────────────────────────────────────────────────
    ac = results.get("air_clearance", {})
    lines += [
        _section("4. Air Clearance Calculation"),
        _row("Minimum air clearance (a)",               f"{ac.get('a',  '—'):.4f}", "cm"),
        _row("Length of cross arm (cl)",                f"{ac.get('cl', '—'):.4f}", "cm"),
        _row("Tower width (b)",                         f"{ac.get('b',  '—'):.4f}", "cm"),
        _row("Insulator string length (l)",             f"{ac.get('l',  '—'):.4f}", "cm"),
        _row("Horizontal separation between phases (c)",f"{ac.get('c',  '—'):.4f}", "cm"),
        _row("Vertical separation between phases (y)",  f"{ac.get('y',  '—'):.4f}", "cm"),
        _row("Height of earth wire from top conductor", f"{ac.get('d',  '—'):.4f}", "cm"),
    ]

    # ── 5. Insulator Selection ────────────────────────────────────────────────
    ins = results.get("insulator", {})
    disc_list = ins.get("no_of_disc", [])
    disc_labels = ["Dry withstand", "Wet withstand", "Temp OV withstand",
                   "Lightning OV withstand", "Switching OV withstand"]
    lines += [
        _section("5. Insulator Selection"),
        _row("1-min dry equivalent withstand",    f"{ins.get('equi_dry',        '—'):.4f}", "kV"),
        _row("1-min wet equivalent withstand",    f"{ins.get('equi_wet',        '—'):.4f}", "kV"),
        _row("Temporary overvoltage withstand",   f"{ins.get('temo_ov_wt',      '—'):.4f}", "kV"),
        _row("Lightning overvoltage withstand",   f"{ins.get('lightning_ov_wt', '—'):.4f}", "kV"),
        _row("Switching overvoltage withstand",   f"{ins.get('switching_ov_wt', '—'):.4f}", "kV"),
        _subsection("Discs required per test criterion"),
    ]
    for label, n in zip(disc_labels, disc_list):
        lines.append(_row(label, n, "discs"))
    lines += [
        "",
        _row("★ Number of insulators required (governing)", ins.get("no_of_insulatorrequired", "—"), "discs"),
    ]

    # ── 6. Conductor Selection ────────────────────────────────────────────────
    cond = results.get("conductor", {})
    lines += [
        _section("6. Conductor Selection"),
        _row("Selected ACSR conductor",    cond.get("name",       "—")),
        _row("Transmission efficiency",    f"{cond.get('efficiency', 0):.3f}", "%"),
        _row("Number of bundle conductors (nb)", cond.get("nb",   "—")),
        _row("Resistance at 65°C (r65)",   f"{cond.get('r65',    0):.6f}", "Ω/km"),
        _row("Operating current",          f"{cond.get('current', 0):.4f}", "A"),
    ]

    # ── 7. Voltage Regulation ─────────────────────────────────────────────────
    vr = results.get("voltage_regulation", {})
    lines += [
        _section("7. Voltage Regulation"),
        _row("GMD used",  f"{vr.get('gmd',  0):.6f}", "m"),
        _row("GMRc used", f"{vr.get('gmrc', 0):.6f}", "m"),
        _row("Voltage regulation",          f"{vr.get('vr',   0):.3f}", "%"),
        _row("Limit",                       "12.000",                   "%"),
        _row("Status", "✓ PASS" if vr.get("vr", 999) <= 12 else "✗ FAIL — see notes"),
    ]

    # ── 8. Corona Check ───────────────────────────────────────────────────────
    cr = results.get("corona", {})
    lines += [
        _section("8. Corona Check"),
        _row("Corona inception voltage (Vcr)", f"{cr.get('vcr',          0):.3f}", "kV"),
        _row("System voltage",                 f"{cr.get('system_voltage',0):.1f}", "kV"),
        _row("Status", "✓ PASS" if cr.get("passed", False) else "✗ FAIL — increase bundles"),
    ]

    # ── 9. Tower Design ──────────────────────────────────────────────────────────
    tw = results.get("tower", {})
    lines += [_section("9. Tower Design")]
    if tw.get("ok"):
        lines += [
            _row("Tower Excel output file", tw.get("excel_path", "—")),
            _row("Spans computed", "200 m, 250 m, 300 m, 350 m"),
            _row("Sheets", "Sag_Results | Bending_Moment | Tower"),
            "",
            "    Open the Excel file for full sag/tension, bending moment,",
            "    tower weight, number of towers and cost per unit length.",
        ]
    elif tw.get("error"):
        lines += [_row("Status", f"FAILED — {tw['error']}")]
    else:
        lines += [_row("Status", "Not computed")]

    # ── 10. Notes / Warnings ──────────────────────────────────────────────────────
    notes = results.get("notes", [])
    if notes:
        lines += [_section("10. Notes & Warnings")]
        for n in notes:
            lines.append(f"    • {n}")

    # ── Footer ────────────────────────────────────────────────────────────────
    lines += [
        "",
        _line("═"),
        "  END OF REPORT",
        _line("═"),
        "",
    ]

    content = "\n".join(lines)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(content)

    return output_path


def append_note(results: dict, note: str):
    """Convenience: add a warning/note to the results dict."""
    results.setdefault("notes", []).append(note)
