# Copyright (c) 2026 rustturf
# All rights reserved.



"""
gui.py  —  Transmission Line Design Tool
Clean tkinter GUI. All calculation logic lives in calculations.py.
Results are written to a report file via report_writer.py.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os
from datetime import datetime

import calculations as calc
import report_writer as rw


# ══════════════════════════════════════════════════════════════════════════════
#  THEME / PALETTE
# ══════════════════════════════════════════════════════════════════════════════

BG       = "#0f1117"
PANEL    = "#1a1d27"
ACCENT   = "#4f8ef7"
ACCENT2  = "#7c3aed"
SUCCESS  = "#22c55e"
WARNING  = "#f59e0b"
DANGER   = "#ef4444"
FG       = "#e2e8f0"
FG_DIM   = "#64748b"
BORDER   = "#2d3148"
FONT     = ("Consolas", 10)
FONT_LG  = ("Consolas", 12, "bold")
FONT_H   = ("Consolas", 14, "bold")
FONT_SM  = ("Consolas", 9)


# ══════════════════════════════════════════════════════════════════════════════
#  WIDGET HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def styled_label(parent, text, font=FONT, fg=FG, **kw):
    return tk.Label(parent, text=text, font=font, fg=fg,
                    bg=kw.pop("bg", PANEL), **kw)

def styled_entry(parent, width=22, **kw):
    e = tk.Entry(parent, width=width, font=FONT,
                 bg="#252836", fg=FG, insertbackground=FG,
                 relief="flat", bd=6,
                 highlightthickness=1, highlightcolor=ACCENT,
                 highlightbackground=BORDER, **kw)
    return e

def styled_button(parent, text, command, color=ACCENT, **kw):
    b = tk.Button(parent, text=text, command=command,
                  font=("Consolas", 10, "bold"),
                  bg=color, fg="white", activebackground=color,
                  activeforeground="white", relief="flat",
                  cursor="hand2", padx=16, pady=6, **kw)
    b.bind("<Enter>", lambda e: b.config(bg=_lighten(color)))
    b.bind("<Leave>", lambda e: b.config(bg=color))
    return b

def _lighten(hex_color):
    """Slightly lighten a hex color for hover."""
    c = hex_color.lstrip("#")
    rgb = tuple(min(255, int(c[i:i+2], 16) + 30) for i in (0, 2, 4))
    return "#{:02x}{:02x}{:02x}".format(*rgb)

def card(parent, **kw):
    f = tk.Frame(parent, bg=PANEL, bd=0,
                 highlightthickness=1, highlightbackground=BORDER,
                 **kw)
    return f


# ══════════════════════════════════════════════════════════════════════════════
#  LOG PANEL
# ══════════════════════════════════════════════════════════════════════════════

class LogPanel(tk.Frame):
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=PANEL, **kw)
        styled_label(self, "  ◈ DESIGN LOG", font=FONT_LG,
                     fg=ACCENT).pack(anchor="w", pady=(8, 4), padx=8)
        self.text = tk.Text(self, bg="#0d1117", fg=FG, font=FONT_SM,
                            relief="flat", bd=0, padx=10, pady=8,
                            wrap="word", state="disabled",
                            highlightthickness=0)
        sb = ttk.Scrollbar(self, command=self.text.yview)
        self.text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.text.pack(fill="both", expand=True, padx=(8, 0), pady=(0, 8))

        # tags
        self.text.tag_configure("info",    foreground=FG)
        self.text.tag_configure("ok",      foreground=SUCCESS)
        self.text.tag_configure("warn",    foreground=WARNING)
        self.text.tag_configure("err",     foreground=DANGER)
        self.text.tag_configure("section", foreground=ACCENT, font=("Consolas", 10, "bold"))
        self.text.tag_configure("dim",     foreground=FG_DIM)

    def log(self, msg, tag="info"):
        self.text.configure(state="normal")
        self.text.insert("end", msg + "\n", tag)
        self.text.see("end")
        self.text.configure(state="disabled")

    def clear(self):
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")

    def section(self, title):
        self.log(f"\n{'─'*60}", "dim")
        self.log(f"  {title}", "section")
        self.log(f"{'─'*60}", "dim")


# ══════════════════════════════════════════════════════════════════════════════
#  STATUS BAR
# ══════════════════════════════════════════════════════════════════════════════

class StatusBar(tk.Frame):
    def __init__(self, parent, **kw):
        super().__init__(parent, bg="#080b12", height=28, **kw)
        self.label = tk.Label(self, text="Ready", font=FONT_SM,
                              fg=FG_DIM, bg="#080b12", anchor="w")
        self.label.pack(side="left", padx=12)
        self.progress = ttk.Progressbar(self, mode="indeterminate", length=120)
        self.progress.pack(side="right", padx=12, pady=4)

    def set(self, msg, color=FG_DIM):
        self.label.config(text=msg, fg=color)

    def busy(self, on):
        if on:
            self.progress.start(12)
        else:
            self.progress.stop()


# ══════════════════════════════════════════════════════════════════════════════
#  GMD INPUT DIALOG
# ══════════════════════════════════════════════════════════════════════════════

class GmdDialog(tk.Toplevel):
    """
    Shown when nc > 2 or nb > 1 — asks user for GMD, GMRl, GMRc.
    result is set to (gmd, gmrl, gmrc) on confirm, or None on cancel.
    """
    def __init__(self, parent, nc, nb, log_fn):
        super().__init__(parent)
        self.title("Manual Geometry Input Required")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.grab_set()
        self.result = None

        # ── Header ────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=ACCENT2, height=4)
        hdr.pack(fill="x")
        tk.Label(self, text="⚠  Manual Geometry Input",
                 font=FONT_H, fg=FG, bg=BG).pack(pady=(14, 2))
        tk.Label(self,
                 text=f"Configuration  nc={nc}, nb={nb}  requires manual GMD / GMR values.\n"
                      "Please calculate and enter them below.",
                 font=FONT_SM, fg=FG_DIM, bg=BG, wraplength=380, justify="center"
                 ).pack(padx=20, pady=(0, 10))

        body = card(self)
        body.pack(padx=20, pady=8, fill="x")

        fields = [
            ("GMD  — Geometric Mean Distance", "m"),
            ("GMRl — GMR for Inductance",      "m"),
            ("GMRc — GMR for Capacitance",     "m"),
        ]
        self._entries = []
        for label, unit in fields:
            row = tk.Frame(body, bg=PANEL)
            row.pack(fill="x", padx=12, pady=5)
            tk.Label(row, text=f"{label}  ({unit})", font=FONT,
                     fg=FG, bg=PANEL, width=38, anchor="w").pack(side="left")
            e = styled_entry(row, width=14)
            e.pack(side="right")
            self._entries.append(e)

        self._status = tk.Label(self, text="", font=FONT_SM, bg=BG, fg=WARNING)
        self._status.pack(pady=4)

        btn_row = tk.Frame(self, bg=BG)
        btn_row.pack(pady=12)
        styled_button(btn_row, "  Confirm  ", self._confirm, color=SUCCESS).pack(side="left", padx=6)
        styled_button(btn_row, "  Cancel   ", self.destroy,  color=DANGER ).pack(side="left", padx=6)

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self._center(parent)

    def _center(self, parent):
        self.update_idletasks()
        pw = parent.winfo_rootx(); ph = parent.winfo_rooty()
        x = pw + (parent.winfo_width()  - self.winfo_width())  // 2
        y = ph + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _confirm(self):
        try:
            vals = [float(e.get()) for e in self._entries]
            if any(v <= 0 for v in vals):
                raise ValueError("All values must be positive.")
            self.result = tuple(vals)
            self.destroy()
        except ValueError as ex:
            self._status.config(text=f"✗  {ex}", fg=DANGER)


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN APPLICATION
# ══════════════════════════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Transmission Line Design Tool")
        self.configure(bg=BG)
        self.minsize(1050, 680)
        self._results = {}          # accumulates data for the report
        self._build_ui()

    # ── UI Layout ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        # top bar
        topbar = tk.Frame(self, bg=ACCENT2, height=3)
        topbar.pack(fill="x")

        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=18, pady=(10, 4))
        tk.Label(header, text="⚡  TRANSMISSION LINE DESIGN",
                 font=("Consolas", 16, "bold"), fg=FG, bg=BG).pack(side="left")
        tk.Label(header, text="Power Systems Engineering Tool",
                 font=FONT_SM, fg=FG_DIM, bg=BG).pack(side="left", padx=12, pady=4)

        # main area: left panel + log
        main = tk.Frame(self, bg=BG)
        main.pack(fill="both", expand=True, padx=12, pady=4)

        left = tk.Frame(main, bg=BG, width=320)
        left.pack(side="left", fill="y", padx=(0, 8))
        left.pack_propagate(False)

        self._build_input_panel(left)

        self.log = LogPanel(main)
        self.log.pack(side="left", fill="both", expand=True)

        self.status = StatusBar(self)
        self.status.pack(fill="x", side="bottom")

    def _build_input_panel(self, parent):
        # ── Primary Inputs ──────────────────────────────────────────────
        c1 = card(parent)
        c1.pack(fill="x", pady=(0, 8))
        styled_label(c1, "  ◈ LINE PARAMETERS", font=FONT_LG,
                     fg=ACCENT).pack(anchor="w", padx=8, pady=(8, 4))

        fields = [
            ("Power  (MW)",         "power_entry",       "e.g. 300"),
            ("Length  (km)",        "length_entry",      "e.g. 200"),
            ("Power Factor",        "pf_entry",          "0.0 – 1.0"),
        ]
        for label, attr, hint in fields:
            self._labeled_entry(c1, label, attr, hint)

        # ── Ice / Wind Load ─────────────────────────────────────────────
        c2 = card(parent)
        c2.pack(fill="x", pady=(0, 8))
        styled_label(c2, "  ◈ ICE LOADING", font=FONT_LG,
                     fg=ACCENT).pack(anchor="w", padx=8, pady=(8, 4))
        self._labeled_entry(c2, "Weight of Ice  (N/m)", "wice_entry", "0 if no ice")

        # ── Output Path ──────────────────────────────────────────────────
        c3 = card(parent)
        c3.pack(fill="x", pady=(0, 8))
        styled_label(c3, "  ◈ OUTPUT FILE", font=FONT_LG,
                     fg=ACCENT).pack(anchor="w", padx=8, pady=(8, 4))
        path_row = tk.Frame(c3, bg=PANEL)
        path_row.pack(fill="x", padx=12, pady=6)
        _desktop = os.path.join(os.path.expanduser("~"), "Desktop", "TL_Design_Report.txt")
        self.path_var = tk.StringVar(value=_desktop)
        path_e = styled_entry(path_row, width=18, textvariable=self.path_var)
        path_e.pack(side="left")
        styled_button(path_row, "…", self._browse_output,
                      color=ACCENT2).pack(side="left", padx=4)

        # ── Action Buttons ────────────────────────────────────────────────
        btn_card = card(parent)
        btn_card.pack(fill="x", pady=(0, 4))
        bf = tk.Frame(btn_card, bg=PANEL)
        bf.pack(padx=12, pady=12, fill="x")
        styled_button(bf, "▶  RUN DESIGN", self._run,
                      color=SUCCESS).pack(fill="x", pady=(0, 6))
        styled_button(bf, "⟳  CLEAR LOG",  self._clear,
                      color=ACCENT2).pack(fill="x")

    def _labeled_entry(self, parent, label, attr, hint=""):
        row = tk.Frame(parent, bg=PANEL)
        row.pack(fill="x", padx=12, pady=4)
        tk.Label(row, text=label, font=FONT, fg=FG, bg=PANEL,
                 anchor="w", width=24).pack(side="left")
        e = styled_entry(row, width=12)
        e.pack(side="left", padx=4)
        if hint:
            tk.Label(row, text=hint, font=FONT_SM, fg=FG_DIM, bg=PANEL).pack(side="left")
        setattr(self, attr, e)

    # ── Callbacks ─────────────────────────────────────────────────────────────

    def _browse_output(self):
        p = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title="Save report as…")
        if p:
            self.path_var.set(p)

    def _clear(self):
        self.log.clear()
        self.status.set("Ready")

    def _run(self):
        # validate inputs
        try:
            power       = float(self.power_entry.get())
            length      = float(self.length_entry.get())
            powerfactor = float(self.pf_entry.get())
            wice_text   = self.wice_entry.get().strip()
            wice        = float(wice_text) if wice_text else 0.0
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numeric values for all fields.")
            return

        if not (0 < powerfactor <= 1):
            messagebox.showerror("Input Error", "Power factor must be between 0 and 1.")
            return

        _fb = os.path.join(os.path.expanduser("~"), "Desktop", "TL_Design_Report.txt")
        output_path = self.path_var.get().strip() or _fb
        self.log.clear()
        self._results = {"inputs": {
            "power": power, "length": length,
            "powerfactor": powerfactor, "wice": wice
        }, "notes": []}

        # run in thread so GUI stays responsive
        self.status.set("Running design calculations…", color=ACCENT)
        self.status.busy(True)
        t = threading.Thread(target=self._design_thread,
                             args=(power, length, powerfactor, wice, output_path),
                             daemon=True)
        t.start()

    # ── Core Design Logic ─────────────────────────────────────────────────────

    def _design_thread(self, power, length, powerfactor, wice, output_path):
        """Runs the full design sequence; calls back to GUI via .after()."""
        try:
            self._run_design(power, length, powerfactor, wice, output_path)
        except Exception as ex:
            self.after(0, self._finish_error, str(ex))

    def _finish_error(self, msg):
        self.log.log(f"\n✗  ERROR: {msg}", "err")
        self.status.set(f"Error — {msg}", color=DANGER)
        self.status.busy(False)

    def _run_design(self, power, length, powerfactor, wice, output_path):
        log  = lambda msg, tag="info": self.after(0, self.log.log, msg, tag)
        sec  = lambda t:               self.after(0, self.log.section, t)
        res  = self._results

        # ── 1. Economical Voltage ────────────────────────────────────────
        sec("1 · ECONOMICAL VOLTAGE")
        veco_orig = calc.most_economical_voltage(power, length, powerfactor)
        log(f"  nc=1: {veco_orig[0]:.2f} kV    nc=2: {veco_orig[1]:.2f} kV")
        Veconomical = calc.selectionofstandardvoltage(veco_orig, use_next_higher=False)
        log(f"  Standard voltages → nc=1: {Veconomical[0]} kV    nc=2: {Veconomical[1]} kV")
        res["voltage_selection"] = {
            "veco_original": veco_orig,
            "veconomical":   Veconomical,
        }

        # ── 2. Technical Analysis / nc selection ─────────────────────────
        sec("2 · TECHNICAL ANALYSIS")
        nc = None
        attempt = 0
        MAX_ATTEMPTS = 10
        while attempt < MAX_ATTEMPTS:
            attempt += 1
            mflimit, SIL, mf, mfmargin = calc.technical_analysis(power, length, Veconomical)
            log(f"  [Attempt {attempt}]  Voltages: nc=1→{Veconomical[0]} kV  nc=2→{Veconomical[1]} kV")
            log(f"  mf Limit: {mflimit:.4f}   SIL: {SIL[0]:.2f}/{SIL[1]:.2f} MW")
            log(f"  mf: {mf[0]:.4f}/{mf[1]:.4f}   Margin: {mfmargin[0]:.4f}/{mfmargin[1]:.4f}")
            nc, needs_recalc = calc.select_nc(mflimit, mfmargin)
            if not needs_recalc:
                log(f"  ✓ Selected nc = {nc}", "ok")
                break
            log("  Both margins > mflimit — trying next higher voltage…", "warn")
            Veconomical = calc.selectionofstandardvoltage(Veconomical, use_next_higher=True)
            if all(v == 765 for v in Veconomical):
                log("  ✗ Reached max standard voltage (765 kV). No valid solution.", "err")
                rw.append_note(res, "Reached 765 kV limit without valid mf margin — manual review needed.")
                break

        res["technical_analysis"] = {
            "mflimit": mflimit, "sil": SIL, "mf": mf, "mfmargin": mfmargin,
            "nc": nc, "selected_voltage": Veconomical[nc - 1] if nc else "—",
        }

        # ── 3. Air Clearance ────────────────────────────────────────────
        sec("3 · AIR CLEARANCE")
        a, cl, b, l, c, y, d = calc.airclearancecalculation(nc, Veconomical)
        log(f"  Minimum air clearance (a)   : {a:.4f} cm")
        log(f"  Cross arm length (cl)        : {cl:.4f} cm")
        log(f"  Tower width (b)              : {b:.4f} cm")
        log(f"  Insulator string length (l)  : {l:.4f} cm")
        log(f"  Horizontal separation (c)    : {c:.4f} cm")
        log(f"  Vertical separation (y)      : {y:.4f} cm")
        log(f"  Earth wire height (d)        : {d:.4f} cm")
        res["air_clearance"] = {"a": a, "cl": cl, "b": b, "l": l, "c": c, "y": y, "d": d}

        # ── 4. Insulator Selection ───────────────────────────────────────
        sec("4 · INSULATOR SELECTION")
        equi_dry, equi_wet, temo_ov_wt, lightning_ov_wt, switching_ov_wt, \
            no_of_disc, no_of_insulatorrequired = calc.insulator_selection(Veconomical[nc - 1])
        log(f"  1-min dry withstand          : {equi_dry:.4f} kV")
        log(f"  1-min wet withstand          : {equi_wet:.4f} kV")
        log(f"  Temporary OV withstand       : {temo_ov_wt:.4f} kV")
        log(f"  Lightning OV withstand       : {lightning_ov_wt:.4f} kV")
        log(f"  Switching OV withstand       : {switching_ov_wt:.4f} kV")
        log(f"  Discs per criterion          : {no_of_disc}")
        log(f"  ★ Insulators required        : {no_of_insulatorrequired}", "ok")
        res["insulator"] = {
            "equi_dry": equi_dry, "equi_wet": equi_wet,
            "temo_ov_wt": temo_ov_wt, "lightning_ov_wt": lightning_ov_wt,
            "switching_ov_wt": switching_ov_wt,
            "no_of_disc": no_of_disc,
            "no_of_insulatorrequired": no_of_insulatorrequired,
        }

        # ── 5. Conductor Selection ───────────────────────────────────────
        sec("5 · CONDUCTOR SELECTION")
        MAX_BUNDLES  = 6
        forced_nb    = None
        gmd_used     = None
        gmrc_used    = None
        vr_result    = None
        conductor_ok = False
        idx = nb = r65 = current = name = efficiency = None

        while True:
            idx, name, efficiency, nb, r65, current = calc.conductorselection(
                Veconomical[nc - 1], length, power, powerfactor, nc, force_nb=forced_nb
            )
            log(f"  Conductor: {name}   Efficiency: {efficiency:.3f}%   Bundles: {nb}")

            needs_manual_gmd = not (nc in [1, 2] and nb == 1)

            if needs_manual_gmd:
                # ── Ask user for GMD values (must happen on main thread) ──
                log(f"  ⚠ nc={nc}, nb={nb} — manual GMD input needed.", "warn")
                gmd_val = gmrl_val = gmrc_val = None
                done_event = threading.Event()

                def _ask_gmd(nc=nc, nb=nb):
                    dlg = GmdDialog(self, nc, nb, log)
                    self.wait_window(dlg)
                    nonlocal gmd_val, gmrl_val, gmrc_val
                    if dlg.result:
                        gmd_val, gmrl_val, gmrc_val = dlg.result
                    done_event.set()

                self.after(0, _ask_gmd)
                done_event.wait()

                if gmd_val is None:
                    log("  ✗ GMD input cancelled by user.", "err")
                    self.after(0, self._finish_error, "GMD input cancelled.")
                    return

                log(f"  GMD={gmd_val:.5f}  GMRl={gmrl_val:.5f}  GMRc={gmrc_val:.5f}")
                vr_result, gmd_used, gmrc_used = calc.voltageregulationcalc(
                    Veconomical[nc - 1], length, r65, current, y, c, idx, nb, nc, powerfactor,
                    gmd_manual=gmd_val, gmrl_manual=gmrl_val, gmrc_manual=gmrc_val
                )
            else:
                vr_result, gmd_used, gmrc_used = calc.voltageregulationcalc(
                    Veconomical[nc - 1], length, r65, current, y, c, idx, nb, nc, powerfactor
                )

            log(f"  Voltage Regulation: {vr_result:.3f}%")
            if vr_result <= 12:
                log(f"  ✓ VR acceptable ({vr_result:.3f}% ≤ 12%)", "ok")
                conductor_ok = True
                break
            next_nb = nb + 1
            if next_nb > MAX_BUNDLES:
                log(f"  ✗ VR still {vr_result:.3f}% at max bundles ({MAX_BUNDLES}). "
                    f"Consider a higher voltage level.", "err")
                rw.append_note(res, f"VR={vr_result:.3f}% exceeds 12% at max bundles — higher voltage recommended.")
                break
            log(f"  VR={vr_result:.3f}% > 12%. Increasing nb → {next_nb}…", "warn")
            forced_nb = next_nb

        res["conductor"] = {
            "name": name, "efficiency": efficiency,
            "nb": nb, "r65": r65, "current": current,
        }

        # ── 6. Voltage Regulation Summary ────────────────────────────────
        sec("6 · VOLTAGE REGULATION SUMMARY")
        log(f"  GMD  : {gmd_used:.6f} m")
        log(f"  GMRc : {gmrc_used:.6f} m")
        log(f"  VR   : {vr_result:.3f} %")
        res["voltage_regulation"] = {"gmd": gmd_used, "gmrc": gmrc_used, "vr": vr_result}

        # ── 7. Corona Check ─────────────────────────────────────────────
        sec("7 · CORONA CHECK")
        corona_passed = False
        while True:
            vcr = calc.corona_inceptionvoltage(gmd_used, gmrc_used)
            sys_v = Veconomical[nc - 1]
            log(f"  Corona inception Vcr : {vcr:.3f} kV")
            log(f"  System voltage       : {sys_v} kV")
            if vcr >= sys_v:
                log(f"  ✓ Corona check PASSED (Vcr={vcr:.3f} ≥ {sys_v} kV)", "ok")
                corona_passed = True
                break
            next_nb = nb + 1
            if next_nb > MAX_BUNDLES:
                log(f"  ✗ Corona still failing at max bundles. Higher voltage level needed.", "err")
                rw.append_note(res, f"Corona check failed at max bundles — higher voltage level recommended.")
                break
            log(f"  ✗ Vcr={vcr:.3f} < {sys_v} kV. Increasing nb → {next_nb}…", "warn")
            forced_nb = next_nb
            idx, name, efficiency, nb, r65, current = calc.conductorselection(
                Veconomical[nc - 1], length, power, powerfactor, nc, force_nb=forced_nb
            )
            needs_manual_gmd = not (nc in [1, 2] and nb == 1)
            if needs_manual_gmd:
                log(f"  ⚠ Corona: nc={nc}, nb={nb} — manual GMD input needed.", "warn")
                gmd_val = gmrl_val = gmrc_val = None
                done_event2 = threading.Event()

                def _ask_gmd2(nc=nc, nb=nb):
                    dlg = GmdDialog(self, nc, nb, log)
                    self.wait_window(dlg)
                    nonlocal gmd_val, gmrl_val, gmrc_val
                    if dlg.result:
                        gmd_val, gmrl_val, gmrc_val = dlg.result
                    done_event2.set()

                self.after(0, _ask_gmd2)
                done_event2.wait()
                if gmd_val is None:
                    log("  ✗ GMD input cancelled.", "err")
                    break
                vr_result, gmd_used, gmrc_used = calc.voltageregulationcalc(
                    Veconomical[nc - 1], length, r65, current, y, c, idx, nb, nc, powerfactor,
                    gmd_manual=gmd_val, gmrl_manual=gmrl_val, gmrc_manual=gmrc_val
                )
            else:
                vr_result, gmd_used, gmrc_used = calc.voltageregulationcalc(
                    Veconomical[nc - 1], length, r65, current, y, c, idx, nb, nc, powerfactor
                )

        res["corona"] = {"vcr": vcr, "system_voltage": Veconomical[nc - 1], "passed": corona_passed}

        # ── 8. Tower Design ──────────────────────────────────────────────
        sec("8 · TOWER DESIGN")
        log("  Running sag/tension and tower height calculations…")
        log("  (Spans: 200, 250, 300, 350 m — results saved to Excel)")
        try:
            tower_xlsx = calc.towerdesign(
                veco  = Veconomical[nc - 1],
                Length= length,
                y     = res["air_clearance"]["y"],   # vertical separation (cm)
                nc    = nc,
                d     = res["air_clearance"]["d"],   # earth wire height (cm)
                wice  = wice,
            )
            log(f"  ✓ Tower Excel saved → {tower_xlsx}", "ok")
            log(f"  Sheets: Sag_Results | Bending_Moment | Tower (cost)")
            res["tower"] = {"excel_path": tower_xlsx, "ok": True}
        except Exception as te:
            log(f"  ✗ Tower design failed: {te}", "err")
            rw.append_note(res, f"Tower design error: {te}")
            res["tower"] = {"excel_path": None, "ok": False, "error": str(te)}

        # ── 9. Write Report ───────────────────────────────────────────────
        sec("9 · SAVING REPORT")
        saved_path = rw.create_report(res, output_path)
        log(f"  ✓ Report saved → {saved_path}", "ok")

        # done
        self.after(0, self._finish_ok, saved_path, res.get("tower", {}).get("excel_path"))

    def _finish_ok(self, path, tower_xlsx=None):
        self.status.set(f"✓  Design complete  —  Report: {path}", color=SUCCESS)
        self.status.busy(False)
        tower_line = f"\n\nTower Excel: {tower_xlsx}" if tower_xlsx else ""
        if messagebox.askyesno("Design Complete",
                               f"Design complete!\n\nText report: {path}{tower_line}\n\nOpen the text report now?"):
            os.startfile(path)
        if tower_xlsx and os.path.exists(tower_xlsx):
            if messagebox.askyesno("Open Tower Excel?",
                                   f"Open tower design Excel file?\n{tower_xlsx}"):
                os.startfile(tower_xlsx)

# ══════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = App()
    app.mainloop()
