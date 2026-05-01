# Copyright (c) 2026 rustturf
# All rights reserved.



import math
import cmath
import pandas as pd
import numpy as np

import os as _os
_DIR = _os.path.dirname(_os.path.abspath(__file__))

# ── Load conductor table once ──────────────────────────────────────────────────
table1 = pd.read_excel(
    _os.path.join(_DIR, 'matlabandexcelfiles', 'british.xlsx'),
    header=1
)

# ══════════════════════════════════════════════════════════════════════════════
#  VOLTAGE SELECTION
# ══════════════════════════════════════════════════════════════════════════════

def most_economical_voltage(power, length, powerfactor):
    Veco = []
    for nc in range(1, 3):          # nc = 1, 2
        V = 5.5 * math.sqrt((length / 1.6) + (power * 1000 / (150 * nc * powerfactor)))
        Veco.append(V)
    return Veco


def selectionofstandardvoltage(veco, use_next_higher=False):
    vst = [66, 132, 220, 400, 500, 765]
    selected_voltages = []
    for v in veco:
        if use_next_higher:
            selected = next((s for s in vst if s > v), vst[-1])
        else:
            nearest = min(vst, key=lambda x: abs(x - v))
            selected = nearest
        selected_voltages.append(selected)
    return selected_voltages


# ══════════════════════════════════════════════════════════════════════════════
#  TECHNICAL ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════

def technical_analysis(power, length, veco):
    Lengthh  = [80, 160, 240, 320, 480, 640]
    mflimit  = [2.75, 2.25, 1.75, 1.35, 1, 0.75]

    if length <= Lengthh[0]:
        y = mflimit[0] + ((mflimit[1] - mflimit[0]) / (Lengthh[1] - Lengthh[0])) * (length - Lengthh[0])
    elif length >= Lengthh[-1]:
        y = mflimit[-2] + ((mflimit[-1] - mflimit[-2]) / (Lengthh[-1] - Lengthh[-2])) * (length - Lengthh[-2])
    else:
        for i in range(len(Lengthh) - 1):
            if Lengthh[i] < length < Lengthh[i + 1]:
                y = mflimit[i] + ((mflimit[i+1] - mflimit[i]) / (Lengthh[i+1] - Lengthh[i])) * (length - Lengthh[i])
                break

    SIL_impedance = [400, 200]
    SIL      = [math.pow(veco[i], 2) / SIL_impedance[i] for i in range(2)]
    mf       = [power / SIL[i]                          for i in range(2)]
    mfmargin = [abs(y - mf[i])                          for i in range(2)]

    return y, SIL, mf, mfmargin


def select_nc(y, mfmargin):
    less_than_y = [(mfmargin[i], i + 1) for i in range(len(mfmargin)) if mfmargin[i] < y]
    if len(less_than_y) == 2:
        min_margin, nc = min(less_than_y)
        return nc, False
    elif len(less_than_y) == 1:
        margin, nc = less_than_y[0]
        return nc, False
    else:
        return None, True


# ══════════════════════════════════════════════════════════════════════════════
#  AIR CLEARANCE
# ══════════════════════════════════════════════════════════════════════════════

def airclearancecalculation(nc, veco):
    a  = ((veco[nc - 1] / math.sqrt(3)) * 1.1 * math.sqrt(2)) + 20
    cl = 2 * a
    b  = 2 * a
    l  = math.sqrt(2) * a              # Shielding angle = 45°
    c  = b + 2 * cl
    xbyy  = 1 / 3
    xbyy2 = math.pow(xbyy, 2)
    ratio1   = (l + a) / cl
    powratio1 = math.pow(ratio1, 2)
    y  = (l + a) / math.sqrt(1 - (powratio1 * xbyy2))
    if nc == 2:
        d = math.sqrt(3) * ((b / 2) + a)
    else:
        d = math.sqrt(3) * cl
    return a, cl, b, l, c, y, d


# ══════════════════════════════════════════════════════════════════════════════
#  INSULATOR SELECTION
# ══════════════════════════════════════════════════════════════════════════════

def insulator_selection(veco):
    FWR = 1.15; NAC = 1.1; FS = 1.2
    SSR = 2.8;  SIR = 1.2; EF = 0.8

    table_ws = pd.read_csv(_os.path.join(_DIR, 'Tables', 'withstand voltage capability for different system voltage.csv'))
    list1 = table_ws['Max_System_Voltage'].tolist()
    v1    = next(s for s in list1 if s >= veco)

    dry     = table_ws[table_ws['Max_System_Voltage'] == v1]['Dry withstand'].values[0]
    wet     = table_ws[table_ws['Max_System_Voltage'] == v1]['Wet Withstand'].values[0]
    impulse = table_ws[table_ws['Max_System_Voltage'] == v1]['Impulse Withstand'].values[0]

    equi_dry        = dry     * FWR * NAC * FS
    equi_wet        = wet     * FWR * NAC * FS
    temo_ov_wt      = EF * math.sqrt(2) * v1 * FWR * NAC * FS
    lightning_ov_wt = impulse * FWR * NAC * FS
    switching_ov_wt = (veco / math.sqrt(3)) * 1.1 * math.sqrt(2) * SSR * SIR * FWR * NAC * FS

    table_disc = pd.read_csv(_os.path.join(_DIR, 'Tables', 'Flashover voltage for 254by154mm disc insulators.csv'))
    list2 = table_disc['1min dry FOV'].tolist()
    list3 = table_disc['1min wet FOV'].tolist()
    list4 = table_disc['Impulse Withstand'].tolist()

    v2 = next(s for s in list2 if s >= equi_dry)
    v3 = next(s for s in list3 if s >= equi_wet)
    v4 = next(s for s in list3 if s >= temo_ov_wt)
    v5 = next(s for s in list4 if s >= lightning_ov_wt)
    v6 = next(s for s in list4 if s >= switching_ov_wt)

    no_of_disc = [
        table_disc[table_disc['1min dry FOV']     == v2]['No of disc'].values[0],
        table_disc[table_disc['1min wet FOV']     == v3]['No of disc'].values[0],
        table_disc[table_disc['1min wet FOV']     == v4]['No of disc'].values[0],
        table_disc[table_disc['Impulse Withstand'] == v5]['No of disc'].values[0],
        table_disc[table_disc['Impulse Withstand'] == v6]['No of disc'].values[0],
    ]
    no_of_insulatorrequired = max(no_of_disc)
    return equi_dry, equi_wet, temo_ov_wt, lightning_ov_wt, switching_ov_wt, no_of_disc, no_of_insulatorrequired


# ══════════════════════════════════════════════════════════════════════════════
#  CONDUCTOR SELECTION
# ══════════════════════════════════════════════════════════════════════════════

def conductorselection(veco, length, power, pf, nc, force_nb=None):
    nb = force_nb if force_nb is not None else 1
    resistivitycoefficient = 0.004
    theta2 = 65
    theta1 = 20

    def currentcalc(nb):
        return power * math.pow(10, 6) / (math.sqrt(3) * veco * math.pow(10, 3) * pf * nc * nb)

    def get_conductor(current):
        conductormask = table1['Current Rating Ampacity'] >= current
        actualvalue   = table1['Current Rating Ampacity'][conductormask].min()
        filtered      = table1[table1['Current Rating Ampacity'] == actualvalue]
        idx           = filtered.index[0]
        return idx, table1.iloc[idx]

    lastrow = table1['Current Rating Ampacity'].iloc[-1]
    current = currentcalc(nb)
    if current >= lastrow:
        nb += 1
    current = currentcalc(nb)

    idx, row = get_conductor(current)
    efficiency = 0

    while efficiency < 94:
        r65        = row['Resistance at 20 degree'] * (1 + resistivitycoefficient * (theta2 - theta1))
        loss       = 3 * math.pow(current, 2) * r65 * length * nc * nb
        percentofloss = loss / (power * math.pow(10, 6))
        efficiency = (1 - percentofloss) * 100

        if efficiency < 94:
            idx += 1
            if idx >= len(table1):
                nb += 1
                current = currentcalc(nb)
                idx, row = get_conductor(current)
            else:
                row = table1.iloc[idx]

    return idx, table1['ACSR'].iloc[idx], efficiency, nb, r65, current


# ══════════════════════════════════════════════════════════════════════════════
#  VOLTAGE REGULATION
# ══════════════════════════════════════════════════════════════════════════════

def voltageregulationcalc(V_R_line, l, r65, current, y, c, idx, nb, nc, power_factor,
                           gmd_manual=None, gmrl_manual=None, gmrc_manual=None):
    frequency = 50
    r65 = r65 * l
    c   = c / 100
    y   = y / 100
    r   = table1['Overall Diameter'].iloc[idx] * math.pow(10, -3) / 2

    GMD_out  = None
    GMRc_out = None

    def inductance_calc(gmd, gmrl):
        L_per_meter = 2 * math.pow(10, -7) * math.log(gmd / gmrl)
        L_total     = L_per_meter * 1000 * l
        return L_total

    def capacitance_calc(gmd, gmrc):
        C_per_meter = (2 * math.pi * 8.85 * math.pow(10, -12)) / math.log(gmd / gmrc)
        C_total     = C_per_meter * 1000 * l
        return C_total

    def voltageregulation(L_total, C_total):
        Xl  = 2 * math.pi * frequency * L_total
        Z   = complex(r65, Xl)
        Bc  = 2 * math.pi * frequency * C_total
        Y   = complex(0, Bc)
        ZY  = Z * Y
        A   = 1 + (ZY / 2)
        B   = Z
        V_R_phase = V_R_line / math.sqrt(3)
        V_R       = complex(V_R_phase * 1000, 0)
        pf_angle  = math.acos(power_factor)
        I_R       = cmath.rect(current, -pf_angle)
        V_S_FL    = A * V_R + B * I_R
        V_R_NL    = V_S_FL / A
        V_R_NL_magnitude = abs(V_R_NL) / 1000
        voltage_regulation = ((V_R_NL_magnitude - V_R_phase) / V_R_phase) * 100
        return voltage_regulation

    if nc == 1 and nb == 1:
        GMD_out  = math.pow((math.pow(c / 2, 2) + math.pow(y, 2)) * c, 1 / 3)
        rl       = 0.7788 * r
        GMRc_out = r
        L        = inductance_calc(GMD_out, rl)
        C        = capacitance_calc(GMD_out, GMRc_out)
        voltage_regulation = voltageregulation(L, C)

    elif nc == 2 and nb == 1:
        n    = math.pow(c * c + y * y, 1 / 2)
        m    = math.pow(c * c + math.pow(2 * y, 2), 1 / 2)
        Gmdry = math.pow(y * y * n * n, 1 / 4)
        Gmdyb = math.pow(y * y * n * n, 1 / 4)
        Gmdbr = math.pow(2 * y * 2 * y * c * c, 1 / 4)
        GMD_out = math.pow(Gmdry * Gmdyb * Gmdbr, 1 / 3)
        rl   = 0.7788 * r
        dr   = math.pow(rl * m * m * rl, 1 / 4)
        dy   = math.pow(rl * c * c * rl, 1 / 4)
        db   = math.pow(rl * m * m * rl, 1 / 4)
        GMRl = math.pow(dr * dy * db, 1 / 3)
        drc  = math.pow(r * m * m * r, 1 / 4)
        dyc  = math.pow(r * c * c * r, 1 / 4)
        dbc  = math.pow(r * m * m * r, 1 / 4)
        GMRc_out = math.pow(drc * dyc * dbc, 1 / 3)
        L = inductance_calc(GMD_out, GMRl)
        C = capacitance_calc(GMD_out, GMRc_out)
        voltage_regulation = voltageregulation(L, C)

    else:
        if gmd_manual is None or gmrl_manual is None or gmrc_manual is None:
            raise ValueError("Manual GMD, GMRl, GMRc must be provided for this configuration.")
        GMD_out  = gmd_manual
        GMRc_out = gmrc_manual
        L        = inductance_calc(gmd_manual, gmrl_manual)
        C        = capacitance_calc(gmd_manual, gmrc_manual)
        voltage_regulation = voltageregulation(L, C)

    return voltage_regulation, GMD_out, GMRc_out


# ══════════════════════════════════════════════════════════════════════════════
#  CORONA
# ══════════════════════════════════════════════════════════════════════════════

def corona_inceptionvoltage(GMD, GMRc):
    vcr = math.sqrt(3) * (30 / math.sqrt(2)) * 100 * GMRc * math.log(GMD / GMRc) * 0.95 * 0.95
    return vcr


# ══════════════════════════════════════════════════════════════════════════════
#  TOWER DESIGN
# ══════════════════════════════════════════════════════════════════════════════

def towerdesign(veco, Length, y, nc, d, wice=0):
    y = y / 100
    d = d / 100

    df       = pd.read_excel(_os.path.join(_DIR, 'matlabandexcelfiles', 'towerheightcalculation.xlsx'))
    dc       = df['Overall Diameter'].tolist()
    area     = df['Overall Area'].tolist()
    alpha    = df['Lactual'].tolist()
    wc       = df['Weight'].tolist()
    uts      = df['UTS'].tolist()
    yme      = df['YME'].tolist()
    conductorname = df['ACSR'].tolist()

    wp     = 100  # wind pressure — fixed per original design
    theta1 = 0; theta2 = 27; theta3 = 65
    # wice passed in from user input
    fs     = 2

    dew   = 14.60; utsew = 6664; ne = 2
    alpha1 = 1  * math.pi / 180
    alpha2 = 7.5 * math.pi / 180
    alpha3 = 15 * math.pi / 180
    sinvalue = (0.8 * math.sin(alpha1)) + (0.15 * math.sin(alpha2)) + (0.05 * math.sin(alpha3))

    coststeel = 150000
    Ww = []; t1 = []; w1 = []

    for i in range(len(dc)):
        Ww.append(round(wp * (2 / 3) * dc[i],  4))
        t1.append(round((uts[i] / 9.81) / fs,  4))
        w1.append(round(math.sqrt(Ww[i] ** 2 + (wc[i] + wice) ** 2), 4))

    l = [200, 250, 300, 350]
    results = []; bendingmoment = []; tttt = []

    for i in range(len(dc)):
        for span in l:
            k1 = (-t1[i]
                  + ((alpha[i] * (theta2 - theta1)) * area[i] * (yme[i] / 9.81))
                  + (((w1[i] ** 2) * ((span / 1000) ** 2)) / (24 * (t1[i] ** 2)) * area[i] * (yme[i] / 9.81)))
            k2     = ((wc[i] ** 2) * ((span / 1000) ** 2) * (yme[i] / 9.81) * area[i]) / 24
            coeff  = [1, k1, 0, -k2]
            roots  = np.roots(coeff)
            t2     = [r.real for r in roots if abs(r.imag) < 1e-6][0]
            k11    = (-t2
                      + ((alpha[i] * (theta3 - theta2)) * area[i] * (yme[i] / 9.81))
                      + (((wc[i] ** 2) * ((span / 1000) ** 2)) / (24 * (t2 ** 2)) * area[i] * (yme[i] / 9.81)))
            coeff  = [1, k11, 0, -k2]
            roots  = np.roots(coeff)
            t3     = [r.real for r in roots if abs(r.imag) < 1e-6][0]
            sagmax = (wc[i] / 1000) * (span ** 2) / (8 * t3)
            ho     = ((((veco * 1.1) - 33) / 33) + 17) * 0.305
            h1     = ho + sagmax
            if nc == 1:
                h2 = h1 + (y / 2); h3 = h2 + (y / 2)
            else:
                h2 = h1 + y;       h3 = h2 + y
            ht = h3 + d
            results.append([conductorname[i], span, round(k1,4), round(k2,4), round(k11,4), t2, t3, sagmax, ho, h1, h2, h3, ht])

            mt   = utsew / 2
            fwp  = (2/3) * dc[i] * span * math.pow(10,-3) * wp
            bmpw = (h1 + h2 + h3) * nc * fwp
            fwe  = wp * dew * span * (2/3) * math.pow(10,-3)
            bmew = ht * ne * fwe
            bmpt = 2 * t1[i] * sinvalue * (h1 + h2 + h3) * nc
            bmet = 2 * mt   * sinvalue * ne * ht
            tbm  = bmpw + bmew + bmpt + bmet
            TW   = 0.000631 * ht * math.sqrt(tbm * fs)
            bendingmoment.append([conductorname[i], span, fwp, bmpw, fwe, bmew, bmpt, bmet, tbm, TW])

            nt           = (Length * math.pow(10, 3) / span) + 1
            costpertower = coststeel * TW
            costtowerpul = costpertower * nt / Length
            tttt.append([conductorname[i], span, TW, math.ceil(nt), costpertower, costtowerpul])

    results_df = pd.DataFrame(results, columns=[
        'Conductor','Span Length','k1','k2','K11','T2','T3','sagmax','ho','h1','h2','h3','ht'])
    results1_df = pd.DataFrame(bendingmoment, columns=[
        'Conductor','Span Length','FWP','BMPW','FWE','BMEW','BMPT','BMET','TBM','Tower Weight'])
    results2_df = pd.DataFrame(tttt, columns=[
        'Conductor','Span','Tower Weight','Number of Towers','Cost per Tower','Cost per Tower per Unit Length'])

    output_path = _os.path.join(_os.path.expanduser('~'), 'Desktop', 'TL_Tower_Design.xlsx')
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        results_df.to_excel(writer,  sheet_name='Sag_Results',     index=False)
        results1_df.to_excel(writer, sheet_name='Bending_Moment',  index=False)
        results2_df.to_excel(writer, sheet_name='Tower',           index=False)

    return output_path
