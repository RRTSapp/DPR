"""
financial_model.py
Pure-Python financial model for solar DPR.
All inputs via a `params` dict. Returns cashflows + IRR + DSCR.
"""

import math


def run_model(params: dict) -> dict:
    cap_ac   = params["project_capacity_ac"]
    cost_lac = params["project_cost_lac"]
    debt_lac = params["debt_lac"]
    equity   = params["equity_lac"]
    r_debt   = params["debt_interest_rate"] / 100
    tenor    = params["debt_tenor_yrs"]
    tariff   = params["ppa_tariff"]
    term     = params["ppa_term"]
    gen_yr1  = params["gen_yr1_lac_per_mwac"] * cap_ac
    deg1     = params["degradation_yr1_pct"] / 100
    deg2     = params["degradation_yr2_pct"] / 100
    om_free  = params["om_yr1_free"]
    om_base  = params["om_rate_lac_per_mwac"] * cap_ac
    om_esc   = params["om_escalation_pct"] / 100
    ins_rate = params["insurance_pct"] / 100
    ins_esc  = params["insurance_esc_pct"] / 100

    dep_slm  = cost_lac / term          # SLM depreciation
    ann_prin = debt_lac / tenor         # Straight-line principal

    def generation(yr):
        if yr == 1:
            return gen_yr1
        return gen_yr1 * (1 - deg1) * ((1 - deg2) ** (yr - 2))

    def revenue(yr):
        return generation(yr) * tariff

    def om(yr):
        if yr == 1 and om_free:
            return 0.0
        base_yr = 2 if om_free else 1
        return om_base * ((1 + om_esc) ** (yr - base_yr))

    def insurance(yr):
        return cost_lac * ins_rate * ((1 + ins_esc) ** (yr - 1))

    def interest(yr):
        if yr > tenor:
            return 0.0
        opening = debt_lac - ann_prin * (yr - 1)
        return opening * r_debt

    def principal(yr):
        return ann_prin if yr <= tenor else 0.0

    # Tax depreciation WDV 40%
    tax_wdv = cost_lac
    carry_loss = 0.0
    cashflows = []

    for yr in range(1, term + 1):
        rev  = revenue(yr)
        o_m  = om(yr)
        ins  = insurance(yr)
        inte = interest(yr)
        prin = principal(yr)
        op_debt = max(debt_lac - ann_prin * (yr - 1), 0) if yr <= tenor else 0
        cl_debt = max(op_debt - prin, 0)

        ebitda = rev - o_m - ins
        ebit   = ebitda - dep_slm
        pbt    = ebit - inte

        # Tax depreciation
        it_dep    = tax_wdv * 0.40
        tax_wdv  -= it_dep

        # Taxable income (normal regime)
        t_inc = rev - o_m - ins - it_dep - inte
        if carry_loss > 0:
            if t_inc > 0:
                offset     = min(t_inc, carry_loss)
                t_inc     -= offset
                carry_loss -= offset
            else:
                carry_loss += abs(t_inc)
                t_inc = 0.0
        elif t_inc < 0:
            carry_loss = abs(t_inc)
            t_inc = 0.0

        normal_tax = t_inc * 0.3338 if t_inc > 0 else 0.0
        mat_tax    = pbt * 0.1669   if pbt > 0   else 0.0
        tax        = max(normal_tax, mat_tax)

        pat         = pbt - tax
        debt_svc    = prin + inte
        cash_for_ds = ebitda - tax
        dscr        = cash_for_ds / debt_svc if debt_svc > 0 else 0.0
        fcfe        = pat + dep_slm - prin

        cashflows.append({
            "year": yr, "gen": generation(yr), "revenue": rev,
            "om": o_m, "insurance": ins, "ebitda": ebitda,
            "depreciation": dep_slm, "ebit": ebit,
            "interest": inte, "pbt": pbt, "tax": tax, "pat": pat,
            "principal": prin, "op_debt": op_debt, "cl_debt": cl_debt,
            "dscr": dscr, "fcfe": fcfe,
            "debt_service": debt_svc, "cash_for_ds": cash_for_ds,
        })

    # IRR calculation
    def npv(rate, cfs):
        return sum(cf / ((1 + rate) ** i) for i, cf in enumerate(cfs))

    def irr(cfs):
        rate = 0.10
        for _ in range(2000):
            n  = npv(rate, cfs)
            dn = sum(-i * cf / ((1 + rate) ** (i + 1))
                     for i, cf in enumerate(cfs) if i > 0)
            if abs(dn) < 1e-10:
                break
            new_rate = rate - n / dn
            if abs(new_rate - rate) < 1e-8:
                rate = new_rate
                break
            rate = new_rate
        return rate

    proj_cfs = [-cost_lac] + [cf["ebitda"] - cf["tax"] for cf in cashflows]
    eq_cfs   = [-equity]   + [cf["fcfe"]               for cf in cashflows]

    dscr_vals = [cf["dscr"] for cf in cashflows
                 if cf["year"] <= tenor and cf["debt_service"] > 0]

    return {
        "cashflows":   cashflows,
        "project_irr": irr(proj_cfs) * 100,
        "equity_irr":  irr(eq_cfs)   * 100,
        "min_dscr":    min(dscr_vals) if dscr_vals else 0,
        "avg_dscr":    sum(dscr_vals) / len(dscr_vals) if dscr_vals else 0,
        "ann_principal": ann_prin,
        "dep_slm":       dep_slm,
    }
