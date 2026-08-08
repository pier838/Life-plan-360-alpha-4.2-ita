
import json
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font


st.set_page_config(page_title="LifePlan 360", page_icon="🏠", layout="wide")


# ============================================================
# Helpers
# ============================================================

def money(x):
    try:
        return f"€{float(x):,.0f}"
    except Exception:
        return str(x)


def progressive_tax(taxable, rates):
    tax = 0.0
    lower = 0.0
    for upper, rate in rates:
        part = max(0.0, min(taxable, upper) - lower)
        tax += part * rate
        lower = upper
        if taxable <= upper:
            break
    return max(0.0, tax)


def net_salary(gross, employee_contrib, local_tax, irpef1, irpef2, irpef3):
    contrib = gross * employee_contrib
    taxable = max(0.0, gross - contrib)
    tax = progressive_tax(
        taxable,
        [(28000.0, irpef1), (50000.0, irpef2), (10**12, irpef3)],
    )
    return max(0.0, gross - contrib - tax - taxable * local_tax)


def net_pension(gross, local_tax, irpef1, irpef2, irpef3):
    tax = progressive_tax(
        gross,
        [(28000.0, irpef1), (50000.0, irpef2), (10**12, irpef3)],
    )
    return max(0.0, gross - tax - gross * local_tax)


def annuity_payment(principal, annual_rate, years):
    if principal <= 0 or years <= 0:
        return 0.0
    r = annual_rate / 12
    n = int(years * 12)
    if abs(r) < 1e-12:
        return principal / n
    return principal * r / (1 - (1 + r) ** (-n))


def mortgage_balance(principal, annual_rate, years, elapsed_years):
    if principal <= 0 or elapsed_years >= years:
        return 0.0
    r = annual_rate / 12
    n = int(years * 12)
    m = int(max(0, elapsed_years * 12))
    if abs(r) < 1e-12:
        return max(0.0, principal * (1 - m / n))
    pmt = principal * r / (1 - (1 + r) ** (-n))
    return max(0.0, principal * (1 + r) ** m - pmt * (((1 + r) ** m - 1) / r))


INPS_COEFF = {
    57: 0.0420, 58: 0.0431, 59: 0.0442, 60: 0.0454, 61: 0.0466,
    62: 0.0480, 63: 0.0494, 64: 0.0509, 65: 0.0525, 66: 0.0542,
    67: 0.0561, 68: 0.0581, 69: 0.0602, 70: 0.0626, 71: 0.0651
}


def inps_coeff(age):
    if age <= 57:
        return INPS_COEFF[57]
    if age >= 71:
        return INPS_COEFF[71]
    return INPS_COEFF[int(age)]


def nuda_reference(age):
    pts = [
        (60, 0.40), (62, 0.45), (65, 0.50), (68, 0.55), (70, 0.60),
        (73, 0.65), (76, 0.70), (80, 0.75), (84, 0.80), (88, 0.85),
    ]
    if age <= pts[0][0]:
        return pts[0][1]
    for (a1, p1), (a2, p2) in zip(pts[:-1], pts[1:]):
        if a1 <= age <= a2:
            w = (age - a1) / (a2 - a1)
            return p1 + (p2 - p1) * w
    return pts[-1][1]


def validate_allocation(name, alloc):
    eq, bd, cash = alloc
    if eq + bd > 1.000001:
        st.error(
            f"{name}: Azionario + Governativi supera il 100% "
            f"({(eq+bd)*100:.1f}%). Riduci una delle due percentuali."
        )
        return False
    total = eq + bd + cash
    if abs(total - 1.0) > 0.001:
        st.error(f"{name}: allocazione non valida ({total*100:.1f}%).")
        return False
    return True


def career_gross(params, age):
    if age >= params["retire_age"]:
        return 0.0, 0.0, 0.0

    start = params["age_start"]
    base = params["gross_salary"]
    growth = params["salary_growth"]
    mode = params["salary_mode"]

    if mode == "Crescita costante":
        ral = base * ((1 + growth) ** (age - start))

    elif mode == "Scatti di carriera":
        ral = base
        steps = sorted(
            [x for x in params["salary_steps"] if x["active"]],
            key=lambda x: x["age"]
        )
        for a in range(start + 1, age + 1):
            ral *= (1 + growth)
            for step in steps:
                if step["age"] == a:
                    ral = step["gross"]

    else:
        pts = sorted(
            [(x["age"], x["gross"]) for x in params["salary_custom"] if x["gross"] > 0],
            key=lambda x: x[0]
        )
        if not pts:
            ral = base * ((1 + growth) ** (age - start))
        elif age <= pts[0][0]:
            ral = pts[0][1] / ((1 + growth) ** max(0, pts[0][0] - age))
        else:
            ral = None
            for (a1, g1), (a2, g2) in zip(pts[:-1], pts[1:]):
                if a1 <= age <= a2:
                    w = (age - a1) / (a2 - a1)
                    ral = g1 + (g2 - g1) * w
                    break
            if ral is None:
                la, lg = pts[-1]
                ral = lg * ((1 + growth) ** max(0, age - la))

    bonus = params["annual_bonus"]
    if params["bonus_growth"]:
        bonus *= (1 + growth) ** max(0, age - start)
    return ral, bonus, ral + bonus


def spending_band(params, age):
    if age < params["band1_end"]:
        return "Senior"
    if age < params["band2_end"]:
        return "Old"
    return "Older"


def annual_living_cost(params, age):
    band = spending_band(params, age)
    inflation_factor = (1 + params["inflation"]) ** (age - params["age_start"])
    values = {
        k: v * 12 * inflation_factor
        for k, v in params["expense_bands"][band].items()
    }
    return band, values, sum(values.values())


def allocation_for_age(params, age):
    alloc = tuple(params["alloc_initial"])
    if params["derisk1"]["active"] and age >= params["derisk1"]["age"]:
        alloc = tuple(params["derisk1"]["alloc"])
    if params["derisk2"]["active"] and age >= params["derisk2"]["age"]:
        alloc = tuple(params["derisk2"]["alloc"])
    return alloc


def rebalance(equity, bonds, cash, target, commission):
    total = equity + bonds + cash
    if total <= 0:
        return equity, bonds, cash, 0.0
    curr = np.array([equity, bonds, cash], dtype=float)
    targ = total * np.array(target, dtype=float)
    turnover = np.abs(targ - curr).sum() / 2
    fee = turnover * commission
    post = max(0.0, total - fee)
    return post * target[0], post * target[1], post * target[2], fee


def active_black_swan_labels(params, age):
    labels = []
    for i, event in enumerate(params["swans"], start=1):
        if event["active"] and event["age"] <= age < event["age"] + max(1, event["duration"]):
            labels.append(f"Cigno nero {i}")
    return labels


def black_swan_apply(equity, bonds, age, event):
    if not event["active"]:
        return equity, bonds, ""

    start = event["age"]
    duration = max(1, event["duration"])
    decline_years = max(1, int(np.ceil(duration / 3)))
    end = start + duration

    if not (start <= age < end):
        return equity, bonds, ""

    if age < start + decline_years:
        annual_drop = 1 - (1 - event["drawdown"]) ** (1 / decline_years)
        equity *= (1 - annual_drop)
        bonds *= (1 + event["bond_effect"])
        return equity, bonds, f"Cigno nero: discesa -{annual_drop*100:.1f}% azionario"

    equity *= (1 + event["recovery_return"])
    bonds *= (1 + event["bond_effect"])
    return equity, bonds, f"Cigno nero: recupero +{event['recovery_return']*100:.1f}% azionario"


def unexpected_cost(params, age):
    total = 0.0
    labels = []
    for idx, event in enumerate(params["unexpected"], start=1):
        if event["active"] and age == event["age"]:
            amount = event["amount_today"] * (
                (1 + params["inflation"]) ** (age - params["age_start"])
            )
            total += amount
            labels.append(event["label"] or f"Imprevisto {idx}")
    return total, "; ".join(labels)


def late_care(params, age):
    if not params["late_care_active"] or age < params["late_care_age"]:
        return 0.0
    return params["late_care_monthly"] * 12 * (
        (1 + params["inflation"]) ** (age - params["age_start"])
    )


def fund_growth(params, fund, age):
    eq, bd, ca = allocation_for_age(params, age)
    eq_net = params["equity_return"] - params["equity_cost"]
    bd_net = params["bond_return"] - params["bond_cost"]
    ca_net = params["cash_return"]
    weighted = eq * eq_net + bd * bd_net + ca * ca_net
    return fund * (1 + weighted)


def lifetime_annuity(capital, age, params):
    if capital <= 0:
        return 0.0
    end_age = max(age + 1, params["annuity_life_expectancy"])
    years = end_age - age
    rate = params["annuity_return"]
    if abs(rate) < 1e-12:
        gross = capital / years
    else:
        gross = capital * rate / (1 - (1 + rate) ** (-years))
    return gross * (1 - params["fund_payout_tax"])


# ============================================================
# Simulation
# ============================================================

def simulate(params, scenario):
    age0 = params["age_start"]
    age_end = params["age_end"]

    eq0, bd0, ca0 = params["alloc_initial"]
    equity = params["initial_wealth"] * eq0
    bonds = params["initial_wealth"] * bd0
    cash = params["initial_wealth"] * ca0

    house = 0.0
    debt = 0.0
    mortgage_principal = 0.0
    mortgage_monthly = 0.0

    nuda_reserve = 0.0
    nuda_done = False
    downsize_done = False

    if scenario["housing"] == "owned":
        house = params["owned_home_value"]

    purchase_done = scenario["housing"] != "buy"
    purchase_age = age0 + int(params.get("purchase_delay_years", 0))
    purchase_shortfall_total = 0.0

    # For buy scenarios, purchase is executed inside the yearly loop.
    # With delay 0 it happens at age_start; with delay N there are N rental years first.

    inps_montante = params["inps_initial_montante"]
    contrib_years = params["contrib_years_initial"]

    fund = params["fund_initial"]
    fund_years = params["fund_years_initial"]

    rita_active = False
    rita_pool = 0.0
    fund_settled = False
    complementary_annuity = 0.0

    inps_locked = False
    inps_net_annual = 0.0

    rows = []

    for age in range(age0, age_end + 1):
        y = age - age0
        events = []

        band, expense_detail, normal_living = annual_living_cost(params, age)
        care_cost = late_care(params, age)
        living = normal_living + care_cost

        ral, bonus, gross = career_gross(params, age)
        work_net = 0.0
        if gross > 0:
            work_net = net_salary(
                gross,
                params["worker_contrib_rate"],
                params["add_rate"],
                params["irpef1"],
                params["irpef2"],
                params["irpef3"],
            )

        if age < params["retire_age"]:
            inps_montante *= (1 + params["inps_revaluation"])
            inps_montante += gross * params["inps_computo_rate"]
            contrib_years += 1

            tfr = gross / 13.5
            fund = fund_growth(params, fund, age) + tfr
            fund_years += 1
        else:
            if not fund_settled and not rita_active:
                fund = fund_growth(params, fund, age)

        if age >= params["inps_age"] and not inps_locked:
            gross_pension = inps_montante * inps_coeff(params["inps_age"])
            inps_net_annual = net_pension(
                gross_pension,
                params["add_rate"],
                params["irpef1"],
                params["irpef2"],
                params["irpef3"],
            )
            inps_locked = True
            events.append("Inizio pensione INPS")

        rita_net = 0.0
        rita_min_age = max(params["retire_age"], params["inps_age"] - 5)
        rita_ok = (
            params["use_rita"]
            and age >= rita_min_age
            and age < params["inps_age"]
            and contrib_years >= 20
            and fund_years >= 5
        )

        if rita_ok and (fund > 0 or rita_pool > 0):
            if not rita_active:
                rita_active = True
                rita_pool = fund * params["rita_share"]
                fund -= rita_pool
                events.append("Inizio RITA")
            years_left = max(1, params["inps_age"] - age)
            gross_rita = rita_pool / years_left
            rita_net = gross_rita * (1 - params["rita_tax"])
            rita_pool = max(0.0, rita_pool - gross_rita)

        capital_from_fund = 0.0
        if age >= params["inps_age"] and not fund_settled:
            total_fund = fund + rita_pool
            fund = 0.0
            rita_pool = 0.0

            mode = params["fund_at_inps_mode"]
            if mode == "Rendita vitalizia":
                complementary_annuity = lifetime_annuity(total_fund, age, params)
            elif mode == "Capitale + rendita":
                capital_from_fund = total_fund * params["fund_capital_share"]
                complementary_annuity = lifetime_annuity(
                    total_fund - capital_from_fund, age, params
                )
            else:
                capital_from_fund = total_fund
                complementary_annuity = 0.0

            fund_settled = True
            if total_fund > 0:
                events.append("Prestazione fondo pensione")

            if capital_from_fund > 0:
                teq, tbd, tca = allocation_for_age(params, age)
                equity += capital_from_fund * teq
                bonds += capital_from_fund * tbd
                cash += capital_from_fund * tca

        inps_income = inps_net_annual if age >= params["inps_age"] else 0.0
        comp_income = complementary_annuity if age >= params["inps_age"] else 0.0
        total_income = work_net + rita_net + inps_income + comp_income

        # Deferred purchase: before purchase the scenario behaves as rent.
        # The house price is expressed in today's/start-year euros and appreciates
        # until the actual purchase age. Buying costs are inflation-adjusted.
        purchase_shortfall = 0.0
        if (
            scenario["housing"] == "buy"
            and not purchase_done
            and age >= purchase_age
        ):
            delay = age - age0
            actual_price = params["house_price"] * ((1 + params["house_app"]) ** delay)
            actual_buying_costs = params["buying_costs"] * ((1 + params["inflation"]) ** delay)
            down = actual_price * scenario["down_pct"]
            upfront_need = down + actual_buying_costs

            # Fund upfront costs from available financial assets:
            # cash -> bonds -> equity. Never allow negative asset balances.
            use = min(max(0.0, cash), upfront_need)
            cash -= use
            upfront_need -= use

            use = min(max(0.0, bonds), upfront_need)
            bonds -= use
            upfront_need -= use

            use = min(max(0.0, equity), upfront_need)
            equity -= use
            upfront_need -= use

            purchase_shortfall = max(0.0, upfront_need)
            purchase_shortfall_total += purchase_shortfall

            if purchase_shortfall <= 1e-9:
                house = actual_price
                mortgage_principal = actual_price - down
                debt = mortgage_principal
                mortgage_monthly = annuity_payment(
                    mortgage_principal,
                    params["mortgage_rate"],
                    params["mortgage_years"],
                )
                purchase_done = True
                events.append(
                    f"Acquisto casa a {age} anni: prezzo {money(actual_price)}"
                )
            else:
                events.append(
                    f"INSOLVENZA ACQUISTO: anticipo/costi non finanziabili "
                    f"({money(purchase_shortfall)})"
                )
                # Prevent repeated attempts every following year.
                purchase_done = True

        # Appreciate house only after it exists; skip appreciation in the same
        # year it is purchased because actual_price already includes it.
        just_bought = (
            scenario["housing"] == "buy"
            and age == purchase_age
            and house > 0
        )
        if house > 0 and y > 0 and not just_bought:
            house *= (1 + params["house_app"])

        if scenario["housing"] == "rent":
            housing_cost = params["rent_monthly"] * 12 * (
                (1 + params["rent_growth"]) ** y
            )

        elif scenario["housing"] == "owned":
            housing_cost = (
                params["owned_home_ordinary_monthly"] * 12
                + params["owned_home_extra_annual"]
            ) * ((1 + params["inflation"]) ** y)
            debt = 0.0

        else:
            # Buy scenario: rent until the purchase is actually completed.
            if house <= 0:
                housing_cost = params["rent_monthly"] * 12 * (
                    (1 + params["rent_growth"]) ** y
                )
                debt = 0.0
            else:
                elapsed_since_purchase = max(0, age - purchase_age)
                debt = mortgage_balance(
                    mortgage_principal,
                    params["mortgage_rate"],
                    params["mortgage_years"],
                    elapsed_since_purchase,
                )
                mortgage_annual = (
                    mortgage_monthly * 12
                    if elapsed_since_purchase < params["mortgage_years"]
                    else 0.0
                )
                housing_cost = (
                    mortgage_annual
                    + params["owner_cost_monthly"] * 12
                    * ((1 + params["inflation"]) ** y)
                )

        if (
            scenario.get("lifecycle") == "nuda"
            and not nuda_done
            and age >= params["nuda_age"]
            and house > 0
        ):
            proceeds = house * params["nuda_pct"]
            nuda_reserve += proceeds
            house = 0.0
            debt = 0.0
            nuda_done = True
            events.append("Vendita nuda proprietà")

        if (
            scenario.get("lifecycle") == "downsize"
            and not downsize_done
            and age >= params["downsize_age"]
            and house > 0
        ):
            smaller = params["smaller_house_today"] * (
                (1 + params["house_app"]) ** y
            )
            released = max(0.0, house - smaller - params["downsize_costs"])
            teq, tbd, tca = allocation_for_age(params, age)
            equity += released * teq
            bonds += released * tbd
            cash += released * tca
            house = smaller
            debt = 0.0
            housing_cost = params["smaller_owner_cost_monthly"] * 12 * (
                (1 + params["inflation"]) ** y
            )
            downsize_done = True
            events.append("Downsizing")

        if params["derisk1"]["active"] and age == params["derisk1"]["age"]:
            equity, bonds, cash, fee = rebalance(
                equity,
                bonds,
                cash,
                params["derisk1"]["alloc"],
                params["derisk1"]["commission"],
            )
            events.append(f"Derisking 1: costo {money(fee)}")

        if params["derisk2"]["active"] and age == params["derisk2"]["age"]:
            equity, bonds, cash, fee = rebalance(
                equity,
                bonds,
                cash,
                params["derisk2"]["alloc"],
                params["derisk2"]["commission"],
            )
            events.append(f"Derisking 2: costo {money(fee)}")

        eq_gain = equity * (params["equity_return"] - params["equity_cost"])
        bd_gain = bonds * (params["bond_return"] - params["bond_cost"])
        cash_gain = cash * params["cash_return"]

        equity += eq_gain * (1 - params["equity_tax"])
        bonds += bd_gain * (1 - params["bond_tax"])
        cash += cash_gain * (1 - params["cash_tax"])

        for swan in params["swans"]:
            equity, bonds, note = black_swan_apply(equity, bonds, age, swan)
            if note:
                events.append(note)

        nuda_reserve *= (1 + params["nuda_return"])

        one_off, one_off_note = unexpected_cost(params, age)
        if one_off_note:
            events.append(one_off_note)

        outflows = living + housing_cost + one_off
        cashflow = total_income - outflows

        teq, tbd, tca = allocation_for_age(params, age)

        uncovered_need = 0.0
        sold_bonds = 0.0
        sold_equity = 0.0
        equity_sale_during_swan = False

        if cashflow >= 0:
            equity += cashflow * teq
            bonds += cashflow * tbd
            cash += cashflow * tca
        else:
            need = -cashflow

            use = min(max(0.0, nuda_reserve), need)
            nuda_reserve = max(0.0, nuda_reserve - use)
            need -= use

            use = min(max(0.0, cash), need)
            cash = max(0.0, cash - use)
            need -= use

            if need > 1e-9:
                use = min(max(0.0, bonds), need)
                bonds = max(0.0, bonds - use)
                need -= use
                sold_bonds += use
                if use > 0:
                    events.append(f"WARNING: disinvestimento governativi {money(use)}")

            if need > 1e-9:
                use = min(max(0.0, equity), need)
                equity = max(0.0, equity - use)
                need -= use
                sold_equity += use
                if use > 0:
                    swans_now = active_black_swan_labels(params, age)
                    if swans_now:
                        equity_sale_during_swan = True
                        events.append(
                            "WARNING GRAVE: vendita forzata di azionario durante "
                            + ", ".join(swans_now)
                            + f" ({money(use)}) — rischio di sequenza dei rendimenti"
                        )
                    else:
                        events.append(
                            f"WARNING: decumulo da azionario {money(use)}"
                        )

            if need > 1e-9:
                uncovered_need = need
                events.append(
                    f"INSOLVENZA: fabbisogno non coperto {money(uncovered_need)}"
                )

        # Never allow financial asset balances to go below zero.
        equity = max(0.0, equity)
        bonds = max(0.0, bonds)
        cash = max(0.0, cash)
        nuda_reserve = max(0.0, nuda_reserve)

        fund_total = fund + rita_pool
        monetary = equity + bonds + cash + nuda_reserve + fund_total
        total_wealth = monetary + house - debt

        rows.append({
            "Età": age,
            "Fascia spese": band,
            "RAL base": ral,
            "Bonus lordo": bonus,
            "RAL complessiva": gross,
            "Netto lavoro": work_net,
            "RITA netta": rita_net,
            "Pensione INPS netta": inps_income,
            "Pensione integrativa": comp_income,
            "Reddito totale": total_income,
            "Cibo": expense_detail["Cibo"],
            "Utenze": expense_detail["Utenze"],
            "Trasporti": expense_detail["Trasporti"],
            "Svago e viaggi": expense_detail["Svago e viaggi"],
            "Salute": expense_detail["Salute"],
            "Assistenza tarda età": care_cost,
            "Costo abitazione": housing_cost,
            "Imprevisti": one_off,
            "Fabbisogno non coperto acquisto casa": purchase_shortfall,
            "Flusso netto": cashflow,
            "Vendita governativi per decumulo": sold_bonds,
            "Vendita azionario per decumulo": sold_equity,
            "Vendita azionario durante cigno nero": equity_sale_during_swan,
            "Fabbisogno non coperto": uncovered_need,
            "Azionario": equity,
            "Obbligazionario": bonds,
            "Liquidità": cash,
            "Riserva nuda proprietà": nuda_reserve,
            "Fondo pensione residuo": fund_total,
            "Montante INPS": inps_montante,
            "Valore casa": house,
            "Debito residuo": debt,
            "Patrimonio monetario": monetary,
            "Patrimonio totale": total_wealth,
            "Eventi": "; ".join(events),
        })

    return pd.DataFrame(rows)


def summarize(df, params, label):
    neg_cf = df[df["Flusso netto"] < 0]
    min_money = float(df["Patrimonio monetario"].min())
    final_money = float(df.iloc[-1]["Patrimonio monetario"])
    final_house = float(df.iloc[-1]["Valore casa"])
    final_total = float(df.iloc[-1]["Patrimonio totale"])

    score = 0
    if (df["Patrimonio monetario"] < 0).any():
        score += 5
    if final_money < params["target_final"]:
        score += 3
    elif final_money < 2 * params["target_final"]:
        score += 2
    elif final_money < 4 * params["target_final"]:
        score += 1
    if min_money < params["target_floor"]:
        score += 2

    if score <= 1:
        risk = "Basso"
    elif score <= 3:
        risk = "Medio"
    elif score <= 5:
        risk = "Medio-alto"
    else:
        risk = "Alto"

    robustness = max(0, 100 - score * 12)

    bond_sell = df[df["Vendita governativi per decumulo"] > 0]
    eq_sell = df[df["Vendita azionario per decumulo"] > 0]
    eq_swan = df[df["Vendita azionario durante cigno nero"] == True]
    insolv = df[df["Fabbisogno non coperto"] > 0]
    purchase_insolv = df[df["Fabbisogno non coperto acquisto casa"] > 0]

    if not insolv.empty or not purchase_insolv.empty:
        risk = "Alto"
        robustness = min(robustness, 15)
    elif not eq_swan.empty and risk in ("Basso", "Medio"):
        risk = "Medio-alto"
        robustness = min(robustness, 55)

    return {
        "Scenario": label,
        "Patrimonio monetario finale": final_money,
        "Valore casa finale": final_house,
        "Patrimonio totale finale": final_total,
        "Patrimonio monetario minimo": min_money,
        "Inizio decumulo": "-" if neg_cf.empty else int(neg_cf.iloc[0]["Età"]),
        "Prima vendita governativi": "-" if bond_sell.empty else int(bond_sell.iloc[0]["Età"]),
        "Prima vendita azionario": "-" if eq_sell.empty else int(eq_sell.iloc[0]["Età"]),
        "Anni con vendita azionario": int((df["Vendita azionario per decumulo"] > 0).sum()),
        "Vendita azionario in cigno nero": "Sì" if not eq_swan.empty else "No",
        "Età esaurimento patrimonio": "-" if insolv.empty else int(insolv.iloc[0]["Età"]),
        "Fabbisogno non coperto cumulato": float(df["Fabbisogno non coperto"].sum()),
        "Fabbisogno acquisto non coperto": float(df["Fabbisogno non coperto acquisto casa"].sum()),
        "Pensione INPS netta annua": float(df.iloc[-1]["Pensione INPS netta"]),
        "Pensione integrativa annua": float(df.iloc[-1]["Pensione integrativa"]),
        "Rischio": risk,
        "Robustezza /100": robustness,
        "Sostenibile": (
            "Sì"
            if insolv.empty
            and purchase_insolv.empty
            and min_money >= 0
            and final_money >= params["target_final"]
            else "No"
        ),
    }


def build_excel(params, results, summary):
    bio = BytesIO()
    wb = Workbook()

    ws = wb.active
    ws.title = "Riepilogo"
    ws["A1"] = "LifePlan 360 — Input e output"
    ws["A1"].font = Font(bold=True, size=16)

    r = 3
    ws.cell(r, 1, "INPUT")
    ws.cell(r, 1).font = Font(bold=True)
    r += 1

    for key, value in params.items():
        ws.cell(r, 1, key)
        if isinstance(value, (dict, list, tuple)):
            value = json.dumps(value, ensure_ascii=False)
        ws.cell(r, 2, value)
        r += 1

    r += 1
    ws.cell(r, 1, "OUTPUT SCENARI")
    ws.cell(r, 1).font = Font(bold=True)
    r += 1

    for c, col in enumerate(summary.columns, start=1):
        ws.cell(r, c, col).font = Font(bold=True)

    for _, rec in summary.iterrows():
        r += 1
        for c, col in enumerate(summary.columns, start=1):
            val = rec[col]
            if hasattr(val, "item"):
                try:
                    val = val.item()
                except Exception:
                    pass
            ws.cell(r, c, val)

    for label, df in results.items():
        sh = wb.create_sheet(label[:31])
        for c, col in enumerate(df.columns, start=1):
            sh.cell(1, c, col).font = Font(bold=True)
        for rr, row in enumerate(df.itertuples(index=False), start=2):
            for c, val in enumerate(row, start=1):
                sh.cell(rr, c, val)

    ch = wb.create_sheet("Grafici")
    ch["A1"] = "Grafici comparativi"
    ch["A1"].font = Font(bold=True, size=16)

    ages = results[next(iter(results))]["Età"].tolist()
    configs = [
        ("Patrimonio monetario", 3, "Patrimonio monetario"),
        ("Patrimonio totale", 60, "Patrimonio totale"),
        ("Flusso netto", 117, "Flusso netto annuale"),
    ]

    for metric, start, title in configs:
        ch.cell(start, 1, "Età")
        for j, label in enumerate(results.keys(), start=2):
            ch.cell(start, j, label)

        for i, age in enumerate(ages, start=start + 1):
            ch.cell(i, 1, age)
            for j, (label, df) in enumerate(results.items(), start=2):
                idx = i - start - 1
                ch.cell(i, j, float(df.iloc[idx][metric]))

        chart = LineChart()
        chart.title = title
        chart.x_axis.title = "Età"
        chart.y_axis.title = "Euro"
        data = Reference(
            ch,
            min_col=2,
            max_col=1 + len(results),
            min_row=start,
            max_row=start + len(ages),
        )
        cats = Reference(
            ch,
            min_col=1,
            min_row=start + 1,
            max_row=start + len(ages),
        )
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.width = 22
        chart.height = 10
        ch.add_chart(chart, f"A{start+50}")

    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


# ============================================================
# Interface
# ============================================================

st.markdown("# 🏠 💰 🚗 ✈️ LifePlan 360")
st.caption(
    "Simulatore della vita finanziaria: reddito, casa, investimenti, pensione, "
    "spese, imprevisti e resilienza."
)

uploaded = st.file_uploader("Importa configurazione JSON", type=["json"])
U = {}
if uploaded is not None:
    try:
        U = json.load(uploaded)
        st.success("Configurazione importata.")
    except Exception as exc:
        st.error(f"Configurazione JSON non valida: {exc}")


def gv(key, default):
    return U.get(key, default)


with st.sidebar:
    st.header("1 · Anagrafica")
    age_start = st.number_input(
        "Età iniziale simulazione", 18, 85, int(gv("age_start", 43))
    )
    age_end = st.number_input(
        "Età finale simulazione",
        max(50, int(age_start) + 1),
        110,
        int(gv("age_end", 90)),
    )
    retire_age = st.number_input(
        "Età uscita dal lavoro",
        int(age_start),
        90,
        int(gv("retire_age", 59)),
    )
    inps_age = st.number_input(
        "Età pensione INPS", 57, 75, int(gv("inps_age", 67))
    )
    target_final = st.number_input(
        "Patrimonio monetario minimo finale desiderato (€)",
        0.0,
        value=float(gv("target_final", 70000.0)),
        step=5000.0,
    )
    target_floor = st.number_input(
        "Patrimonio monetario minimo di sicurezza (€)",
        0.0,
        value=float(gv("target_floor", 30000.0)),
        step=5000.0,
    )

    st.header("2 · Reddito e carriera")
    gross_salary = st.number_input(
        "RAL iniziale (€)",
        0.0,
        value=float(gv("gross_salary", 68000.0)),
        step=1000.0,
    )
    salary_growth = st.number_input(
        "Crescita RAL annua (%)",
        0.0,
        10.0,
        float(gv("salary_growth", 0.02)) * 100,
        0.1,
        help="Guida: 1% prudente · 2% centrale · 3% dinamica",
    ) / 100

    salary_modes = ["Crescita costante", "Scatti di carriera", "Tabella personalizzata"]
    salary_mode = st.selectbox(
        "Modalità carriera",
        salary_modes,
        index=salary_modes.index(gv("salary_mode", "Crescita costante")),
    )

    salary_steps = []
    saved_steps = gv("salary_steps", [])
    if salary_mode == "Scatti di carriera":
        for i in range(3):
            saved = saved_steps[i] if isinstance(saved_steps, list) and len(saved_steps) > i else {}
            with st.expander(f"Scatto {i+1}"):
                active = st.checkbox(
                    "Attivo", bool(saved.get("active", False)), key=f"step_on_{i}"
                )
                step_age = st.number_input(
                    "Età",
                    int(age_start),
                    int(age_end),
                    int(saved.get("age", min(int(age_end), int(age_start) + 5 * (i + 1)))),
                    key=f"step_age_{i}",
                )
                step_gross = st.number_input(
                    "Nuova RAL (€)",
                    0.0,
                    value=float(saved.get("gross", 75000 + 10000 * i)),
                    step=1000.0,
                    key=f"step_gross_{i}",
                )
                salary_steps.append(
                    {"active": active, "age": int(step_age), "gross": step_gross}
                )

    salary_custom = []
    if salary_mode == "Tabella personalizzata":
        saved_custom = gv("salary_custom", [])
        rows = []
        for i in range(6):
            saved = (
                saved_custom[i]
                if isinstance(saved_custom, list) and len(saved_custom) > i
                else {}
            )
            rows.append({
                "Età": int(saved.get("age", min(int(age_end), int(age_start) + 5 * i))),
                "RAL": (
                    float(saved.get("gross", gross_salary * ((1 + salary_growth) ** (5 * i))))
                    if i < 4
                    else 0.0
                ),
            })
        edited = st.data_editor(
            pd.DataFrame(rows),
            num_rows="fixed",
            use_container_width=True,
            key="career_table",
        )
        for _, row in edited.iterrows():
            if float(row["RAL"]) > 0:
                salary_custom.append(
                    {"age": int(row["Età"]), "gross": float(row["RAL"])}
                )

    annual_bonus = st.number_input(
        "Bonus annuo lordo iniziale (€)",
        0.0,
        value=float(gv("annual_bonus", 0.0)),
        step=1000.0,
    )
    bonus_growth = st.checkbox(
        "Bonus cresce con la RAL", bool(gv("bonus_growth", True))
    )

    st.header("3 · Fiscalità lavoro")
    worker_contrib_rate = st.number_input(
        "Contributi lavoratore in busta (%)",
        0.0,
        20.0,
        float(gv("worker_contrib_rate", 0.0919)) * 100,
        0.1,
        help="Serve alla stima del netto; è distinto dall'aliquota di computo INPS.",
    ) / 100
    add_rate = st.number_input(
        "Addizionali regionali/comunali (%)",
        0.0,
        5.0,
        float(gv("add_rate", 0.02)) * 100,
        0.1,
    ) / 100
    irpef1 = st.number_input(
        "IRPEF primo scaglione (%)", 0.0, 60.0, float(gv("irpef1", 0.23)) * 100, 0.5
    ) / 100
    irpef2 = st.number_input(
        "IRPEF secondo scaglione (%)", 0.0, 60.0, float(gv("irpef2", 0.35)) * 100, 0.5
    ) / 100
    irpef3 = st.number_input(
        "IRPEF terzo scaglione (%)", 0.0, 60.0, float(gv("irpef3", 0.43)) * 100, 0.5
    ) / 100

    st.header("4 · Previdenza INPS")
    inps_initial_montante = st.number_input(
        "Montante INPS iniziale (€)",
        0.0,
        value=float(gv("inps_initial_montante", 0.0)),
        step=5000.0,
    )
    contrib_years_initial = st.number_input(
        "Anni contributivi già maturati",
        0.0,
        60.0,
        float(gv("contrib_years_initial", 15.0)),
        0.5,
    )
    inps_computo_rate = st.number_input(
        "Aliquota di computo INPS (%)",
        0.0,
        50.0,
        float(gv("inps_computo_rate", 0.33)) * 100,
        0.5,
        help="Default 33% per dipendente; non coincide con la sola trattenuta in busta.",
    ) / 100
    inps_revaluation = st.number_input(
        "Rivalutazione montante INPS annua (%)",
        0.0,
        10.0,
        float(gv("inps_revaluation", 0.015)) * 100,
        0.1,
    ) / 100

    st.header("5 · Fondo pensione / TFR")
    fund_initial = st.number_input(
        "Fondo pensione iniziale (€)",
        0.0,
        value=float(gv("fund_initial", 0.0)),
        step=5000.0,
    )
    fund_years_initial = st.number_input(
        "Anni già nel fondo pensione",
        0.0,
        60.0,
        float(gv("fund_years_initial", 0.0)),
        0.5,
    )
    use_rita = st.checkbox(
        "Utilizzare la RITA quando possibile",
        bool(gv("use_rita", True)),
    )
    rita_share = st.slider(
        "% fondo destinato alla RITA",
        0,
        100,
        int(gv("rita_share", 1.0) * 100),
    ) / 100
    rita_tax = st.number_input(
        "Tassazione RITA stimata (%)",
        0.0,
        30.0,
        float(gv("rita_tax", 0.15)) * 100,
        0.5,
    ) / 100

    fund_modes = [
        "Rendita vitalizia",
        "Capitale + rendita",
        "100% capitale (se consentito)",
    ]
    fund_at_inps_mode = st.selectbox(
        "Utilizzo fondo residuo alla pensione INPS",
        fund_modes,
        index=fund_modes.index(gv("fund_at_inps_mode", "Rendita vitalizia")),
    )
    fund_capital_share = st.slider(
        "% capitale nella modalità mista",
        0,
        100,
        int(gv("fund_capital_share", 0.50) * 100),
    ) / 100
    annuity_life_expectancy = st.number_input(
        "Età attesa per calcolo rendita vitalizia",
        70,
        110,
        int(gv("annuity_life_expectancy", 90)),
    )
    annuity_return = st.number_input(
        "Rendimento tecnico rendita (%)",
        0.0,
        10.0,
        float(gv("annuity_return", 0.015)) * 100,
        0.1,
    ) / 100
    fund_payout_tax = st.number_input(
        "Tassazione media rendita fondo (%)",
        0.0,
        30.0,
        float(gv("fund_payout_tax", 0.15)) * 100,
        0.5,
    ) / 100

    st.header("6 · Investimenti")
    initial_wealth = st.number_input(
        "Patrimonio monetario iniziale (€)",
        0.0,
        value=float(gv("initial_wealth", 260000.0)),
        step=5000.0,
    )

    saved_alloc = gv("alloc_initial", [0.60, 0.30, 0.10])
    ai_eq = st.number_input(
        "Azionario iniziale (%)", 0.0, 100.0,
        float(saved_alloc[0] * 100), 1.0
    ) / 100
    ai_bd = st.number_input(
        "Governativi iniziali (%)", 0.0, 100.0,
        float(saved_alloc[1] * 100), 1.0
    ) / 100
    ai_ca = max(0.0, 1.0 - ai_eq - ai_bd)
    st.caption(f"Liquidità automatica: {ai_ca*100:.1f}%")

    equity_return = st.number_input(
        "Rendimento ETF azionario lordo (%)",
        -20.0,
        30.0,
        float(gv("equity_return", 0.065)) * 100,
        0.1,
        help="Guida: 5% prudente · 6,5% centrale · 8% favorevole",
    ) / 100
    equity_tax = st.number_input(
        "Tassazione rendimenti azionari (%)",
        0.0,
        40.0,
        float(gv("equity_tax", 0.26)) * 100,
        0.5,
    ) / 100
    equity_cost = st.number_input(
        "TER/costo annuo ETF azionario (%)",
        0.0,
        5.0,
        float(gv("equity_cost", 0.002)) * 100,
        0.05,
    ) / 100

    bond_return = st.number_input(
        "Rendimento governativi lordo (%)",
        -10.0,
        20.0,
        float(gv("bond_return", 0.03)) * 100,
        0.1,
        help="Guida: 2% prudente · 3% centrale · 4% favorevole",
    ) / 100
    bond_tax = st.number_input(
        "Tassazione governativi (%)",
        0.0,
        30.0,
        float(gv("bond_tax", 0.125)) * 100,
        0.5,
    ) / 100
    bond_cost = st.number_input(
        "Costo annuo obbligazionario (%)",
        0.0,
        5.0,
        float(gv("bond_cost", 0.001)) * 100,
        0.05,
    ) / 100

    cash_return = st.number_input(
        "Rendimento liquidità (%)",
        -5.0,
        10.0,
        float(gv("cash_return", 0.005)) * 100,
        0.1,
    ) / 100
    cash_tax = st.number_input(
        "Tassazione rendimento liquidità (%)",
        0.0,
        40.0,
        float(gv("cash_tax", 0.26)) * 100,
        0.5,
    ) / 100

    d1 = gv("derisk1", {})
    st.subheader("Derisking 1")
    d1_active = st.checkbox("Attivo derisking 1", bool(d1.get("active", True)))
    d1_age = st.number_input(
        "Età derisking 1",
        int(age_start),
        int(age_end),
        int(d1.get("age", 62)),
    )
    d1_saved = d1.get("alloc", [0.50, 0.40, 0.10])
    d1_eq = st.number_input(
        "Azionario dopo derisking 1 (%)", 0.0, 100.0,
        float(d1_saved[0] * 100), 1.0
    ) / 100
    d1_bd = st.number_input(
        "Governativi dopo derisking 1 (%)", 0.0, 100.0,
        float(d1_saved[1] * 100), 1.0
    ) / 100
    d1_ca = max(0.0, 1.0 - d1_eq - d1_bd)
    st.caption(f"Liquidità automatica dopo derisking 1: {d1_ca*100:.1f}%")
    d1_comm = st.number_input(
        "Commissione riallocazione 1 (%)",
        0.0,
        5.0,
        float(d1.get("commission", 0.002)) * 100,
        0.05,
    ) / 100

    d2 = gv("derisk2", {})
    st.subheader("Derisking 2")
    d2_active = st.checkbox("Attivo derisking 2", bool(d2.get("active", True)))
    d2_age = st.number_input(
        "Età derisking 2",
        int(age_start),
        int(age_end),
        int(d2.get("age", 70)),
    )
    d2_saved = d2.get("alloc", [0.30, 0.50, 0.20])
    d2_eq = st.number_input(
        "Azionario dopo derisking 2 (%)", 0.0, 100.0,
        float(d2_saved[0] * 100), 1.0
    ) / 100
    d2_bd = st.number_input(
        "Governativi dopo derisking 2 (%)", 0.0, 100.0,
        float(d2_saved[1] * 100), 1.0
    ) / 100
    d2_ca = max(0.0, 1.0 - d2_eq - d2_bd)
    st.caption(f"Liquidità automatica dopo derisking 2: {d2_ca*100:.1f}%")
    d2_comm = st.number_input(
        "Commissione riallocazione 2 (%)",
        0.0,
        5.0,
        float(d2.get("commission", 0.002)) * 100,
        0.05,
    ) / 100

    st.header("7 · Inflazione e spese")
    inflation = st.number_input(
        "Inflazione annua (%)",
        0.0,
        15.0,
        float(gv("inflation", 0.02)) * 100,
        0.1,
        help="Guida: 1,5% bassa · 2% centrale · 3% stress · 4% severa",
    ) / 100

    default_band1 = min(max(int(age_start) + 1, 62), int(age_end) - 1)
    band1_end = st.number_input(
        "Fine fascia Senior",
        int(age_start) + 1,
        int(age_end) - 1,
        int(gv("band1_end", default_band1)),
    )

    default_band2 = min(max(int(band1_end) + 1, 70), int(age_end))
    band2_end = st.number_input(
        "Fine fascia Old",
        int(band1_end) + 1,
        int(age_end),
        int(gv("band2_end", default_band2)),
    )

    st.caption(
        f"Senior: {age_start}–{band1_end-1} · "
        f"Old: {band1_end}–{band2_end-1} · "
        f"Older: {band2_end}–{age_end}"
    )

    guide_defaults = {
        "Senior": {
            "Cibo": 450.0,
            "Utenze": 220.0,
            "Trasporti": 350.0,
            "Svago e viaggi": 500.0,
            "Salute": 120.0,
        },
        "Old": {
            "Cibo": 450.0,
            "Utenze": 230.0,
            "Trasporti": 300.0,
            "Svago e viaggi": 400.0,
            "Salute": 220.0,
        },
        "Older": {
            "Cibo": 450.0,
            "Utenze": 240.0,
            "Trasporti": 200.0,
            "Svago e viaggi": 200.0,
            "Salute": 400.0,
        },
    }

    saved_bands = gv("expense_bands", {})
    expense_bands = {}

    for band in ["Senior", "Old", "Older"]:
        vals = {}
        with st.expander(f"Spese {band}"):
            for cat in ["Cibo", "Utenze", "Trasporti", "Svago e viaggi", "Salute"]:
                default = float(
                    saved_bands.get(band, {}).get(cat, guide_defaults[band][cat])
                )
                vals[cat] = st.number_input(
                    f"{cat} €/mese",
                    0.0,
                    value=default,
                    step=25.0,
                    key=f"{band}_{cat}",
                )
        expense_bands[band] = vals

    st.subheader("Assistenza non autosufficienza")
    late_care_active = st.checkbox(
        "Considera assistenza in tarda età",
        bool(gv("late_care_active", False)),
    )
    late_care_age = st.number_input(
        "Età inizio assistenza",
        int(age_start),
        int(age_end),
        int(gv("late_care_age", min(85, int(age_end)))),
    )
    late_care_monthly = st.number_input(
        "Costo mensile in euro di oggi (€)",
        0.0,
        value=float(gv("late_care_monthly", 2300.0)),
        step=100.0,
        help="Voce guida per assistenza domiciliare/RSA; completamente modificabile.",
    )

    st.header("8 · Casa e affitto")
    owns_home_initially = st.checkbox(
        "Possiedo già la casa in cui vivo",
        bool(gv("owns_home_initially", False)),
    )
    owned_home_value = st.number_input(
        "Valore casa già posseduta (€)",
        0.0,
        value=float(gv("owned_home_value", 260000.0)),
        step=5000.0,
        help="Usato solo nello scenario 'Casa già posseduta'.",
    )
    owned_home_ordinary_monthly = st.number_input(
        "Manutenzione ordinaria casa posseduta €/mese",
        0.0,
        value=float(gv("owned_home_ordinary_monthly", 200.0)),
        step=25.0,
    )
    owned_home_extra_annual = st.number_input(
        "Manutenzione straordinaria media €/anno",
        0.0,
        value=float(gv("owned_home_extra_annual", 1800.0)),
        step=250.0,
    )

    rent_monthly = st.number_input(
        "Affitto mensile iniziale (€)",
        0.0,
        value=float(gv("rent_monthly", 1300.0)),
        step=50.0,
    )
    rent_growth = st.number_input(
        "Crescita affitto annua (%)",
        0.0,
        15.0,
        float(gv("rent_growth", 0.03)) * 100,
        0.1,
    ) / 100

    house_price = st.number_input(
        "Valore/prezzo casa da acquistare nell'anno iniziale (€)",
        0.0,
        value=float(gv("house_price", 260000.0)),
        step=5000.0,
        help=(
            "È il prezzo della casa espresso all'anno iniziale della simulazione. "
            "Se l'acquisto viene posticipato, il simulatore rivaluta questo prezzo "
            "fino all'anno effettivo di acquisto."
        ),
    )
    purchase_delay_years = st.number_input(
        "Ritardo acquisto casa (anni)",
        0,
        max(0, int(age_end) - int(age_start)),
        int(gv("purchase_delay_years", 0)),
        help=(
            "0 = acquisto nell'anno iniziale. "
            "Se > 0, prima dell'acquisto si vive in affitto per il numero di anni indicato."
        ),
    )
    if purchase_delay_years == 0:
        st.caption("Scenario acquisto puro: la casa viene comprata subito.")
    else:
        st.caption(
            f"Scenario misto: affitto per {purchase_delay_years} anni, "
            f"poi acquisto a {int(age_start) + int(purchase_delay_years)} anni."
        )
    house_app = st.number_input(
        "Rivalutazione casa annua (%)",
        -10.0,
        15.0,
        float(gv("house_app", 0.015)) * 100,
        0.1,
        help="Guida: 0% molto prudente · 0,5% prudente · 1,5% centrale · 2–3% favorevole",
    ) / 100
    mortgage_rate = st.number_input(
        "Tasso mutuo (%)",
        0.0,
        20.0,
        float(gv("mortgage_rate", 0.033)) * 100,
        0.1,
    ) / 100
    mortgage_years = st.number_input(
        "Durata mutuo (anni)", 1, 40, int(gv("mortgage_years", 20))
    )
    buying_costs = st.number_input(
        "Costi acquisto escluso anticipo (€)",
        0.0,
        value=float(gv("buying_costs", 32000.0)),
        step=1000.0,
    )
    owner_cost_monthly = st.number_input(
        "Manutenzione/costi proprietà €/mese",
        0.0,
        value=float(gv("owner_cost_monthly", 350.0)),
        step=25.0,
    )

    st.subheader("Nuda proprietà")
    nuda_age = st.number_input(
        "Età vendita nuda proprietà",
        50,
        100,
        int(gv("nuda_age", 70)),
    )
    nuda_guide = nuda_reference(nuda_age)
    st.caption(
        f"Guida indicativa alla percentuale sul valore pieno: circa {nuda_guide*100:.0f}%."
    )
    nuda_pct = st.number_input(
        "% valore casa incassato",
        0.0,
        100.0,
        float(gv("nuda_pct", nuda_guide)) * 100,
        1.0,
    ) / 100
    nuda_return = st.number_input(
        "Rendimento capitale nuda proprietà (%)",
        -10.0,
        20.0,
        float(gv("nuda_return", 0.015)) * 100,
        0.1,
    ) / 100

    st.subheader("Downsizing")
    downsize_age = st.number_input(
        "Età downsizing", 50, 100, int(gv("downsize_age", 70))
    )
    smaller_house_today = st.number_input(
        "Valore oggi casa più piccola (€)",
        0.0,
        value=float(gv("smaller_house_today", 180000.0)),
        step=5000.0,
    )
    downsize_costs = st.number_input(
        "Costi compravendita downsizing (€)",
        0.0,
        value=float(gv("downsize_costs", 25000.0)),
        step=1000.0,
    )
    smaller_owner_cost_monthly = st.number_input(
        "Costi nuova casa €/mese",
        0.0,
        value=float(gv("smaller_owner_cost_monthly", 250.0)),
        step=25.0,
    )

    st.header("9 · Tre cigni neri")
    swans = []
    saved_swans = gv("swans", [])
    for i in range(3):
        saved = (
            saved_swans[i]
            if isinstance(saved_swans, list) and len(saved_swans) > i
            else {}
        )
        with st.expander(f"Cigno nero {i+1}"):
            active = st.checkbox(
                "Attivo",
                bool(saved.get("active", False)),
                key=f"swan_active_{i}",
            )
            event_age = st.number_input(
                "Età inizio",
                int(age_start),
                int(age_end),
                int(saved.get("age", min(int(age_end), 55 + 10 * i))),
                key=f"swan_age_{i}",
            )
            duration = st.number_input(
                "Durata totale (anni)",
                1,
                15,
                int(saved.get("duration", 3)),
                key=f"swan_duration_{i}",
                help="Circa 1/3 discesa e 2/3 recupero.",
            )
            drawdown = st.number_input(
                "Perdita massima azionario (%)",
                0.0,
                90.0,
                float(saved.get("drawdown", 0.35)) * 100,
                1.0,
                key=f"swan_drawdown_{i}",
            ) / 100
            recovery = st.number_input(
                "Rendimento annuo durante recupero (%)",
                -20.0,
                50.0,
                float(saved.get("recovery_return", 0.10)) * 100,
                0.5,
                key=f"swan_recovery_{i}",
            ) / 100
            bond_effect = st.number_input(
                "Effetto annuo sui governativi (%)",
                -30.0,
                30.0,
                float(saved.get("bond_effect", 0.0)) * 100,
                0.5,
                key=f"swan_bond_{i}",
            ) / 100
            swans.append({
                "active": active,
                "age": int(event_age),
                "duration": int(duration),
                "drawdown": drawdown,
                "recovery_return": recovery,
                "bond_effect": bond_effect,
            })

    st.header("10 · Imprevisti economici")
    unexpected = []
    saved_unexpected = gv("unexpected", [])
    for i in range(2):
        saved = (
            saved_unexpected[i]
            if isinstance(saved_unexpected, list) and len(saved_unexpected) > i
            else {}
        )
        with st.expander(f"Imprevisto {i+1}"):
            active = st.checkbox(
                "Attivo",
                bool(saved.get("active", False)),
                key=f"un_active_{i}",
            )
            event_age = st.number_input(
                "Età evento",
                int(age_start),
                int(age_end),
                int(saved.get("age", min(int(age_end), 65 + 10 * i))),
                key=f"un_age_{i}",
            )
            amount = st.number_input(
                "Perdita in euro di oggi (€)",
                0.0,
                value=float(saved.get("amount_today", 50000.0)),
                step=5000.0,
                key=f"un_amount_{i}",
            )
            label = st.text_input(
                "Descrizione",
                str(saved.get("label", "")),
                key=f"un_label_{i}",
            )
            unexpected.append({
                "active": active,
                "age": int(event_age),
                "amount_today": amount,
                "label": label,
            })


params = {
    "age_start": int(age_start),
    "age_end": int(age_end),
    "retire_age": int(retire_age),
    "inps_age": int(inps_age),
    "target_final": target_final,
    "target_floor": target_floor,
    "gross_salary": gross_salary,
    "salary_growth": salary_growth,
    "salary_mode": salary_mode,
    "salary_steps": salary_steps,
    "salary_custom": salary_custom,
    "annual_bonus": annual_bonus,
    "bonus_growth": bonus_growth,
    "worker_contrib_rate": worker_contrib_rate,
    "add_rate": add_rate,
    "irpef1": irpef1,
    "irpef2": irpef2,
    "irpef3": irpef3,
    "inps_initial_montante": inps_initial_montante,
    "contrib_years_initial": contrib_years_initial,
    "inps_computo_rate": inps_computo_rate,
    "inps_revaluation": inps_revaluation,
    "fund_initial": fund_initial,
    "fund_years_initial": fund_years_initial,
    "use_rita": use_rita,
    "rita_share": rita_share,
    "rita_tax": rita_tax,
    "fund_at_inps_mode": fund_at_inps_mode,
    "fund_capital_share": fund_capital_share,
    "annuity_life_expectancy": int(annuity_life_expectancy),
    "annuity_return": annuity_return,
    "fund_payout_tax": fund_payout_tax,
    "initial_wealth": initial_wealth,
    "alloc_initial": [ai_eq, ai_bd, ai_ca],
    "equity_return": equity_return,
    "equity_tax": equity_tax,
    "equity_cost": equity_cost,
    "bond_return": bond_return,
    "bond_tax": bond_tax,
    "bond_cost": bond_cost,
    "cash_return": cash_return,
    "cash_tax": cash_tax,
    "derisk1": {
        "active": d1_active,
        "age": int(d1_age),
        "alloc": [d1_eq, d1_bd, d1_ca],
        "commission": d1_comm,
    },
    "derisk2": {
        "active": d2_active,
        "age": int(d2_age),
        "alloc": [d2_eq, d2_bd, d2_ca],
        "commission": d2_comm,
    },
    "inflation": inflation,
    "band1_end": int(band1_end),
    "band2_end": int(band2_end),
    "expense_bands": expense_bands,
    "late_care_active": late_care_active,
    "late_care_age": int(late_care_age),
    "late_care_monthly": late_care_monthly,
    "owns_home_initially": owns_home_initially,
    "owned_home_value": owned_home_value,
    "owned_home_ordinary_monthly": owned_home_ordinary_monthly,
    "owned_home_extra_annual": owned_home_extra_annual,
    "rent_monthly": rent_monthly,
    "rent_growth": rent_growth,
    "house_price": house_price,
    "purchase_delay_years": int(purchase_delay_years),
    "house_app": house_app,
    "mortgage_rate": mortgage_rate,
    "mortgage_years": int(mortgage_years),
    "buying_costs": buying_costs,
    "owner_cost_monthly": owner_cost_monthly,
    "nuda_age": int(nuda_age),
    "nuda_pct": nuda_pct,
    "nuda_return": nuda_return,
    "downsize_age": int(downsize_age),
    "smaller_house_today": smaller_house_today,
    "downsize_costs": downsize_costs,
    "smaller_owner_cost_monthly": smaller_owner_cost_monthly,
    "swans": swans,
    "unexpected": unexpected,
}


valid = True
valid &= validate_allocation("Allocazione iniziale", params["alloc_initial"])
if params["derisk1"]["active"]:
    valid &= validate_allocation("Derisking 1", params["derisk1"]["alloc"])
if params["derisk2"]["active"]:
    valid &= validate_allocation("Derisking 2", params["derisk2"]["alloc"])

if params["derisk1"]["active"] and params["derisk2"]["active"]:
    if params["derisk2"]["age"] <= params["derisk1"]["age"]:
        st.warning("Il derisking 2 dovrebbe normalmente avvenire dopo il derisking 1.")


st.subheader("Assunzioni principali e affidabilità")
assumptions = pd.DataFrame([
    ["RAL iniziale", money(gross_salary), "Alta", "Dato inserito dall'utente"],
    ["Crescita RAL", f"{salary_growth*100:.1f}%", "Media", "Carriera futura incerta"],
    ["Stima netto", "Modello parametrico", "Media", "Non include ogni detrazione/deduzione personale"],
    ["Inflazione", f"{inflation*100:.1f}%", "Bassa-Media", "Ipotesi di lungo periodo"],
    ["Rendimento azionario", f"{equity_return*100:.1f}%", "Bassa-Media", "Rendimento atteso"],
    ["Rendimento governativi", f"{bond_return*100:.1f}%", "Media", "Dipende dai tassi futuri"],
    ["Rivalutazione casa", f"{house_app*100:.1f}%", "Bassa-Media", "Dipende da zona e immobile"],
    ["Pensione INPS", "Montante × coefficiente", "Media", "Stima di pianificazione"],
], columns=["Parametro", "Valore", "Affidabilità", "Nota"])
st.dataframe(assumptions, use_container_width=True, hide_index=True)


st.markdown("---")
st.caption("La simulazione non viene ricalcolata mentre modifichi gli input: parte solo quando premi il pulsante.")
run_simulation = st.button(
    "▶️ Esegui simulazione",
    type="primary",
    use_container_width=True,
    disabled=not valid,
)

if run_simulation:
    st.session_state["lifeplan_run_params"] = params

if "lifeplan_run_params" in st.session_state:
    if not run_simulation:
        st.info(
            "Hai modificato degli input? Premi **Esegui simulazione** per aggiornare i risultati. "
            "Qui sotto restano visibili i risultati dell'ultima esecuzione."
        )
    params = st.session_state["lifeplan_run_params"]

    purchase_suffix = (
        ""
        if params["purchase_delay_years"] == 0
        else f" (dopo {params['purchase_delay_years']} anni in affitto)"
    )

    scenarios = [
        {"label": "Affitto", "housing": "rent", "down_pct": 0.0, "lifecycle": "keep"},
        {"label": "Mutuo 20%" + purchase_suffix, "housing": "buy", "down_pct": 0.20, "lifecycle": "keep"},
        {"label": "Mutuo 25%" + purchase_suffix, "housing": "buy", "down_pct": 0.25, "lifecycle": "keep"},
        {"label": "Mutuo 30%" + purchase_suffix, "housing": "buy", "down_pct": 0.30, "lifecycle": "keep"},
        {"label": "Mutuo 25% + Nuda" + purchase_suffix, "housing": "buy", "down_pct": 0.25, "lifecycle": "nuda"},
        {"label": "Mutuo 25% + Downsizing" + purchase_suffix, "housing": "buy", "down_pct": 0.25, "lifecycle": "downsize"},
    ]

    if owns_home_initially:
        scenarios.insert(
            1,
            {"label": "Casa già posseduta", "housing": "owned", "down_pct": 0.0, "lifecycle": "keep"},
        )

    results = {}
    summaries = []

    for scenario in scenarios:
        df = simulate(params, scenario)
        results[scenario["label"]] = df
        summaries.append(summarize(df, params, scenario["label"]))

    summary = pd.DataFrame(summaries)

    st.subheader("Specchietto finale scenari")
    st.dataframe(
        summary.style.format({
            "Patrimonio monetario finale": "€{:,.0f}",
            "Valore casa finale": "€{:,.0f}",
            "Patrimonio totale finale": "€{:,.0f}",
            "Patrimonio monetario minimo": "€{:,.0f}",
            "Pensione INPS netta annua": "€{:,.0f}",
            "Pensione integrativa annua": "€{:,.0f}",
            "Fabbisogno non coperto cumulato": "€{:,.0f}",
            "Fabbisogno acquisto non coperto": "€{:,.0f}",
        }),
        use_container_width=True,
    )

    tab1, tab2, tab3, tab4 = st.tabs([
        "Patrimonio monetario",
        "Patrimonio totale",
        "Flusso netto",
        "Asset allocation",
    ])

    with tab1:
        st.line_chart(
            pd.DataFrame({
                k: v.set_index("Età")["Patrimonio monetario"]
                for k, v in results.items()
            })
        )

    with tab2:
        st.line_chart(
            pd.DataFrame({
                k: v.set_index("Età")["Patrimonio totale"]
                for k, v in results.items()
            })
        )

    with tab3:
        st.line_chart(
            pd.DataFrame({
                k: v.set_index("Età")["Flusso netto"]
                for k, v in results.items()
            })
        )

    with tab4:
        asset_scenario = st.selectbox(
            "Scenario asset allocation",
            list(results.keys()),
            key="asset_scenario",
        )
        st.line_chart(
            results[asset_scenario]
            .set_index("Età")[["Azionario", "Obbligazionario", "Liquidità"]]
        )

    st.subheader("Timeline eventi")
    timeline_scenario = st.selectbox(
        "Scenario timeline",
        list(results.keys()),
        key="timeline_scenario",
    )
    timeline = results[timeline_scenario]
    timeline = timeline[timeline["Eventi"] != ""][["Età", "Eventi"]]
    st.dataframe(timeline, use_container_width=True, hide_index=True)

    st.subheader("Dettaglio annuale")
    detail_scenario = st.selectbox(
        "Scenario dettagliato",
        list(results.keys()),
        key="detail_scenario",
    )
    st.dataframe(results[detail_scenario], use_container_width=True)

    st.download_button(
        "Esporta configurazione JSON",
        json.dumps(params, ensure_ascii=False, indent=2).encode("utf-8"),
        file_name="lifeplan360_config.json",
        mime="application/json",
    )

    st.download_button(
        "Scarica CSV scenario selezionato",
        results[detail_scenario].to_csv(index=False).encode("utf-8"),
        file_name=f"{detail_scenario.replace(' ', '_')}.csv",
        mime="text/csv",
    )

    excel_bytes = build_excel(params, results, summary)
    st.download_button(
        "Scarica report Excel completo",
        excel_bytes,
        file_name="LifePlan360_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


st.caption(
    "LifePlan 360 è uno strumento di simulazione e pianificazione. "
    "Non sostituisce consulenza fiscale, previdenziale o finanziaria professionale."
)
