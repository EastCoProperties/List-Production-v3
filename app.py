
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from xlsxwriter.utility import xl_rowcol_to_cell

st.set_page_config(page_title="Underwriting Engine V2.1", layout="wide")
st.title("Underwriting Engine V2.1")
st.caption("Confidence-weighted vacant land underwriting with distorted-assessment handling")

BUCKETS = [
    (1.00, 1.99, "1.00-1.99"),
    (2.00, 2.99, "2.00-2.99"),
    (3.00, 4.99, "3.00-4.99"),
    (5.00, 7.49, "5.00-7.49"),
    (7.50, 9.99, "7.50-9.99"),
    (10.00, 14.99, "10.00-14.99"),
    (15.00, 24.99, "15.00-24.99"),
    (25.00, 49.99, "25.00-49.99"),
    (50.00, 99.99, "50.00-99.99"),
    (100.00, float("inf"), "100+"),
]

def to_num(series):
    return pd.to_numeric(
        pd.Series(series).astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.strip()
        .replace({"": np.nan, "nan": np.nan, "None": np.nan}),
        errors="coerce",
    )

def acre_bucket(acres):
    if pd.isna(acres) or acres <= 0:
        return "Unknown"
    for low, high, label in BUCKETS:
        if low <= acres <= high:
            return label
    return "Unknown"

def roi_target(mv):
    if pd.isna(mv) or mv <= 0:
        return np.nan
    return min(0.75, max(0.55, 0.55 + ((mv - 50000) / 200000) * 0.20))

def assessed_weight_and_label(assessed_reliability):
    if pd.isna(assessed_reliability):
        return 0.10, "Unknown"
    if assessed_reliability >= 0.70:
        return 0.20, "Normal"
    if assessed_reliability >= 0.40:
        return 0.10, "Possibly Low"
    return 0.00, "Likely Deferred/Distorted"

def weighted_mv(tlp, parent_median, adjusted_county, tlp_ratio, parent_comp_count, assessed_reliability):
    vals = {
        "tlp": tlp if pd.notna(tlp) and tlp > 0 else np.nan,
        "parent": parent_median if pd.notna(parent_median) and parent_median > 0 else np.nan,
        "county": adjusted_county if pd.notna(adjusted_county) and adjusted_county > 0 else np.nan,
    }

    county_weight, assessed_label = assessed_weight_and_label(assessed_reliability)

    strong_comps = pd.notna(parent_comp_count) and parent_comp_count >= 10
    usable_comps = pd.notna(parent_comp_count) and parent_comp_count >= 5

    if pd.isna(tlp_ratio):
        tlp_weight = 0.45
        parent_weight = 0.55 - county_weight
        confidence = "Unknown TLP"
    elif tlp_ratio <= 1.40 and strong_comps:
        tlp_weight = 0.70
        parent_weight = 0.30 - county_weight
        confidence = "High TLP Confidence"
    elif tlp_ratio <= 1.80 and usable_comps:
        tlp_weight = 0.40
        parent_weight = 0.60 - county_weight
        confidence = "Medium TLP Confidence"
    elif tlp_ratio <= 2.50 and usable_comps:
        tlp_weight = 0.25
        parent_weight = 0.75 - county_weight
        confidence = "Low TLP Confidence"
    else:
        tlp_weight = 0.15
        parent_weight = 0.85 - county_weight
        confidence = "TLP Likely Inflated"

    weights = {"tlp": tlp_weight, "parent": max(parent_weight, 0), "county": county_weight}

    # V2.1: if county assessed appears distorted, remove adjusted county MV entirely.
    if county_weight <= 0:
        vals["county"] = np.nan

    available = {k: v for k, v in vals.items() if pd.notna(v) and v > 0}
    if not available:
        return np.nan, "No MV Source", assessed_label, county_weight

    total_weight = sum(weights[k] for k in available)
    if total_weight <= 0:
        return np.nan, "No MV Source", assessed_label, county_weight

    mv = sum(available[k] * weights[k] for k in available) / total_weight
    return mv, confidence, assessed_label, county_weight

def offer_floor(mv, assessed, assessed_reliability):
    # V2.1: if assessed is distorted, do not use it in the offer floor.
    if pd.notna(assessed_reliability) and assessed_reliability < 0.40:
        return mv * 0.60
    if pd.isna(assessed) or assessed <= 0:
        return mv * 0.60
    return max(mv * 0.60, assessed * 0.60)

def make_offer_set(strategy, acres, frontage, nav, mv, assessed, assessed_reliability, min_profit=20000):
    if pd.isna(nav) or pd.isna(mv) or mv <= 0:
        return ("No Offer", "No Offer", "No Offer", np.nan)

    roi = roi_target(mv)
    if pd.isna(roi):
        return ("No Offer", "No Offer", "No Offer", np.nan)

    floor = offer_floor(mv, assessed, assessed_reliability)

    road_cost = (acres * 43560 / frontage) * 100 if pd.notna(acres) and pd.notna(frontage) and frontage > 0 else np.nan
    offer20 = min(nav / (1 + roi) - 20000, nav - 40000)
    offer7 = min(nav / (1 + roi) - 7000, nav - 27000)

    if strategy == "Road":
        if pd.isna(road_cost):
            return ("No Offer", "No Offer", "No Offer", np.nan)
        base = min(nav / (1 + roi) - road_cost, nav - road_cost - min_profit)
        cost = road_cost
    else:
        if offer20 >= floor:
            base, cost = offer20, 20000
        elif offer7 >= floor:
            base, cost = offer7, 7000
        else:
            return ("No Offer", "No Offer", "No Offer", np.nan)

    if pd.isna(base) or base < floor or nav - (base + cost) < min_profit:
        return ("No Offer", "No Offer", "No Offer", cost)

    standard = min(base, mv * 1.10) if base / mv > 1.10 else base
    offers = [max(standard - mv * 0.06, floor), max(standard - mv * 0.03, floor), standard]

    final = []
    for offer in offers:
        if offer < floor or nav - (offer + cost) < min_profit:
            final.append("No Offer")
        else:
            final.append(round(float(offer), 2))
    return final[0], final[1], final[2], cost

def formula_templates(rownum, col_letters):
    c = lambda name: f"{col_letters[name]}{rownum}"

    acres = c("_Calc Acres")
    frontage = c("_Calc Frontage")
    assessed = c("_Calc Assessed")
    ass_rel = c("Assessed Reliability Ratio")
    nav = c("New After Value")
    strategy = c("Strategy")
    mv = c("Offer MV Used")
    devcost = c("Development Cost Used")
    first = c("1st Offer")
    second = c("2nd Offer")
    third = c("3rd Offer")

    roi = f"MIN(0.75,MAX(0.55,0.55+(({mv}-50000)/200000)*0.2))"
    floor = f"IF({ass_rel}<0.4,{mv}*0.6,MAX({mv}*0.6,{assessed}*0.6))"
    road_cost = f"({acres}*43560/{frontage})*100"
    offer20 = f"MIN({nav}/(1+{roi})-20000,{nav}-40000)"
    offer7 = f"MIN({nav}/(1+{roi})-7000,{nav}-27000)"
    road_offer = f"MIN({nav}/(1+{roi})-{road_cost},{nav}-{road_cost}-20000)"
    cost = f'IF({strategy}="Road",{road_cost},IF({offer20}>={floor},20000,IF({offer7}>={floor},7000,0)))'
    base = f'IF({strategy}="Road",{road_offer},IF({offer20}>={floor},{offer20},IF({offer7}>={floor},{offer7},"No Offer")))'
    standard = f'IF({base}="No Offer","No Offer",IF({base}/{mv}>1.1,{mv}*1.1,{base}))'

    f1 = f'IF({standard}="No Offer","No Offer",MAX({standard}-({mv}*0.06),{floor}))'
    f2 = f'IF({standard}="No Offer","No Offer",MAX({standard}-({mv}*0.03),{floor}))'
    f3 = standard

    formulas = {}
    formulas["Development Cost Used"] = f'=IF({mv}="No Offer","No Offer",{cost})'
    formulas["1st Offer"] = f'=IF({mv}="No Offer","No Offer",LET(final,{f1},cost,{cost},IF(final="No Offer","No Offer",IF(OR(final<{floor},{nav}-(final+cost)<20000),"No Offer",final))))'
    formulas["2nd Offer"] = f'=IF({mv}="No Offer","No Offer",LET(final,{f2},cost,{cost},IF(final="No Offer","No Offer",IF(OR(final<{floor},{nav}-(final+cost)<20000),"No Offer",final))))'
    formulas["3rd Offer"] = f'=IF({mv}="No Offer","No Offer",LET(final,{f3},cost,{cost},IF(final="No Offer","No Offer",IF(OR(final<{floor},{nav}-(final+cost)<20000),"No Offer",final))))'

    formulas["% MV 1st Offer"] = f'=IF(OR({first}="No Offer",{mv}="No Offer"),"No Offer",{first}/{mv})'
    formulas["% MV 2nd Offer"] = f'=IF(OR({second}="No Offer",{mv}="No Offer"),"No Offer",{second}/{mv})'
    formulas["% MV 3rd Offer"] = f'=IF(OR({third}="No Offer",{mv}="No Offer"),"No Offer",{third}/{mv})'

    formulas["Profit - 1st Offer"] = f'=IF(OR({first}="No Offer",{devcost}="No Offer"),"No Offer",{nav}-({first}+{devcost}))'
    formulas["Profit - 2nd Offer"] = f'=IF(OR({second}="No Offer",{devcost}="No Offer"),"No Offer",{nav}-({second}+{devcost}))'
    formulas["Profit - 3rd Offer"] = f'=IF(OR({third}="No Offer",{devcost}="No Offer"),"No Offer",{nav}-({third}+{devcost}))'

    p1 = c("Profit - 1st Offer")
    p2 = c("Profit - 2nd Offer")
    p3 = c("Profit - 3rd Offer")

    formulas["ROI - 1st Offer"] = f'=IF(OR({first}="No Offer",{devcost}="No Offer"),"No Offer",{p1}/({first}+{devcost}))'
    formulas["ROI - 2nd Offer"] = f'=IF(OR({second}="No Offer",{devcost}="No Offer"),"No Offer",{p2}/({second}+{devcost}))'
    formulas["ROI - 3rd Offer"] = f'=IF(OR({third}="No Offer",{devcost}="No Offer"),"No Offer",{p3}/({third}+{devcost}))'

    roi2 = c("ROI - 2nd Offer")
    tlp_ratio = c("TLP Reliability Ratio")
    formulas["Review Flag"] = (
        f'=IF({third}="No Offer","No Offer",'
        f'IF({tlp_ratio}>2.5,"Review - TLP High",'
        f'IF({ass_rel}<0.4,"Review - Likely Deferred Assessed",'
        f'IF(AND({assessed}>0,{third}/{assessed}>=6),"Review - Offer 600%+ of Assessed",'
        f'IF({strategy}="Road","Review - Road","OK")))))'
    )
    formulas["Deal Quality Score"] = f'=IF(OR({mv}="No Offer",{roi2}="No Offer"),"No Offer",({nav}/{mv})+({roi2}*0.5)-IF({strategy}="Road",0.25,0))'
    return formulas

uploaded = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx", "xls"])

if uploaded:
    if uploaded.name.lower().endswith(".csv"):
        raw = pd.read_csv(uploaded, low_memory=False)
    else:
        raw = pd.read_excel(uploaded)

    st.success(f"Loaded {len(raw):,} rows and {len(raw.columns):,} columns.")

    cols = list(raw.columns)
    def best_index(names):
        for name in names:
            if name in cols:
                return cols.index(name)
        lower = {c.lower(): i for i, c in enumerate(cols)}
        for name in names:
            if name.lower() in lower:
                return lower[name.lower()]
        return 0

    with st.expander("Column Mapping", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        acres_col = c1.selectbox("Lot Acres", cols, index=best_index(["Lot Acres", "Calc Acreage", "Acres", "Acreage"]))
        frontage_col = c2.selectbox("Road Frontage", cols, index=best_index(["Road Frontage", "Frontage", "Road_Frontage"]))
        tlp_col = c3.selectbox("TLP Estimate / Current MV", cols, index=best_index(["TLP Estimate", "TLP", "Estimated Value"]))
        assessed_col = c4.selectbox("County Assessed Value", cols, index=best_index(["Total Assessed Value", "Assessed Value", "Total Market Value"]))

    with st.expander("Assumptions", expanded=False):
        a1, a2, a3, a4 = st.columns(4)
        max_lots = a1.number_input("Max Lot Count", value=5, min_value=1, max_value=50)
        min_frontage = a2.number_input("Min Frontage Per Lot", value=150, min_value=1)
        road_cost_lf = a3.number_input("Road Cost / Linear Foot", value=100, min_value=1)
        selling_factor = a4.number_input("Net Selling Factor", value=0.93, min_value=0.1, max_value=1.0)

    if st.button("Process File", type="primary"):
        with st.spinner("Processing underwriting model..."):
            calc = pd.DataFrame(index=raw.index)
            calc["_Calc Acres"] = to_num(raw[acres_col])
            calc["_Calc Frontage"] = to_num(raw[frontage_col])
            calc["_Calc TLP"] = to_num(raw[tlp_col])
            calc["_Calc Assessed"] = to_num(raw[assessed_col])

            calc["Acreage Bucket"] = calc["_Calc Acres"].apply(acre_bucket)

            valid = calc[(calc["_Calc Acres"] > 0) & (calc["_Calc TLP"] > 0)].copy()
            valid["TLP PPA"] = valid["_Calc TLP"] / valid["_Calc Acres"]

            assessed_valid = calc[(calc["_Calc Acres"] > 0) & (calc["_Calc Assessed"] > 0)].copy()
            assessed_valid["Assessed PPA"] = assessed_valid["_Calc Assessed"] / assessed_valid["_Calc Acres"]

            tlp_bucket_median = valid.groupby("Acreage Bucket")["TLP PPA"].median().to_dict()
            tlp_bucket_count = valid.groupby("Acreage Bucket")["TLP PPA"].count().to_dict()
            assessed_bucket_median = assessed_valid.groupby("Acreage Bucket")["Assessed PPA"].median().to_dict()

            lots = []
            for _, r in calc.iterrows():
                a, f = r["_Calc Acres"], r["_Calc Frontage"]
                if pd.isna(a) or pd.isna(f) or a <= 0 or f <= 0:
                    lots.append(np.nan)
                else:
                    lots.append(max(1, min(max_lots, np.floor(f / min_frontage), np.floor((3 * f * f) / (a * 43560)))))
            calc["Estimated Lot Count"] = lots
            calc["Avg Resulting Lot Size"] = calc["_Calc Acres"] / calc["Estimated Lot Count"]
            calc["Resulting Lot Bucket"] = calc["Avg Resulting Lot Size"].apply(acre_bucket)

            calc["Original $/Acre"] = calc["_Calc TLP"] / calc["_Calc Acres"]
            calc["Median Comp $/Acre"] = calc["Resulting Lot Bucket"].map(tlp_bucket_median)
            calc["Parent Median $/Acre"] = calc["Acreage Bucket"].map(tlp_bucket_median)
            calc["Parent Comp Count"] = calc["Acreage Bucket"].map(tlp_bucket_count).fillna(0)
            calc["Parent Median Value"] = calc["Parent Median $/Acre"] * calc["_Calc Acres"]
            calc["Value Increase %"] = np.maximum(0, (calc["Median Comp $/Acre"] / calc["Original $/Acre"]) - 1)

            calc["County Assessed $/Acre"] = calc["_Calc Assessed"] / calc["_Calc Acres"]
            calc["Median Assessed $/Acre"] = calc["Acreage Bucket"].map(assessed_bucket_median)
            calc["Assessed Reliability Ratio"] = calc["County Assessed $/Acre"] / calc["Median Assessed $/Acre"]
            calc["TLP Reliability Ratio"] = calc["Original $/Acre"] / calc["Parent Median $/Acre"]

            reliable = (
                calc["Assessed Reliability Ratio"].between(0.4, 1.8) &
                (calc["_Calc Assessed"] > 0) &
                (calc["_Calc TLP"] > 0)
            )
            ratios = (calc.loc[reliable, "_Calc TLP"] / calc.loc[reliable, "_Calc Assessed"]).replace([np.inf, -np.inf], np.nan).dropna()
            if len(ratios) < 5:
                ratios = (calc["_Calc TLP"] / calc["_Calc Assessed"]).replace([np.inf, -np.inf], np.nan).dropna()
            cal_factor = float(ratios.median()) if len(ratios) else 1.5

            calc["County Calibration Factor"] = cal_factor
            calc["Adjusted County MV"] = calc["_Calc Assessed"] * cal_factor

            road_ppa = tlp_bucket_median.get("1.00-1.99", np.nan)
            if pd.isna(road_ppa) and len(tlp_bucket_median):
                road_ppa = np.nanmedian(list(tlp_bucket_median.values()))

            strategies = []
            after_values = []
            for _, r in calc.iterrows():
                acres = r["_Calc Acres"]
                frontage = r["_Calc Frontage"]
                mv = r["_Calc TLP"]
                clean_ppa = r["Median Comp $/Acre"]
                parent_ppa = r["Parent Median $/Acre"]
                orig_ppa = r["Original $/Acre"]
                lots_i = r["Estimated Lot Count"]

                premium = 1.0
                if pd.notna(orig_ppa) and pd.notna(parent_ppa) and parent_ppa > 0:
                    premium = min(1.35, max(1.0, orig_ppa / parent_ppa))

                clean_value_raw = (clean_ppa if pd.notna(clean_ppa) else 0) * premium * (acres if pd.notna(acres) else 0) * selling_factor
                clean_cap = (mv if pd.notna(mv) else 0) * 2.0
                clean_value = max(mv if pd.notna(mv) else 0, min(clean_value_raw, clean_cap if clean_cap > 0 else clean_value_raw))

                road_value_raw = (road_ppa if pd.notna(road_ppa) else 0) * (acres if pd.notna(acres) else 0) * 0.75 * selling_factor
                road_cap = (mv if pd.notna(mv) else 0) * 1.75
                road_value = min(road_value_raw, road_cap if road_cap > 0 else road_value_raw)

                if pd.notna(lots_i) and lots_i < 3 and pd.notna(acres) and acres >= 10 and pd.notna(frontage) and frontage >= 50 and road_value > clean_value:
                    strategies.append("Road")
                    after_values.append(max(clean_value, road_value, mv if pd.notna(mv) else 0))
                else:
                    strategies.append("Frontage")
                    after_values.append(clean_value)

            calc["New After Value"] = after_values
            calc["Strategy"] = strategies

            weighted_mvs = []
            mv_conf = []
            assessed_labels = []
            assessed_weights = []
            for _, r in calc.iterrows():
                mv, label, ass_label, ass_weight = weighted_mv(
                    r["_Calc TLP"], r["Parent Median Value"], r["Adjusted County MV"],
                    r["TLP Reliability Ratio"], r["Parent Comp Count"], r["Assessed Reliability Ratio"]
                )

                if pd.notna(r["Assessed Reliability Ratio"]) and r["Assessed Reliability Ratio"] < 0.4:
                    candidates = [mv, r["Parent Median Value"], r["_Calc TLP"]]
                else:
                    candidates = [mv, r["Parent Median Value"], r["_Calc TLP"], r["Adjusted County MV"], r["_Calc Assessed"]]

                chosen = "No Offer"
                for candidate in candidates:
                    if pd.isna(candidate) or candidate <= 0:
                        continue
                    offers = make_offer_set(
                        r["Strategy"], r["_Calc Acres"], r["_Calc Frontage"], r["New After Value"],
                        candidate, r["_Calc Assessed"], r["Assessed Reliability Ratio"]
                    )
                    if offers[2] != "No Offer":
                        chosen = round(float(candidate), 2)
                        break

                weighted_mvs.append(chosen)
                mv_conf.append(label)
                assessed_labels.append(ass_label)
                assessed_weights.append(ass_weight)

            calc["MV Confidence Label"] = mv_conf
            calc["Assessed Reliability Label"] = assessed_labels
            calc["Assessed Weight Used"] = assessed_weights
            calc["Offer MV Used"] = weighted_mvs

            for col in [
                "Development Cost Used", "1st Offer", "2nd Offer", "3rd Offer",
                "% MV 1st Offer", "% MV 2nd Offer", "% MV 3rd Offer",
                "Profit - 1st Offer", "Profit - 2nd Offer", "Profit - 3rd Offer",
                "ROI - 1st Offer", "ROI - 2nd Offer", "ROI - 3rd Offer",
                "Review Flag", "Deal Quality Score"
            ]:
                calc[col] = ""

            out = pd.concat([raw, calc], axis=1)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                out.to_excel(writer, index=False, sheet_name="Offer Model")
                wb = writer.book
                ws = writer.sheets["Offer Model"]
                headers = list(out.columns)
                col_letters = {h: xl_rowcol_to_cell(0, idx).rstrip("1") for idx, h in enumerate(headers)}

                for rownum in range(2, len(out) + 2):
                    for colname, formula in formula_templates(rownum, col_letters).items():
                        ws.write_formula(rownum - 1, headers.index(colname), formula)

                header_fmt = wb.add_format({"bold": True, "bg_color": "#0F766E", "font_color": "white", "border": 1})
                money_fmt = wb.add_format({"num_format": "$#,##0"})
                pct_fmt = wb.add_format({"num_format": "0.0%"})

                for idx, h in enumerate(headers):
                    ws.write(0, idx, h, header_fmt)
                    ws.set_column(idx, idx, 14)

                for h in [
                    "Parent Median Value", "Adjusted County MV", "New After Value", "Offer MV Used",
                    "Development Cost Used", "1st Offer", "2nd Offer", "3rd Offer",
                    "Profit - 1st Offer", "Profit - 2nd Offer", "Profit - 3rd Offer"
                ]:
                    if h in headers:
                        ws.set_column(headers.index(h), headers.index(h), 16, money_fmt)

                for h in [
                    "Value Increase %", "Assessed Reliability Ratio", "TLP Reliability Ratio",
                    "Assessed Weight Used", "% MV 1st Offer", "% MV 2nd Offer", "% MV 3rd Offer",
                    "ROI - 1st Offer", "ROI - 2nd Offer", "ROI - 3rd Offer"
                ]:
                    if h in headers:
                        ws.set_column(headers.index(h), headers.index(h), 14, pct_fmt)

                ws.freeze_panes(1, 0)
                ws.autofilter(0, 0, len(out), len(headers) - 1)

                summary = pd.DataFrame({
                    "Metric": [
                        "Rows Processed", "Frontage Strategy Count", "Road Strategy Count", "No Offer Count",
                        "Likely Deferred/Distorted Assessed Count", "County Calibration Factor",
                        "Median TLP Reliability Ratio", "Median Assessed Reliability Ratio"
                    ],
                    "Value": [
                        len(out), (calc["Strategy"] == "Frontage").sum(), (calc["Strategy"] == "Road").sum(),
                        (calc["Offer MV Used"] == "No Offer").sum(),
                        (calc["Assessed Reliability Label"] == "Likely Deferred/Distorted").sum(),
                        cal_factor,
                        calc["TLP Reliability Ratio"].replace([np.inf, -np.inf], np.nan).median(),
                        calc["Assessed Reliability Ratio"].replace([np.inf, -np.inf], np.nan).median()
                    ]
                })
                summary.to_excel(writer, index=False, sheet_name="Summary")

                assumptions = pd.DataFrame({
                    "Assumption": [
                        "Version", "Acreage Buckets", "Max Lot Count", "Min Frontage Per Lot",
                        "Ideal Depth/Width Ratio", "Road Trigger", "Road Land Loss", "Road Lot Size",
                        "Road Cost", "Selling Factor", "Minimum Profit", "ROI Range", "Offer Spread",
                        "Manual Adjustment", "Dynamic Assessed Weighting", "Distorted Assessed Rule", "Offer Floor Rule"
                    ],
                    "Value": [
                        "V2.1", ", ".join([b[2] for b in BUCKETS]), max_lots, min_frontage,
                        "3.0", "Lots < 3, acres >= 10, frontage >= 50", "25%", "1.5 acres",
                        f"${road_cost_lf}/linear foot", selling_factor, "$20,000", "55% to 75%",
                        "3rd standard, 2nd -3%, 1st -6%",
                        "Edit Offer MV Used to recalculate downstream columns",
                        "County weight reduced when assessed is below assessed PPA trendline",
                        "If Assessed Reliability Ratio < 0.40, assessed gets 0% weight and is excluded from offer floor and 600% assessed review.",
                        "If assessed distorted: 60% of Offer MV Used only. Otherwise: max(60% of Offer MV Used, 60% of Assessed)."
                    ]
                })
                assumptions.to_excel(writer, index=False, sheet_name="Assumptions")

            output.seek(0)
            st.success("Processing complete.")
            st.download_button(
                "Download processed workbook",
                data=output,
                file_name="processed_underwriting_engine_v2_1.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
