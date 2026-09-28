import io
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd
import streamlit as st

from recon import (cost_rollup, employee_cost_detail, load_bank, load_production,
                   clean_name, load_salary, map_names, name_score, production_summary, reconcile_bank)

HERE = Path(__file__).parent
DEFAULTS = {
    "salary": HERE / "JEWELLERY AUG,2026.xlsx",
    "bank": HERE / "SHREEJAA DIAMOND & JEWELLERY MANUFACTURING LLP.xls",
    "production": HERE / "Worker Production - Jewel - Aug26.xlsx",
}

st.set_page_config(page_title="Salary Check - Jewellery", layout="wide")
st.title("Jewellery Salary / Bank / Production Check")

# ---------------------------------------------------------------- inputs
with st.sidebar:
    st.header("Files")
    up_sal = st.file_uploader("Salary sheet (.xlsx)", type=["xlsx", "xls"])
    up_bank = st.file_uploader("Bank statement (.xls/.xlsx/.html)", type=["xls", "xlsx", "html", "htm"])
    up_prod = st.file_uploader("Worker production (.xlsx / .csv)", type=["xlsx", "xls", "csv"])
    st.caption("If a file is not uploaded, the one in the project folder is used.")
    st.header("Settings")
    batch_label = st.text_input("Bulk salary text in bank description", "SAL")
    threshold = st.slider("Name match threshold", 0.5, 1.0, 0.8, 0.05)


def pick(upload, key):
    if upload is not None:
        return upload
    if DEFAULTS[key].exists():
        return str(DEFAULTS[key])
    return None


src_sal, src_bank, src_prod = pick(up_sal, "salary"), pick(up_bank, "bank"), pick(up_prod, "production")
missing = [n for n, s in [("salary", src_sal), ("bank", src_bank), ("production", src_prod)] if s is None]
if missing:
    st.warning(f"Please upload: {', '.join(missing)}")
    st.stop()

def load(label, fn, src):
    try:
        return fn(src)
    except Exception as e:
        st.error(f"Could not read the {label} file: {e}")
        st.stop()


salary = load("salary", load_salary, src_sal)
bank = load("bank statement", load_bank, src_bank)
prod = load("production", load_production, src_prod)
psum = production_summary(prod)
rec, bulk_info, used_bank = reconcile_bank(salary, bank, batch_label)

# ---------------------------------------------------------------- KPIs
paid = rec["Paid Status"].str.startswith("Paid")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Employees on salary sheet", len(rec))
c2.metric("Salary total", f"₹{rec['Salary'].sum():,.0f}")
c3.metric("Confirmed in bank", f"₹{rec.loc[paid, 'Salary'].sum():,.0f}", f"{paid.sum()} of {len(rec)}")
c4.metric("Not found / mismatch", int((~paid).sum()))
c5.metric("Production received (g)", f"{psum['Receive Wt'].sum():,.3f}")

tab_ins, tab_cost, tab_staff, tab_bank, tab_prod, tab_link, tab_raw = st.tabs(
    ["📊 Insights", "Workers (in production)", "Staff (salary & bank only)", "Salary vs Bank", "Production",
     "Salary vs Production", "Bank statement"])

money = {"Salary": "₹{:,.0f}", "Bank Amount": "₹{:,.0f}"}

# ---------------------------------------------------------------- tab 1
with tab_bank:
    st.subheader("Bulk transfer batch")
    if bulk_info["Matched"]:
        st.success(f"Bulk total ₹{bulk_info['Sheet bulk total']:,.0f} found in bank on "
                   f"{bulk_info['Bank date']:%d-%m-%Y %H:%M} — \"{bulk_info['Bank description']}\"")
    else:
        st.error(f"Bulk total ₹{bulk_info['Sheet bulk total']:,.0f} NOT found as a single bank debit.")
        cand = bank[(bank["Cr/Dr"] == "DR") & bank["Description"].str.contains(batch_label, case=False, na=False)]
        if not cand.empty:
            st.write("Salary-like debits in the statement:")
            st.dataframe(cand, hide_index=True)

    st.subheader("Per employee")
    status_filter = st.multiselect("Show status", sorted(rec["Paid Status"].unique()),
                                   default=sorted(rec["Paid Status"].unique()))
    view = rec[rec["Paid Status"].isin(status_filter)]

    def colour(row):
        c = "" if row["Paid Status"].startswith("Paid") else "background-color: rgba(255,80,80,.25)"
        return [c] * len(row)

    st.dataframe(view.style.apply(colour, axis=1).format(money, na_rep=""),
                 hide_index=True, width="stretch")

    by_mode = rec.groupby(["Mode", "Paid Status"])["Salary"].agg(["count", "sum"]).reset_index()
    st.dataframe(by_mode.style.format({"sum": "₹{:,.0f}"}), hide_index=True)

# ---------------------------------------------------------------- tab 2
with tab_prod:
    st.subheader("Production by worker")
    st.dataframe(psum.style.format({"Issue Wt": "{:,.3f}", "Receive Wt": "{:,.3f}",
                                    "Loss Wt": "{:,.3f}", "PLoss Wt": "{:,.3f}", "Loss %": "{:.2f}"}),
                 hide_index=True, width="stretch")
    worker = st.selectbox("Detail for worker", psum["Employee"])
    st.dataframe(prod[prod["Employee"] == worker], hide_index=True, width="stretch")

# ---------------------------------------------------------------- tab 3
with tab_link:
    st.subheader("1. Check name mapping")
    st.caption("Production names are matched to salary names automatically. "
               "Fix or clear the 'Salary Name' column where it is wrong.")
    auto = map_names(psum["Employee"], rec["Name"].tolist(), threshold)
    mapping = st.data_editor(
        auto, hide_index=True, width="stretch", key=f"map_{threshold}",
        column_config={
            "Production Name": st.column_config.TextColumn(disabled=True),
            "Salary Name": st.column_config.SelectboxColumn(options=[""] + sorted(rec["Name"])),
            "Score": st.column_config.NumberColumn(disabled=True, format="%.2f"),
        })

    m = mapping[mapping["Salary Name"].fillna("") != ""]
    prod_by_sal = (psum.merge(m[["Production Name", "Salary Name"]], left_on="Employee", right_on="Production Name")
                   .groupby("Salary Name", as_index=False)[["Qty", "Issue Wt", "Receive Wt", "Loss Wt"]].sum())
    combined = rec.merge(prod_by_sal, left_on="Name", right_on="Salary Name", how="left").drop(columns="Salary Name")
    combined["Loss %"] = (combined["Loss Wt"] / combined["Issue Wt"] * 100).round(2)
    combined["Salary per g"] = (combined["Salary"] / combined["Receive Wt"]).round(2)
    combined["Salary per pc"] = (combined["Salary"] / combined["Qty"]).round(2)

    st.subheader("2. Salary with production")
    only_prod = st.checkbox("Only employees with production", value=False)
    show = combined[combined["Qty"].notna()] if only_prod else combined
    st.dataframe(
        show[["Name", "Mode", "Salary", "Paid Status", "Qty", "Issue Wt", "Receive Wt",
              "Loss Wt", "Loss %", "Salary per g", "Salary per pc"]]
        .sort_values("Salary per g", ascending=False)
        .style.format({"Salary": "₹{:,.0f}", "Issue Wt": "{:,.3f}", "Receive Wt": "{:,.3f}",
                       "Loss Wt": "{:,.3f}", "Loss %": "{:.2f}", "Salary per g": "₹{:,.2f}",
                       "Salary per pc": "₹{:,.2f}", "Qty": "{:,.0f}"}, na_rep="—"),
        hide_index=True, width="stretch")

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**On salary sheet, no production entry**")
        no_prod = combined[combined["Qty"].isna()][["Name", "Salary"]]
        st.dataframe(no_prod.style.format({"Salary": "₹{:,.0f}"}), hide_index=True, width="stretch")
    with col_b:
        st.markdown("**In production, not on salary sheet** (contractors / unmapped)")
        unmapped = psum[~psum["Employee"].isin(m["Production Name"])][["Employee", "Qty", "Receive Wt"]]
        st.dataframe(unmapped, hide_index=True, width="stretch")

# ---------------------------------------------------------------- tab: employee cost
cost_fmt = {"Pcs": "{:,.0f}", "Issue Wt": "{:,.3f}", "Receive Wt": "{:,.3f}", "Share %": "{:.1f}%",
            "Salary": "₹{:,.0f}", "Cost": "₹{:,.0f}", "Cost per pc": "₹{:,.2f}", "Cost per g": "₹{:,.2f}"}

with tab_cost:
    st.caption("Employees found in the production file (name mapping is in the 'Salary vs Production' tab). "
               "Each worker's salary is split across the metal/category rows they worked on. "
               "Contractors not on the salary sheet are excluded.")
    basis_label = st.radio("Split salary by", ["Issue weight", "Pieces"], horizontal=True)
    basis = "Issue Wt" if basis_label == "Issue weight" else "Qty"
    detail = employee_cost_detail(prod, mapping, rec, basis)

    f1, f2, f3 = st.columns(3)
    emps = f1.multiselect("Employee", sorted(detail["Employee"].unique()), placeholder="All")
    metals = f2.multiselect("Metal", sorted(detail["Metal"].unique()), placeholder="All")
    cats = f3.multiselect("Category", sorted(detail["Category"].unique()), placeholder="All")
    fd = detail
    if emps:
        fd = fd[fd["Employee"].isin(emps)]
    if metals:
        fd = fd[fd["Metal"].isin(metals)]
    if cats:
        fd = fd[fd["Category"].isin(cats)]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Pcs", f"{fd['Pcs'].sum():,.0f}")
    k2.metric("Issue wt (g)", f"{fd['Issue Wt'].sum():,.3f}")
    k3.metric("Cost per pc", f"₹{fd['Cost'].sum() / max(fd['Pcs'].sum(), 1):,.2f}")
    k4.metric("Cost per g", f"₹{fd['Cost'].sum() / max(fd['Issue Wt'].sum(), 1e-9):,.2f}")

    st.subheader("Employee × Metal × Category")
    st.dataframe(fd[["Employee", "Metal", "Category", "Pcs", "Issue Wt", "Share %", "Salary",
                     "Cost", "Cost per pc", "Cost per g"]].style.format(cost_fmt, na_rep="—"),
                 hide_index=True, width="stretch", height=450)

    st.subheader("Per employee total")
    st.dataframe(cost_rollup(fd, ["Employee"]).sort_values("Cost per g", ascending=False)
                 .style.format(cost_fmt, na_rep="—"), hide_index=True, width="stretch")

    ca, cb = st.columns(2)
    with ca:
        st.subheader("By Metal × Category")
        st.dataframe(cost_rollup(fd, ["Metal", "Category"]).style.format(cost_fmt, na_rep="—"),
                     hide_index=True, width="stretch")
    with cb:
        st.subheader("By Metal")
        st.dataframe(cost_rollup(fd, ["Metal"]).style.format(cost_fmt, na_rep="—"),
                     hide_index=True, width="stretch")

# ---------------------------------------------------------------- tab: staff
staff = rec[~rec["Name"].isin(detail["Employee"])].reset_index(drop=True)

with tab_staff:
    st.caption("Employees on the salary sheet with no entry in the production file. "
               "If a worker shows up here, fix their name in the 'Salary vs Production' tab.")
    staff_paid = staff["Paid Status"].str.startswith("Paid")
    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Staff", len(staff))
    s2.metric("Salary total", f"₹{staff['Salary'].sum():,.0f}")
    s3.metric("Confirmed in bank", f"₹{staff.loc[staff_paid, 'Salary'].sum():,.0f}",
              f"{staff_paid.sum()} of {len(staff)}")
    s4.metric("Not found / mismatch", int((~staff_paid).sum()))

    st.dataframe(staff[["Name", "Account", "Mode", "Salary", "Paid Status", "Bank Date", "Bank Ref", "Bank Amount"]]
                 .style.apply(colour, axis=1).format(money, na_rep=""),
                 hide_index=True, width="stretch", height=min(40 + 35 * len(staff), 700))

# ---------------------------------------------------------------- tab 4
with tab_raw:
    st.subheader("Bank statement")
    only_used = st.checkbox("Only rows matched to salary", value=False)
    b = used_bank if only_used else bank
    st.dataframe(b.style.format({"Amount": "₹{:,.2f}", "Balance": "₹{:,.2f}"}),
                 hide_index=True, width="stretch")

# ---------------------------------------------------------------- tab: insights
MIN_WT = 50  # ignore tiny rows (g) when ranking, they give noisy ratios

with tab_ins:
    st.caption("Calculated from the loaded files; updates when you upload a new month or fix name mapping.")
    actions = []

    # pcs done vs salary
    st.subheader("Pcs done vs Salary")
    ws = cost_rollup(detail, ["Employee"]).drop(columns="Cost").merge(
        rec[["Name", "Salary"]], left_on="Employee", right_on="Name").drop(columns="Name")
    ws["Type"] = "Salaried"
    un_all = psum[~psum["Employee"].isin(m["Production Name"])].rename(columns={"Qty": "Pcs"})
    un_all = un_all[["Employee", "Pcs", "Issue Wt"]].assign(Type="Not on salary sheet")
    pv = pd.concat([ws, un_all], ignore_index=True).sort_values(["Type", "Pcs"], ascending=[False, False])
    pv = pv[["Employee", "Type", "Pcs", "Issue Wt", "Salary", "Cost per pc", "Cost per g"]]

    t1, t2, t3, t4 = st.columns(4)
    t1.metric("Total pcs done", f"{pv['Pcs'].sum():,.0f}")
    t2.metric("Pcs by salaried workers", f"{ws['Pcs'].sum():,.0f}",
              f"{ws['Pcs'].sum() / max(pv['Pcs'].sum(), 1):.0%} of total", delta_color="off")
    t3.metric("Salary of these workers", f"₹{ws['Salary'].sum():,.0f}")
    t4.metric("Salary per pc", f"₹{ws['Salary'].sum() / max(ws['Pcs'].sum(), 1):,.2f}",
              f"₹{ws['Salary'].sum() / max(ws['Issue Wt'].sum(), 1e-9):,.2f} per g", delta_color="off")

    total = pd.DataFrame([{"Employee": "TOTAL (salaried)", "Type": "", "Pcs": ws["Pcs"].sum(),
                           "Issue Wt": ws["Issue Wt"].sum(), "Salary": ws["Salary"].sum(),
                           "Cost per pc": ws["Salary"].sum() / max(ws["Pcs"].sum(), 1),
                           "Cost per g": ws["Salary"].sum() / max(ws["Issue Wt"].sum(), 1e-9)}])
    pv_show = pd.concat([pv, total], ignore_index=True).rename(
        columns={"Cost per pc": "Salary per pc", "Cost per g": "Salary per g"})
    st.dataframe(
        pv_show.style.format({"Pcs": "{:,.0f}", "Issue Wt": "{:,.3f}", "Salary": "₹{:,.0f}",
                              "Salary per pc": "₹{:,.2f}", "Salary per g": "₹{:,.2f}"}, na_rep="—")
        .apply(lambda r: ["font-weight: bold" if r["Employee"].startswith("TOTAL") else ""] * len(r), axis=1),
        hide_index=True, width="stretch", height=min(40 + 35 * len(pv_show), 800))
    st.caption("'Not on salary sheet' = contractors or names not matched (fix in 'Salary vs Production'). "
               "Their salary isn't in this file, so it shows as —.")

    # salaried but not in production
    st.subheader("On salary sheet but not in production file")
    np_ = rec[~rec["Name"].isin(ws["Employee"])].copy()
    if np_.empty:
        st.success("Every salaried employee appears in the production file.")
    else:
        free = un_all["Employee"].tolist()  # production names not yet linked to anyone
        def first_ok(a, b):  # first names must agree, surname alone is not enough
            fa, fb = clean_name(a).split()[:1], clean_name(b).split()[:1]
            return bool(fa and fb) and SequenceMatcher(None, fa[0], fb[0]).ratio() >= 0.8

        pairs = sorted(((name_score(n, f), n, f) for n in np_["Name"] for f in free if first_ok(n, f)),
                       reverse=True)
        hint, taken = {}, set()
        for sc, n, f in pairs:  # each production name suggested for one person only
            if sc >= 0.5 and n not in hint and f not in taken:
                hint[n] = f
                taken.add(f)
        np_["Possible production name"] = np_["Name"].map(hint).fillna("")
        n1, n2, n3 = st.columns(3)
        n1.metric("People", len(np_), f"{len(np_) / len(rec):.0%} of employees", delta_color="off")
        n2.metric("Their salary", f"₹{np_['Salary'].sum():,.0f}",
                  f"{np_['Salary'].sum() / rec['Salary'].sum():.0%} of payroll", delta_color="off")
        n3.metric("Average salary", f"₹{np_['Salary'].mean():,.0f}",
                  f"workers avg ₹{ws['Salary'].mean():,.0f}" if len(ws) else None, delta_color="off")
        top3 = np_.nlargest(3, "Salary")
        bullets = [f"Highest paid: " + ", ".join(f"**{r.Name}** ₹{r.Salary:,.0f}" for r in top3.itertuples())]
        likely = np_[np_["Possible production name"] != ""]
        if not likely.empty:
            bullets.append("May be the same person under a different name in production: " + ", ".join(
                f"**{a}** → `{b}`" for a, b in zip(likely["Name"], likely["Possible production name"])) +
                " — link them in 'Salary vs Production' to move them to workers.")
            actions.append("Confirm possible name matches: " + ", ".join(likely["Name"]) + ".")
        bullets.append("The rest are likely office / support staff (no pieces recorded), "
                       "so their salary is overhead on top of the per-piece labour cost above.")
        st.markdown("\n".join(f"- {b}" for b in bullets))
        with st.expander(f"Show all {len(np_)} names"):
            st.dataframe(np_[["Name", "Mode", "Salary", "Paid Status", "Possible production name"]]
                         .sort_values("Salary", ascending=False).style.format({"Salary": "₹{:,.0f}"}),
                         hide_index=True, width="stretch")
    st.divider()

    # 1. payments
    st.subheader("1. Salary payment check")
    n_bad = int((~paid).sum())
    if n_bad == 0 and bulk_info["Matched"]:
        st.success(f"All {len(rec)} employees (₹{rec['Salary'].sum():,.0f}) are confirmed in the bank statement. "
                   "No mismatches or missing payments.")
    else:
        st.error(f"{n_bad} employee(s) worth ₹{rec.loc[~paid, 'Salary'].sum():,.0f} not confirmed in the bank. "
                 "See the 'Salary vs Bank' tab.")
        actions.append("Investigate salary lines not found in the bank statement.")
    n_bulk = int((rec["Mode"] == "BULK").sum())
    if n_bulk:
        st.info(f"Bulk batch (₹{bulk_info['Sheet bulk total']:,.0f}, {n_bulk} people) can only be checked as a total. "
                "To verify each account, get the bank's bulk-payment success report.")

    # 2. payroll split
    st.subheader("2. Payroll split")
    w_sal, s_sal, tot = detail.drop_duplicates("Employee")["Salary"].sum(), staff["Salary"].sum(), rec["Salary"].sum()
    p1, p2, p3 = st.columns(3)
    p1.metric("Workers (in production)", f"{detail['Employee'].nunique()} people",
              f"₹{w_sal:,.0f} · {w_sal / tot:.0%}", delta_color="off")
    p2.metric("Staff (no production)", f"{len(staff)} people",
              f"₹{s_sal:,.0f} · {s_sal / tot:.0%}", delta_color="off")
    top2 = rec.nlargest(2, "Salary")
    ratio = top2["Salary"].iloc[0] / top2["Salary"].iloc[1]
    p3.metric("Highest salary", f"₹{top2['Salary'].iloc[0]:,.0f}", top2["Name"].iloc[0], delta_color="off")
    if s_sal > w_sal:
        st.markdown(f"- Staff with no production take **{s_sal / tot:.0%}** of payroll. Check whether some are "
                    "workers whose production names didn't match (see 'Salary vs Production').")
    if ratio >= 1.8:
        st.markdown(f"- **{top2['Name'].iloc[0]}** earns {ratio:.1f}× the next highest salary — confirm role.")

    # 3. production by people not on the salary sheet
    st.subheader("3. Production by people not on the salary sheet")
    un = psum[~psum["Employee"].isin(m["Production Name"])]
    u1, u2, u3 = st.columns(3)
    u1.metric("Share of issue weight", f"{un['Issue Wt'].sum() / psum['Issue Wt'].sum():.0%}")
    u2.metric("Share of pieces", f"{un['Qty'].sum() / psum['Qty'].sum():.0%}")
    u3.metric("Share of metal loss", f"{un['Loss Wt'].sum() / psum['Loss Wt'].sum():.0%}")
    big_un = un[un["Issue Wt"] >= MIN_WT]
    if not big_un.empty:
        st.markdown(f"- Main names: **{', '.join(big_un['Employee'])}**. If these are contractors paid separately, "
                    "the real labour cost per gram is higher than shown here.")
        st.dataframe(big_un[["Employee", "Qty", "Issue Wt", "Loss Wt", "Loss %"]]
                     .style.format({"Issue Wt": "{:,.3f}", "Loss Wt": "{:,.3f}", "Loss %": "{:.2f}"}),
                     hide_index=True)
        actions.append(f"Map or add contractor cost for: {', '.join(big_un['Employee'])}.")

    # 4. worker cost spread
    st.subheader("4. Labour cost per gram by worker")
    wc = cost_rollup(detail, ["Employee"])
    wc = wc[wc["Issue Wt"] > 0]
    if not wc.empty:
        med = wc["Cost per g"].median()
        c_avg = detail["Cost"].sum() / detail["Issue Wt"].sum()
        lo, hi = wc.nsmallest(1, "Cost per g").iloc[0], wc.nlargest(1, "Cost per g").iloc[0]
        q1, q2, q3 = st.columns(3)
        q1.metric("Average cost per g", f"₹{c_avg:,.2f}")
        q2.metric("Lowest", f"₹{lo['Cost per g']:,.2f}", lo["Employee"], delta_color="off")
        q3.metric("Highest", f"₹{hi['Cost per g']:,.2f}", hi["Employee"], delta_color="off")
        outl = wc[wc["Cost per g"] > 2 * med].sort_values("Cost per g", ascending=False)
        if not outl.empty:
            st.markdown(f"- **{len(outl)} worker(s)** cost more than 2× the median (₹{med:,.2f}/g). "
                        "Either their work isn't recorded in production, or output is low:")
            st.dataframe(outl.style.format(cost_fmt, na_rep="—"), hide_index=True)
            actions.append(f"Review roles/output of: {', '.join(outl['Employee'])}.")

    # 5. loss
    st.subheader("5. Metal loss")
    overall = psum["Loss Wt"].sum() / psum["Issue Wt"].sum() * 100
    st.metric("Overall loss", f"{overall:.2f}%", f"{psum['Loss Wt'].sum():,.3f} g", delta_color="off")
    l1, l2, l3 = st.columns(3)
    ranked = psum[psum["Issue Wt"] >= MIN_WT]
    with l1:
        st.markdown("**Highest loss % (workers)**")
        st.dataframe(ranked.nlargest(5, "Loss %")[["Employee", "Issue Wt", "Loss %"]]
                     .style.format({"Issue Wt": "{:,.1f}", "Loss %": "{:.2f}"}), hide_index=True)
    with l2:
        st.markdown("**Lowest loss % (workers)**")
        st.dataframe(ranked.nsmallest(5, "Loss %")[["Employee", "Issue Wt", "Loss %"]]
                     .style.format({"Issue Wt": "{:,.1f}", "Loss %": "{:.2f}"}), hide_index=True)
    with l3:
        st.markdown("**By category**")
        cat = prod.groupby("Category", as_index=False)[["Issue Wt", "Loss Wt"]].sum()
        cat["Loss %"] = cat["Loss Wt"] / cat["Issue Wt"] * 100
        st.dataframe(cat.sort_values("Loss %", ascending=False)[["Category", "Issue Wt", "Loss %"]]
                     .style.format({"Issue Wt": "{:,.1f}", "Loss %": "{:.2f}"}), hide_index=True)
    high_cat = cat[(cat["Issue Wt"] >= 10) & (cat["Loss %"] > overall + 1.5)]
    if not high_cat.empty:
        st.markdown(f"- Categories well above average loss: **{', '.join(high_cat['Category'])}**.")
        actions.append(f"Look into loss on {', '.join(high_cat['Category'])} and the highest-loss workers.")
    met = prod.groupby("Metal", as_index=False)[["Issue Wt", "Loss Wt", "PLoss Wt"]].sum()
    met["Loss %"] = met["Loss Wt"] / met["Issue Wt"] * 100
    st.markdown("- By metal: " + ", ".join(f"**{r.Metal}** {r['Loss %']:.2f}%" for _, r in met.iterrows()))

    # 6. product cost
    st.subheader("6. Labour cost by product (salaried workers only)")
    mc = cost_rollup(detail, ["Metal", "Category"])
    mc = mc[mc["Issue Wt"] >= MIN_WT]
    d1, d2 = st.columns(2)
    with d1:
        st.markdown("**Most expensive per gram**")
        st.dataframe(mc.nlargest(5, "Cost per g").style.format(cost_fmt, na_rep="—"), hide_index=True)
    with d2:
        st.markdown("**Cheapest per gram**")
        st.dataframe(mc.nsmallest(5, "Cost per g").style.format(cost_fmt, na_rep="—"), hide_index=True)
    st.caption(f"Salary split by {basis_label.lower()} (change on the Workers tab). "
               f"Rows under {MIN_WT} g are left out of rankings.")

    # next steps
    st.subheader("Suggested next steps")
    if actions:
        st.markdown("\n".join(f"{i}. {a}" for i, a in enumerate(actions, 1)))
    else:
        st.markdown("Nothing needs attention.")

# ---------------------------------------------------------------- export
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as xw:
    rec.to_excel(xw, sheet_name="Salary vs Bank", index=False)
    combined.to_excel(xw, sheet_name="Salary vs Production", index=False)
    psum.to_excel(xw, sheet_name="Production Summary", index=False)
    detail.to_excel(xw, sheet_name="Workers Cost Detail", index=False)
    staff.to_excel(xw, sheet_name="Staff", index=False)
    cost_rollup(detail, ["Metal", "Category"]).to_excel(xw, sheet_name="Cost by Metal-Category", index=False)
    mapping.to_excel(xw, sheet_name="Name Mapping", index=False)
st.sidebar.download_button("Download report (.xlsx)", buf.getvalue(), "salary_check_report.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
