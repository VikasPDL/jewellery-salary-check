"""Parsing and matching logic for salary / bank / production reconciliation."""
import io
import re
from difflib import SequenceMatcher

import pandas as pd


# ---------------------------------------------------------------- helpers
def clean_name(s) -> str:
    s = str(s or "").upper()
    s = re.sub(r"[^A-Z ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _tok_sim(a: str, b: str) -> float:
    if a == b:
        return 1.0
    # initials: "M" matches "MOHAN"
    if len(a) == 1 or len(b) == 1:
        return 0.9 if a[0] == b[0] else 0.0
    return SequenceMatcher(None, a, b).ratio()


def name_score(a: str, b: str) -> float:
    """Similarity 0-1 between two person names, tolerant to missing middle
    names, spelling differences (AYRE/AYARE) and initials."""
    ta, tb = clean_name(a).split(), clean_name(b).split()
    if not ta or not tb:
        return 0.0
    short, long_ = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
    used, total = set(), 0.0
    for t in short:
        best, bi = 0.0, None
        for i, u in enumerate(long_):
            if i in used:
                continue
            s = _tok_sim(t, u)
            if s > best:
                best, bi = s, i
        if best >= 0.75:
            used.add(bi)
            total += best
    score = total / len(short)
    # a single matching first name is weak evidence
    if len(short) == 1:
        score *= 0.6
    return round(score, 3)


def best_match(name: str, candidates: list[str], threshold: float = 0.8):
    scored = sorted(((name_score(name, c), c) for c in candidates), reverse=True)
    if not scored or scored[0][0] < threshold:
        return None, scored[0][0] if scored else 0.0
    return scored[0][1], scored[0][0]


def _src(f):
    """Accept a path or an uploaded file object."""
    if hasattr(f, "getvalue"):
        return io.BytesIO(f.getvalue())
    return f


# ---------------------------------------------------------------- salary
def load_salary(f) -> pd.DataFrame:
    raw = pd.read_excel(_src(f), header=None)
    rows = []
    for _, r in raw.iterrows():
        name, acct, amt = r.iloc[1], r.iloc[2], r.iloc[3]
        if pd.isna(name) or pd.isna(amt):
            continue
        name = str(name).replace("�", "").strip()
        if name.upper() in ("NAME", "TOTAL"):
            continue
        try:
            amt = float(amt)
        except (TypeError, ValueError):
            continue
        acct = "" if pd.isna(acct) else str(acct).strip()
        mode = "NEFT" if acct.upper() == "NEFT" else "BULK"
        rows.append({"Name": name, "Account": acct if mode == "BULK" else "",
                     "Mode": mode, "Salary": amt})
    return pd.DataFrame(rows)


# ---------------------------------------------------------------- bank
# header keywords -> standard column; first matching rule wins for each column
_BANK_COLS = [
    ("Date", ("txn date", "transaction date", "tran date", "value date", "date")),
    ("Description", ("description", "narration", "particulars", "remarks", "details")),
    ("Debit", ("debit", "withdrawal", "dr amount", "dr.")),
    ("Credit", ("credit", "deposit", "cr amount", "cr.")),
    ("Cr/Dr", ("cr/dr", "dr/cr", "type")),
    ("Amount", ("amount",)),
    ("Balance", ("balance",)),
    ("No.", ("no.", "sr", "s.no", "sl")),
]


def _map_bank_cols(cols) -> dict:
    out = {}
    for std, keys in _BANK_COLS:
        for c in cols:
            lc = str(c).strip().lower()
            if c not in out and any(lc.startswith(k) or (len(k) > 3 and k in lc) for k in keys):
                out[c] = std
                break
    return out


def _find_header(raw: pd.DataFrame) -> pd.DataFrame:
    """Statements often have a title block above the table; find the header row."""
    for i in range(min(len(raw), 40)):
        m = _map_bank_cols(raw.iloc[i].tolist())
        if {"Date", "Description"} <= set(m.values()) and (
                "Amount" in m.values() or {"Debit", "Credit"} & set(m.values())):
            df = raw.iloc[i + 1:].copy()
            df.columns = [str(c).strip() for c in raw.iloc[i]]
            return df
    df = raw.iloc[1:].copy()  # no recognisable header: assume first row
    df.columns = [str(c).strip() for c in raw.iloc[0]]
    return df


def _num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.replace(r"[^\d.\-]", "", regex=True).replace("", None),
                         errors="coerce")


def load_bank(f) -> pd.DataFrame:
    src = _src(f)
    data = src.read() if hasattr(src, "read") else open(src, "rb").read()
    head = data[:200].lstrip().lower()
    name = str(getattr(f, "name", f)).lower()
    if head.startswith(b"<"):
        tables = pd.read_html(io.StringIO(data.decode("utf-8", "ignore")))
        df = max(tables, key=len)
        if not {"Date", "Description"} <= set(_map_bank_cols(df.columns).values()):
            df = _find_header(pd.concat([pd.DataFrame([list(df.columns)], columns=df.columns), df]))
    elif name.endswith(".csv"):
        df = _find_header(pd.read_csv(io.BytesIO(data), header=None, dtype=str))
    else:
        df = _find_header(pd.read_excel(io.BytesIO(data), header=None))

    colmap = _map_bank_cols(df.columns)
    df = df.rename(columns=colmap)
    df = df.loc[:, ~df.columns.duplicated()]
    missing = [c for c in ("Date", "Description") if c not in df.columns]
    if missing or not ("Amount" in df.columns or {"Debit", "Credit"} <= set(df.columns)):
        raise ValueError("Could not recognise the bank statement columns. Need a date, a description/narration, "
                         "and either Amount + Cr/Dr or Debit + Credit columns. "
                         f"Columns found: {[str(c) for c in df.columns]}")

    if {"Debit", "Credit"} <= set(df.columns):
        dr, cr = _num(df["Debit"]).fillna(0), _num(df["Credit"]).fillna(0)
        df["Amount"] = dr.where(dr > 0, cr)
        df["Cr/Dr"] = pd.Series("CR", index=df.index).where(dr <= 0, "DR")
        df.loc[(dr <= 0) & (cr <= 0), "Amount"] = float("nan")
    else:
        df["Amount"] = _num(df["Amount"])
        if "Cr/Dr" not in df.columns:
            df["Cr/Dr"] = "DR"  # unknown: treat all as debits
    df = df.dropna(subset=["Amount"]).copy()
    df["Date"] = pd.to_datetime(df["Date"], dayfirst=True, errors="coerce", format="mixed")
    df = df.dropna(subset=["Date"])
    df["Balance"] = _num(df["Balance"]) if "Balance" in df.columns else float("nan")
    if "No." not in df.columns:
        df["No."] = range(1, len(df) + 1)
    df["Description"] = df["Description"].astype(str).str.strip()
    df["Cr/Dr"] = df["Cr/Dr"].astype(str).str.strip().str.upper().str.replace(".", "", regex=False)
    df["Payee"] = df["Description"].str.extract(r"^(?:IB)?NEFT/(?:[A-Z]{4}/)?(.+)$")[0]
    return df[["No.", "Date", "Description", "Cr/Dr", "Amount", "Balance", "Payee"]].reset_index(drop=True)


def reconcile_bank(salary: pd.DataFrame, bank: pd.DataFrame, batch_label: str = "SAL"):
    """Returns (salary with payment status, bulk batch summary, bank rows used)."""
    sal = salary.copy()
    sal["Paid Status"], sal["Bank Date"], sal["Bank Ref"], sal["Bank Amount"] = "Not found", pd.NaT, "", float("nan")
    dr = bank[bank["Cr/Dr"] == "DR"].copy()
    used = set()

    # bulk: the whole batch should appear as one (or more) debits
    bulk_total = sal.loc[sal["Mode"] == "BULK", "Salary"].sum()
    batch_rows = dr[dr["Description"].str.contains(batch_label, case=False, na=False)]
    exact = batch_rows[(batch_rows["Amount"] - bulk_total).abs() < 1]
    if exact.empty:  # fall back to any debit with that exact amount
        exact = dr[(dr["Amount"] - bulk_total).abs() < 1]
    bulk_info = {"Sheet bulk total": bulk_total, "Matched": not exact.empty}
    if not exact.empty:
        b = exact.iloc[0]
        used.add(b.name)
        bulk_info.update({"Bank date": b["Date"], "Bank description": b["Description"],
                          "Bank amount": b["Amount"]})
        m = sal["Mode"] == "BULK"
        sal.loc[m, "Paid Status"] = "Paid (bulk batch)"
        sal.loc[m, "Bank Date"] = b["Date"]
        sal.loc[m, "Bank Ref"] = b["Description"]
        sal.loc[m, "Bank Amount"] = sal.loc[m, "Salary"]

    # NEFT: match each person by name + amount
    neft = dr[dr["Payee"].notna()]
    for i, r in sal[sal["Mode"] == "NEFT"].iterrows():
        best, best_s = None, 0.0
        for j, b in neft.iterrows():
            if j in used:
                continue
            s = name_score(r["Name"], b["Payee"])
            if abs(b["Amount"] - r["Salary"]) < 1:
                s += 0.5
            if s > best_s:
                best, best_s = j, s
        if best is not None and best_s >= 0.9:
            b = neft.loc[best]
            used.add(best)
            same_amt = abs(b["Amount"] - r["Salary"]) < 1
            sal.loc[i, ["Paid Status", "Bank Date", "Bank Ref", "Bank Amount"]] = [
                "Paid (NEFT)" if same_amt else "Amount mismatch", b["Date"], b["Description"], b["Amount"]]
    return sal, bulk_info, bank.loc[sorted(used)]


# ---------------------------------------------------------------- production
def _is_csv(f) -> bool:
    return str(getattr(f, "name", f)).lower().endswith(".csv")


def _read_csv_raw(f) -> pd.DataFrame:
    for enc in ("utf-8-sig", "cp1252"):   # Excel "Save as CSV" may use either
        try:
            return pd.read_csv(_src(f), header=None, dtype=object, encoding=enc, skip_blank_lines=False)
        except UnicodeDecodeError:
            continue
    raise ValueError("Could not read the CSV file (unknown text encoding).")


def load_production(f) -> pd.DataFrame:
    """Same report layout as the Excel export (header row starts with 'Category'), as .xlsx or .csv."""
    raw = _read_csv_raw(f) if _is_csv(f) else pd.read_excel(_src(f), header=None)
    hdr = raw.index[raw.iloc[:, 0].astype(str).str.strip().eq("Category")][0]
    df = raw.iloc[hdr + 1:, :8].copy()
    df.columns = ["Category", "Metal", "Employee", "Qty", "Issue Wt", "Receive Wt", "Loss Wt", "PLoss Wt"]
    df[["Category", "Metal"]] = df[["Category", "Metal"]].ffill()
    df = df[df["Employee"].notna()]  # subtotal rows have no employee
    df["Employee"] = df["Employee"].astype(str).str.strip()
    for c in ["Qty", "Issue Wt", "Receive Wt", "Loss Wt", "PLoss Wt"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df.reset_index(drop=True)


def production_summary(prod: pd.DataFrame) -> pd.DataFrame:
    g = prod.groupby("Employee", as_index=False)[["Qty", "Issue Wt", "Receive Wt", "Loss Wt", "PLoss Wt"]].sum()
    g["Loss %"] = (g["Loss Wt"] / g["Issue Wt"] * 100).round(2)
    return g.sort_values("Receive Wt", ascending=False).reset_index(drop=True)


def map_names(prod_names, salary_names, threshold=0.8) -> pd.DataFrame:
    rows = []
    for p in prod_names:
        m, s = best_match(p, salary_names, threshold)
        rows.append({"Production Name": p, "Salary Name": m or "", "Score": s})
    return pd.DataFrame(rows)


def employee_cost_detail(prod: pd.DataFrame, mapping: pd.DataFrame, salary: pd.DataFrame,
                         basis: str = "Issue Wt") -> pd.DataFrame:
    """Per employee x metal x category: pcs, issue wt, and salary allocated
    to each row in proportion to `basis` ("Issue Wt" or "Qty")."""
    m = mapping[mapping["Salary Name"].fillna("") != ""][["Production Name", "Salary Name"]]
    d = prod.merge(m, left_on="Employee", right_on="Production Name")
    d = (d.groupby(["Salary Name", "Metal", "Category"], as_index=False)[["Qty", "Issue Wt", "Receive Wt"]].sum()
         .rename(columns={"Salary Name": "Employee", "Qty": "Pcs"}))
    col = "Pcs" if basis == "Qty" else basis
    d["Share %"] = d[col] / d.groupby("Employee")[col].transform("sum") * 100
    d = d.merge(salary[["Name", "Salary"]], left_on="Employee", right_on="Name").drop(columns="Name")
    d["Cost"] = d["Salary"] * d["Share %"] / 100
    d["Cost per pc"] = d["Cost"] / d["Pcs"].where(d["Pcs"] > 0)
    d["Cost per g"] = d["Cost"] / d["Issue Wt"].where(d["Issue Wt"] > 0)
    return d.sort_values(["Employee", "Metal", "Category"]).reset_index(drop=True)


def cost_rollup(detail: pd.DataFrame, by: list[str]) -> pd.DataFrame:
    g = detail.groupby(by, as_index=False)[["Pcs", "Issue Wt", "Cost"]].sum()
    g["Cost per pc"] = g["Cost"] / g["Pcs"].where(g["Pcs"] > 0)
    g["Cost per g"] = g["Cost"] / g["Issue Wt"].where(g["Issue Wt"] > 0)
    return g
