import sys
import os
import requests
import pandas as pd
import json
import datetime as dt
from pathlib import Path
import numpy as np
import traceback

def make_logger(base_path: Path):
    log_path = base_path / "app_log.txt"
    def log(msg, err=False):
        timestamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{timestamp}] {msg}"
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass
        print(line, flush=True)
        if err:
            pass
    return log

if getattr(sys, "frozen", False):
    BASE_PATH = Path(sys.executable).parent
else:
    BASE_PATH = Path(__file__).parent

log = make_logger(BASE_PATH)
log("Starting mf_rolling_returns script.")

json_path = BASE_PATH / "scheme_codes.json"
if not json_path.exists():
    log(f"ERROR: scheme_codes.json not found at {json_path}", err=True)
    raise FileNotFoundError(f"scheme_codes.json not found at {json_path}")

with open(json_path, "r", encoding="utf-8") as f:
    scheme_codes = json.load(f)

log(f"Loaded scheme_codes.json ({len(scheme_codes)} AMCs).")

def get_nav_data(scheme_code):
    url = f"https://api.mfapi.in/mf/{scheme_code}"
    resp = requests.get(url, timeout=30)
    if resp.status_code != 200:
        raise RuntimeError(f"HTTP {resp.status_code} for code {scheme_code}")
    data = resp.json().get("data", [])
    if not data:
        raise RuntimeError(f"No data returned for scheme {scheme_code}")
    df = pd.DataFrame(data)
    # parse date & nav, handle bad rows robustly
    df["date"] = pd.to_datetime(df["date"], dayfirst=True, errors="coerce")
    df["nav"] = pd.to_numeric(df["nav"], errors="coerce")
    df = df.dropna(subset=["date", "nav"]).sort_values("date").reset_index(drop=True)
    return df

def last_nav_on_or_before(df, target_date: dt.date):
    df_before = df[df["date"].dt.date <= target_date]
    if df_before.empty:
        return None  
    row = df_before.iloc[-1]  
    return {"date": row["date"].date(), "nav": float(row["nav"])}

def monthly_rolling_return(df):
    today = dt.date.today()
    last_day_prev_month = (today.replace(day=1) - dt.timedelta(days=1))
    last_day_two_months_ago = (last_day_prev_month.replace(day=1) - dt.timedelta(days=1))

    t1 = last_nav_on_or_before(df, last_day_prev_month)
    t0 = last_nav_on_or_before(df, last_day_two_months_ago)

    if t0 is None or t1 is None:
        return None

    start_nav = t0["nav"]
    end_nav = t1["nav"]
    days = (t1["date"] - t0["date"]).days
    if days <= 0 or start_nav == 0:
        return None

    rr = ((end_nav - start_nav) / start_nav) * (365.0 / days)
    return rr  

def build_results(scheme_codes, logger):
    results = {}  
    total_amcs = len(scheme_codes)
    total_tasks = sum(len(cat_dict.keys()) * 2 for cat_dict in scheme_codes.values())
    task_counter = 0

    for amc, cat_dict in scheme_codes.items():
        logger(f"Processing AMC: {amc}")
        for category, plans in cat_dict.items():
            task_counter += 0  
            if category not in results:
                results[category] = []
            row = {"AMC": amc}
            for plan in ["Direct", "Regular"]:
                task_counter += 1
                code = plans.get(plan)
                if code is None:
                    logger(f"[{task_counter}/{total_tasks}] {amc} - {category} - {plan}: no scheme code, skipping")
                    row[f"Rolling Return - {plan}"] = None
                    continue
                logger(f"[{task_counter}/{total_tasks}] Fetching {amc} - {category} - {plan} (code={code}) ...")
                try:
                    df_nav = get_nav_data(code)
                    rr = monthly_rolling_return(df_nav)
                    row[f"Rolling Return - {plan}"] = rr
                    if rr is None:
                        logger(f"[{task_counter}/{total_tasks}] {amc} - {category} - {plan}: insufficient NAV history for t0/t1")
                    else:
                        logger(f"[{task_counter}/{total_tasks}] {amc} - {category} - {plan}: RR={rr:.6f}")
                except Exception as e:
                    logger(f"[{task_counter}/{total_tasks}] ERROR fetching {amc} - {category} - {plan}: {e}")
                    logger(traceback.format_exc())
                    row[f"Rolling Return - {plan}"] = None
            results[category].append(row)
    return results

def export_to_excel(results: dict, base_path: Path, logger):
    output_dir = base_path / "outputs"
    output_dir.mkdir(exist_ok=True)
    today = dt.date.today()
    report_month = (today.replace(day=1) - dt.timedelta(days=1)).strftime("%b-%Y")
    filename = output_dir / f"Rolling_Returns_{report_month}.xlsx"

    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        workbook = writer.book

        header_fmt = workbook.add_format({
            "bold": True, "text_wrap": True, "valign": "top", "align": "center",
            "fg_color": "#D7E4BC", "border": 1
        })
        percent_fmt = workbook.add_format({"num_format": "0.00%", "border": 1, "align": "right"})
        normal_fmt = workbook.add_format({"border": 1, "align": "left"})
        thin_border_fmt = workbook.add_format({"border": 1})
        thick_outer_fmt = workbook.add_format({"top": 2, "bottom": 2, "left": 2, "right": 2})

        padding_row_height = 6   
        padding_col_width = 2.0 

        for category, rows in results.items():
            df = pd.DataFrame(rows)
            if df.empty:
                logger(f"No rows for category {category}, skipping sheet.")
                continue

            expected_cols = ["AMC", "Rolling Return - Direct", "Rolling Return - Regular"]
            for c in expected_cols:
                if c not in df.columns:
                    df[c] = None
            df = df[expected_cols]

            df.to_excel(writer, sheet_name=category, index=False, startrow=1, startcol=1)
            worksheet = writer.sheets[category]

            worksheet.set_row(0, padding_row_height)
            worksheet.set_column(0, 0, padding_col_width)

            for col_idx, col_name in enumerate(df.columns, start=1):
                worksheet.write(1, col_idx, col_name, header_fmt)

            col_widths = {col: len(str(col)) for col in df.columns}
            for r_idx, row in enumerate(rows, start=2):  
                for c_idx, col in enumerate(df.columns, start=1):
                    raw_val = row.get(col)
                    if col.startswith("Rolling Return"):
                        if raw_val is None or (isinstance(raw_val, float) and (np.isnan(raw_val) or np.isinf(raw_val))):
                            worksheet.write(r_idx, c_idx, None, normal_fmt)
                            display_text = ""
                        else:
                            worksheet.write(r_idx, c_idx, float(raw_val), percent_fmt)
                            display_text = f"{raw_val * 100:.2f}%"
                    else:
                        if raw_val is None:
                            worksheet.write(r_idx, c_idx, "", normal_fmt)
                            display_text = ""
                        else:
                            worksheet.write(r_idx, c_idx, str(raw_val), normal_fmt)
                            display_text = str(raw_val)
                    col_widths[col] = max(col_widths[col], len(display_text))

            n_rows = len(df) + 1  
            n_cols = len(df.columns)
            worksheet.conditional_format(1, 1, n_rows + 1, n_cols, {"type": "no_errors", "format": thin_border_fmt})

            worksheet.conditional_format(1, 1, n_rows + 1, n_cols, {"type": "no_errors", "format": thick_outer_fmt})

            for idx, col in enumerate(df.columns, start=1):
                width = col_widths.get(col, len(col)) + 2
                if width < 8:
                    width = 8
                worksheet.set_column(idx, idx, width)

            logger(f"Wrote sheet '{category}' with {len(df)} rows.")

    logger(f"Excel file written: {filename}")
    return filename

def run_all():
    try:
        results = build_results(scheme_codes, log)
        if not results:
            log("No results produced. Exiting.", err=True)
            return
        xlsx_path = export_to_excel(results, BASE_PATH, log)
        log(f"SUCCESS: Report generated at {xlsx_path}")
    except Exception as ex:
        log("FATAL ERROR: " + str(ex), err=True)
        log(traceback.format_exc(), err=True)

if __name__ == "__main__":
    run_all()
