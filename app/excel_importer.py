"""
Module for parsing incoming Excel/CSV files, performing smart auto-matching,
data normalization (Persian dates, Rial/Toman, Persian digits),
automatic unmapped columns concatenation into notes, and specialized preview generation.
"""

import os
import re
import io
import openpyxl
import pandas as pd
from typing import Dict, List, Any, Optional
from app.schemas import TARGET_SCHEMAS, BANK_OPERATIONS


def to_english_digits(text: Any) -> str:
    """Converts Persian and Arabic digits to ASCII digits."""
    if text is None:
        return ""
    s = str(text)
    persian_digits = "۰۱۲۳۴۵۶۷۸۹"
    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    for i in range(10):
        s = s.replace(persian_digits[i], str(i)).replace(arabic_digits[i], str(i))
    return s.strip()


def normalize_persian_date(date_val: Any) -> str:
    """Standardizes various Persian date formats to YYYY/MM/DD."""
    if not date_val:
        return ""
    clean = to_english_digits(date_val).replace("-", "/").replace(".", "/").replace("\\", "/")
    
    # Check if purely digits like 14050701
    num_match = re.match(r"^(\d{4})(\d{2})(\d{2})$", clean)
    if num_match:
        return f"{num_match.group(1)}/{num_match.group(2)}/{num_match.group(3)}"
    
    # Check if separated with slashes
    slash_match = re.match(r"^(\d{2,4})/(\d{1,2})/(\d{1,2})$", clean)
    if slash_match:
        year = slash_match.group(1)
        month = slash_match.group(2).zfill(2)
        day = slash_match.group(3).zfill(2)
        if len(year) == 2:
            year = "14" + year
        return f"{year}/{month}/{day}"
    return clean


def parse_uploaded_file(file_bytes: bytes, filename: str) -> Dict[str, Any]:
    """Parses an uploaded file (XLSX, XLS, CSV) and returns sheet names, columns, and rows."""
    ext = os.path.splitext(filename)[1].lower()
    sheets_data = {}
    
    if ext in [".xlsx", ".xlsm"]:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        for sname in wb.sheetnames:
            ws = wb[sname]
            all_rows = list(ws.iter_rows(values_only=True))
            if not all_rows:
                sheets_data[sname] = {"columns": [], "rows": [], "total_rows": 0}
                continue
            
            # Find the best header row
            header_idx = 0
            best_score = 0
            for idx in range(min(5, len(all_rows))):
                row = all_rows[idx]
                score = sum(1 for c in row if c is not None and str(c).strip() != "")
                if score > best_score:
                    best_score = score
                    header_idx = idx
            
            header_row = all_rows[header_idx]
            columns = []
            for col_i, c in enumerate(header_row):
                val = str(c).strip() if c is not None else ""
                col_name = val if val else f"ستون_{col_i+1}"
                original = col_name
                counter = 1
                while col_name in columns:
                    col_name = f"{original}_{counter}"
                    counter += 1
                columns.append(col_name)
            
            data_rows = []
            for r_idx in range(header_idx + 1, len(all_rows)):
                r = all_rows[r_idx]
                if any(v is not None and str(v).strip() != "" for v in r):
                    row_dict = {}
                    for col_i, col_name in enumerate(columns):
                        val = r[col_i] if col_i < len(r) else None
                        row_dict[col_name] = val
                    data_rows.append(row_dict)
            
            sheets_data[sname] = {
                "columns": columns,
                "rows": data_rows,
                "total_rows": len(data_rows)
            }
            
    elif ext == ".csv":
        df = None
        for enc in ["utf-8-sig", "utf-8", "cp1256", "latin1"]:
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc)
                break
            except Exception:
                continue
        if df is None:
            raise ValueError("قالب فایل CSV قابل خواندن نبود.")
        
        columns = [str(c).strip() for c in df.columns]
        df = df.where(pd.notnull(df), None)
        data_rows = df.to_dict(orient="records")
        sheets_data["Sheet1"] = {
            "columns": columns,
            "rows": data_rows,
            "total_rows": len(data_rows)
        }
    else:
        raise ValueError("فرمت فایل نامعتبر است. لطفاً فایل اکسل (.xlsx) یا (.csv) انتخاب فرمایید.")
    
    first_sheet = list(sheets_data.keys())[0] if sheets_data else "Sheet1"
    return {
        "filename": filename,
        "sheet_names": list(sheets_data.keys()),
        "current_sheet": first_sheet,
        "sheets": sheets_data
    }


def auto_match_columns(source_columns: List[str], target_module_id: str) -> Dict[str, str]:
    """Computes best auto-matching between source columns and target schema columns."""
    schema = TARGET_SCHEMAS.get(target_module_id)
    if not schema:
        return {}
    
    mapping = {}
    used_sources = set()
    
    # 1. Exact match pass
    for t_col in schema["columns"]:
        t_key = t_col["key"]
        t_title = t_col["title"].lower()
        for s_col in source_columns:
            if s_col in used_sources:
                continue
            s_clean = s_col.strip().lower()
            if s_clean == t_title or s_clean in t_title or t_title in s_clean:
                mapping[t_key] = s_col
                used_sources.add(s_col)
                break
    
    # 2. Keyword matching pass
    for t_col in schema["columns"]:
        t_key = t_col["key"]
        if t_key in mapping and mapping[t_key]:
            continue
        
        keywords = t_col.get("keywords", [])
        best_match = None
        best_score = 0
        
        for s_col in source_columns:
            if s_col in used_sources:
                continue
            s_clean = s_col.strip().lower()
            for kw in keywords:
                kw_clean = kw.lower()
                if kw_clean in s_clean:
                    score = len(kw_clean) / max(len(s_clean), 1) * 10
                    if score > best_score:
                        best_score = score
                        best_match = s_col
        
        if best_match:
            mapping[t_key] = best_match
            used_sources.add(best_match)
        else:
            mapping[t_key] = ""
            
    return mapping


def generate_mapped_preview(
    raw_rows: List[Dict[str, Any]],
    target_module_id: str,
    column_mapping: Dict[str, str],
    custom_defaults: Dict[str, Any],
    options: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Transforms raw source rows into target schema rows according to column_mapping.
    Supports single-column vs dual-column bank amounts, automatically concatenates
    all unmapped columns into notes, and calculates real-time net weights and amounts.
    """
    schema = TARGET_SCHEMAS.get(target_module_id)
    if not schema:
        raise ValueError(f"ماژول نامعتبر: {target_module_id}")
    
    normalize_dates = options.get("normalize_dates", True)
    toman_to_rial = options.get("toman_to_rial", False)
    clean_persian_digits = options.get("clean_digits", True)
    concat_unmapped = options.get("concat_unmapped", True)
    bank_amount_mode = options.get("bank_amount_mode", "dual")  # 'dual' or 'single'
    single_amount_col = options.get("single_amount_col", "")
    direction_col = options.get("direction_col", "")
    selected_bank_code = options.get("selected_bank", "BANK01")
    
    # Identify all mapped source columns to determine which are UNMAPPED
    mapped_source_cols = set()
    for col_name in column_mapping.values():
        if col_name and str(col_name).strip():
            mapped_source_cols.add(str(col_name).strip())
    if bank_amount_mode == "single":
        if single_amount_col: mapped_source_cols.add(single_amount_col)
        if direction_col: mapped_source_cols.add(direction_col)
        
    mapped_rows = []
    total_warnings = 0
    total_deposit = 0.0
    total_withdrawal = 0.0
    total_weight = 0.0
    
    for row_idx, src_row in enumerate(raw_rows):
        mapped_item = {
            "_id": row_idx + 1,
            "_selected": True,
            "_warnings": [],
            "_bank_code": selected_bank_code
        }
        
        # 1. Process regular columns from schema
        for t_col in schema["columns"]:
            t_key = t_col["key"]
            source_col_name = column_mapping.get(t_key)
            default_val = custom_defaults.get(t_key, t_col.get("default", ""))
            
            raw_val = None
            if source_col_name and source_col_name in src_row:
                raw_val = src_row[source_col_name]
            
            final_val = default_val if (raw_val is None or str(raw_val).strip() == "") else raw_val
            col_type = t_col.get("type", "text")
            
            # Persian digit cleanup
            if clean_persian_digits and final_val is not None and isinstance(final_val, str):
                final_val = to_english_digits(final_val)
            
            # Date formatting
            if col_type == "date":
                if normalize_dates and final_val:
                    final_val = normalize_persian_date(final_val)
                if t_col.get("required") and (not final_val or len(str(final_val)) < 8):
                    mapped_item["_warnings"].append(f"تاریخ {t_col['title']} ناقص است")
            
            # Number formatting
            elif col_type == "number":
                if final_val is not None and str(final_val).strip() != "":
                    try:
                        num_str = to_english_digits(str(final_val)).replace(",", "").replace("ریال", "").replace("تومان", "").strip()
                        num_val = float(num_str)
                        if toman_to_rial and ("مبلغ" in t_col["title"] or "ریال" in t_col["title"]):
                            num_val = num_val * 10
                        final_val = int(num_val) if num_val.is_integer() else round(num_val, 2)
                    except Exception:
                        mapped_item["_warnings"].append(f"مقدار غیرعددی: {final_val}")
                else:
                    final_val = 0 if t_col.get("required") else ""
            
            mapped_item[t_key] = final_val

        # 2. Bank Special Amount Mode Handling (Single amount column + Direction)
        if target_module_id == "transactions" and bank_amount_mode == "single" and single_amount_col:
            raw_amt_val = src_row.get(single_amount_col, 0)
            raw_dir_val = str(src_row.get(direction_col, "")).lower() if direction_col else ""
            
            try:
                amt_str = to_english_digits(str(raw_amt_val)).replace(",", "").strip()
                amt_num = float(amt_str) if amt_str else 0.0
                if toman_to_rial: amt_num *= 10
            except Exception:
                amt_num = 0.0
                mapped_item["_warnings"].append("مبلغ تراکنش تک‌ستونی قابل خواندن نبود")
            
            # Check direction / sign
            is_deposit = False
            is_withdrawal = False
            
            deposit_keywords = ["واریز", "بستانکار", "ورود", "+", "cr", "credit", "deposit", "دریافت", "وارده"]
            withdrawal_keywords = ["برداشت", "بدهکار", "خروج", "-", "dr", "debit", "withdraw", "پرداخت", "صادره"]
            
            if any(k in raw_dir_val for k in deposit_keywords):
                is_deposit = True
            elif any(k in raw_dir_val for k in withdrawal_keywords):
                is_withdrawal = True
            elif amt_num < 0:
                is_withdrawal = True
                amt_num = abs(amt_num)
            else:
                is_deposit = True  # default positive to deposit
                
            if is_deposit:
                mapped_item["col13"] = int(amt_num) if amt_num.is_integer() else amt_num
                mapped_item["col14"] = 0
                mapped_item["_dir"] = "واریز"
            else:
                mapped_item["col13"] = 0
                mapped_item["col14"] = int(amt_num) if amt_num.is_integer() else amt_num
                mapped_item["_dir"] = "برداشت"

        # 3. Automatic Unmapped Columns Concatenation into "توضیحات"
        if concat_unmapped:
            unmapped_parts = []
            for s_col, s_val in src_row.items():
                if s_col not in mapped_source_cols and s_val is not None:
                    s_str = str(s_val).strip()
                    if s_str and s_str != "0" and s_str.lower() != "nan":
                        unmapped_parts.append(f"[{s_col}: {s_str}]")
            
            if unmapped_parts:
                notes_key = "col18" if target_module_id == "transactions" else ("col15" if target_module_id == "baskol" else "col18")
                base_notes = str(mapped_item.get(notes_key, "") or "").strip()
                extra_notes = " | ".join(unmapped_parts)
                combined = f"{base_notes} | {extra_notes}" if base_notes else extra_notes
                mapped_item[notes_key] = combined

        # 4. Smart Operation and Account Defaulting for Transactions
        if target_module_id == "transactions":
            dep_val = float(mapped_item.get("col13") or 0)
            wit_val = float(mapped_item.get("col14") or 0)
            total_deposit += dep_val
            total_withdrawal += wit_val
            
            # Suggest operation based on direction
            if wit_val > 0 and dep_val == 0:
                # If notes mentions personal manager expenses or Qasem Zahedi
                notes_text = str(mapped_item.get("col18", "")).lower()
                if "قاسم" in notes_text or "زاهدی" in notes_text or "شخصی" in notes_text:
                    mapped_item["col10"] = "هزینه‌های شخصی مدیر (قاسم زاهدی) - جاری شرکا"
                    mapped_item["col7"] = "CAP002"
                elif "کارمزد" in notes_text or "fee" in notes_text:
                    mapped_item["col10"] = "کارمزد و هزینه‌های بانکی"
                    mapped_item["col7"] = "X001"
                elif "باغدار" in notes_text or "خرید" in notes_text:
                    mapped_item["col10"] = "خرید مرکبات / تسویه حساب باغدار"
                    if not mapped_item.get("col7") or mapped_item["col7"] == "C001":
                        mapped_item["col7"] = "S001"
            elif dep_val > 0:
                notes_text = str(mapped_item.get("col18", "")).lower()
                if ("اموال" in notes_text or "زمین" in notes_text or "ملک" in notes_text) and ("قاسم" in notes_text or "زاهدی" in notes_text):
                    mapped_item["col10"] = "فروش اموال مدیر (قاسم زاهدی) - ورود به چرخه شرکت"
                    mapped_item["col7"] = "CAP003"
                elif "پوز" in notes_text or "pos" in notes_text or "کارتخوان" in notes_text:
                    mapped_item["col10"] = "سرجمع دستگاه پوز (POS)"
                    mapped_item["col7"] = "X004"

        # 5. Live Net Weight Calculation for Baskol
        if target_module_id == "baskol":
            gross = float(mapped_item.get("col11") or 0)
            baskets = float(mapped_item.get("col12") or 0)
            tare_weight = baskets * 1.5
            net_before_drop = max(0, gross - tare_weight)
            net_final = round(net_before_drop * 0.99, 1)  # 1% drop deduction
            mapped_item["_net_calculated"] = net_final
            mapped_item["_tare_calculated"] = tare_weight
            total_weight += gross

        if mapped_item["_warnings"]:
            total_warnings += len(mapped_item["_warnings"])
            
        mapped_rows.append(mapped_item)
        
    return {
        "target_module": target_module_id,
        "target_title": schema["title"],
        "schema_columns": schema["columns"],
        "mapped_rows": mapped_rows,
        "bank_operations": BANK_OPERATIONS,
        "summary": {
            "total_rows": len(mapped_rows),
            "total_deposit": total_deposit,
            "total_withdrawal": total_withdrawal,
            "total_weight": total_weight,
            "total_warnings": total_warnings
        }
    }
