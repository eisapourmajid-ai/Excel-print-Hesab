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
from typing import Dict, List, Any, Optional, Tuple
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


def split_persian_datetime(val: Any) -> Tuple[str, str]:
    """
    Splits composite datetime into (date_str, time_str).
    Handles formats like:
    - '1405/06/04 11:29'
    - '1405/06/04 11:29:45'
    - '1405-06-04 11:29'
    - '1405/06/04T11:29'
    - openpyxl / pandas datetime timestamps
    """
    if val is None or str(val).strip() == "":
        return "", ""
    
    # Check if object is a datetime
    if isinstance(val, (pd.Timestamp, openpyxl.packaging.core.datetime.datetime)):
        try:
            time_part = val.strftime("%H:%M")
            date_part = str(val.date()).replace("-", "/")
            return date_part, time_part
        except Exception:
            pass

    s = to_english_digits(str(val)).strip()
    
    # Extract time pattern HH:MM or HH:MM:SS
    time_str = ""
    time_match = re.search(r"(\b\d{1,2}:\d{2}(?::\d{2})?\b)", s)
    if time_match:
        raw_time = time_match.group(1)
        time_parts = raw_time.split(":")
        hh = time_parts[0].zfill(2)
        mm = time_parts[1].zfill(2)
        time_str = f"{hh}:{mm}"
        # Remove time from s to isolate date
        s = s[:time_match.start()] + " " + s[time_match.end():]
        s = s.strip()
    
    # Clean date
    s = s.replace("-", "/").replace(".", "/").replace("\\", "/").strip()
    date_str = ""
    date_match = re.search(r"(\b\d{2,4}/\d{1,2}/\d{1,2}\b)", s)
    if date_match:
        parts = date_match.group(1).split("/")
        year = parts[0]
        month = parts[1].zfill(2)
        day = parts[2].zfill(2)
        if len(year) == 2:
            year = "14" + year
        date_str = f"{year}/{month}/{day}"
    else:
        num_m = re.search(r"\b(\d{4})(\d{2})(\d{2})\b", s)
        if num_m:
            date_str = f"{num_m.group(1)}/{num_m.group(2)}/{num_m.group(3)}"
        elif not time_str:
            date_str = s

    return date_str, time_str


def normalize_persian_date(date_val: Any) -> str:
    """Standardizes various Persian date formats to YYYY/MM/DD."""
    if not date_val:
        return ""
    d, _ = split_persian_datetime(date_val)
    return d if d else to_english_digits(str(date_val)).strip()


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


def auto_match_columns(source_columns: List[str], target_module_id: str) -> Dict[str, Any]:
    """
    Computes best auto-matching between source columns and target schema columns.
    Also detects single-column amount modes (type column + amount column) and composite date/time.
    """
    schema = TARGET_SCHEMAS.get(target_module_id)
    if not schema:
        return {"mapping": {}, "detected_mode": "dual"}
    
    mapping = {}
    used_sources = set()
    
    # Specific Bank Detection: check if there's 'نوع تراکنش' and 'مبلغ تراکنش'
    detected_mode = "dual"
    suggested_single_col = ""
    suggested_dir_col = ""
    
    if target_module_id == "transactions":
        has_type_col = any("نوع" in c and "تراکنش" in c for c in source_columns) or any(c in ["نوع", "ماهیت", "نوع تراکنش", "وضعیت تراکنش"] for c in source_columns)
        has_amount_col = any("مبلغ" in c for c in source_columns)
        has_sep_dep = any("واریز" in c for c in source_columns)
        has_sep_wit = any("برداشت" in c for c in source_columns)
        
        if (has_type_col or not (has_sep_dep and has_sep_wit)) and has_amount_col:
            detected_mode = "single"
            for c in source_columns:
                if "مبلغ" in c or "amount" in c.lower():
                    suggested_single_col = c
                    break
            for c in source_columns:
                if ("نوع" in c and "تراکنش" in c) or c in ["نوع", "ماهیت", "نوع تراکنش", "بدهکار/بستانکار"]:
                    suggested_dir_col = c
                    break
    
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
            
    return {
        "mapping": mapping,
        "detected_mode": detected_mode,
        "suggested_single_col": suggested_single_col,
        "suggested_dir_col": suggested_dir_col
    }


def generate_mapped_preview(
    raw_rows: List[Dict[str, Any]],
    target_module_id: str,
    column_mapping: Dict[str, str],
    custom_defaults: Dict[str, Any],
    options: Dict[str, Any]
) -> Dict[str, Any]:
    """
    Transforms raw source rows into target schema rows according to column_mapping.
    Supports:
    - Auto splitting combined Date & Time columns (e.g. 1405/06/04 11:29)
    - Mode A: Dual column (واریز / برداشت)
    - Mode B: Amount column + Type column (نوع تراکنش + مبلغ تراکنش)
    - Mode C: Single signed column (+/-)
    - Automatically concatenates all unmapped columns into notes.
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
    auto_split_datetime = options.get("auto_split_datetime", True)
    
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
            
            # Date formatting and Date/Time separation
            if col_type == "date":
                if auto_split_datetime and raw_val:
                    d_part, t_part = split_persian_datetime(raw_val)
                    if d_part:
                        final_val = d_part
                    if t_part:
                        # Store extracted time for time column if not already populated
                        mapped_item["_extracted_time"] = t_part
                elif normalize_dates and final_val:
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

        # Auto-fill Time from Extracted Datetime if available
        if auto_split_datetime and "_extracted_time" in mapped_item:
            time_val = mapped_item.get("_extracted_time")
            time_key = "col4" if target_module_id == "transactions" else "col3"
            time_src = column_mapping.get(time_key, "")
            # If time column was unmapped, or mapped to same col as date, or empty/default
            if (not time_src or time_src == column_mapping.get("col3") or not mapped_item.get(time_key) or mapped_item.get(time_key) in ["10:00", "08:30"]):
                mapped_item[time_key] = time_val

        # 2. Bank Special Amount Mode Handling (Single amount column + Type / Direction)
        if target_module_id == "transactions" and bank_amount_mode == "single" and single_amount_col:
            raw_amt_val = src_row.get(single_amount_col, 0)
            raw_dir_val = str(src_row.get(direction_col, "")).lower() if direction_col else ""
            
            try:
                amt_str = to_english_digits(str(raw_amt_val)).replace(",", "").replace("ریال", "").replace("تومان", "").strip()
                amt_num = float(amt_str) if amt_str else 0.0
                if toman_to_rial:
                    amt_num *= 10
            except Exception:
                amt_num = 0.0
                mapped_item["_warnings"].append("مبلغ تراکنش تک‌ستونی قابل خواندن نبود")
            
            # Check direction / sign
            is_deposit = False
            is_withdrawal = False
            
            deposit_keywords = ["واریز", "بستانکار", "ورود", "+", "cr", "credit", "deposit", "دریافت", "وارده"]
            withdrawal_keywords = ["برداشت", "بدهکار", "خروج", "-", "dr", "debit", "withdraw", "پرداخت", "صادره"]
            
            if any(k in raw_dir_val for k in withdrawal_keywords):
                is_withdrawal = True
            elif any(k in raw_dir_val for k in deposit_keywords):
                is_deposit = True
            elif amt_num < 0:
                is_withdrawal = True
                amt_num = abs(amt_num)
            else:
                # Default positive to deposit
                is_deposit = True
                
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
            
            notes_text = str(mapped_item.get("col18", "")).lower()
            
            # Suggest operation based on direction & description content
            if wit_val > 0 and dep_val == 0:
                if "قاسم" in notes_text or "زاهدی" in notes_text or "شخصی" in notes_text:
                    mapped_item["col10"] = "هزینه‌های شخصی مدیر (قاسم زاهدی) - جاری شرکا"
                    mapped_item["col7"] = "CAP002"
                elif "کارمزد" in notes_text or "fee" in notes_text:
                    mapped_item["col10"] = "کارمزد و هزینه‌های بانکی"
                    mapped_item["col7"] = "X001"
                elif "باغدار" in notes_text or "خرید" in notes_text or "تامسون" in notes_text:
                    mapped_item["col10"] = "خرید مرکبات / تسویه حساب باغدار"
                    if not mapped_item.get("col7") or mapped_item["col7"] == "C001":
                        mapped_item["col7"] = "S001"
                else:
                    mapped_item["col10"] = "هزینه‌های جاری و سورتینگ"
            elif dep_val > 0:
                if ("اموال" in notes_text or "زمین" in notes_text or "ملک" in notes_text) and ("قاسم" in notes_text or "زاهدی" in notes_text):
                    mapped_item["col10"] = "فروش اموال مدیر (قاسم زاهدی) - ورود به چرخه شرکت"
                    mapped_item["col7"] = "CAP003"
                elif "پوز" in notes_text or "pos" in notes_text or "کارتخوان" in notes_text:
                    mapped_item["col10"] = "سرجمع دستگاه پوز (POS)"
                    mapped_item["col7"] = "X004"
                elif "میدان" in notes_text or "بارفروش" in notes_text or "حجره" in notes_text:
                    mapped_item["col10"] = "واریز بارفروشان (دریافت بابت فروش مرکبات)"
                    if not mapped_item.get("col7") or mapped_item["col7"] == "S001":
                        mapped_item["col7"] = "C001"

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
