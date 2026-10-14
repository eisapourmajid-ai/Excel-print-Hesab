"""
FastAPI application for Mazandasht Excel Import, Column Sync, and Safe Merging Platform.
"""

import os
import io
import shutil
import openpyxl
from datetime import datetime
from typing import Dict, Any, List, Optional
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Query
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from app.schemas import TARGET_SCHEMAS, BANK_OPERATIONS
from app.excel_importer import (
    parse_uploaded_file,
    auto_match_columns,
    generate_mapped_preview
)
from app.excel_merger import merge_rows_into_system_excel

app = FastAPI(
    title="Mazandasht Excel Sync & Merge Studio",
    description="سامانه هوشمند ورود، تطبیق و ادغام فایل‌های اکسل سورتینگ مازندشت",
    version="1.2.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

os.makedirs("app/static", exist_ok=True)
os.makedirs("app/templates", exist_ok=True)
os.makedirs("backups", exist_ok=True)
os.makedirs("samples", exist_ok=True)

app.mount("/static", StaticFiles(directory="app/static"), name="static")


class AutoMatchRequest(BaseModel):
    source_columns: List[str]
    target_module: str


class PreviewMappedRequest(BaseModel):
    raw_rows: List[Dict[str, Any]]
    target_module: str
    column_mapping: Dict[str, str]
    custom_defaults: Optional[Dict[str, Any]] = {}
    options: Optional[Dict[str, Any]] = {
        "normalize_dates": True,
        "clean_digits": True,
        "toman_to_rial": False,
        "concat_unmapped": True,
        "bank_amount_mode": "dual",
        "single_amount_col": "",
        "direction_col": "",
        "selected_bank": "BANK01"
    }


class MergeRequest(BaseModel):
    target_module: str
    rows: List[Dict[str, Any]]
    user_name: Optional[str] = "کاربر سیستم"


class BankItem(BaseModel):
    code: str
    title: str
    account_no: Optional[str] = ""
    notes: Optional[str] = ""
    opening_balance: Optional[float] = 0
    active: Optional[bool] = True


class SaveBanksRequest(BaseModel):
    banks: List[BankItem]


@app.api_route("/", methods=["GET", "HEAD"], response_class=HTMLResponse)
async def serve_index():
    with open("app/templates/index.html", "r", encoding="utf-8") as f:
        return f.read()


@app.api_route("/health", methods=["GET", "HEAD"])
async def health_check():
    return {"status": "ok"}


@app.get("/api/schemas")
async def get_schemas():
    """Returns available destination schemas and bank operations."""
    return {
        "schemas": TARGET_SCHEMAS,
        "bank_operations": BANK_OPERATIONS
    }


@app.get("/api/accounts-list")
async def get_accounts_list():
    """Returns active accounts chart from INFO.xlsx."""
    if not os.path.exists("INFO.xlsx"):
        return {"accounts": []}
    wb = openpyxl.load_workbook("INFO.xlsx", data_only=True)
    ws = wb["Accounts"]
    accounts = []
    for r in range(7, ws.max_row + 1):
        code = ws.cell(r, 1).value
        nature = ws.cell(r, 2).value
        sub = ws.cell(r, 3).value
        name = ws.cell(r, 4).value
        if code and name:
            accounts.append({
                "code": str(code).strip(),
                "name": str(name).strip(),
                "nature": str(nature).strip() if nature else "",
                "subgroup": str(sub).strip() if sub else ""
            })
    return {"accounts": accounts}


@app.get("/api/banks-list")
async def get_banks_list():
    """Returns company bank accounts from INFO.xlsx including Account No / IBAN."""
    if not os.path.exists("INFO.xlsx"):
        return {"banks": []}
    wb = openpyxl.load_workbook("INFO.xlsx", data_only=True)
    ws = wb["Banks"]
    banks = []
    for r in range(7, 17):
        code = ws.cell(r, 1).value
        title = ws.cell(r, 2).value
        bal = ws.cell(r, 3).value
        active = ws.cell(r, 4).value
        notes = ws.cell(r, 5).value
        account_no = ws.cell(r, 6).value if ws.max_column >= 6 else ""
        if code:
            banks.append({
                "code": str(code).strip(),
                "title": str(title).strip() if title else "",
                "opening_balance": bal or 0,
                "active": str(active).strip() == "بله",
                "notes": str(notes).strip() if notes else "",
                "account_no": str(account_no).strip() if account_no else ""
            })
    return {"banks": banks}


@app.post("/api/save-banks")
async def save_banks(req: SaveBanksRequest):
    """Saves updated bank list into INFO.xlsx Banks sheet with automated backup."""
    if not os.path.exists("INFO.xlsx"):
        raise HTTPException(status_code=404, detail="فایل INFO.xlsx یافت نشد.")
    
    os.makedirs("backups", exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"backups/INFO_before_banks_update_{timestamp}.xlsx"
    shutil.copy2("INFO.xlsx", backup_path)
    
    wb = openpyxl.load_workbook("INFO.xlsx", data_only=False)
    if "Banks" not in wb.sheetnames:
        raise HTTPException(status_code=400, detail="برگه Banks در INFO.xlsx یافت نشد.")
    
    ws = wb["Banks"]
    ws.cell(5, 1, "کد حساب")
    ws.cell(5, 2, "عنوان نمایشی (روزمره)")
    ws.cell(5, 3, "مانده اولیه (ریال)")
    ws.cell(5, 4, "فعال")
    ws.cell(5, 5, "توضیحات")
    ws.cell(5, 6, "شماره حساب / شبا")
    
    for idx, b in enumerate(req.banks):
        r = 7 + idx
        ws.cell(r, 1, b.code.strip())
        ws.cell(r, 2, b.title.strip())
        ws.cell(r, 3, b.opening_balance)
        ws.cell(r, 4, "بله" if b.active else "خیر")
        ws.cell(r, 5, b.notes.strip() if b.notes else "")
        ws.cell(r, 6, b.account_no.strip() if b.account_no else "")
        
    wb.save("INFO.xlsx")
    wb.close()
    
    return {
        "success": True,
        "message": f"اطلاعات {len(req.banks)} حساب بانکی با موفقیت در فایل مادر INFO.xlsx ثبت و ذخیره شد.",
        "backup_name": os.path.basename(backup_path)
    }


@app.get("/api/sample-files")
async def get_sample_files():
    """Lists pre-built sample Excel files."""
    samples = [
        {
            "id": "sample_bank_single_col",
            "name": "صورت‌حساب بانکی تک‌ستونی (یک ستون مبلغ + تشخیص واریز/برداشت)",
            "file": "sample_bank_single_col.xlsx",
            "target_suggestion": "transactions",
            "desc": "حاوی ردیف، تاریخ، ساعت، یک ستون مبلغ کل، ستون تشخیص (واریز/برداشت)، پیگیری و شرح لاتین",
            "mode": "single"
        },
        {
            "id": "sample_bank_dual_col",
            "name": "صورت‌حساب بانکی دو‌ستونی (ستون مجزای واریز و برداشت)",
            "file": "sample_bank_dual_col.xlsx",
            "target_suggestion": "transactions",
            "desc": "حاوی شماره ارجاع، تاریخ، ساعت، ستون واریز، ستون برداشت، نام شخص و شرح تراکنش",
            "mode": "dual"
        },
        {
            "id": "sample_baskol_weighbridge",
            "name": "خروجی باسکول توزین دیجیتال (بار باغداران و فروش)",
            "file": "sample_baskol_weighbridge.xlsx",
            "target_suggestion": "baskol",
            "desc": "حاوی شماره قبض، تاریخ، نوع عملیات، نام باغدار، پلاک خودرو، وزن پر و تعداد سبد"
        },
        {
            "id": "sample_fruit_sales",
            "name": "فاکتورهای خرید و فروش مرکبات",
            "file": "sample_fruit_sales.xlsx",
            "target_suggestion": "trade",
            "desc": "حاوی شماره فاکتور، تاریخ، نام خریدار/باغدار، وزن خالص، فی کیلو و مبلغ کل"
        }
    ]
    return {"samples": samples}


@app.post("/api/load-sample")
async def load_sample(sample_id: str = Form(...)):
    """Loads one of the pre-built sample files directly."""
    sample_files = {
        "sample_bank_single_col": "samples/sample_bank_single_col.xlsx",
        "sample_bank_dual_col": "samples/sample_bank_dual_col.xlsx",
        "sample_baskol_weighbridge": "samples/sample_baskol_weighbridge.xlsx",
        "sample_fruit_sales": "samples/sample_fruit_sales.xlsx"
    }
    file_path = sample_files.get(sample_id)
    if not file_path or not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="فایل نمونه یافت نشد.")
    
    with open(file_path, "rb") as f:
        file_bytes = f.read()
    
    parsed = parse_uploaded_file(file_bytes, os.path.basename(file_path))
    return parsed


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    """Handles uploaded XLSX, XLS, or CSV file."""
    try:
        content = await file.read()
        parsed = parse_uploaded_file(content, file.filename)
        return parsed
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/auto-match")
async def auto_match(req: AutoMatchRequest):
    """Calculates auto-matching between source columns and target schema."""
    mapping = auto_match_columns(req.source_columns, req.target_module)
    return {"mapping": mapping}


@app.post("/api/preview-mapped")
async def preview_mapped(req: PreviewMappedRequest):
    """Generates mapped data preview with validations and summaries."""
    try:
        res = generate_mapped_preview(
            raw_rows=req.raw_rows,
            target_module_id=req.target_module,
            column_mapping=req.column_mapping,
            custom_defaults=req.custom_defaults or {},
            options=req.options or {}
        )
        return res
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/merge")
async def merge_records(req: MergeRequest):
    """Safely merges mapped rows into the target system Excel workbook."""
    try:
        if not req.rows:
            raise HTTPException(status_code=400, detail="هیچ سطری برای ادغام ارسال نشده است.")
        
        result = merge_rows_into_system_excel(
            target_module_id=req.target_module,
            rows_to_merge=req.rows,
            user_name=req.user_name or "کاربر سیستم"
        )
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/system-status")
async def get_system_status():
    """Returns overview of the 9 main system workbooks."""
    files_info = []
    order = [
        ("INFO.xlsx", "Accounts", "اطلاعات پایه و حساب‌ها (قاسم زاهدی)"),
        ("Transactions.xlsx", "Entry", "دفتر بانک و صندوق (تراکنش‌ها)"),
        ("Baskol.xlsx", "Baskol", "دفتر باسکول (توزین بار)"),
        ("Purchases.xlsx", "Report", "گزارش و صورتحساب خرید"),
        ("Sales.xlsx", "Report", "گزارش و صورتحساب فروش"),
        ("Employees.xlsx", "Wages", "دستمزد و کارکرد پرسنل"),
        ("Expenses.xlsx", "Report", "گزارش هزینه‌های جاری"),
        ("Cheques.xlsx", "Issues", "دفتر چک‌های صادره"),
        ("Baskets.xlsx", "Entry", "گردش و خرید سبد")
    ]
    
    for fname, main_sheet, label in order:
        if os.path.exists(fname):
            stat = os.stat(fname)
            size_kb = round(stat.st_size / 1024, 1)
            mod_time = datetime.fromtimestamp(stat.st_mtime).strftime("%Y/%m/%d %H:%M")
            files_info.append({
                "file": fname,
                "label": label,
                "main_sheet": main_sheet,
                "size_kb": size_kb,
                "modified": mod_time,
                "status": "ready"
            })
        else:
            files_info.append({
                "file": fname,
                "label": label,
                "main_sheet": main_sheet,
                "size_kb": 0,
                "modified": "-",
                "status": "missing"
            })
            
    backups = [f for f in os.listdir("backups") if f.endswith(".xlsx")] if os.path.exists("backups") else []
    return {
        "files": files_info,
        "backup_count": len(backups)
    }


@app.get("/api/preview-system/{filename}/{sheet_name}")
async def preview_system_sheet(filename: str, sheet_name: str, limit: int = 50):
    """Returns current data from any system Excel sheet for verification."""
    if not os.path.exists(filename):
        raise HTTPException(status_code=404, detail="فایل یافت نشد.")
    
    wb = openpyxl.load_workbook(filename, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise HTTPException(status_code=404, detail="برگه در فایل یافت نشد.")
    
    ws = wb[sheet_name]
    rows = []
    for r in range(1, min(limit + 1, ws.max_row + 1)):
        row_vals = [ws.cell(r, c).value for c in range(1, min(20, ws.max_column + 1))]
        if any(v is not None for v in row_vals):
            rows.append({
                "row_num": r,
                "values": [str(v) if v is not None else "" for v in row_vals]
            })
            
    return {
        "filename": filename,
        "sheet": sheet_name,
        "total_rows": ws.max_row,
        "total_cols": ws.max_column,
        "rows": rows
    }


@app.get("/api/download/{filename}")
async def download_file(filename: str, is_backup: bool = False):
    """Allows downloading system workbooks or backups directly."""
    target_dir = "backups" if is_backup else "."
    path = os.path.join(target_dir, filename)
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="فایل برای دانلود یافت نشد.")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=filename
    )


@app.get("/api/backups")
async def list_backups():
    """Lists all available backups with timestamps and sizes."""
    if not os.path.exists("backups"):
        return {"backups": []}
    files = [f for f in os.listdir("backups") if f.endswith(".xlsx")]
    items = []
    for f in sorted(files, reverse=True):
        p = os.path.join("backups", f)
        st = os.stat(p)
        items.append({
            "filename": f,
            "size_kb": round(st.st_size / 1024, 1),
            "created": datetime.fromtimestamp(st.st_ctime).strftime("%Y/%m/%d %H:%M:%S")
        })
    return {"backups": items}
