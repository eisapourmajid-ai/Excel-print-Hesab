"""
Schema definitions for Mazandasht Sorting & Accounting Excel Files.
Defines target modules, their sheets, formula columns to preserve,
input columns to map, descriptions, types, and fuzzy-matching keywords.
"""

# Specialized Accounting Operations for Bank Statements
BANK_OPERATIONS = [
    {
        "id": "fruit_buyer_deposit",
        "title": "واریز بارفروشان (دریافت بابت فروش مرکبات)",
        "default_type": "واریز",
        "nature": "بستانکار مشتری",
        "category": "فروش مرکبات",
        "suggested_party_prefix": "C",
        "desc": "دریافت وجه از خریداران و بارفروشان میدان بار بابت فاکتور فروش مرکبات"
    },
    {
        "id": "interbank_transfer",
        "title": "جابه‌جایی بین حساب‌ها (انتقال داخلی بانکی)",
        "default_type": "واریز",
        "nature": "انتقال خنثی",
        "category": "انتقال داخلی",
        "suggested_party_prefix": "BANK",
        "desc": "انتقال نقدینگی بین حساب‌های بانکی شرکت (بدون اثر سود و زیان)"
    },
    {
        "id": "misc_deposit",
        "title": "واریز متفرقه",
        "default_type": "واریز",
        "nature": "درآمد متفرقه",
        "category": "سایر درآمدها",
        "suggested_party_prefix": "X",
        "desc": "سایر درآمدهای غیرعملیاتی شرکت"
    },
    {
        "id": "pos_settlement",
        "title": "سرجمع دستگاه پوز (POS)",
        "default_type": "واریز",
        "nature": "فروش نقدی",
        "category": "فروش نقدی",
        "suggested_party_prefix": "X004",
        "desc": "واریزی تجمیعی پایان روز دستگاه‌های کارتخوان باسکول و کارگاه"
    },
    {
        "id": "bank_fees",
        "title": "کارمزد و هزینه‌های بانکی",
        "default_type": "برداشت",
        "nature": "هزینه مالی",
        "category": "هزینه مالی",
        "suggested_party_prefix": "X001",
        "desc": "کارمزد حواله ساتنا، پایا، صدور چک و آبونمان پیامک بانکی"
    },
    {
        "id": "loan_received",
        "title": "وام و تسهیلات دریافتی",
        "default_type": "واریز",
        "nature": "افزایش بدهی تسهیلات",
        "category": "تسهیلات بانکی",
        "suggested_party_prefix": "L",
        "desc": "واریز وام و تسهیلات بانکی یا قرض تجاری به حساب شرکت (ثبت در بدهی‌های ترازنامه)"
    },
    {
        "id": "loan_repayment",
        "title": "بازپرداخت وام، تسهیلات و اقساط",
        "default_type": "برداشت",
        "nature": "کاهش بدهی تسهیلات",
        "category": "اقساط وام",
        "suggested_party_prefix": "L",
        "desc": "پرداخت اقساط اصل و کارمزد تسهیلات بانکی"
    },
    {
        "id": "director_asset_sale",
        "title": "فروش اموال مدیر (قاسم زاهدی) - ورود به چرخه شرکت",
        "default_type": "واریز",
        "nature": "آورده و جاری شرکا (ترازنامه)",
        "category": "جاری شرکا — قاسم زاهدی",
        "suggested_party_prefix": "CAP003",
        "desc": "نقدینگی حاصل از فروش اموال و دارایی شخصی مدیر و ورود به حساب شرکت (افزایش جاری شرکا در ترازنامه)"
    },
    {
        "id": "director_personal_expense",
        "title": "هزینه‌های شخصی مدیر (قاسم زاهدی) - جاری شرکا",
        "default_type": "برداشت",
        "nature": "برداشت شخصی شرکا (ترازنامه)",
        "category": "جاری شرکا — قاسم زاهدی",
        "suggested_party_prefix": "CAP002",
        "desc": "پرداخت هزینه‌های شخصی مدیر از حساب شرکت (برداشت جاری شرکا در ترازنامه - عدم انتساب به هزینه‌های شرکت)"
    },
    {
        "id": "citrus_purchase_payment",
        "title": "خرید مرکبات / تسویه حساب باغدار",
        "default_type": "برداشت",
        "nature": "کاهش بدهی باغدار",
        "category": "خرید مرکبات",
        "suggested_party_prefix": "S",
        "desc": "پرداخت علی‌الحساب یا تسویه نهایی به باغداران و تأمین‌کنندگان میوه"
    },
    {
        "id": "operating_expenses",
        "title": "هزینه‌های جاری و سورتینگ",
        "default_type": "برداشت",
        "nature": "هزینه عملیاتی",
        "category": "هزینه‌های سورتینگ",
        "suggested_party_prefix": "SG",
        "desc": "هزینه‌های بسته‌بندی، کارتن، گازوئیل، برق، تعمیرات و باربری"
    },
    {
        "id": "payroll_wages",
        "title": "حقوق و دستمزد پرسنل",
        "default_type": "برداشت",
        "nature": "هزینه دستمزد",
        "category": "دستمزد پرسنل",
        "suggested_party_prefix": "E",
        "desc": "پرداخت دستمزد کارگران خط سورتینگ، انباردار و حسابداری"
    }
]

TARGET_SCHEMAS = {
    "transactions": {
        "id": "transactions",
        "file": "Transactions.xlsx",
        "sheet": "Entry",
        "title": "صورت‌حساب و تراکنش‌های بانک",
        "icon": "credit-card",
        "start_row": 9,
        "max_capacity": 2008,
        "date_col": 3,
        "check_empty_cols": [3, 5, 7],
        "description": "ثبت گردش مالی حساب‌های بانکی، واریز بارفروشان، پوز، هزینه‌های شخصی و آورده مدیر (قاسم زاهدی)، وام‌ها و مغایرت‌گیری",
        "preserved_formulas": {
            1: "=IF($C{r}=\"\",\"\",COUNTA($C$9:$C{r}))",
            2: "=IF($C{r}=\"\",\"\",\"TXN-\"&RIGHT(\"0000\"&{r_counter},4))",
            6: "=IF($E{r}=\"\",\"\",IFERROR(IF(VLOOKUP(LEFT(UPPER(TRIM($E{r})),6),BankTbl,2,0)=\"\",\"(بدون عنوان)\",VLOOKUP(LEFT(UPPER(TRIM($E{r})),6),BankTbl,2,0)),\"بانک ناشناخته\"))",
            8: "=IF($G{r}=\"\",\"\",IFERROR(IF(VLOOKUP($G{r},AcctTbl,4,0)=\"\",\"(بدون نام)\",VLOOKUP($G{r},AcctTbl,4,0)),\"کد ناشناخته\"))",
            9: "=IF($G{r}=\"\",\"\",IFERROR(IF(VLOOKUP($G{r},AcctTbl,3,0)=\"\",IF(VLOOKUP($G{r},AcctTbl,2,0)=\"\",\"—\",VLOOKUP($G{r},AcctTbl,2,0)),VLOOKUP($G{r},AcctTbl,3,0)),\"—\"))",
            16: "=IF(OR($C{r}=\"\",$E{r}=\"\",Calc!$H{r}=\"\"),\"\",IFERROR(VLOOKUP(LEFT(UPPER(TRIM($E{r})),6),BankTbl,3,0),0)+SUMIFS(CalcDep,CalcBank,LEFT(UPPER(TRIM($E{r})),6),CalcKey,\"<=\"&Calc!$H{r})-SUMIFS(CalcWit,CalcBank,LEFT(UPPER(TRIM($E{r})),6),CalcKey,\"<=\"&Calc!$H{r}))",
            17: "=IF(OR($C{r}=\"\",$O{r}=\"\"),\"\",IF($P{r}=\"\",\"؟\",IF(ROUND($O{r}-$P{r},0)=0,\"√\",\"×\")))"
        },
        "columns": [
            {
                "key": "col3",
                "col_index": 3,
                "title": "تاریخ تراکنش",
                "required": True,
                "type": "date",
                "locked": True,
                "hint": "قالب استاندارد: ۱۴۰۵/۰۷/۰۱ (قفل‌شده)",
                "keywords": ["تاریخ", "date", "tarikh", "روز"]
            },
            {
                "key": "col4",
                "col_index": 4,
                "title": "ساعت تراکنش",
                "required": False,
                "type": "text",
                "locked": True,
                "default": "10:00",
                "hint": "قالب: ۱۰:۳۰ (قفل‌شده)",
                "keywords": ["ساعت", "time", "زمان"]
            },
            {
                "key": "col12",
                "col_index": 12,
                "title": "شماره سند / پیگیری",
                "required": False,
                "type": "text",
                "locked": False,
                "default": "",
                "hint": "شماره ارجاع، شماره فیش یا پیگیری",
                "keywords": ["پیگیری", "ارجاع", "سند", "فیش", "شماره فیش", "ref", "trace", "doc", "tracking", "سریال"]
            },
            {
                "key": "col13",
                "col_index": 13,
                "title": "مبلغ واریز (ریال)",
                "required": False,
                "type": "number",
                "locked": True,
                "default": 0,
                "hint": "مبلغ بستانکار / دریافتی به ریال (قفل‌شده)",
                "keywords": ["واریز", "دریافت", "بستانکار", "deposit", "credit", "مبلغ واریز"]
            },
            {
                "key": "col14",
                "col_index": 14,
                "title": "مبلغ برداشت (ریال)",
                "required": False,
                "type": "number",
                "locked": True,
                "default": 0,
                "hint": "مبلغ بدهکار / پرداختی به ریال (قفل‌شده)",
                "keywords": ["برداشت", "پرداخت", "بدهکار", "withdraw", "debit", "مبلغ برداشت"]
            },
            {
                "key": "col10",
                "col_index": 10,
                "title": "نوع عملیات مالی",
                "required": True,
                "type": "select",
                "options": [op["title"] for op in BANK_OPERATIONS],
                "locked": False,
                "default": "واریز بارفروشان (دریافت بابت فروش مرکبات)",
                "hint": "انتخاب دقیق سرفصل مالی جهت ترازنامه و سود و زیان",
                "keywords": ["عملیات", "نوع عملیات", "نوع", "تراکنش", "شرح", "بابت"]
            },
            {
                "key": "col7",
                "col_index": 7,
                "title": "کد و نام طرف حساب",
                "required": True,
                "type": "party",
                "locked": False,
                "default": "C001",
                "hint": "انتخاب از حساب‌های INFO (مشتری، باغدار، جاری قاسم زاهدی...)",
                "keywords": ["طرف", "شخص", "حساب", "مشتری", "باغدار", "صاحب حساب", "نام", "party", "account"]
            },
            {
                "key": "col11",
                "col_index": 11,
                "title": "سرفصل / بابت",
                "required": False,
                "type": "text",
                "locked": False,
                "default": "فروش مرکبات",
                "hint": "فروش، تسویه باغدار، جاری قاسم زاهدی، کارمزد...",
                "keywords": ["بابت", "سرفصل", "عنوان", "موضوع", "category"]
            },
            {
                "key": "col18",
                "col_index": 18,
                "title": "توضیحات تکمیلی (حاوی ستون‌های نگاشت‌نشده)",
                "required": False,
                "type": "text",
                "locked": False,
                "default": "",
                "hint": "ستون‌های اضافه فایل بانک خودکار اینجا الحاق می‌شوند",
                "keywords": ["توضیحات", "شرح", "یادداشت", "notes", "description"]
            }
        ]
    },
    "baskol": {
        "id": "baskol",
        "file": "Baskol.xlsx",
        "sheet": "Baskol",
        "title": "ورود و خروج بار و باسکول",
        "icon": "truck",
        "start_row": 9,
        "max_capacity": 2008,
        "date_col": 2,
        "check_empty_cols": [2, 4, 7],
        "description": "ثبت قبض‌های توزین باسکول برای خرید و فروش مرکبات با محاسبه زنده وزن خالص، افت و تار سبد",
        "preserved_formulas": {
            1: "=IF($B{r}=\"\",\"\",COUNTA($B$9:$B{r}))",
            6: "=IF($E{r}=\"\",\"\",IFERROR(VLOOKUP($E{r},OpTbl,2,0),\"\"))",
            8: "=IF($G{r}=\"\",\"\",IFERROR(IF(VLOOKUP($G{r},AcctTbl,4,0)=\"\",\"(بدون نام)\",VLOOKUP($G{r},AcctTbl,4,0)),\"کد ناشناخته\"))",
            13: "=IF(OR($B{r}=\"\",$K{r}=\"\"),\"\",($K{r}-IF($L{r}=\"\",0,$L{r})*IFERROR(VLOOKUP(\"P01\",ParamTbl,3,0),1.5))*(1-IF(IFERROR(VLOOKUP($E{r},OpTbl,3,0),\"\")=\"خرید\",IFERROR(VLOOKUP(\"P02\",ParamTbl,3,0),1),0)/100))"
        },
        "columns": [
            {
                "key": "col2",
                "col_index": 2,
                "title": "تاریخ توزین",
                "required": True,
                "type": "date",
                "locked": False,
                "default": "",
                "hint": "قالب استاندارد: ۱۴۰۵/۰۷/۰۱",
                "keywords": ["تاریخ", "date", "tarikh", "زمان", "روز", "توزین"]
            },
            {
                "key": "col3",
                "col_index": 3,
                "title": "ساعت",
                "required": False,
                "type": "text",
                "locked": False,
                "default": "08:30",
                "hint": "قالب: ۰۸:۳۰",
                "keywords": ["ساعت", "time", "زمان"]
            },
            {
                "key": "col4",
                "col_index": 4,
                "title": "شماره سند / قبض باسکول",
                "required": True,
                "type": "text",
                "locked": False,
                "default": "",
                "hint": "شماره فیش باسکول یا قبض",
                "keywords": ["قبض", "سند", "فیش", "شماره", "doc", "ticket", "bill", "no"]
            },
            {
                "key": "col5",
                "col_index": 5,
                "title": "نوع عملیات",
                "required": True,
                "type": "select",
                "options": ["خرید مرکبات", "فروش", "مرجوعی خرید", "مرجوعی فروش", "خدمات سورتینگ"],
                "locked": False,
                "default": "خرید مرکبات",
                "hint": "خرید مرکبات / فروش",
                "keywords": ["عملیات", "نوع", "operation", "نوع بار"]
            },
            {
                "key": "col7",
                "col_index": 7,
                "title": "کد و نام طرف حساب",
                "required": True,
                "type": "party",
                "locked": False,
                "default": "S001",
                "hint": "باغدار (S...) یا خریدار (C... / SL...)",
                "keywords": ["باغدار", "طرف", "مشتری", "کد", "شخص", "صاحب بار", "party"]
            },
            {
                "key": "col10",
                "col_index": 10,
                "title": "نوع مرکبات",
                "required": False,
                "type": "text",
                "locked": False,
                "default": "پرتقال تامسون",
                "hint": "پرتقال تامسون، خونی، نارنگی، لیمو...",
                "keywords": ["مرکبات", "محصول", "میوه", "پرتقال", "نارنگی", "citrus"]
            },
            {
                "key": "col11",
                "col_index": 11,
                "title": "وزن ناخالص / باسکول (kg)",
                "required": True,
                "type": "number",
                "locked": False,
                "default": "",
                "hint": "وزن ثبت‌شده روی باسکول",
                "keywords": ["وزن", "ناخالص", "باسکول", "وزن پر", "gross", "scale", "weight"]
            },
            {
                "key": "col12",
                "col_index": 12,
                "title": "تعداد سبد / جعبه",
                "required": False,
                "type": "number",
                "locked": False,
                "default": 0,
                "hint": "تعداد سبد جهت محاسبه تار (هر سبد ۱.۵ kg)",
                "keywords": ["سبد", "جعبه", "تعداد", "box", "crates", "basket"]
            },
            {
                "key": "col15",
                "col_index": 15,
                "title": "توضیحات و شرح بار (پلاک/راننده)",
                "required": False,
                "type": "text",
                "locked": False,
                "default": "",
                "hint": "شماره پلاک، نام راننده، ستون‌های الحاقی",
                "keywords": ["توضیحات", "پلاک", "راننده", "شرح", "notes"]
            }
        ]
    },
    "trade": {
        "id": "trade",
        "file": "Purchases.xlsx",
        "sheet": "Calc",
        "title": "خرید و فروش مرکبات",
        "icon": "scale",
        "start_row": 9,
        "max_capacity": 2008,
        "date_col": 1,
        "check_empty_cols": [1, 7],
        "description": "فاکتورهای قطعی خرید از باغداران، فروش به بارفروشان، فی، مبالغ و صورت‌حساب اشخاص",
        "preserved_formulas": {},
        "columns": [
            {
                "key": "col1",
                "col_index": 1,
                "title": "تاریخ فاکتور",
                "required": True,
                "type": "date",
                "locked": False,
                "default": "",
                "hint": "۱۴۰۵/۰۷/۰۱",
                "keywords": ["تاریخ", "date", "tarikh"]
            },
            {
                "key": "col4",
                "col_index": 4,
                "title": "شماره فاکتور / سند",
                "required": False,
                "type": "text",
                "locked": False,
                "default": "",
                "hint": "شماره فاکتور خرید یا فروش",
                "keywords": ["فاکتور", "شماره", "سند", "invoice", "doc"]
            },
            {
                "key": "col5",
                "col_index": 5,
                "title": "نوع سند",
                "required": True,
                "type": "select",
                "options": ["فروش مرکبات", "خرید مرکبات", "مرجوعی"],
                "locked": False,
                "default": "فروش مرکبات",
                "hint": "خرید یا فروش",
                "keywords": ["نوع", "نوع فاکتور", "خرید", "فروش"]
            },
            {
                "key": "col7",
                "col_index": 7,
                "title": "کد و نام طرف حساب",
                "required": True,
                "type": "party",
                "locked": False,
                "default": "C001",
                "hint": "مشتری یا تأمین‌کننده",
                "keywords": ["طرف", "مشتری", "باغدار", "نام", "کد"]
            },
            {
                "key": "col11",
                "col_index": 11,
                "title": "نوع مرکبات",
                "required": False,
                "type": "text",
                "locked": False,
                "default": "پرتقال تامسون",
                "hint": "پرتقال تامسون، خونی، نارنگی...",
                "keywords": ["محصول", "میوه", "مرکبات", "نوع محصول"]
            },
            {
                "key": "col14",
                "col_index": 14,
                "title": "وزن خالص (kg)",
                "required": True,
                "type": "number",
                "locked": False,
                "default": "",
                "hint": "وزن خالص بار به کیلوگرم",
                "keywords": ["خالص", "وزن خالص", "وزن", "net", "kg"]
            },
            {
                "key": "col15",
                "col_index": 15,
                "title": "فی هر کیلو (ریال)",
                "required": False,
                "type": "number",
                "locked": False,
                "default": 0,
                "hint": "نرخ هر کیلو به ریال",
                "keywords": ["فی", "قیمت", "نرخ", "price", "rate"]
            },
            {
                "key": "col16",
                "col_index": 16,
                "title": "مبلغ کل فاکتور (ریال)",
                "required": True,
                "type": "number",
                "locked": False,
                "default": 0,
                "hint": "مبلغ کل فاکتور به ریال",
                "keywords": ["مبلغ", "جمع", "مبلغ کل", "amount", "total"]
            },
            {
                "key": "col18",
                "col_index": 18,
                "title": "توضیحات و شرایط تسویه",
                "required": False,
                "type": "text",
                "locked": False,
                "default": "",
                "hint": "شرایط تسویه، چک، ستون‌های الحاقی",
                "keywords": ["توضیحات", "تسویه", "شرح", "notes"]
            }
        ]
    }
}
