
# -*- coding: utf-8 -*-
"""
將六套(A~F)文件依固定順序合併成六個獨立PDF
順序：Payment → Dell TMHK → Dell e-signed → SSH → 2025 Warranty extension → Invoice(一或多份)
來源目錄：input_pdfs/
輸出目錄：output/
A 套有兩張發票 (1401098209, 1401098750)
"""

import re
from pathlib import Path
from typing import List
from PyPDF2 import PdfMerger

BASE_DIR = Path(".")
INPUT_DIR = BASE_DIR / "input_pdfs"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

SETS = [
    {"label":"A","tmhk":"tmhk251218a","order":"dell-25120008","ssh":"ssh-25121801-twt","invoices":["1401098209","1401098750"]},
    {"label":"B","tmhk":"tmhk251218b","order":"dell-25120010","ssh":"ssh-25121802-meg","invoices":["1401098748"]},
    {"label":"C","tmhk":"tmhk251218c","order":"dell-25120009","ssh":"ssh-25121803-thl","invoices":["1401098749"]},
    {"label":"D","tmhk":"tmhk251218d","order":"dell-25120006","ssh":"ssh-25121804-pht","invoices":["1401098751"]},
    {"label":"E","tmhk":"tmhk251218e","order":"dell-25110013","ssh":"ssh-25121805-jp","invoices":["1401096860"]},
    {"label":"F","tmhk":"tmhk251218f","order":"dell-25110014","ssh":"ssh-25121806-int","invoices":["1401096723"]},
]

def list_pdf_files() -> List[Path]:
    return [p for p in INPUT_DIR.iterdir() if p.is_file() and p.suffix.lower()==".pdf"]

def find_one_by_tokens(tokens: List[str]) -> Path | None:
    tokens = [t.lower() for t in tokens]
    candidates = []
    for f in list_pdf_files():
        name = f.name.lower()
        if all(tok in name for tok in tokens):
            candidates.append(f)
    if not candidates:
        return None
    candidates.sort(key=lambda p: len(p.name))
    return candidates[0]

def find_payment(tmhk: str):      return find_one_by_tokens(["payment application_vendor","202512","warranty",tmhk])
def find_dell_tmhk(order:str,tmhk:str): return find_one_by_tokens([order,"tm",tmhk])
def find_dell_esigned(order:str):
    for v in ["e-signed","e_signed","esigned"]:
        p = find_one_by_tokens([order,"tm",v])
        if p: return p
    return None
def find_ssh(ssh_name:str):       return find_one_by_tokens([ssh_name])
def find_warranty_ext(order:str): return find_one_by_tokens(["2025","warranty extension",order])

def find_invoices(invoice_ids: List[str]) -> List[Path]:
    files = list_pdf_files()
    found = []
    for inv in invoice_ids:
        inv_lower = inv.lower()
        hit = [f for f in files if inv_lower in f.name.lower()]
        if hit:
            hit.sort(key=lambda p: len(p.name))
            found.append(hit[0])
        else:
            for f in files:
                if re.search(rf"{re.escape(inv)}", f.name, flags=re.IGNORECASE):
                    found.append(f); break
    return found

def merge_one_set(idx:int, s:dict):
    label, tmhk, order, ssh, invoices = s["label"], s["tmhk"], s["order"], s["ssh"], s["invoices"]
    steps = [
        ("Payment",           find_payment(tmhk)),
        ("Dell TMHK",         find_dell_tmhk(order, tmhk)),
        ("Dell e-signed",     find_dell_esigned(order)),
        ("SSH",               find_ssh(ssh)),
        ("Warranty extension",find_warranty_ext(order)),
    ]
    for i, f in enumerate(find_invoices(invoices), start=1):
        steps.append((f"Invoice #{i}", f))

    missing = [tag for tag, p in steps if p is None]
    if missing:
        print(f"[WARN] 套別 {label}: 找不到 → {missing}")

    merger = PdfMerger()
    appended = 0
    for tag, p in steps:
        if p and p.exists():
            merger.append(str(p)); appended += 1
            print(f"  [+] {label}-{tag}: {p.name}")
        else:
            print(f"  [ ] {label}-{tag}: (缺檔)")

    out_name = f"{idx:02d}__Package_{label}_{tmhk.upper()}_{order}.pdf"
    out_path = OUTPUT_DIR / out_name
    if appended>0:
        merger.write(str(out_path)); merger.close()
        print(f"[OK] 套別 {label} 合併完成 → {out_path}")
    else:
        print(f"[SKIP] 套別 {label} 無可合併頁面")

def main():
    if not INPUT_DIR.exists():
        print(f"[ERR] 找不到 input_pdfs/，請先上傳 PDF 到此資料夾"); return
    print(f"[INFO] 來源：{INPUT_DIR.resolve()}\n[INFO] 輸出：{OUTPUT_DIR.resolve()}\n")
    for i, s in enumerate(SETS, start=1):
        print(f"=== 合併套別 {s['label']} ===")
        merge_one_set(i, s)

ifif __name__ == "__main__":
