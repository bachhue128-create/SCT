# -*- coding: utf-8 -*-
"""
HỆ THỐNG CẬP NHẬT INFOGRAPHIC THI ĐUA SIM FPT (BẢN PYTHON)
Dành cho người dùng có cài đặt Python và openpyxl.
"""

import os
import glob
import base64
import subprocess
from datetime import datetime
import openpyxl

def main():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    input_dir = os.path.join(base_dir, "input_data")
    output_dir = os.path.join(base_dir, "output")
    template_path = os.path.join(base_dir, "template", "infographic_template.html")
    assets_dir = os.path.join(base_dir, "template", "assets")
    
    os.makedirs(output_dir, exist_ok=True)
    
    # 1. Tìm file excel mới nhất
    excel_files = glob.glob(os.path.join(input_dir, "*.xlsx"))
    if not excel_files:
        excel_files = glob.glob(os.path.join(base_dir, "*.xlsx"))
        
    if not excel_files:
        print("[LỖI] Không tìm thấy file .xlsx nào trong input_data!")
        return

    excel_files.sort(key=os.path.getmtime, reverse=True)
    target_file = excel_files[0]
    print(f"[1/4] Đang đọc file: {os.path.basename(target_file)}")
    
    # 2. Đọc sheet "Thi đua Sim"
    wb = openpyxl.load_workbook(target_file, data_only=True)
    target_sheet_name = None
    for name in wb.sheetnames:
        if "thi đua sim" in name.lower() or "thi dua sim" in name.lower() or "sim" in name.lower():
            target_sheet_name = name
            break
            
    if not target_sheet_name:
        print("[LỖI] Không tìm thấy sheet 'Thi đua Sim'!")
        return
        
    ws = wb[target_sheet_name]
    print(f"      Đã tìm thấy sheet: {target_sheet_name}")
    
    # Tìm ngày báo cáo
    report_date = datetime.now().strftime("%d/%m/%Y")
    if "Report" in wb.sheetnames:
        r_ws = wb["Report"]
        val_b2 = r_ws["B2"].value
        if isinstance(val_b2, datetime):
            report_date = val_b2.strftime("%d/%m/%Y")
            
    # 3. Phân tích các shop
    shops = []
    for row in range(4, ws.max_row + 1):
        shop = ws.cell(row=row, column=1).value
        if not shop or "total" in str(shop).lower() or "tổng" in str(shop).lower():
            continue
            
        try:
            sl_may = float(ws.cell(row=row, column=3).value or 0)
            sl_sim = float(ws.cell(row=row, column=4).value or 0)
            tl_sim = float(ws.cell(row=row, column=5).value or 0)
            tang_truong = float(ws.cell(row=row, column=6).value or 0)
            
            shops.append({
                "shop": str(shop).strip(),
                "sl_may": sl_may,
                "sl_sim": sl_sim,
                "tl_sim": tl_sim,
                "tang_truong": tang_truong
            })
        except Exception:
            continue
            
    print(f"      Tổng số shop đã tổng hợp: {len(shops)} shop")
    
    # 4. Xếp hạng
    vua_sim = sorted(shops, key=lambda x: x["sl_sim"], reverse=True)[:5]
    cao_thu = sorted(shops, key=lambda x: x["tl_sim"], reverse=True)[:5]
    but_pha = sorted(shops, key=lambda x: x["tang_truong"], reverse=True)[:5]
    
    # Base64 assets
    header_b64 = ""
    header_path = os.path.join(assets_dir, "header_banner.jpg")
    if os.path.exists(header_path):
        with open(header_path, "rb") as f:
            header_b64 = "data:image/jpeg;base64," + base64.b64encode(f.read()).decode("utf-8")
            
    mascot_b64 = ""
    mascot_path = os.path.join(assets_dir, "mascot.png")
    if os.path.exists(mascot_path):
        with open(mascot_path, "rb") as f:
            mascot_b64 = "data:image/png;base64," + base64.b64encode(f.read()).decode("utf-8")
            
    # 5. Cập nhật HTML
    with open(template_path, "r", encoding="utf-8") as f:
        html = f.read()
        
    replacements = {
        "__HEADER_BANNER_BASE64__": header_b64,
        "__MASCOT_BASE64__": mascot_b64,
        "__NGAY_CAP_NHAT__": report_date,
        
        "__VUA_SIM_TOP1_NAME__": vua_sim[0]["shop"],
        "__VUA_SIM_TOP1_VALUE__": f"{int(vua_sim[0]['sl_sim']):,} SIM",
        "__VUA_SIM_TOP1_MAY__": f"{int(vua_sim[0]['sl_may']):,}",
        "__VUA_SIM_TOP1_RATE__": f"{vua_sim[0]['tl_sim']*100:.1f}%",
        
        "__VUA_SIM_TOP2_NAME__": vua_sim[1]["shop"],
        "__VUA_SIM_TOP2_VALUE__": f"{int(vua_sim[1]['sl_sim']):,} SIM",
        "__VUA_SIM_TOP2_MAY__": f"{int(vua_sim[1]['sl_may']):,}",
        "__VUA_SIM_TOP2_RATE__": f"{vua_sim[1]['tl_sim']*100:.1f}%",
        
        "__VUA_SIM_TOP3_NAME__": vua_sim[2]["shop"],
        "__VUA_SIM_TOP3_VALUE__": f"{int(vua_sim[2]['sl_sim']):,} SIM",
        "__VUA_SIM_TOP4_NAME__": vua_sim[3]["shop"],
        "__VUA_SIM_TOP4_VALUE__": f"{int(vua_sim[3]['sl_sim']):,} SIM",
        "__VUA_SIM_TOP5_NAME__": vua_sim[4]["shop"],
        "__VUA_SIM_TOP5_VALUE__": f"{int(vua_sim[4]['sl_sim']):,} SIM",
        
        "__TY_LE_TOP1_NAME__": cao_thu[0]["shop"],
        "__TY_LE_TOP1_VALUE__": f"{cao_thu[0]['tl_sim']*100:.1f}%",
        "__TY_LE_TOP1_SIM__": f"{int(cao_thu[0]['sl_sim']):,} SIM",
        "__TY_LE_TOP1_MAY__": f"{int(cao_thu[0]['sl_may']):,}",
        
        "__TY_LE_TOP2_NAME__": cao_thu[1]["shop"],
        "__TY_LE_TOP2_VALUE__": f"{cao_thu[1]['tl_sim']*100:.1f}%",
        "__TY_LE_TOP2_SIM__": f"{int(cao_thu[1]['sl_sim']):,} SIM",
        "__TY_LE_TOP2_MAY__": f"{int(cao_thu[1]['sl_may']):,}",
        
        "__TY_LE_TOP3_NAME__": cao_thu[2]["shop"],
        "__TY_LE_TOP3_VALUE__": f"{cao_thu[2]['tl_sim']*100:.1f}%",
        "__TY_LE_TOP4_NAME__": cao_thu[3]["shop"],
        "__TY_LE_TOP4_VALUE__": f"{cao_thu[3]['tl_sim']*100:.1f}%",
        "__TY_LE_TOP5_NAME__": cao_thu[4]["shop"],
        "__TY_LE_TOP5_VALUE__": f"{cao_thu[4]['tl_sim']*100:.1f}%",
        
        "__TANG_TRUONG_TOP1_NAME__": but_pha[0]["shop"],
        "__TANG_TRUONG_TOP1_VALUE__": f"{'+' if but_pha[0]['tang_truong']>0 else ''}{but_pha[0]['tang_truong']*100:.1f}%",
        "__TANG_TRUONG_TOP1_SIM__": f"{int(but_pha[0]['sl_sim']):,} SIM",
        "__TANG_TRUONG_TOP1_RATE__": f"{but_pha[0]['tl_sim']*100:.1f}%",
        
        "__TANG_TRUONG_TOP2_NAME__": but_pha[1]["shop"],
        "__TANG_TRUONG_TOP2_VALUE__": f"{'+' if but_pha[1]['tang_truong']>0 else ''}{but_pha[1]['tang_truong']*100:.1f}%",
        "__TANG_TRUONG_TOP2_SIM__": f"{int(but_pha[1]['sl_sim']):,} SIM",
        "__TANG_TRUONG_TOP2_RATE__": f"{but_pha[1]['tl_sim']*100:.1f}%",
        
        "__TANG_TRUONG_TOP3_NAME__": but_pha[2]["shop"],
        "__TANG_TRUONG_TOP3_VALUE__": f"{'+' if but_pha[2]['tang_truong']>0 else ''}{but_pha[2]['tang_truong']*100:.1f}%",
        "__TANG_TRUONG_TOP4_NAME__": but_pha[3]["shop"],
        "__TANG_TRUONG_TOP4_VALUE__": f"{'+' if but_pha[3]['tang_truong']>0 else ''}{but_pha[3]['tang_truong']*100:.1f}%",
        "__TANG_TRUONG_TOP5_NAME__": but_pha[4]["shop"],
        "__TANG_TRUONG_TOP5_VALUE__": f"{'+' if but_pha[4]['tang_truong']>0 else ''}{but_pha[4]['tang_truong']*100:.1f}%",
    }
    
    for k, v in replacements.items():
        html = html.replace(k, v)
        
    out_html = os.path.join(output_dir, "KetQua_ThiDua_Sim.html")
    with open(out_html, "w", encoding="utf-8") as f:
        f.write(html)
        
    print(f"[3/4] Đã lưu file HTML: {out_html}")
    
    # 6. Chụp ảnh PNG bằng Chrome/Edge
    out_png = os.path.join(output_dir, "KetQua_ThiDua_Sim.png")
    chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
    edge_path = "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
    
    browser = chrome_path if os.path.exists(chrome_path) else (edge_path if os.path.exists(edge_path) else None)
    if browser:
        print("[4/4] Đang render ảnh PNG...")
        cmd = [
            browser,
            "--headless",
            f"--screenshot={out_png}",
            "--no-sandbox",
            "--disable-gpu",
            "--window-size=1200,1440",
            "--hide-scrollbars",
            out_html
        ]
        subprocess.run(cmd, check=True)
        print(f"[THÀNH CÔNG] File ảnh đã tạo: {out_png}")
        os.startfile(out_png)

if __name__ == "__main__":
    main()
