import io
import pandas as pd
from typing import List, Dict, Any
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def export_to_excel(results: List[Dict[str, Any]]) -> io.BytesIO:
    """
    Mengekspor daftar hasil analisis laporan keuangan ke file Excel dengan format rapi dan formula aktif.
    """
    output = io.BytesIO()
    
    # 1. Siapkan data baris untuk sheet ringkasan
    rows_data = []
    for idx, item in enumerate(results, start=1):
        roa_dec = (item.get("roa_percent") / 100.0) if item.get("roa_percent") is not None else None
        
        rows_data.append({
            "No": idx,
            "Kode Saham": item.get("ticker", "N/A"),
            "Nama Perusahaan": item.get("entity_name", "N/A"),
            "Periode Laporan": item.get("period", "N/A"),
            "Auditor Tahun Berjalan (KAP)": item.get("current_auditor", "N/A"),
            "KAP Tier": item.get("kap_tier", "N/A"),
            "Nama Partner Audit (Signing Partner)": item.get("audit_partner", "N/A"),
            "Auditor Tahun Sebelumnya": item.get("prior_auditor", "N/A"),
            "Status Pergantian KAP": item.get("auditor_switching", "N/A"),
            "Skala Pembulatan": item.get("rounding_text", "Satuan Penuh"),
            "Total Aset (Riil Rp)": item.get("total_assets_scaled"),
            "Laba Bersih (Riil Rp)": item.get("net_income_scaled"),
            "ROA (%)": roa_dec,
            "Kategori ROA": item.get("roa_status", "N/A"),
            "Nama File": item.get("filename", "")
        })

    df_summary = pd.DataFrame(rows_data)

    # 2. Siapkan data baris untuk sheet jejak audit
    audit_rows = []
    for idx, item in enumerate(results, start=1):
        trail = item.get("audit_trail", {})
        audit_rows.append({
            "No": idx,
            "Kode Saham": item.get("ticker", "N/A"),
            "Sheet Info Umum": trail.get("general_sheet", "N/A"),
            "Sheet Posisi Keuangan": trail.get("balance_sheet", "N/A"),
            "Baris Total Aset": trail.get("balance_row_matched", "N/A"),
            "Label Baris Aset": trail.get("balance_label_matched", "N/A"),
            "Sheet Laba Rugi": trail.get("income_sheet", "N/A"),
            "Baris Laba Bersih": trail.get("income_row_matched", "N/A"),
            "Label Baris Laba": trail.get("income_label_matched", "N/A"),
            "Formula ROA": item.get("roa_formula", "N/A")
        })
    df_audit = pd.DataFrame(audit_rows)

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name="Ringkasan Analisis ROA & KAP", index=False)
        df_audit.to_excel(writer, sheet_name="Jejak Audit & Verifikasi", index=False)

        # Ambil workbook dan styling
        wb = writer.book
        ws_summary = wb["Ringkasan Analisis ROA & KAP"]
        ws_audit = wb["Jejak Audit & Verifikasi"]

        # Style Header
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Dark Blue
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

        thin_border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        )

        for ws in [ws_summary, ws_audit]:
            ws.views.sheetView[0].showGridLines = True
            ws.row_dimensions[1].height = 28

            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_align
                cell.border = thin_border

            # Styling data baris
            for r_idx in range(2, ws.max_row + 1):
                ws.row_dimensions[r_idx].height = 22
                for c_idx in range(1, ws.max_column + 1):
                    d_cell = ws.cell(row=r_idx, column=c_idx)
                    d_cell.font = Font(name="Segoe UI", size=10)
                    d_cell.border = thin_border
                    
                    # Zebra striping
                    if r_idx % 2 == 1:
                        d_cell.fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

        # Number Formats pada ws_summary
        # Kolom 11: Total Aset, Kolom 12: Laba Bersih, Kolom 13: ROA (%)
        for r_idx in range(2, ws_summary.max_row + 1):
            # No
            ws_summary.cell(row=r_idx, column=1).alignment = Alignment(horizontal="center")
            # Ticker
            ws_summary.cell(row=r_idx, column=2).alignment = Alignment(horizontal="center")
            ws_summary.cell(row=r_idx, column=2).font = Font(name="Segoe UI", size=10, bold=True)
            # Periode
            ws_summary.cell(row=r_idx, column=4).alignment = Alignment(horizontal="center")
            # Tier & Rotasi
            ws_summary.cell(row=r_idx, column=6).alignment = Alignment(horizontal="center")
            ws_summary.cell(row=r_idx, column=9).alignment = Alignment(horizontal="center")
            ws_summary.cell(row=r_idx, column=10).alignment = Alignment(horizontal="center")

            # Format Angka Total Aset
            c_asset = ws_summary.cell(row=r_idx, column=11)
            c_asset.number_format = '#,##0'
            c_asset.alignment = Alignment(horizontal="right")

            # Format Angka Laba Bersih
            c_income = ws_summary.cell(row=r_idx, column=12)
            c_income.number_format = '#,##0'
            c_income.alignment = Alignment(horizontal="right")

            # Format Persentase ROA
            c_roa = ws_summary.cell(row=r_idx, column=13)
            c_roa.number_format = '0.00%'
            c_roa.alignment = Alignment(horizontal="right")
            c_roa.font = Font(name="Segoe UI", size=10, bold=True)

            # Kategori ROA
            ws_summary.cell(row=r_idx, column=14).alignment = Alignment(horizontal="center")

        # Auto-adjust column width
        for ws in [ws_summary, ws_audit]:
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    val_str = str(cell.value or '')
                    if len(val_str) > max_len:
                        max_len = len(val_str)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    output.seek(0)
    return output

def export_to_csv(results: List[Dict[str, Any]]) -> str:
    """Mengekspor hasil analisis ke string format CSV."""
    rows_data = []
    for idx, item in enumerate(results, start=1):
        rows_data.append({
            "No": idx,
            "Kode Saham": item.get("ticker", "N/A"),
            "Nama Perusahaan": item.get("entity_name", "N/A"),
            "Periode": item.get("period", "N/A"),
            "Auditor Tahun Berjalan (KAP)": item.get("current_auditor", "N/A"),
            "KAP Tier": item.get("kap_tier", "N/A"),
            "Partner Audit": item.get("audit_partner", "N/A"),
            "Auditor Sebelumnya": item.get("prior_auditor", "N/A"),
            "Status Rotasi KAP": item.get("auditor_switching", "N/A"),
            "Pembulatan": item.get("rounding_text", "Satuan Penuh"),
            "Total Aset Riil (Rp)": item.get("total_assets_scaled"),
            "Laba Bersih Riil (Rp)": item.get("net_income_scaled"),
            "ROA (%)": item.get("roa_percent"),
            "Kategori ROA": item.get("roa_status", "N/A"),
            "Formula ROA": item.get("roa_formula", "N/A")
        })
    df = pd.DataFrame(rows_data)
    return df.to_csv(index=False)
