import os
import io
import pandas as pd
from typing import Dict, List, Any

# Data master emiten realistis sesuai format e-XBRL BEI / IDX
SAMPLE_COMPANIES = [
    {
        "ticker": "BBCA",
        "entity_name": "PT Bank Central Asia Tbk",
        "period": "31 Desember 2023",
        "current_auditor": "Tanudiredja, Wibisana, Rintis & Rekan (PricewaterhouseCoopers)",
        "audit_partner": "Lucy Luciana Suhenda",
        "prior_auditor": "Tanudiredja, Wibisana, Rintis & Rekan (PricewaterhouseCoopers)",
        "rounding": "Satuan Penuh",
        "rounding_multiplier": 1,
        "total_assets": 1408105000000000, # Rp 1.408 Triliun
        "total_assets_prior": 1314732000000000,
        "net_income": 48639000000000,      # Rp 48,6 Triliun
        "net_income_prior": 40736000000000,
        "sheet_style": "idx_standard"
    },
    {
        "ticker": "BBRI",
        "entity_name": "PT Bank Rakyat Indonesia (Persero) Tbk",
        "period": "31 Desember 2023",
        "current_auditor": "Purwantono, Sungkoro & Surja (Ernst & Young)",
        "audit_partner": "Deden Riyadi",
        "prior_auditor": "Purwantono, Sungkoro & Surja (Ernst & Young)",
        "rounding": "Satuan Penuh",
        "rounding_multiplier": 1,
        "total_assets": 1965000000000000, # Rp 1.965 Triliun
        "total_assets_prior": 1865639000000000,
        "net_income": 60400000000000,      # Rp 60,4 Triliun
        "net_income_prior": 51408000000000,
        "sheet_style": "idx_standard"
    },
    {
        "ticker": "TLKM",
        "entity_name": "PT Telkom Indonesia (Persero) Tbk",
        "period": "31 Desember 2023",
        "current_auditor": "Purwantono, Sungkoro & Surja (Ernst & Young)",
        "audit_partner": "Sherly Jokom",
        "prior_auditor": "Purwantono, Sungkoro & Surja (Ernst & Young)",
        "rounding": "Dalam Jutaan Rupiah",
        "rounding_multiplier": 1000000,
        "total_assets": 287042000,         # Rp 287 Triliun (skala jutaan)
        "total_assets_prior": 275192000,
        "net_income": 24560000,            # Rp 24,5 Triliun (skala jutaan)
        "net_income_prior": 20753000,
        "sheet_style": "idx_standard"
    },
    {
        "ticker": "ASII",
        "entity_name": "PT Astra International Tbk",
        "period": "31 Desember 2023",
        "current_auditor": "Tanudiredja, Wibisana, Rintis & Rekan (PricewaterhouseCoopers)",
        "audit_partner": "Djohan Pinnarwan",
        "prior_auditor": "Siddharta Widjaja & Rekan (KPMG)", # Rotasi KAP!
        "rounding": "Dalam Jutaan Rupiah",
        "rounding_multiplier": 1000000,
        "total_assets": 445000000,         # Rp 445 Triliun (skala jutaan)
        "total_assets_prior": 413294000,
        "net_income": 33800000,            # Rp 33,8 Triliun (skala jutaan)
        "net_income_prior": 28944000,
        "sheet_style": "idx_standard"
    },
    {
        "ticker": "ICBP",
        "entity_name": "PT Indofood CBP Sukses Makmur Tbk",
        "period": "31 Desember 2023",
        "current_auditor": "Purwantono, Sungkoro & Surja (Ernst & Young)",
        "audit_partner": "Franky Jamin",
        "prior_auditor": "Purwantono, Sungkoro & Surja (Ernst & Young)",
        "rounding": "Dalam Jutaan Rupiah",
        "rounding_multiplier": 1000000,
        "total_assets": 119500000,         # Rp 119,5 Triliun (skala jutaan)
        "total_assets_prior": 115305000,
        "net_income": 9800000,             # Rp 9,8 Triliun (skala jutaan)
        "net_income_prior": 7430000,
        "sheet_style": "idx_descriptive"   # Menggunakan nama sheet deskriptif
    },
    {
        "ticker": "GOTO",
        "entity_name": "PT GoTo Gojek Tokopedia Tbk",
        "period": "31 Desember 2023",
        "current_auditor": "Imelda & Rekan (Deloitte Touche Tohmatsu)",
        "audit_partner": "Rosita Uli Sinaga",
        "prior_auditor": "Imelda & Rekan (Deloitte Touche Tohmatsu)",
        "rounding": "Dalam Jutaan Rupiah",
        "rounding_multiplier": 1000000,
        "total_assets": 54000000,          # Rp 54 Triliun (skala jutaan)
        "total_assets_prior": 61200000,
        "net_income": -90500000,           # Rugi Bersih (skala jutaan)
        "net_income_prior": -40400000,
        "sheet_style": "idx_standard"
    }
]

def create_idx_excel(company_data: Dict[str, Any]) -> io.BytesIO:
    """
    Membuat file Excel dalam memori yang mereplikasi secara presisi format e-XBRL IDX.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        is_standard = company_data.get("sheet_style") == "idx_standard"
        sheet_info = "1000000" if is_standard else "Informasi Umum"
        sheet_balance = "1110000" if is_standard else "Laporan Posisi Keuangan"
        sheet_income = "1311000" if is_standard else "Laporan Laba Rugi"

        # 1. Sheet Informasi Umum
        info_rows = [
            ["Informasi Umum Perusahaan / General Company Information", "", "", ""],
            ["Nama entitas / Entity name", company_data["entity_name"], "", ""],
            ["Kode entitas / Entity code", company_data["ticker"], "", ""],
            ["Nomor pokok wajib pajak / Taxpayer identification number", "01.000.123.4-056.000", "", ""],
            ["Industri / Industry", "Keuangan & Korporasi", "", ""],
            ["Sektor / Sector", "Terbuka (Listed)", "", ""],
            ["Auditor tahun berjalan / Current year auditor", company_data["current_auditor"], "", ""],
            ["Nama partner audit tahun berjalan / Name of current year audit signing partner", company_data["audit_partner"], "", ""],
            ["Auditor tahun sebelumnya / Prior year auditor", company_data["prior_auditor"], "", ""],
            ["Pembulatan yang digunakan dalam penyajian laporan keuangan / Rounding used in presentation of financial statements", company_data["rounding"], "", ""],
            ["Tanggal akhir periode berjalan / End date of current period", company_data["period"], "", ""],
            ["Tanggal akhir tahun sebelumnya / End date of prior year", "31 Desember 2022", "", ""]
        ]
        df_info = pd.DataFrame(info_rows)
        df_info.to_excel(writer, sheet_name=sheet_info, index=False, header=False)

        # 2. Sheet Posisi Keuangan (Neraca / Balance Sheet)
        balance_rows = [
            ["Laporan Posisi Keuangan / Statement of Financial Position", "", "31 Desember 2023", "31 Desember 2022"],
            ["Aset / Assets", "", "", ""],
            ["Jumlah aset lancar / Total current assets", "", company_data["total_assets"] * 0.45, company_data["total_assets_prior"] * 0.44],
            ["Jumlah aset tidak lancar / Total non-current assets", "", company_data["total_assets"] * 0.55, company_data["total_assets_prior"] * 0.56],
            ["Total aset / Total assets", "", company_data["total_assets"], company_data["total_assets_prior"]],
            ["Liabilitas dan Ekuitas / Liabilities and Equity", "", "", ""],
            ["Jumlah liabilitas / Total liabilities", "", company_data["total_assets"] * 0.52, company_data["total_assets_prior"] * 0.54],
            ["Jumlah ekuitas / Total equity", "", company_data["total_assets"] * 0.48, company_data["total_assets_prior"] * 0.46],
            ["Total liabilitas dan ekuitas / Total liabilities and equity", "", company_data["total_assets"], company_data["total_assets_prior"]]
        ]
        df_balance = pd.DataFrame(balance_rows)
        df_balance.to_excel(writer, sheet_name=sheet_balance, index=False, header=False)

        # 3. Sheet Laba Rugi (Income Statement)
        income_rows = [
            ["Laporan Laba Rugi Komprehensif / Statement of Profit or Loss", "", "31 Desember 2023", "31 Desember 2022"],
            ["Pendapatan usaha / Revenues", "", abs(company_data["net_income"]) * 3.5, abs(company_data["net_income_prior"]) * 3.4],
            ["Beban pokok pendapatan / Cost of revenues", "", abs(company_data["net_income"]) * 2.1, abs(company_data["net_income_prior"]) * 2.0],
            ["Laba kotor / Gross profit", "", abs(company_data["net_income"]) * 1.4, abs(company_data["net_income_prior"]) * 1.4],
            ["Beban usaha / Operating expenses", "", abs(company_data["net_income"]) * 0.35, abs(company_data["net_income_prior"]) * 0.36],
            ["Laba (rugi) sebelum pajak / Profit (loss) before tax", "", company_data["net_income"] * 1.25, company_data["net_income_prior"] * 1.25],
            ["Beban pajak penghasilan / Income tax expense", "", abs(company_data["net_income"]) * 0.25, abs(company_data["net_income_prior"]) * 0.25],
            ["Laba (rugi) periode berjalan / Profit (loss) for the period", "", company_data["net_income"], company_data["net_income_prior"]],
            ["Total laba (rugi) komprehensif periode berjalan / Total comprehensive income", "", company_data["net_income"], company_data["net_income_prior"]]
        ]
        df_income = pd.DataFrame(income_rows)
        df_income.to_excel(writer, sheet_name=sheet_income, index=False, header=False)

    output.seek(0)
    return output

def generate_sample_files_on_disk(target_dir: str) -> List[str]:
    """Menyimpan file sampel Excel IDX ke folder target."""
    os.makedirs(target_dir, exist_ok=True)
    generated_paths = []
    for comp in SAMPLE_COMPANIES:
        file_bytes = create_idx_excel(comp)
        file_path = os.path.join(target_dir, f"FinancialStatement-2023-Tahunan-{comp['ticker']}.xlsx")
        with open(file_path, "wb") as f:
            f.write(file_bytes.getvalue())
        generated_paths.append(file_path)
    return generated_paths
