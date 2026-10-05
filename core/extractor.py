import re
import math
import numpy as np
import pandas as pd
from typing import Dict, Any, List, Optional, Tuple

class FinancialExtractor:
    """
    Ekstraktor Laporan Keuangan Terstruktur / e-XBRL IDX.
    Mengekstraksi:
    - Informasi Entitas (Kode Ticker, Nama Perusahaan, Periode)
    - Informasi Auditor & KAP (KAP Berjalan, KAP Sebelumnya, Partner Audit, Status Rotasi)
    - Data Keuangan (Total Aset, Laba Bersih, Faktor Pembulatan)
    - Perhitungan ROA (Return on Assets) dan Analisis Kesehatan
    """

    # Identifikasi Sheet Informasi Umum
    GENERAL_INFO_SHEET_PATTERNS = [
        r"^1000000",
        r"informasi\s*umum",
        r"general\s*information",
        r"info\s*perusahaan",
        r"profil\s*perusahaan",
        r"entity\s*information"
    ]

    # Identifikasi Sheet Posisi Keuangan (Neraca / Balance Sheet)
    BALANCE_SHEET_PATTERNS = [
        r"^1110000",
        r"^1210000",
        r"posisi\s*keuangan",
        r"financial\s*position",
        r"neraca",
        r"balance\s*sheet"
    ]

    # Identifikasi Sheet Laba Rugi (Income Statement)
    INCOME_STATEMENT_PATTERNS = [
        r"^1220000",
        r"^1311000",
        r"^1312000",
        r"^1321000",
        r"^1322000",
        r"laba\s*rugi",
        r"profit\s*or\s*loss",
        r"income\s*statement",
        r"komprehensif"
    ]

    # Mapping Faktor Pembulatan (Urutkan dari spesifik / panjang ke umum)
    ROUNDING_PATTERNS = [
        (r"triliun|trillions|dalam\s*triliun", 1000000000000, "Dalam Triliun Rupiah"),
        (r"miliaran|billions|dalam\s*miliaran", 1000000000, "Dalam Miliaran Rupiah"),
        (r"jutaan|millions|dalam\s*jutaan", 1000000, "Dalam Jutaan Rupiah"),
        (r"ribuan|thousands|dalam\s*ribuan", 1000, "Dalam Ribuan Rupiah"),
        (r"satuan\s*penuh|full\s*amount|penuh|dalam\s*rupiah|rupiah\s*penuh", 1, "Satuan Penuh")
    ]

    def __init__(self, file_path_or_bytes, filename: str = ""):
        self.file_source = file_path_or_bytes
        self.filename = filename
        self.excel_file = None
        self.sheet_names = []
        self.sheets_data: Dict[str, pd.DataFrame] = {}

    def load_file(self):
        """Membaca file Excel dan memuat semua sheet ke dalam memori."""
        self.excel_file = pd.ExcelFile(self.file_source)
        self.sheet_names = self.excel_file.sheet_names
        for sheet in self.sheet_names:
            try:
                # Baca sheet tanpa header awal agar fleksibel menemukan label di baris mana saja
                df = self.excel_file.parse(sheet, header=None)
                self.sheets_data[sheet] = df
            except Exception as e:
                self.sheets_data[sheet] = pd.DataFrame()

    def _find_sheet_by_pattern(self, patterns: List[str]) -> Optional[str]:
        """Mencari nama sheet yang cocok dengan regex pola."""
        for pattern in patterns:
            for sheet in self.sheet_names:
                if re.search(pattern, sheet, re.IGNORECASE):
                    return sheet
        return None

    def _clean_text(self, val: Any) -> str:
        """Membersihkan teks dari spasi berlebih dan karakter khusus."""
        if val is None or pd.isna(val):
            return ""
        return re.sub(r"\s+", " ", str(val)).strip()

    def _parse_numeric(self, val: Any) -> Optional[float]:
        """Mengubah berbagai format teks angka menjadi float."""
        if val is None or pd.isna(val):
            return None
        if isinstance(val, (int, float)):
            if math.isnan(val) or math.isinf(val):
                return None
            return float(val)
        
        s = str(val).strip()
        if not s or s.lower() in ["-", "–", "—", "nil", "null", "none", "n/a"]:
            return 0.0

        # Tangani format tanda kurung negatif, misal: (1.234.567) atau (1,234.56)
        is_negative = False
        if s.startswith("(") and s.endswith(")"):
            is_negative = True
            s = s[1:-1].strip()
        elif s.startswith("-"):
            is_negative = True
            s = s[1:].strip()

        # Bersihkan mata uang & simbol
        s = re.sub(r"[^\d,\.]", "", s)
        if not s:
            return None

        # Deteksi format pemisah ribuan & desimal (Indonesia: 1.000,50 vs US: 1,000.50)
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                # Format Indonesia: titik = ribuan, koma = desimal
                s = s.replace(".", "").replace(",", ".")
            else:
                # Format US: koma = ribuan, titik = desimal
                s = s.replace(",", "")
        elif "," in s:
            # Jika koma di 2-3 digit terakhir, mungkin desimal; jika 3 digit tepat dari akhir dan panjang, bisa ribuan
            parts = s.split(",")
            if len(parts) == 2 and len(parts[1]) in [1, 2]:
                s = s.replace(",", ".")
            else:
                s = s.replace(",", "")
        
        try:
            val_float = float(s)
            return -val_float if is_negative else val_float
        except ValueError:
            return None

    def _search_label_in_df(self, df: pd.DataFrame, label_patterns: List[str], max_col_scan: int = 4) -> List[Tuple[int, int, str]]:
        """Mencari posisi sel (baris, kolom, teks) yang cocok dengan pola label."""
        matches = []
        if df.empty:
            return matches

        cols_to_check = min(len(df.columns), max_col_scan)
        for r_idx in range(len(df)):
            for c_idx in range(cols_to_check):
                cell_val = self._clean_text(df.iloc[r_idx, c_idx])
                if not cell_val:
                    continue
                for pattern in label_patterns:
                    if re.search(pattern, cell_val, re.IGNORECASE):
                        matches.append((r_idx, c_idx, cell_val))
                        break
        return matches

    def _get_row_value_after_col(self, df: pd.DataFrame, r_idx: int, c_idx: int, prefer_numeric: bool = False) -> Tuple[Any, Optional[int]]:
        """Mengambil nilai data pada baris yang sama setelah kolom label."""
        if r_idx >= len(df):
            return None, None

        row = df.iloc[r_idx]
        candidates = []
        
        # Periksa kolom setelah c_idx
        for col_idx in range(c_idx + 1, len(row)):
            raw_val = row.iloc[col_idx]
            cleaned = self._clean_text(raw_val)
            if not cleaned or cleaned.lower() in ["nil", "null", "none"]:
                continue
            
            num_val = self._parse_numeric(raw_val)
            if prefer_numeric:
                if num_val is not None:
                    return num_val, col_idx
            else:
                # Untuk teks informasi umum (nama KAP, partner, entitas)
                # Lewati jika hanya tanda titik dua ":" atau deskripsi duplikat dalam bahasa inggris
                if cleaned in [":", "-", "="] or len(cleaned) <= 1:
                    continue
                candidates.append((cleaned, col_idx))
        
        if candidates:
            return candidates[0][0], candidates[0][1]
        return None, None

    def extract_general_info(self) -> Dict[str, Any]:
        """Mengekstraksi Informasi Umum: Entitas, Ticker, KAP, Partner Audit, Pembulatan."""
        info = {
            "entity_name": None,
            "ticker": None,
            "current_auditor": None,
            "audit_partner": None,
            "prior_auditor": None,
            "auditor_switching": "Tidak Diketahui",
            "rounding_factor": 1,
            "rounding_text": "Satuan Penuh",
            "period": None,
            "source_sheet": None,
            "debug_trace": {}
        }

        # Cari sheet informasi umum (default IDX: 1000000)
        target_sheet = self._find_sheet_by_pattern(self.GENERAL_INFO_SHEET_PATTERNS)
        sheets_to_try = [target_sheet] if target_sheet else []
        # Tambahkan sheet pertama dan sheet lainnya sebagai fallback
        for s in self.sheet_names:
            if s not in sheets_to_try:
                sheets_to_try.append(s)

        # Pola kata kunci
        patterns_map = {
            "entity_name": [
                r"nama\s*entitas",
                r"entity\s*name",
                r"nama\s*perusahaan",
                r"company\s*name",
                r"nama\s*emiten"
            ],
            "ticker": [
                r"kode\s*entitas",
                r"entity\s*code",
                r"kode\s*saham",
                r"ticker\s*code",
                r"stock\s*code",
                r"simbol\s*saham"
            ],
            "current_auditor": [
                r"auditor\s*tahun\s*berjalan",
                r"current\s*year\s*auditor",
                r"kantor\s*akuntan\s*publik\s*tahun\s*berjalan",
                r"nama\s*kantor\s*akuntan\s*publik",
                r"nama\s*kap\b",
                r"independent\s*auditor",
                r"auditor\s*independen"
            ],
            "audit_partner": [
                r"nama\s*partner\s*audit\s*tahun\s*berjalan",
                r"name\s*of\s*current\s*year\s*audit\s*signing\s*partner",
                r"partner\s*audit",
                r"rekan\s*audit",
                r"akuntan\s*publik\s*yang\s*menandatangani",
                r"signing\s*partner",
                r"nama\s*akuntan\s*publik"
            ],
            "prior_auditor": [
                r"auditor\s*tahun\s*sebelumnya",
                r"prior\s*year\s*auditor",
                r"kantor\s*akuntan\s*publik\s*tahun\s*lalu",
                r"kap\s*tahun\s*lalu",
                r"auditor\s*tahun\s*lalu"
            ],
            "rounding": [
                r"pembulatan\s*yang\s*digunakan",
                r"rounding\s*used\s*in\s*presentation",
                r"tingkat\s*pembulatan",
                r"dalam\s*ribuan|dalam\s*jutaan|dalam\s*miliaran",
                r"presentation\s*currency"
            ],
            "period": [
                r"tanggal\s*akhir\s*periode\s*berjalan",
                r"end\s*date\s*of\s*current\s*period",
                r"tanggal\s*akhir\s*tahun\s*berjalan",
                r"periode\s*laporan\s*keuangan",
                r"financial\s*statement\s*period"
            ]
        }

        # Scan sheet
        for sheet_name in sheets_to_try:
            if not sheet_name or sheet_name not in self.sheets_data:
                continue
            df = self.sheets_data[sheet_name]
            if df.empty:
                continue

            # Cari setiap elemen
            for key, patterns in patterns_map.items():
                if key == "rounding" and info.get("rounding_factor", 1) != 1:
                    continue
                if key != "rounding" and info.get(key) is not None:
                    continue
                matches = self._search_label_in_df(df, patterns, max_col_scan=4)
                for r_idx, c_idx, matched_label in matches:
                    val, found_col = self._get_row_value_after_col(df, r_idx, c_idx, prefer_numeric=False)
                    if val is not None:
                        val_str = str(val).strip()
                        if key == "rounding":
                            val_lower = val_str.lower()
                            matched_rounding = False
                            for pattern_rx, factor, label_desc in self.ROUNDING_PATTERNS:
                                if re.search(pattern_rx, val_lower, re.IGNORECASE):
                                    info["rounding_factor"] = factor
                                    info["rounding_text"] = val_str if val_str else label_desc
                                    matched_rounding = True
                                    break
                            if matched_rounding:
                                break
                        elif key == "ticker":
                            # Bersihkan ticker, ambil 4 karakter uppercase jika ada
                            clean_t = re.sub(r"[^A-Za-z0-9]", "", val_str).upper()
                            info["ticker"] = clean_t[:6] if clean_t else val_str
                            info["debug_trace"]["ticker_row"] = (sheet_name, r_idx, matched_label)
                            break
                        else:
                            info[key] = val_str
                            info["debug_trace"][f"{key}_row"] = (sheet_name, r_idx, matched_label)
                            break

            if info["entity_name"] and info["current_auditor"]:
                info["source_sheet"] = sheet_name
                break

        # Fallback Ticker dari nama file jika belum ditemukan
        if not info["ticker"] and self.filename:
            ticker_match = re.search(r"[-_]([A-Z]{4})[-_\.]", self.filename)
            if ticker_match:
                info["ticker"] = ticker_match.group(1)
            else:
                ticker_match2 = re.search(r"([A-Z]{4})", self.filename)
                if ticker_match2:
                    info["ticker"] = ticker_match2.group(1)

        # Fallback Entity Name jika kosong tapi ticker ada
        if not info["entity_name"] and info["ticker"]:
            info["entity_name"] = f"Entitas {info['ticker']}"

        # Analisis Auditor Switching (Pergantian KAP)
        curr_aud = (info["current_auditor"] or "").strip().lower()
        prior_aud = (info["prior_auditor"] or "").strip().lower()

        if curr_aud and prior_aud:
            # Bandingkan kemiripan / kesamaan
            clean_curr = re.sub(r"[^\w]", "", curr_aud)
            clean_prior = re.sub(r"[^\w]", "", prior_aud)
            if clean_curr == clean_prior:
                info["auditor_switching"] = "Tetap (Sama)"
            else:
                # Cek jika salah satu merupakan sub-string dari lainnya (misal beda penulisan nama rekan)
                if clean_curr in clean_prior or clean_prior in clean_curr:
                    info["auditor_switching"] = "Tetap (Afiliasi Sama)"
                else:
                    info["auditor_switching"] = "Berganti (Rotasi KAP)"
        elif curr_aud and not prior_aud:
            info["auditor_switching"] = "Tetap (Tidak Ada Catatan Rotasi)"
        else:
            info["auditor_switching"] = "Tidak Diketahui"

        return info

    def extract_balance_sheet(self) -> Dict[str, Any]:
        """Mengekstraksi Total Aset dari Laporan Posisi Keuangan (Neraca)."""
        res = {
            "total_assets_raw": None,
            "total_assets_prior_raw": None,
            "source_sheet": None,
            "matched_row": None,
            "matched_label": None
        }

        # Cari sheet posisi keuangan (default IDX: 1110000 atau 1210000)
        target_sheet = self._find_sheet_by_pattern(self.BALANCE_SHEET_PATTERNS)
        sheets_to_try = [target_sheet] if target_sheet else []
        for s in self.sheet_names:
            if s not in sheets_to_try:
                sheets_to_try.append(s)

        # Pola prioritas tinggi untuk Total Aset
        primary_patterns = [
            r"^(jumlah\s*aset|total\s*aset|total\s*assets|jumlah\s*aktiva|total\s*aktiva)$",
            r"^(jumlah|total)\s*aset\s*\/",
            r"^(total|jumlah)\s*assets\s*\/",
            r"total\s*aset\s*lancar\s*dan\s*tidak\s*lancar",
            r"total\s*aset\b",
            r"total\s*assets\b",
            r"jumlah\s*aset\b"
        ]

        # Pola negatif yang TIDAK boleh dianggap Total Aset
        negative_patterns = [
            r"aset\s*lancar",
            r"current\s*assets",
            r"aset\s*tidak\s*lancar",
            r"non-current\s*assets",
            r"aset\s*tetap",
            r"fixed\s*assets",
            r"aset\s*pajak",
            r"tax\s*assets",
            r"aset\s*keuangan",
            r"financial\s*assets",
            r"aset\s*biologis",
            r"aset\s*lainnya",
            r"aset\s*hak\s*guna",
            r"aset\s*takberwujud"
        ]

        for sheet_name in sheets_to_try:
            if not sheet_name or sheet_name not in self.sheets_data:
                continue
            df = self.sheets_data[sheet_name]
            if df.empty:
                continue

            for r_idx in range(len(df)):
                row_labels = []
                for c_idx in range(min(4, len(df.columns))):
                    cell_val = self._clean_text(df.iloc[r_idx, c_idx])
                    if cell_val:
                        row_labels.append(cell_val)
                
                full_row_label = " ".join(row_labels).strip()
                if not full_row_label:
                    continue

                # Cek apakah terkena negative patterns
                is_negative = False
                for neg in negative_patterns:
                    if re.search(neg, full_row_label, re.IGNORECASE):
                        # Kecuali jika jelas ada kata "jumlah aset" atau "total aset" secara spesifik
                        if not re.search(r"^(jumlah|total)\s*aset\s*(\/|\(|$)", full_row_label, re.IGNORECASE):
                            is_negative = True
                            break
                if is_negative:
                    continue

                # Cek apakah cocok dengan primary patterns
                matched = False
                for pattern in primary_patterns:
                    if re.search(pattern, full_row_label, re.IGNORECASE):
                        matched = True
                        break

                if matched:
                    # Ambil angka di baris ini
                    numeric_cells = []
                    for c_idx in range(min(4, len(df.columns)), len(df.columns)):
                        val = df.iloc[r_idx, c_idx]
                        num = self._parse_numeric(val)
                        if num is not None:
                            numeric_cells.append(num)
                    
                    # Jika tidak dapat dari col 4 ke atas, scan semua kolom
                    if not numeric_cells:
                        for c_idx in range(len(df.columns)):
                            val = df.iloc[r_idx, c_idx]
                            num = self._parse_numeric(val)
                            if num is not None and abs(num) > 100: # hindari indeks baris kecil
                                numeric_cells.append(num)

                    if numeric_cells:
                        res["total_assets_raw"] = numeric_cells[0]
                        if len(numeric_cells) > 1:
                            res["total_assets_prior_raw"] = numeric_cells[1]
                        res["source_sheet"] = sheet_name
                        res["matched_row"] = r_idx + 1
                        res["matched_label"] = full_row_label
                        return res

        return res

    def extract_income_statement(self) -> Dict[str, Any]:
        """Mengekstraksi Laba Bersih (Net Income) dari Laporan Laba Rugi."""
        res = {
            "net_income_raw": None,
            "net_income_prior_raw": None,
            "source_sheet": None,
            "matched_row": None,
            "matched_label": None
        }

        # Cari sheet laba rugi (default IDX: 1220000, 1311000, atau 1321000)
        target_sheet = self._find_sheet_by_pattern(self.INCOME_STATEMENT_PATTERNS)
        sheets_to_try = [target_sheet] if target_sheet else []
        for s in self.sheet_names:
            if s not in sheets_to_try:
                sheets_to_try.append(s)

        primary_patterns = [
            r"^(laba\s*\(rugi\)\s*periode\s*berjalan|profit\s*\(loss\)\s*for\s*the\s*period)$",
            r"laba\s*\(rugi\)\s*periode\s*berjalan\b",
            r"laba\s*\(rugi\)\s*tahun\s*berjalan\b",
            r"profit\s*\(loss\)\s*for\s*the\s*(period|year)\b",
            r"laba\s*bersih\s*periode\s*berjalan\b",
            r"laba\s*bersih\s*tahun\s*berjalan\b",
            r"net\s*profit\s*for\s*the\s*period\b",
            r"net\s*income\s*for\s*the\s*period\b",
            r"laba\s*\(rugi\)\s*bersih\b",
            r"laba\s*bersih\b",
            r"pemilik\s*entitas\s*induk" # fallback jika diatribusikan ke pemilik entitas induk
        ]

        negative_patterns = [
            r"sebelum\s*pajak",
            r"before\s*tax",
            r"usaha",
            r"operasional",
            r"operating",
            r"kotor",
            r"gross",
            r"per\s*saham",
            r"per\s*share",
            r"komprehensif\s*lain",
            r"other\s*comprehensive"
        ]

        for sheet_name in sheets_to_try:
            if not sheet_name or sheet_name not in self.sheets_data:
                continue
            df = self.sheets_data[sheet_name]
            if df.empty:
                continue

            for r_idx in range(len(df)):
                row_labels = []
                for c_idx in range(min(4, len(df.columns))):
                    cell_val = self._clean_text(df.iloc[r_idx, c_idx])
                    if cell_val:
                        row_labels.append(cell_val)
                
                full_row_label = " ".join(row_labels).strip()
                if not full_row_label:
                    continue

                # Cek negatif
                is_negative = False
                for neg in negative_patterns:
                    if re.search(neg, full_row_label, re.IGNORECASE):
                        is_negative = True
                        break
                if is_negative:
                    continue

                matched = False
                for pattern in primary_patterns:
                    if re.search(pattern, full_row_label, re.IGNORECASE):
                        matched = True
                        break

                if matched:
                    numeric_cells = []
                    for c_idx in range(min(4, len(df.columns)), len(df.columns)):
                        val = df.iloc[r_idx, c_idx]
                        num = self._parse_numeric(val)
                        if num is not None:
                            numeric_cells.append(num)

                    if not numeric_cells:
                        for c_idx in range(len(df.columns)):
                            val = df.iloc[r_idx, c_idx]
                            num = self._parse_numeric(val)
                            if num is not None:
                                numeric_cells.append(num)

                    if numeric_cells:
                        res["net_income_raw"] = numeric_cells[0]
                        if len(numeric_cells) > 1:
                            res["net_income_prior_raw"] = numeric_cells[1]
                        res["source_sheet"] = sheet_name
                        res["matched_row"] = r_idx + 1
                        res["matched_label"] = full_row_label
                        return res

        return res

    def analyze(self) -> Dict[str, Any]:
        """Melakukan ekstraksi lengkap dan menghitung metrik ROA."""
        self.load_file()

        general_info = self.extract_general_info()
        balance_info = self.extract_balance_sheet()
        income_info = self.extract_income_statement()

        rounding_factor = general_info.get("rounding_factor", 1)
        total_assets_raw = balance_info.get("total_assets_raw")
        net_income_raw = income_info.get("net_income_raw")

        # Hitung Nilai Terskalakan (Nominal Riil Rupiah)
        total_assets_scaled = (total_assets_raw * rounding_factor) if total_assets_raw is not None else None
        net_income_scaled = (net_income_raw * rounding_factor) if net_income_raw is not None else None

        # Perhitungan ROA:
        # ROA (%) = (Laba Bersih / Total Aset) * 100%
        # Catatan: Karena Laba Bersih & Total Aset berskala sama dari laporan yang sama, rasio mentah maupun terskalakan identik
        roa_pct = None
        roa_status = "Tidak Lengkap"
        roa_badge_color = "gray"
        roa_formula_display = "Data tidak memadai"

        if total_assets_raw is not None and net_income_raw is not None:
            if total_assets_raw != 0:
                roa_pct = (net_income_raw / total_assets_raw) * 100.0
                roa_formula_display = f"({net_income_raw:,.2f} / {total_assets_raw:,.2f}) × 100% = {roa_pct:.2f}%"
                
                # Klasifikasi Kesehatan ROA (Standar Analisis Keuangan & OJK/BI)
                if roa_pct >= 10.0:
                    roa_status = "Sangat Unggul (≥ 10%)"
                    roa_badge_color = "emerald"
                elif roa_pct >= 5.0:
                    roa_status = "Sehat / Baik (5% - 10%)"
                    roa_badge_color = "teal"
                elif roa_pct >= 2.0:
                    roa_status = "Cukup / Standar (2% - 5%)"
                    roa_badge_color = "blue"
                elif roa_pct >= 0.0:
                    roa_status = "Rendah / Kurang (0% - 2%)"
                    roa_badge_color = "amber"
                else:
                    roa_status = "Rugi / Negatif (< 0%)"
                    roa_badge_color = "rose"
            else:
                roa_status = "Total Aset Nol"
                roa_badge_color = "gray"
        elif total_assets_raw is None:
            roa_status = "Total Aset Tidak Ditemukan"
            roa_badge_color = "amber"
        elif net_income_raw is None:
            roa_status = "Laba Bersih Tidak Ditemukan"
            roa_badge_color = "amber"

        # Deteksi status Big Four KAP
        big_four_keywords = ["pwc", "pricewaterhouse", "tanudiredja", "kpmg", "siddharta", "ey", "ernst", "purwantono", "deloitte", "imelda", "satrio"]
        curr_aud_lower = (general_info.get("current_auditor") or "").lower()
        is_big_four = any(k in curr_aud_lower for k in big_four_keywords)

        result = {
            "filename": self.filename,
            "ticker": general_info.get("ticker") or "N/A",
            "entity_name": general_info.get("entity_name") or "Perusahaan Tidak Dikenal",
            "period": general_info.get("period") or "31 Desember (Tahunan)",
            
            # Auditor & KAP
            "current_auditor": general_info.get("current_auditor") or "Tidak Disebutkan",
            "prior_auditor": general_info.get("prior_auditor") or "Tidak Disebutkan",
            "audit_partner": general_info.get("audit_partner") or "Tidak Disebutkan",
            "auditor_switching": general_info.get("auditor_switching", "Tidak Diketahui"),
            "is_big_four": is_big_four,
            "kap_tier": "Big Four" if is_big_four else ("Non-Big Four" if general_info.get("current_auditor") else "N/A"),
            
            # Data Keuangan
            "rounding_text": general_info.get("rounding_text", "Satuan Penuh"),
            "rounding_factor": rounding_factor,
            
            "total_assets_raw": total_assets_raw,
            "total_assets_scaled": total_assets_scaled,
            "total_assets_prior_raw": balance_info.get("total_assets_prior_raw"),
            
            "net_income_raw": net_income_raw,
            "net_income_scaled": net_income_scaled,
            "net_income_prior_raw": income_info.get("net_income_prior_raw"),
            
            # ROA
            "roa_percent": round(roa_pct, 4) if roa_pct is not None else None,
            "roa_formatted": f"{roa_pct:.2f}%" if roa_pct is not None else "N/A",
            "roa_status": roa_status,
            "roa_badge_color": roa_badge_color,
            "roa_formula": roa_formula_display,
            
            # Traceability Audit
            "audit_trail": {
                "general_sheet": general_info.get("source_sheet"),
                "balance_sheet": balance_info.get("source_sheet"),
                "balance_row_matched": balance_info.get("matched_row"),
                "balance_label_matched": balance_info.get("matched_label"),
                "income_sheet": income_info.get("source_sheet"),
                "income_row_matched": income_info.get("matched_row"),
                "income_label_matched": income_info.get("matched_label"),
                "total_sheets_analyzed": len(self.sheet_names),
                "sheet_names": self.sheet_names
            }
        }

        return result
