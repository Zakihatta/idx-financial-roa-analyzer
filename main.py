import os
import io
import json
from typing import List, Optional
from fastapi import FastAPI, File, UploadFile, HTTPException, Body
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from core.extractor import FinancialExtractor
from core.sample_data import SAMPLE_COMPANIES, create_idx_excel
from core.exporter import export_to_excel, export_to_csv

app = FastAPI(
    title="IDX Financial Statement & ROA Analyzer",
    description="Sistem Analisis Laporan Keuangan IDX e-XBRL & Kalkulator ROA Otomatis",
    version="1.0.0"
)

# CORS Middleware jika dibutuhkan
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Direktori dasar
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
SAMPLES_DIR = os.path.join(BASE_DIR, "samples")

os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(SAMPLES_DIR, exist_ok=True)

# Mount folder static
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    """Menyajikan halaman utama dashboard analitik."""
    index_path = os.path.join(STATIC_DIR, "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path)
    return HTMLResponse("<h2>Aplikasi sedang memuat aset frontend... Silakan muat ulang.</h2>")

@app.post("/api/upload")
async def upload_financial_statements(files: List[UploadFile] = File(...)):
    """
    Menerima unggahan satu atau banyak file Excel (.xlsx, .xls)
    dan mengekstraksi informasi umum, KAP, partner audit, serta perhitungan ROA.
    """
    if not files:
        raise HTTPException(status_code=400, detail="Tidak ada file yang diunggah.")

    results = []
    errors = []

    for file in files:
        filename = file.filename
        if not (filename.endswith(".xlsx") or filename.endswith(".xls")):
            errors.append({"filename": filename, "error": "Format file tidak didukung. Harap unggah .xlsx atau .xls"})
            continue

        try:
            content = await file.read()
            extractor = FinancialExtractor(io.BytesIO(content), filename=filename)
            analysis = extractor.analyze()
            results.append(analysis)
        except Exception as e:
            errors.append({
                "filename": filename,
                "error": f"Gagal membaca file: {str(e)}"
            })

    return {
        "success": True,
        "total_files": len(files),
        "processed_count": len(results),
        "error_count": len(errors),
        "data": results,
        "errors": errors
    }

@app.get("/api/demo-data")
async def get_demo_data():
    """
    Menghasilkan data demo otomatis dari contoh emiten BEI (BBCA, BBRI, TLKM, ASII, ICBP, GOTO).
    """
    results = []
    for comp in SAMPLE_COMPANIES:
        try:
            excel_bytes = create_idx_excel(comp)
            filename = f"FinancialStatement-2023-Tahunan-{comp['ticker']}.xlsx"
            extractor = FinancialExtractor(excel_bytes, filename=filename)
            analysis = extractor.analyze()
            results.append(analysis)
        except Exception as e:
            print(f"Error generating demo for {comp['ticker']}: {e}")

    return {
        "success": True,
        "total_files": len(results),
        "processed_count": len(results),
        "data": results
    }

@app.post("/api/export/excel")
async def export_excel_endpoint(payload: dict = Body(...)):
    """
    Mengunduh data hasil analisis ke file Excel (.xlsx) dengan formula dan format profesional.
    """
    data = payload.get("data", [])
    if not data:
        raise HTTPException(status_code=400, detail="Tidak ada data untuk diekspor.")

    try:
        excel_buffer = export_to_excel(data)
        filename = "Analisis_Laporan_Keuangan_ROA_IDX.xlsx"
        return StreamingResponse(
            excel_buffer,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal membuat file Excel: {str(e)}")

@app.post("/api/export/csv")
async def export_csv_endpoint(payload: dict = Body(...)):
    """
    Mengunduh data hasil analisis ke file CSV.
    """
    data = payload.get("data", [])
    if not data:
        raise HTTPException(status_code=400, detail="Tidak ada data untuk diekspor.")

    try:
        csv_content = export_to_csv(data)
        csv_buffer = io.BytesIO(csv_content.encode("utf-8-sig"))
        filename = "Analisis_Laporan_Keuangan_ROA_IDX.csv"
        return StreamingResponse(
            csv_buffer,
            media_type="text/csv",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal membuat file CSV: {str(e)}")

@app.get("/api/download-sample/{ticker}")
async def download_sample_file(ticker: str):
    """
    Mengunduh file Excel contoh e-XBRL IDX asli untuk emiten tertentu.
    """
    ticker = ticker.upper()
    comp = next((c for c in SAMPLE_COMPANIES if c["ticker"] == ticker), None)
    if not comp:
        raise HTTPException(status_code=404, detail=f"Sampel emiten {ticker} tidak ditemukan.")

    excel_buffer = create_idx_excel(comp)
    filename = f"FinancialStatement-2023-Tahunan-{ticker}.xlsx"
    return StreamingResponse(
        excel_buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    host = os.environ.get("HOST", "0.0.0.0")
    uvicorn.run("main:app", host=host, port=port, reload=False)
