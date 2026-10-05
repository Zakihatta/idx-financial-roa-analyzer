import os
import sys
import webbrowser
import uvicorn

if __name__ == "__main__":
    print("=" * 65)
    print("  ALIFIL FINANCIAL INTELLIGENCE - IDX XBRL & ROA ANALYZER")
    print("=" * 65)
    print("  Memulai server FastAPI di: http://127.0.0.1:8000")
    print("  Membuka browser otomatis...")
    print("  Tekan Ctrl+C di terminal ini untuk berhenti.")
    print("=" * 65)
    
    try:
        webbrowser.open("http://127.0.0.1:8000")
    except Exception:
        pass

    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=False)
