import requests
import json

def test_full_pipeline():
    base_url = "http://127.0.0.1:8000"
    
    print("1. Testing GET /")
    r_index = requests.get(f"{base_url}/")
    assert r_index.status_code == 200
    assert "IDX Financial Statement & ROA Analyzer" in r_index.text
    print("   -> OK (Status 200)")

    print("2. Testing GET /api/demo-data")
    r_demo = requests.get(f"{base_url}/api/demo-data")
    assert r_demo.status_code == 200
    data = r_demo.json()
    assert data["success"] is True
    assert len(data["data"]) == 6
    print(f"   -> OK ({len(data['data'])} companies extracted)")

    print("3. Testing POST /api/upload (Multi-file upload)")
    files = [
        ("files", ("FinancialStatement-BBCA.xlsx", open("samples/FinancialStatement-2023-Tahunan-BBCA.xlsx", "rb"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")),
        ("files", ("FinancialStatement-ASII.xlsx", open("samples/FinancialStatement-2023-Tahunan-ASII.xlsx", "rb"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
    ]
    r_upload = requests.post(f"{base_url}/api/upload", files=files)
    assert r_upload.status_code == 200
    up_res = r_upload.json()
    assert up_res["processed_count"] == 2
    for item in up_res["data"]:
        print(f"   -> Extracted: {item['ticker']} ({item['entity_name']}) | KAP: {item['current_auditor']} | ROA: {item['roa_formatted']} | Rotasi: {item['auditor_switching']}")

    print("4. Testing POST /api/export/excel")
    r_excel = requests.post(f"{base_url}/api/export/excel", json={"data": up_res["data"]})
    assert r_excel.status_code == 200
    assert len(r_excel.content) > 5000
    print(f"   -> OK (Excel generated, {len(r_excel.content)} bytes)")

    print("5. Testing POST /api/export/csv")
    r_csv = requests.post(f"{base_url}/api/export/csv", json={"data": up_res["data"]})
    assert r_csv.status_code == 200
    assert "Kode Saham" in r_csv.text
    print(f"   -> OK (CSV generated, {len(r_csv.text.splitlines())} lines)")

    print("6. Testing GET /api/download-sample/TLKM")
    r_sample = requests.get(f"{base_url}/api/download-sample/TLKM")
    assert r_sample.status_code == 200
    assert len(r_sample.content) > 3000
    print(f"   -> OK (Sample Excel downloaded, {len(r_sample.content)} bytes)")

    print("\nALL API PIPELINE TESTS PASSED 100% SUCCESSFULLY!")

if __name__ == "__main__":
    test_full_pipeline()
