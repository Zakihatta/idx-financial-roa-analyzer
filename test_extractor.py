import json
from core.sample_data import SAMPLE_COMPANIES, create_idx_excel
from core.extractor import FinancialExtractor

def test_all():
    print("Testing extraction on all sample IDX companies...")
    for comp in SAMPLE_COMPANIES:
        excel_buf = create_idx_excel(comp)
        extractor = FinancialExtractor(excel_buf, filename=f"FinancialStatement-2023-Tahunan-{comp['ticker']}.xlsx")
        res = extractor.analyze()
        print(f"\n--- {res['ticker']} ({res['entity_name']}) ---")
        print(f"  Period           : {res['period']}")
        print(f"  Auditor (KAP)    : {res['current_auditor']} ({res['kap_tier']})")
        print(f"  Audit Partner    : {res['audit_partner']}")
        print(f"  Prior Auditor    : {res['prior_auditor']}")
        print(f"  Auditor Switching: {res['auditor_switching']}")
        print(f"  Total Assets     : {res['total_assets_raw']:,.0f} ({res['rounding_text']}) -> Real: Rp {res['total_assets_scaled']:,.0f}")
        print(f"  Net Income       : {res['net_income_raw']:,.0f} -> Real: Rp {res['net_income_scaled']:,.0f}")
        print(f"  ROA Result       : {res['roa_formatted']} ({res['roa_status']})")
        print(f"  Formula          : {res['roa_formula']}")

if __name__ == "__main__":
    test_all()
