import pandas as pd
import GRIR

def test_po_numbers():
    """Test script to verify PO numbers are correct."""
    try:
        # Run the analysis
        summary_df = GRIR.run_analysis(
            "data/EKBE.xlsx",
            "data/EKPO.XLSX", 
            "data/email.xlsx",
            send_emails=False
        )
        
        print("✅ Analysis completed successfully!")
        print(f"Total records: {len(summary_df)}")
        print(f"Unique PO numbers: {len(summary_df['PO'].unique())}")
        print("\nFirst 10 PO numbers:")
        for po in summary_df['PO'].unique()[:10]:
            print(f"  - {po}")
        
        print("\nSample records:")
        print(summary_df[['PO', 'Line', 'Line/Shade', 'Description', 'Plant']].head(10))
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_po_numbers() 