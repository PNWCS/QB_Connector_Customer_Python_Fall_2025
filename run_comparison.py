"""Script to run customer comparison between Excel and QuickBooks."""

import sys
import os
from src.quickbook_connector.customer_excel_qb_sync import process_customers

if __name__ == "__main__":
    # Default Excel file in the current project directory
    default_excel_file = os.path.join(os.path.dirname(__file__), "company_data.xlsx")

    # Allow optional custom path
    excel_file = sys.argv[1] if len(sys.argv) > 1 else default_excel_file

    if not os.path.exists(excel_file):
        print(f"Excel file not found: {excel_file}")
        print("Please ensure 'company_data.xlsx' exists in the project folder, or specify a path:")
        print("Example: python run_customer_comparison.py C:\\path\\to\\customers.xlsx")
        sys.exit(1)

    print(f"Starting customer comparison for: {excel_file}\n")

    try:
        result = process_customers(excel_file)

        print("\n=== Summary ===")
        print("Completed successfully!")
        print(f"- Matching customers (same ID & data): {result.matching_count}")
        print(f"- Conflicts (same ID, different data): {len(result.same_id_diff_data)}")
        print(f"- Customers only in Excel (added to QB): {len(result.only_in_excel)}")
        print(f"- Customers only in QB: {len(result.only_in_qb)}")

    except Exception as e:
        print(f"\n Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
