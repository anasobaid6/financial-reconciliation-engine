"""
================================================================================
Financial Source Data Generator
Author: Expert Financial Systems & Python Automation Consultant
Description: Generates realistic multi-source financial transactional exports:
             1. gateway_charges.csv (Stripe-like Payment Gateway Export - 2,000 records)
             2. bank_ledger.csv (ERP / Bank Settlement Ledger - 2,000 records)
             Injects realistic enterprise discrepancies:
             - Fee mismatches (unexpected gateway surcharge)
             - Unreconciled missing funds (unsettled charges)
             - Delayed postings (settlement lag > 3 business days)
             - Matched transactions (exact reconciliation)
================================================================================
"""

from datetime import datetime, timedelta
import logging
from pathlib import Path
import random
import numpy as np
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
GATEWAY_FILE = BASE_DIR / "gateway_charges.csv"
BANK_FILE = BASE_DIR / "bank_ledger.csv"

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s]: %(message)s")
logger = logging.getLogger("FinDataGen")


def generate_financial_datasets(total_records: int = 2000) -> None:
    logger.info(f"Generating {total_records:,} financial transactional records...")
    random.seed(42)
    np.random.seed(42)

    # 1. Base transaction identifiers
    charge_ids = [f"ch_3N{i:06d}xL7qA" for i in range(100001, 100001 + total_records)]
    
    # 2. Created Timestamps (across Q1 2024)
    start_date = datetime(2024, 1, 1, 8, 30, 0)
    random_minutes = np.random.randint(0, 90 * 24 * 60, size=total_records)
    created_dates = [start_date + timedelta(minutes=int(m)) for m in random_minutes]

    # 3. Monetary values (Simulating B2B & high-volume eCommerce checkouts)
    # Gross amounts between $25.00 and $4,500.00
    gross_amounts = np.round(np.random.exponential(scale=180.0, size=total_records) + 25.0, 2)
    gross_amounts = np.clip(gross_amounts, 19.99, 4999.00)

    # Standard Stripe Fee: 2.9% + $0.30
    standard_fees = np.round(gross_amounts * 0.029 + 0.30, 2)
    gateway_nets = np.round(gross_amounts - standard_fees, 2)

    customers = [f"user_{random.randint(1000, 9999)}@clientcorp.com" for _ in range(total_records)]
    payment_methods = np.random.choice(["visa", "mastercard", "amex", "apple_pay"], size=total_records, p=[0.48, 0.32, 0.12, 0.08])

    # Build Gateway DataFrame
    gateway_df = pd.DataFrame({
        "Charge_ID": charge_ids,
        "Created_Date": [d.strftime("%Y-%m-%d %H:%M:%S") for d in created_dates],
        "Customer_Email": customers,
        "Payment_Method": payment_methods,
        "Gross_Amount": gross_amounts,
        "Gateway_Fee": standard_fees,
        "Net_Amount": gateway_nets,
        "Currency": ["USD"] * total_records,
    })

    # Build Bank Ledger by applying realistic settlement variations
    bank_records = []
    
    # Distribution of scenarios:
    # 82% Exact Match (Normal 2-day batch payout)
    # 7% Fee Discrepancy (e.g. international card +1.5% cross-border fee not logged in CRM)
    # 6% Unreconciled / Missing Funds (Charge successful on gateway, but payout withheld/missing)
    # 5% Delayed Posting (Payout delayed 5-9 days due to risk review or bank holidays)

    for idx in range(total_records):
        charge_id = charge_ids[idx]
        created_dt = created_dates[idx]
        net_amt = gateway_nets[idx]
        dice = random.random()

        if dice < 0.82:
            # SCENARIO A: EXACT MATCH (Standard 2-day settlement)
            posting_dt = created_dt + timedelta(days=2)
            bank_records.append({
                "Reference_ID": charge_id,
                "Posting_Date": posting_dt.strftime("%Y-%m-%d"),
                "Credit_Amount": net_amt,
                "Bank_Account": "JPM-Operating-8841",
                "Description": f"STRIPE TRANSFER {charge_id[-8:]}",
                "Currency": "USD",
            })

        elif dice < 0.89:
            # SCENARIO B: FEE MISMATCH (Extra 1.5% cross-border / dispute surcharge)
            # Gateway net is lower than expected internal ERP credit or vice versa
            fee_leak = round(random.uniform(2.50, 28.00), 2)
            altered_credit = round(net_amt - fee_leak, 2)
            posting_dt = created_dt + timedelta(days=2)
            bank_records.append({
                "Reference_ID": charge_id,
                "Posting_Date": posting_dt.strftime("%Y-%m-%d"),
                "Credit_Amount": altered_credit,
                "Bank_Account": "JPM-Operating-8841",
                "Description": f"STRIPE TRANSFER {charge_id[-8:]} ADJ",
                "Currency": "USD",
            })

        elif dice < 0.95:
            # SCENARIO C: UNRECONCILED MISSING FUNDS
            # The transaction occurred on Gateway, but no settlement credit ever appeared in bank
            # We omit this transaction from bank_ledger entirely to test missing fund detection!
            continue

        else:
            # SCENARIO D: DELAYED POSTING / LIQUIDITY MISMATCH
            # Normal net amount, but settlement delayed by 6-9 days
            delayed_days = random.randint(5, 9)
            posting_dt = created_dt + timedelta(days=delayed_days)
            bank_records.append({
                "Reference_ID": charge_id,
                "Posting_Date": posting_dt.strftime("%Y-%m-%d"),
                "Credit_Amount": net_amt,
                "Bank_Account": "JPM-Operating-8841",
                "Description": f"STRIPE LATE SETTLEMENT {charge_id[-8:]}",
                "Currency": "USD",
            })

    # Add a few unexpected bank credits (e.g., direct bank refund reverse or unidentified wire)
    for extra_id in range(1, 16):
        dummy_ref = f"ACH_DIRECT_{extra_id:04d}"
        bank_records.append({
            "Reference_ID": dummy_ref,
            "Posting_Date": "2024-02-15",
            "Credit_Amount": round(random.uniform(150.00, 1200.00), 2),
            "Bank_Account": "JPM-Operating-8841",
            "Description": f"MISC UNMAPPED SETTLEMENT {dummy_ref}",
            "Currency": "USD",
        })

    bank_df = pd.DataFrame(bank_records)

    # Save both files
    gateway_df.to_csv(GATEWAY_FILE, index=False)
    bank_df.to_csv(BANK_FILE, index=False)

    logger.info(f"Successfully exported Gateway Charges: {GATEWAY_FILE} ({len(gateway_df):,} rows)")
    logger.info(f"Successfully exported Bank Ledger: {BANK_FILE} ({len(bank_df):,} rows)")


if __name__ == "__main__":
    generate_financial_datasets(total_records=2000)
