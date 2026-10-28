# Execution & Upwork Visual Capture Guide
> How to run the reconciliation engine and capture visual assets of the styled Excel workbook that demonstrate elite financial engineering to CFOs and finance leaders.

---

## 🛠️ Step 1: Script Execution

```bash
# 1. Navigate to the project directory
cd C:\Users\Anas\.gemini\antigravity-ide\scratch\financial-reconciliation-engine

# 2. Generate the two source datasets (Gateway & Bank Ledger)
python generate_financial_sources.py

# 3. Run the automated reconciliation and Excel modeling engine
python reconcile_engine.py
```

Upon completion, open [`Executive_Reconciliation_Report.xlsx`](Executive_Reconciliation_Report.xlsx) in Microsoft Excel or Google Sheets.

---

## 📸 Step 2: Exact Areas to Capture for Upwork Portfolio

CFOs and enterprise financial controllers care deeply about aesthetics, structure, and dynamic financial integrity. Follow these framing tips:

### Screen 1: The C-Suite Dashboard (Tab 1: "Executive Summary")
* **Framing:** Zoom to **110%** in Excel. Frame rows 1 through 18.
* **Key Elements to Highlight:**
  1. **Corporate Title Banner:** Navy Blue (`#0F172A`) header with clean white Segoe UI typography and UTC timestamp.
  2. **The 4 KPI Scorecards:**
     * `TOTAL GATEWAY GROSS: $412,850.20`
     * `ACTUAL BANK DEPOSITS: $375,190.40`
     * `RECONCILIATION RATE: 81.9%`
     * `UNRECONCILED AT-RISK CAPITAL: $18,572.39`
  3. **The Discrepancy Breakdown Table:** Highlighting the distinct categories with count and exposure.
* **Pro-Tip:** Click on the **Reconciliation Rate** cell so Excel's **Formula Bar** shows:
  `=COUNTIF('All Reconciled Records'!J2:J2016, "EXACT_MATCH") / COUNTA('All Reconciled Records'!J2:J2016)`
  *This visual detail proves you build live, dynamic financial models rather than hardcoding static numbers.*

### Screen 2: Actionable Audit List (Tab 2: "Discrepancy Action List")
* **Framing:** Capture the top 15 rows of Tab 2.
* **Key Elements to Highlight:**
  1. **Color-Coded Alert Logic:**
     * **Soft Red Rows (`#FEE2E2` / Text `#991B1B`):** Demonstrating `UNRECONCILED_MISSING_FUNDS` with action: *"Open Stripe Payout Inquiry; Funds not settled"*.
     * **Soft Amber Rows (`#FEF3C7` / Text `#92400E`):** Demonstrating `FEE_DISCREPANCY` with action: *"Audit fee contract: Variance of $14.20 detected"*.
     * **Soft Blue Rows (`#E0F2FE` / Text `#0369A1`):** Demonstrating `TIMING_DELAY` (Settlement > 3 days).
  2. **Zero Truncation:** Notice that every column width is auto-fitted, with zero `###` display errors.

### Screen 3: Complete Audit Trail (Tab 3: "All Reconciled Records")
* **Framing:** Capture the clean zebra-striped layout showing columns from `Transaction ID` through `Reconciliation Status`.
* **Key Elements to Highlight:**
  * Auto-filter drop-downs enabled across row 1.
  * Top header row frozen (`Freeze Panes`), maintaining column visibility while scrolling through 2,000+ records.
  * Formatted monetary numbers (`$#,##0.00`) aligned cleanly to the right.

---

## 🎨 Step 3: Upwork Portfolio Cover Collage (16:9 Thumbnail)

Using Canva or Figma, combine:
* **Background:** Deep navy blue or slate gray (`#0F172A`).
* **Left Image:** Angled screenshot of the **Executive Summary KPI Cards**.
* **Right Image:** Angled screenshot of the **Discrepancy Action List with red/yellow highlights**.
* **Headline:** `Zero-Touch Payment Gateway Reconciliation Engine [Python + OpenPyXL]`.
* **Sub-badges:** `Multi-Tab C-Suite Model` • `Dynamic Excel Formulas` • `Instant Leak Detection`.
