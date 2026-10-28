"""
================================================================================
Zero-Touch Payment Gateway Reconciliation & Executive Reporting Engine
Author: Expert Financial Systems & Python Automation Consultant
Description: Ingests raw Stripe gateway charges and bank settlement ledger exports,
             reconciles transactions, isolates financial discrepancies, and outputs
             a C-suite grade multi-tab Excel financial model using OpenPyXL with
             native Excel formulas (=SUM, =SUMIF, =COUNTIF), KPI cards, and auto-fit columns.
================================================================================
"""

from datetime import datetime, timezone
import logging
from pathlib import Path
import sys
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd

# Ensure Windows terminal standard streams handle UTF-8 properly
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

BASE_DIR = Path(__file__).resolve().parent
GATEWAY_FILE = BASE_DIR / "gateway_charges.csv"
BANK_FILE = BASE_DIR / "bank_ledger.csv"
OUTPUT_REPORT_FILE = BASE_DIR / "Executive_Reconciliation_Report.xlsx"
LOG_FILE = BASE_DIR / "reconciliation.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s]: %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8", mode="w")],
)
logger = logging.getLogger("ReconcileEngine")


# -----------------------------------------------------------------------------
# 1. RECONCILIATION & DISCREPANCY CLASSIFICATION
# -----------------------------------------------------------------------------
class ReconciliationProcessor:
    """
    Automated financial matching engine. Ingests disparate gateway exports
    and bank records, performing outer-join alignment and heuristic categorization.
    """

    @staticmethod
    def process_and_reconcile(gateway_path: Path, bank_path: Path) -> pd.DataFrame:
        logger.info(f"Ingesting gateway export: {gateway_path}")
        df_gateway = pd.read_csv(gateway_path)
        logger.info(f"Ingesting bank ledger: {bank_path}")
        df_bank = pd.read_csv(bank_path)

        # Harmonize key columns
        df_gateway["Join_ID"] = df_gateway["Charge_ID"].astype(str).str.strip()
        df_bank["Join_ID"] = df_bank["Reference_ID"].astype(str).str.strip()

        # Perform Full Outer Join to capture unilateral records
        merged = pd.merge(
            df_gateway,
            df_bank,
            on="Join_ID",
            how="outer",
            suffixes=("_gw", "_bank"),
        )

        # Standardize dates & compute settlement turnaround time
        merged["Created_Date_parsed"] = pd.to_datetime(merged["Created_Date"], errors="coerce")
        merged["Posting_Date_parsed"] = pd.to_datetime(merged["Posting_Date"], errors="coerce")
        merged["Days_to_Settle"] = (merged["Posting_Date_parsed"] - merged["Created_Date_parsed"]).dt.days

        # Ensure numeric columns
        merged["Gross_Amount"] = merged["Gross_Amount"].fillna(0.0).round(2)
        merged["Gateway_Fee"] = merged["Gateway_Fee"].fillna(0.0).round(2)
        merged["Net_Amount"] = merged["Net_Amount"].fillna(0.0).round(2)
        merged["Credit_Amount"] = merged["Credit_Amount"].fillna(0.0).round(2)

        # Variance: Bank Settled Credit vs Gateway Net Expected
        merged["Variance"] = (merged["Credit_Amount"] - merged["Net_Amount"]).round(2)

        # Categorization Logic
        statuses = []
        action_recs = []

        for _, row in merged.iterrows():
            has_gw = pd.notna(row["Charge_ID"])
            has_bank = pd.notna(row["Reference_ID"])
            variance = row["Variance"]
            days = row["Days_to_Settle"]

            if has_gw and not has_bank:
                statuses.append("UNRECONCILED_MISSING_FUNDS")
                action_recs.append("Open Stripe Payout Inquiry; Funds not settled to bank account")
            elif not has_gw and has_bank:
                statuses.append("UNMATCHED_BANK_CREDIT")
                action_recs.append("Investigate unidentified bank deposit; missing gateway charge ID")
            elif abs(variance) > 0.01:
                statuses.append("FEE_DISCREPANCY")
                action_recs.append(f"Audit fee contract: Variance of ${abs(variance):.2f} detected")
            elif days > 3:
                statuses.append("TIMING_DELAY")
                action_recs.append(f"Flagged liquidity delay: Settlement took {int(days)} days (>3 days SLA)")
            else:
                statuses.append("EXACT_MATCH")
                action_recs.append("Reconciled; No action needed")

        merged["Reconciliation_Status"] = statuses
        merged["Action_Required"] = action_recs
        merged["Identifier"] = merged["Join_ID"]

        logger.info(f"Reconciliation completed across {len(merged):,} aggregate records.")
        return merged


# -----------------------------------------------------------------------------
# 2. OPENPYXL ENTERPRISE STYLING & EXCEL MODELING ENGINE
# -----------------------------------------------------------------------------
class ExecutiveExcelReportBuilder:
    """
    Builds a presentation-ready financial reconciliation model featuring:
    - Tab 1: Executive Summary with KPI metric cards, dynamic =SUM and =COUNTIF formulas.
    - Tab 2: Discrepancy Action List with conditional alert fills and action items.
    - Tab 3: All Reconciled Records audit log with zebra striping and number formats.
    - Automated column width adjustment eliminating '###' overflow errors.
    """

    def __init__(self, data: pd.DataFrame, output_path: Path):
        self.df = data
        self.output_path = output_path
        self.wb = openpyxl.Workbook()
        # Remove default sheet
        self.wb.remove(self.wb.active)

        # Color Palette Design Tokens
        self.NAVY_DARK = "0F172A"       # Headers / Primary Titles
        self.NAVY_LIGHT = "1E293B"      # Secondary Headers
        self.ACCENT_BLUE = "2563EB"     # Highlights / Key Totals
        self.CARD_BG = "F8FAFC"         # KPI Card background
        self.CARD_BORDER = "CBD5E1"     # Card border
        self.TEXT_DARK = "0F172A"
        self.TEXT_MUTED = "64748B"
        self.ALERT_RED_FILL = "FEE2E2"
        self.ALERT_RED_TEXT = "991B1B"
        self.ALERT_AMBER_FILL = "FEF3C7"
        self.ALERT_AMBER_TEXT = "92400E"
        self.ALERT_BLUE_FILL = "E0F2FE"
        self.ALERT_BLUE_TEXT = "0369A1"
        self.ZEBRA_FILL = "F8FAFC"

        # Fonts
        self.font_title = Font(name="Segoe UI", size=16, bold=True, color="FFFFFF")
        self.font_subtitle = Font(name="Segoe UI", size=10, italic=True, color="E2E8F0")
        self.font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
        self.font_card_num = Font(name="Segoe UI", size=16, bold=True, color="0F172A")
        self.font_card_label = Font(name="Segoe UI", size=9, bold=True, color="475569")
        self.font_regular = Font(name="Segoe UI", size=10, color="0F172A")
        self.font_bold = Font(name="Segoe UI", size=10, bold=True, color="0F172A")

        # Borders
        self.border_thin = Border(
            left=Side(style="thin", color="E2E8F0"),
            right=Side(style="thin", color="E2E8F0"),
            top=Side(style="thin", color="E2E8F0"),
            bottom=Side(style="thin", color="E2E8F0"),
        )
        self.border_card = Border(
            left=Side(style="medium", color="CBD5E1"),
            right=Side(style="medium", color="CBD5E1"),
            top=Side(style="medium", color="CBD5E1"),
            bottom=Side(style="medium", color="CBD5E1"),
        )
        self.border_double_bottom = Border(
            top=Side(style="thin", color="CBD5E1"),
            bottom=Side(style="double", color="0F172A"),
        )

    def build_all(self) -> None:
        self._build_tab_executive_summary()
        self._build_tab_discrepancy_list()
        self._build_tab_all_reconciled()
        self._auto_fit_columns()
        self.wb.save(self.output_path)
        logger.info(f"Generated Executive Reconciliation Excel Model: {self.output_path}")

    # -------------------------------------------------------------------------
    # TAB 1: EXECUTIVE SUMMARY
    # -------------------------------------------------------------------------
    def _build_tab_executive_summary(self) -> None:
        ws = self.wb.create_sheet(title="Executive Summary", index=0)
        ws.views.sheetView[0].showGridLines = True

        # 1. Corporate Header Banner (Row 2 to 3)
        ws.merge_cells("B2:H2")
        title_cell = ws["B2"]
        title_cell.value = "PAYMENT GATEWAY RECONCILIATION & SETTLEMENT AUDIT"
        title_cell.font = self.font_title
        title_cell.fill = PatternFill(start_color=self.NAVY_DARK, end_color=self.NAVY_DARK, fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.merge_cells("B3:H3")
        sub_cell = ws["B3"]
        sub_cell.value = f"Financial Period: Q1 2024 | Model Run: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')} | Confidential - CFO / Finance Committee"
        sub_cell.font = self.font_subtitle
        sub_cell.fill = PatternFill(start_color=self.NAVY_LIGHT, end_color=self.NAVY_LIGHT, fill_type="solid")
        sub_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[2].height = 28
        ws.row_dimensions[3].height = 20

        # Total rows in records tab
        records_count = len(self.df) + 1  # 2 to records_count + 1

        # 2. Executive KPI Cards (Row 5 to 7)
        # Card 1: Total Gateway Processed Volume
        # Card 2: Actual Bank Settlements
        # Card 3: Reconciliation Rate %
        # Card 4: Unreconciled At-Risk Capital
        cards = [
            ("B", "C", "TOTAL GATEWAY GROSS", f"=SUM('All Reconciled Records'!C2:C{records_count})", '"$"#,##0.00'),
            ("D", "D", "ACTUAL BANK DEPOSITS", f"=SUM('All Reconciled Records'!G2:G{records_count})", '"$"#,##0.00'),
            ("E", "F", "RECONCILIATION RATE", f"=COUNTIF('All Reconciled Records'!J2:J{records_count}, \"EXACT_MATCH\") / COUNTA('All Reconciled Records'!J2:J{records_count})", "0.0%"),
            ("G", "H", "UNRECONCILED AT-RISK CAPITAL", f"=SUMIF('All Reconciled Records'!J2:J{records_count}, \"UNRECONCILED_MISSING_FUNDS\", 'All Reconciled Records'!E2:E{records_count})", '"$"#,##0.00'),
        ]

        ws.row_dimensions[5].height = 18
        ws.row_dimensions[6].height = 30

        for col_start, col_end, label, formula, num_format in cards:
            if col_start != col_end:
                ws.merge_cells(f"{col_start}5:{col_end}5")
                ws.merge_cells(f"{col_start}6:{col_end}6")

            top_cell = ws[f"{col_start}5"]
            top_cell.value = label
            top_cell.font = self.font_card_label
            top_cell.alignment = Alignment(horizontal="center", vertical="center")
            top_cell.fill = PatternFill(start_color=self.CARD_BG, end_color=self.CARD_BG, fill_type="solid")

            val_cell = ws[f"{col_start}6"]
            val_cell.value = formula
            val_cell.font = self.font_card_num
            val_cell.alignment = Alignment(horizontal="center", vertical="center")
            val_cell.number_format = num_format
            val_cell.fill = PatternFill(start_color=self.CARD_BG, end_color=self.CARD_BG, fill_type="solid")

            # Apply borders around card
            for col in range(openpyxl.utils.column_index_from_string(col_start), openpyxl.utils.column_index_from_string(col_end) + 1):
                ws.cell(row=5, column=col).border = self.border_thin
                ws.cell(row=6, column=col).border = self.border_thin

        # 3. Discrepancy Breakdown Table (Row 9 onwards)
        ws.merge_cells("B9:H9")
        sec1 = ws["B9"]
        sec1.value = "DISCREPANCY VARIANCE BREAKDOWN"
        sec1.font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        sec1.fill = PatternFill(start_color=self.NAVY_LIGHT, end_color=self.NAVY_LIGHT, fill_type="solid")
        sec1.alignment = Alignment(horizontal="left", indent=1, vertical="center")
        ws.row_dimensions[9].height = 22

        table_headers = ["Category / Status", "Description & Operational Impact", "Count", "Formula Reference", "Financial Exposure ($)", "Action Protocol"]
        header_cols = ["B", "C", "D", "E", "F", "G"]

        ws.row_dimensions[10].height = 22
        for col_letter, text in zip(["B", "C", "D", "E", "F", "G"], table_headers):
            c = ws[f"{col_letter}10"]
            c.value = text
            c.font = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
            c.fill = PatternFill(start_color=self.NAVY_DARK, end_color=self.NAVY_DARK, fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = self.border_thin

        status_rows = [
            ("EXACT_MATCH", "Settled cleanly within 2-day standard SLA", f"=COUNTIF('All Reconciled Records'!J2:J{records_count}, \"EXACT_MATCH\")", "Normal", f"=SUMIF('All Reconciled Records'!J2:J{records_count}, \"EXACT_MATCH\", 'All Reconciled Records'!G2:G{records_count})", "Automated Clear"),
            ("UNRECONCILED_MISSING_FUNDS", "Gateway captured charge but never deposited to bank account", f"=COUNTIF('All Reconciled Records'!J2:J{records_count}, \"UNRECONCILED_MISSING_FUNDS\")", "CRITICAL RISK", f"=SUMIF('All Reconciled Records'!J2:J{records_count}, \"UNRECONCILED_MISSING_FUNDS\", 'All Reconciled Records'!E2:E{records_count})", "Open Gateway Ticket"),
            ("FEE_DISCREPANCY", "Bank credit does not match gateway expected net (over-deduction)", f"=COUNTIF('All Reconciled Records'!J2:J{records_count}, \"FEE_DISCREPANCY\")", "HIGH AUDIT", f"=SUMIF('All Reconciled Records'!J2:J{records_count}, \"FEE_DISCREPANCY\", 'All Reconciled Records'!H2:H{records_count})", "Audit Gateway Surcharges"),
            ("TIMING_DELAY", "Settled cleanly but required 4+ business days to hit ledger", f"=COUNTIF('All Reconciled Records'!J2:J{records_count}, \"TIMING_DELAY\")", "LIQUIDITY LAG", f"=SUMIF('All Reconciled Records'!J2:J{records_count}, \"TIMING_DELAY\", 'All Reconciled Records'!G2:G{records_count})", "Review Float Terms"),
            ("UNMATCHED_BANK_CREDIT", "Bank credit received with unmapped gateway charge identifier", f"=COUNTIF('All Reconciled Records'!J2:J{records_count}, \"UNMATCHED_BANK_CREDIT\")", "ACCOUNTING", f"=SUMIF('All Reconciled Records'!J2:J{records_count}, \"UNMATCHED_BANK_CREDIT\", 'All Reconciled Records'!G2:G{records_count})", "Reconcile ERP Journal"),
        ]

        current_row = 11
        for cat, desc, count_formula, impact, sum_formula, protocol in status_rows:
            ws[f"B{current_row}"] = cat
            ws[f"C{current_row}"] = desc
            ws[f"D{current_row}"] = count_formula
            ws[f"E{current_row}"] = impact
            ws[f"F{current_row}"] = sum_formula
            ws[f"G{current_row}"] = protocol

            # Alignments & formats
            ws[f"B{current_row}"].font = self.font_bold
            ws[f"C{current_row}"].font = self.font_regular
            ws[f"D{current_row}"].font = self.font_bold
            ws[f"D{current_row}"].alignment = Alignment(horizontal="right")
            ws[f"D{current_row}"].number_format = "#,##0"
            ws[f"E{current_row}"].alignment = Alignment(horizontal="center")
            ws[f"F{current_row}"].font = self.font_bold
            ws[f"F{current_row}"].alignment = Alignment(horizontal="right")
            ws[f"F{current_row}"].number_format = '"$"#,##0.00'
            ws[f"G{current_row}"].alignment = Alignment(horizontal="center")

            # Conditional Highlights
            if cat == "UNRECONCILED_MISSING_FUNDS":
                fill = PatternFill(start_color=self.ALERT_RED_FILL, end_color=self.ALERT_RED_FILL, fill_type="solid")
                for col in ["B", "C", "D", "E", "F", "G"]:
                    ws[f"{col}{current_row}"].fill = fill
                    ws[f"{col}{current_row}"].font = Font(name="Segoe UI", size=10, bold=True, color=self.ALERT_RED_TEXT)
            elif cat == "FEE_DISCREPANCY":
                fill = PatternFill(start_color=self.ALERT_AMBER_FILL, end_color=self.ALERT_AMBER_FILL, fill_type="solid")
                for col in ["B", "C", "D", "E", "F", "G"]:
                    ws[f"{col}{current_row}"].fill = fill
                    ws[f"{col}{current_row}"].font = Font(name="Segoe UI", size=10, bold=True, color=self.ALERT_AMBER_TEXT)

            for col in ["B", "C", "D", "E", "F", "G"]:
                ws[f"{col}{current_row}"].border = self.border_thin

            ws.row_dimensions[current_row].height = 20
            current_row += 1

        # Total Summary Row
        ws[f"B{current_row}"] = "TOTAL ACTIVE RECONCILIATION AUDIT"
        ws[f"B{current_row}"].font = self.font_bold
        ws[f"D{current_row}"] = f"=SUM(D11:D{current_row-1})"
        ws[f"D{current_row}"].font = self.font_bold
        ws[f"D{current_row}"].alignment = Alignment(horizontal="right")
        ws[f"D{current_row}"].number_format = "#,##0"
        ws[f"F{current_row}"] = f"=SUM(F11:F{current_row-1})"
        ws[f"F{current_row}"].font = self.font_bold
        ws[f"F{current_row}"].alignment = Alignment(horizontal="right")
        ws[f"F{current_row}"].number_format = '"$"#,##0.00'

        for col in ["B", "C", "D", "E", "F", "G"]:
            ws[f"{col}{current_row}"].border = self.border_double_bottom
        ws.row_dimensions[current_row].height = 22

    # -------------------------------------------------------------------------
    # TAB 2: DISCREPANCY ACTION LIST
    # -------------------------------------------------------------------------
    def _build_tab_discrepancy_list(self) -> None:
        ws = self.wb.create_sheet(title="Discrepancy Action List", index=1)
        ws.views.sheetView[0].showGridLines = True

        headers = [
            "Transaction ID / Ref",
            "Customer Account",
            "Gateway Gross",
            "Gateway Net Expected",
            "Bank Credit Settled",
            "Net Variance ($)",
            "Days to Settle",
            "Discrepancy Classification",
            "Remediation Action Required",
            "Assigned Owner",
        ]
        ws.append(headers)

        # Style Header
        header_fill = PatternFill(start_color=self.NAVY_DARK, end_color=self.NAVY_DARK, fill_type="solid")
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = self.font_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = self.border_thin
        ws.row_dimensions[1].height = 26

        # Filter discrepancies
        discrepancies = self.df[self.df["Reconciliation_Status"] != "EXACT_MATCH"].copy()

        fill_red = PatternFill(start_color=self.ALERT_RED_FILL, end_color=self.ALERT_RED_FILL, fill_type="solid")
        font_red = Font(name="Segoe UI", size=10, bold=True, color=self.ALERT_RED_TEXT)
        fill_amber = PatternFill(start_color=self.ALERT_AMBER_FILL, end_color=self.ALERT_AMBER_FILL, fill_type="solid")
        font_amber = Font(name="Segoe UI", size=10, bold=True, color=self.ALERT_AMBER_TEXT)
        fill_blue = PatternFill(start_color=self.ALERT_BLUE_FILL, end_color=self.ALERT_BLUE_FILL, fill_type="solid")
        font_blue = Font(name="Segoe UI", size=10, color=self.ALERT_BLUE_TEXT)

        row_idx = 2
        for _, r in discrepancies.iterrows():
            status = r["Reconciliation_Status"]
            owner = "FinOps Treasury" if status == "UNRECONCILED_MISSING_FUNDS" else ("Merchant Billing" if status == "FEE_DISCREPANCY" else "Payments Team")

            ws.append([
                str(r["Identifier"]),
                str(r["Customer_Email"]) if pd.notna(r["Customer_Email"]) else "N/A",
                float(r["Gross_Amount"]),
                float(r["Net_Amount"]),
                float(r["Credit_Amount"]),
                float(r["Variance"]),
                int(r["Days_to_Settle"]) if pd.notna(r["Days_to_Settle"]) else 0,
                status,
                str(r["Action_Required"]),
                owner,
            ])

            # Apply row styling based on status
            ws.row_dimensions[row_idx].height = 20
            chosen_fill = fill_red if status == "UNRECONCILED_MISSING_FUNDS" else (fill_amber if status == "FEE_DISCREPANCY" else fill_blue)
            chosen_font = font_red if status == "UNRECONCILED_MISSING_FUNDS" else (font_amber if status == "FEE_DISCREPANCY" else font_blue)

            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.fill = chosen_fill
                cell.font = chosen_font
                cell.border = self.border_thin

                # Number formats
                if col_idx in (3, 4, 5, 6):
                    cell.number_format = '"$"#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_idx == 7:
                    cell.number_format = '#,##0'
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx in (1, 8, 10):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

            row_idx += 1

        # Freeze header and apply auto-filter
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    # -------------------------------------------------------------------------
    # TAB 3: ALL RECONCILED RECORDS
    # -------------------------------------------------------------------------
    def _build_tab_all_reconciled(self) -> None:
        ws = self.wb.create_sheet(title="All Reconciled Records", index=2)
        ws.views.sheetView[0].showGridLines = True

        headers = [
            "Transaction ID",
            "Created Date",
            "Gross Amount",
            "Gateway Fee",
            "Net Expected",
            "Bank Posting Date",
            "Bank Credit Settled",
            "Variance ($)",
            "Turnaround Days",
            "Reconciliation Status",
        ]
        ws.append(headers)

        header_fill = PatternFill(start_color=self.NAVY_LIGHT, end_color=self.NAVY_LIGHT, fill_type="solid")
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = self.font_header
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = self.border_thin
        ws.row_dimensions[1].height = 26

        zebra_fill = PatternFill(start_color=self.ZEBRA_FILL, end_color=self.ZEBRA_FILL, fill_type="solid")

        row_idx = 2
        for _, r in self.df.iterrows():
            ws.append([
                str(r["Identifier"]),
                str(r["Created_Date"])[:10] if pd.notna(r["Created_Date"]) else "N/A",
                float(r["Gross_Amount"]),
                float(r["Gateway_Fee"]),
                float(r["Net_Amount"]),
                str(r["Posting_Date"]) if pd.notna(r["Posting_Date"]) else "NOT_POSTED",
                float(r["Credit_Amount"]),
                float(r["Variance"]),
                int(r["Days_to_Settle"]) if pd.notna(r["Days_to_Settle"]) else 0,
                str(r["Reconciliation_Status"]),
            ])

            ws.row_dimensions[row_idx].height = 19
            is_even = (row_idx % 2 == 0)

            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = self.font_regular
                cell.border = self.border_thin
                if is_even:
                    cell.fill = zebra_fill

                if col_idx in (3, 4, 5, 7, 8):
                    cell.number_format = '"$"#,##0.00'
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                elif col_idx in (2, 6, 9):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx == 10:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    # Subtle highlight on status
                    if cell.value == "UNRECONCILED_MISSING_FUNDS":
                        cell.font = Font(name="Segoe UI", size=10, bold=True, color="DC2626")
                    elif cell.value == "FEE_DISCREPANCY":
                        cell.font = Font(name="Segoe UI", size=10, bold=True, color="D97706")
                    elif cell.value == "EXACT_MATCH":
                        cell.font = Font(name="Segoe UI", size=10, color="16A34A")

            row_idx += 1

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

    # -------------------------------------------------------------------------
    # AUTO-FIT COLUMN WIDTHS ACROSS ALL SHEETS
    # -------------------------------------------------------------------------
    def _auto_fit_columns(self) -> None:
        """
        Dynamically calculates max string width across all cells to prevent
        the dreaded '###' truncation error in Excel.
        """
        for ws in self.wb.worksheets:
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                # Skip merged summary title rows for width calculation
                for cell in col:
                    if ws.title == "Executive Summary" and cell.row in (2, 3):
                        continue
                    val_str = str(cell.value or "")
                    if str(val_str).startswith("="):
                        val_str = "$123,456.78"
                    max_len = max(max_len, len(val_str))
                ws.column_dimensions[col_letter].width = max(max_len + 4, 13)


# -----------------------------------------------------------------------------
# 3. SCRIPT EXECUTION ENTRYPOINT
# -----------------------------------------------------------------------------
def run_reconciliation() -> None:
    print("\n================================================================================")
    print(" ZERO-TOUCH PAYMENT RECONCILIATION & EXECUTIVE REPORTING ENGINE")
    print("================================================================================")

    # 1. Reconcile
    recon_df = ReconciliationProcessor.process_and_reconcile(GATEWAY_FILE, BANK_FILE)

    # 2. Build Excel Model
    builder = ExecutiveExcelReportBuilder(recon_df, OUTPUT_REPORT_FILE)
    builder.build_all()

    # 3. Print Executive Terminal Summary
    summary = recon_df["Reconciliation_Status"].value_counts()
    total_recs = len(recon_df)
    exact_matches = summary.get("EXACT_MATCH", 0)
    match_rate = (exact_matches / total_recs) * 100.0

    missing_capital = recon_df[recon_df["Reconciliation_Status"] == "UNRECONCILED_MISSING_FUNDS"]["Net_Amount"].sum()
    fee_leakage = abs(recon_df[recon_df["Reconciliation_Status"] == "FEE_DISCREPANCY"]["Variance"].sum())

    print(f"\n[+] Processing Status: COMPLETED SUCCESSFULLY")
    print(f"[+] Total Transactions Audited:    {total_recs:,}")
    print(f"[+] Clean Reconciliation Rate:      {match_rate:.1f}% ({exact_matches:,} transactions)")
    print(f"[!] Unreconciled Missing Capital:   ${missing_capital:,.2f} (Action Required)")
    print(f"[!] Total Fee Overcharges/Leak:     ${fee_leakage:,.2f}")
    print(f"[+] Output Executive Spreadsheet:   {OUTPUT_REPORT_FILE}")
    print("================================================================================\n")


if __name__ == "__main__":
    run_reconciliation()
