#!/usr/bin/env python3
"""
Financiële Cockpit - Excel Workbook Generator

Generates a comprehensive financial dashboard workbook with:
- Live price tracking (via Excel Power Query)
- Holdings tracking (BTC, DOGE, Real Estate, Cash)
- Loan management with LTV calculations
- Automated alerts based on user-defined thresholds
- Stress testing scenarios

Usage:
    python generate_workbook.py [output_path]

Output:
    Creates Financiele_Cockpit.xlsx in the /dist folder (or specified path)
"""

import sys
import os
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.worksheet import Worksheet

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from model_config import (
    CRYPTO_ASSETS,
    SETTINGS_PARAMETERS,
    SHEETS_CONFIG,
    TABLE_DEFINITIONS,
    SAMPLE_DATA,
    DATA_VALIDATIONS,
    COLORS,
)
from excel_builder import (
    create_named_styles,
    create_excel_table,
    add_dropdown_validation,
    add_status_conditional_formatting,
    add_severity_conditional_formatting,
    set_column_widths,
    auto_column_widths,
    freeze_panes,
    add_sheet_title,
    add_section_header,
    add_instruction_text,
    create_kpi_cell,
    create_dashboard_section,
)


class FinancialCockpitGenerator:
    """Generates the Financial Cockpit Excel workbook."""

    def __init__(self):
        self.wb = Workbook()
        self.sheets: Dict[str, Worksheet] = {}

    def generate(self, output_path: str) -> None:
        """Generate the complete workbook and save to file."""
        print("Generating Financiële Cockpit...")

        # Initialize styles
        create_named_styles(self.wb)

        # Remove default sheet
        default_sheet = self.wb.active
        self.wb.remove(default_sheet)

        # Create all sheets in order
        for sheet_config in sorted(SHEETS_CONFIG, key=lambda x: x["order"]):
            ws = self.wb.create_sheet(title=sheet_config["name"])
            self.sheets[sheet_config["name"]] = ws
            print(f"  Created sheet: {sheet_config['name']}")

        # Build each sheet
        self._build_instellingen_sheet()
        self._build_koersen_sheet()
        self._build_koers_setup_sheet()
        self._build_btc_holdings_sheet()
        self._build_doge_holdings_sheet()
        self._build_vastgoed_sheet()
        self._build_leningen_sheet()
        self._build_btc_onderpand_sheet()
        self._build_cash_sheet()
        self._build_alerts_sheet()
        self._build_dashboard_sheet()

        # Set Dashboard as active sheet
        self.wb.active = self.sheets["Dashboard"]

        # Save workbook
        self.wb.save(output_path)
        print(f"\nWorkbook saved to: {output_path}")

    # =========================================================================
    # INSTELLINGEN (Settings) Sheet
    # =========================================================================

    def _build_instellingen_sheet(self) -> None:
        """Build the Settings sheet with all configurable parameters."""
        ws = self.sheets["Instellingen"]

        # Title
        add_sheet_title(ws, "Instellingen - Configuratie Parameters", row=1, col=1)

        # Settings table
        settings_cols = ["Parameter", "Waarde", "Eenheid", "Toelichting"]
        settings_data = [
            {
                "Parameter": p["parameter"],
                "Waarde": p["value"],
                "Eenheid": p["unit"],
                "Toelichting": p["description"],
            }
            for p in SETTINGS_PARAMETERS
        ]

        create_excel_table(
            ws,
            table_name="tblSettings",
            columns=settings_cols,
            start_row=3,
            start_col=1,
            data=settings_data,
        )

        # Loan Collateral Mapping table (separate section)
        add_section_header(ws, "Lening-Onderpand Koppeling", row=3, col=7, span=4)

        mapping_cols = ["LoanID", "CollateralType", "CollateralRef", "Notes"]
        mapping_data = SAMPLE_DATA.get("tblLoanCollateralMap", [])

        create_excel_table(
            ws,
            table_name="tblLoanCollateralMap",
            columns=mapping_cols,
            start_row=4,
            start_col=7,
            data=mapping_data,
            num_empty_rows=10,
        )

        # Add data validation for CollateralType
        add_dropdown_validation(ws, "collateral_type", "H5:H50")

        # Column widths
        set_column_widths(ws, {
            1: 25,  # Parameter
            2: 12,  # Waarde
            3: 12,  # Eenheid
            4: 55,  # Toelichting
            6: 3,   # Spacer
            7: 12,  # LoanID
            8: 15,  # CollateralType
            9: 20,  # CollateralRef
            10: 30, # Notes
        })

        freeze_panes(ws, row=4, col=1)

    # =========================================================================
    # KOERSEN (Prices) Sheet
    # =========================================================================

    def _build_koersen_sheet(self) -> None:
        """Build the Prices sheet."""
        ws = self.sheets["Koersen"]

        add_sheet_title(ws, "Koersen - Live Prijzen", row=1, col=1)

        # Prices table
        price_cols = ["Asset", "Symbol", "EUR prijs", "USD prijs", "Laatste update"]
        price_data = SAMPLE_DATA.get("tblKoersen", [])

        create_excel_table(
            ws,
            table_name="tblKoersen",
            columns=price_cols,
            start_row=3,
            start_col=1,
            data=price_data,
        )

        # Apply number formats to price columns
        for row in range(4, 10):
            ws.cell(row=row, column=3).number_format = '€ #,##0.00'
            ws.cell(row=row, column=4).number_format = '$ #,##0.00'
            ws.cell(row=row, column=5).number_format = 'DD-MM-YYYY HH:MM'

        # Column widths
        set_column_widths(ws, {
            1: 15,  # Asset
            2: 10,  # Symbol
            3: 15,  # EUR prijs
            4: 15,  # USD prijs
            5: 20,  # Laatste update
        })

        # Instructions
        ws.cell(row=10, column=1, value="Let op: Koersen worden bijgewerkt via Power Query. Zie 'Koers-Setup' voor instructies.")
        ws.cell(row=10, column=1).font = Font(italic=True, color="666666")

        freeze_panes(ws, row=4, col=1)

    # =========================================================================
    # KOERS-SETUP (Price Setup Instructions) Sheet
    # =========================================================================

    def _build_koers_setup_sheet(self) -> None:
        """Build the Price Setup instructions sheet."""
        ws = self.sheets["Koers-Setup"]

        add_sheet_title(ws, "Koers-Setup - Power Query Instructies", row=1, col=1)

        instructions = """
INSTRUCTIES VOOR AUTOMATISCHE KOERSEN VIA POWER QUERY

1. POWER QUERY INSTELLEN
   a) Ga naar Data > Get Data > From Web
   b) Voer de volgende URL in:
      https://api.coingecko.com/api/v3/simple/price?ids=bitcoin,dogecoin&vs_currencies=eur,usd
   c) Klik OK en wacht tot de data is geladen

2. DATA TRANSFORMEREN
   a) In Power Query Editor: Transform > To Table
   b) Expand de kolommen om bitcoin.eur, bitcoin.usd, dogecoin.eur, dogecoin.usd te krijgen
   c) Hernoem kolommen naar: Asset, EUR prijs, USD prijs
   d) Voeg een kolom toe voor Laatste update: = DateTime.LocalNow()
   e) Klik Close & Load to... > selecteer de Koersen sheet

3. AUTOMATISCHE REFRESH INSTELLEN
   a) Selecteer de query in het Queries & Connections panel
   b) Rechtermuisklik > Properties
   c) Vink aan: "Refresh every X minutes" (X = waarde uit Instellingen!RefreshMinutes)
   d) Optioneel: vink "Refresh data when opening the file" aan

4. ALTERNATIEVE API'S (indien CoinGecko niet werkt)
   - CoinCap: https://api.coincap.io/v2/assets?ids=bitcoin,dogecoin
   - Binance: https://api.binance.com/api/v3/ticker/price?symbols=["BTCEUR","DOGEEUR"]

BELANGRIJKE OPMERKINGEN
- De gratis CoinGecko API heeft rate limits (10-30 calls/minuut)
- Bij problemen: verhoog RefreshMinutes in Instellingen
- Test de connectie eerst handmatig voordat je automatische refresh aanzet

HANDMATIGE UPDATE
Als Power Query niet werkt, kun je de koersen handmatig invoeren in de Koersen sheet.
De berekeningen en alerts werken dan nog steeds correct.
"""

        current_row = 3
        for line in instructions.strip().split('\n'):
            ws.cell(row=current_row, column=1, value=line)
            if line.startswith('1.') or line.startswith('2.') or line.startswith('3.') or line.startswith('4.'):
                ws.cell(row=current_row, column=1).font = Font(bold=True, size=11)
            elif line.startswith('INSTRUCTIES') or line.startswith('BELANGRIJKE') or line.startswith('HANDMATIGE'):
                ws.cell(row=current_row, column=1).font = Font(bold=True, size=12, color=COLORS["header_bg"])
            current_row += 1

        # Set column width
        ws.column_dimensions['A'].width = 100

    # =========================================================================
    # BITCOIN HOLDINGS Sheet
    # =========================================================================

    def _build_btc_holdings_sheet(self) -> None:
        """Build the Bitcoin Holdings sheet."""
        ws = self.sheets["Bitcoin Holdings"]

        add_sheet_title(ws, "Bitcoin Holdings", row=1, col=1)

        # Holdings table
        btc_cols = ["Datum", "Wallet/Platform", "BTC Aantal", "Type", "Opmerking"]
        btc_data = SAMPLE_DATA.get("tblBTC", [])

        create_excel_table(
            ws,
            table_name="tblBTC",
            columns=btc_cols,
            start_row=3,
            start_col=1,
            data=btc_data,
            num_empty_rows=20,
        )

        # Data validation for Type column
        add_dropdown_validation(ws, "holding_type", "D4:D100")

        # Number formats
        for row in range(4, 100):
            ws.cell(row=row, column=1).number_format = 'YYYY-MM-DD'
            ws.cell(row=row, column=3).number_format = '0.00000000'

        # Summary section
        summary_row = 2
        ws.cell(row=summary_row, column=7, value="Totaal BTC:")
        ws.cell(row=summary_row, column=7).font = Font(bold=True)
        ws.cell(row=summary_row, column=8, value="=SUM(tblBTC[BTC Aantal])")
        ws.cell(row=summary_row, column=8).number_format = '0.00000000'
        ws.cell(row=summary_row, column=8).font = Font(bold=True)

        ws.cell(row=summary_row + 1, column=7, value="Spot:")
        ws.cell(row=summary_row + 1, column=8, value='=SUMIF(tblBTC[Type],"Spot",tblBTC[BTC Aantal])')
        ws.cell(row=summary_row + 1, column=8).number_format = '0.00000000'

        ws.cell(row=summary_row + 2, column=7, value="Collateral:")
        ws.cell(row=summary_row + 2, column=8, value='=SUMIF(tblBTC[Type],"Collateral",tblBTC[BTC Aantal])')
        ws.cell(row=summary_row + 2, column=8).number_format = '0.00000000'

        # Column widths
        set_column_widths(ws, {
            1: 12,  # Datum
            2: 20,  # Wallet/Platform
            3: 15,  # BTC Aantal
            4: 12,  # Type
            5: 30,  # Opmerking
            6: 3,   # Spacer
            7: 12,  # Label
            8: 15,  # Value
        })

        freeze_panes(ws, row=4, col=1)

    # =========================================================================
    # DOGECOIN HOLDINGS Sheet
    # =========================================================================

    def _build_doge_holdings_sheet(self) -> None:
        """Build the Dogecoin Holdings sheet."""
        ws = self.sheets["Dogecoin Holdings"]

        add_sheet_title(ws, "Dogecoin Holdings", row=1, col=1)

        # Holdings table
        doge_cols = ["Datum", "Wallet/Platform", "DOGE Aantal", "Type", "Opmerking"]
        doge_data = SAMPLE_DATA.get("tblDOGE", [])

        create_excel_table(
            ws,
            table_name="tblDOGE",
            columns=doge_cols,
            start_row=3,
            start_col=1,
            data=doge_data,
            num_empty_rows=20,
        )

        # Data validation for Type column
        add_dropdown_validation(ws, "holding_type", "D4:D100")

        # Number formats
        for row in range(4, 100):
            ws.cell(row=row, column=1).number_format = 'YYYY-MM-DD'
            ws.cell(row=row, column=3).number_format = '#,##0.00'

        # Summary section
        summary_row = 2
        ws.cell(row=summary_row, column=7, value="Totaal DOGE:")
        ws.cell(row=summary_row, column=7).font = Font(bold=True)
        ws.cell(row=summary_row, column=8, value="=SUM(tblDOGE[DOGE Aantal])")
        ws.cell(row=summary_row, column=8).number_format = '#,##0.00'
        ws.cell(row=summary_row, column=8).font = Font(bold=True)

        # Column widths
        set_column_widths(ws, {
            1: 12,  # Datum
            2: 20,  # Wallet/Platform
            3: 15,  # DOGE Aantal
            4: 12,  # Type
            5: 30,  # Opmerking
            6: 3,   # Spacer
            7: 12,  # Label
            8: 15,  # Value
        })

        freeze_panes(ws, row=4, col=1)

    # =========================================================================
    # VASTGOED (Real Estate) Sheet
    # =========================================================================

    def _build_vastgoed_sheet(self) -> None:
        """Build the Real Estate sheet."""
        ws = self.sheets["Vastgoed"]

        add_sheet_title(ws, "Vastgoed - Onroerend Goed", row=1, col=1)

        # Real estate table
        vastgoed_cols = [
            "Pand", "Adres", "Type", "Waarde(EUR)",
            "Hypotheek(EUR)", "Overige Leningen(EUR)", "Netto Waarde"
        ]
        vastgoed_data = SAMPLE_DATA.get("tblVastgoed", [])

        # Add Netto Waarde calculation to data
        for item in vastgoed_data:
            if "Netto Waarde" not in item:
                item["Netto Waarde"] = None  # Will be formula

        create_excel_table(
            ws,
            table_name="tblVastgoed",
            columns=vastgoed_cols,
            start_row=3,
            start_col=1,
            data=vastgoed_data,
            num_empty_rows=10,
        )

        # Add Netto Waarde formula to each row
        for row in range(4, 20):
            cell = ws.cell(row=row, column=7)
            cell.value = f"=IF(D{row}=\"\",\"\",D{row}-E{row}-F{row})"

        # Data validation for Type
        add_dropdown_validation(ws, "vastgoed_type", "C4:C50")

        # Number formats
        for row in range(4, 50):
            for col in [4, 5, 6, 7]:  # Currency columns
                ws.cell(row=row, column=col).number_format = '€ #,##0'

        # Summary section
        summary_row = 2
        ws.cell(row=summary_row, column=9, value="Totale Waarde:")
        ws.cell(row=summary_row, column=9).font = Font(bold=True)
        ws.cell(row=summary_row, column=10, value="=SUM(tblVastgoed[Waarde(EUR)])")
        ws.cell(row=summary_row, column=10).number_format = '€ #,##0'
        ws.cell(row=summary_row, column=10).font = Font(bold=True)

        ws.cell(row=summary_row + 1, column=9, value="Totale Schuld:")
        ws.cell(row=summary_row + 1, column=10, value="=SUM(tblVastgoed[Hypotheek(EUR)])+SUM(tblVastgoed[Overige Leningen(EUR)])")
        ws.cell(row=summary_row + 1, column=10).number_format = '€ #,##0'

        ws.cell(row=summary_row + 2, column=9, value="Netto Waarde:")
        ws.cell(row=summary_row + 2, column=10, value="=SUM(tblVastgoed[Netto Waarde])")
        ws.cell(row=summary_row + 2, column=10).number_format = '€ #,##0'

        # Column widths
        set_column_widths(ws, {
            1: 25,  # Pand
            2: 35,  # Adres
            3: 15,  # Type
            4: 15,  # Waarde
            5: 15,  # Hypotheek
            6: 18,  # Overige Leningen
            7: 15,  # Netto Waarde
            8: 3,   # Spacer
            9: 15,  # Label
            10: 15, # Value
        })

        freeze_panes(ws, row=4, col=1)

    # =========================================================================
    # LENINGEN (Loans) Sheet
    # =========================================================================

    def _build_leningen_sheet(self) -> None:
        """Build the Loans sheet."""
        ws = self.sheets["Leningen"]

        add_sheet_title(ws, "Leningen - Overzicht", row=1, col=1)

        # Loans table
        leningen_cols = [
            "LoanID", "Datum", "OnderpandType", "OnderpandRef", "Lener",
            "Hoofdsom(EUR)", "Openstaand(EUR)", "Rente%", "Looptijd",
            "LiquidatieLTV", "Opmerking"
        ]
        leningen_data = SAMPLE_DATA.get("tblLeningen", [])

        create_excel_table(
            ws,
            table_name="tblLeningen",
            columns=leningen_cols,
            start_row=3,
            start_col=1,
            data=leningen_data,
            num_empty_rows=15,
        )

        # Data validation for OnderpandType
        add_dropdown_validation(ws, "collateral_type", "C4:C50")

        # Number formats
        for row in range(4, 50):
            ws.cell(row=row, column=2).number_format = 'YYYY-MM-DD'  # Datum
            ws.cell(row=row, column=6).number_format = '€ #,##0.00'  # Hoofdsom
            ws.cell(row=row, column=7).number_format = '€ #,##0.00'  # Openstaand
            ws.cell(row=row, column=8).number_format = '0.00%'       # Rente%
            ws.cell(row=row, column=10).number_format = '0.00%'      # LiquidatieLTV

        # Summary section
        summary_row = 2
        ws.cell(row=summary_row, column=13, value="Totaal Openstaand:")
        ws.cell(row=summary_row, column=13).font = Font(bold=True)
        ws.cell(row=summary_row, column=14, value="=SUM(tblLeningen[Openstaand(EUR)])")
        ws.cell(row=summary_row, column=14).number_format = '€ #,##0.00'
        ws.cell(row=summary_row, column=14).font = Font(bold=True)

        ws.cell(row=summary_row + 1, column=13, value="BTC-leningen:")
        ws.cell(row=summary_row + 1, column=14, value='=SUMIF(tblLeningen[OnderpandType],"BTC",tblLeningen[Openstaand(EUR)])')
        ws.cell(row=summary_row + 1, column=14).number_format = '€ #,##0.00'

        ws.cell(row=summary_row + 2, column=13, value="Vastgoed-leningen:")
        ws.cell(row=summary_row + 2, column=14, value='=SUMIF(tblLeningen[OnderpandType],"Vastgoed",tblLeningen[Openstaand(EUR)])')
        ws.cell(row=summary_row + 2, column=14).number_format = '€ #,##0.00'

        # Column widths
        set_column_widths(ws, {
            1: 10,  # LoanID
            2: 12,  # Datum
            3: 14,  # OnderpandType
            4: 18,  # OnderpandRef
            5: 15,  # Lener
            6: 14,  # Hoofdsom
            7: 14,  # Openstaand
            8: 10,  # Rente%
            9: 12,  # Looptijd
            10: 12, # LiquidatieLTV
            11: 25, # Opmerking
            12: 3,  # Spacer
            13: 18, # Label
            14: 15, # Value
        })

        freeze_panes(ws, row=4, col=1)

    # =========================================================================
    # BTC ALS ONDERPAND (BTC as Collateral View) Sheet
    # =========================================================================

    def _build_btc_onderpand_sheet(self) -> None:
        """Build the BTC as Collateral view sheet with LTV calculations."""
        ws = self.sheets["BTC als Onderpand"]

        add_sheet_title(ws, "BTC als Onderpand - LTV Monitor", row=1, col=1)

        # Column headers
        onderpand_cols = [
            "LoanID", "Platform/Lener", "BTC Onderpand", "Openstaand(EUR)",
            "BTC prijs(EUR)", "CollateralValue(EUR)", "CurrentLTV",
            "LiquidationLTV", "DistanceToLiq", "Status",
            "RequiredRepayEUR", "RequiredAddBTC"
        ]

        # Create table header
        for col_idx, header in enumerate(onderpand_cols, start=1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.style = "header_style"

        # Add formula rows for BTC loans (up to 10)
        for row_idx in range(4, 14):
            row_num = row_idx - 3  # 1, 2, 3...

            # LoanID - lookup from Leningen where OnderpandType = BTC
            ws.cell(row=row_idx, column=1, value=f'=IFERROR(INDEX(tblLeningen[LoanID],SMALL(IF(tblLeningen[OnderpandType]="BTC",ROW(tblLeningen[OnderpandType])-ROW(tblLeningen[[#Headers],[OnderpandType]])),{row_num})),"")')

            # Platform/Lener
            ws.cell(row=row_idx, column=2, value=f'=IF(A{row_idx}="","",XLOOKUP(A{row_idx},tblLeningen[LoanID],tblLeningen[Lener],""))')

            # BTC Onderpand - lookup from mapping or calculate from collateral
            ws.cell(row=row_idx, column=3, value=f'=IF(A{row_idx}="","",IFERROR(SUMIF(tblBTC[Wallet/Platform],XLOOKUP(A{row_idx},tblLeningen[LoanID],tblLeningen[OnderpandRef],""),tblBTC[BTC Aantal]),0))')

            # Openstaand(EUR)
            ws.cell(row=row_idx, column=4, value=f'=IF(A{row_idx}="","",XLOOKUP(A{row_idx},tblLeningen[LoanID],tblLeningen[Openstaand(EUR)],0))')

            # BTC prijs(EUR)
            ws.cell(row=row_idx, column=5, value=f'=IF(A{row_idx}="","",XLOOKUP("BTC",tblKoersen[Symbol],tblKoersen[EUR prijs],0))')

            # CollateralValue(EUR)
            ws.cell(row=row_idx, column=6, value=f'=IF(A{row_idx}="","",C{row_idx}*E{row_idx})')

            # CurrentLTV
            ws.cell(row=row_idx, column=7, value=f'=IF(OR(A{row_idx}="",F{row_idx}=0),"",D{row_idx}/F{row_idx})')

            # LiquidationLTV - from loan or default from settings
            ws.cell(row=row_idx, column=8, value=f'=IF(A{row_idx}="","",IF(XLOOKUP(A{row_idx},tblLeningen[LoanID],tblLeningen[LiquidatieLTV],"")="",XLOOKUP("DefaultLiquidationLTV_BTC",tblSettings[Parameter],tblSettings[Waarde]),XLOOKUP(A{row_idx},tblLeningen[LoanID],tblLeningen[LiquidatieLTV])))')

            # DistanceToLiq
            ws.cell(row=row_idx, column=9, value=f'=IF(OR(A{row_idx}="",G{row_idx}=""),"",H{row_idx}-G{row_idx})')

            # Status
            status_formula = (
                f'=IF(A{row_idx}="","",'
                f'IF(I{row_idx}<=XLOOKUP("RiskBuffer_BTC",tblSettings[Parameter],tblSettings[Waarde]),"RISK",'
                f'IF(I{row_idx}<=XLOOKUP("AlertBuffer_BTC",tblSettings[Parameter],tblSettings[Waarde]),"ALERT","OK")))'
            )
            ws.cell(row=row_idx, column=10, value=status_formula)

            # RequiredRepayEUR - amount to repay to reach TargetLTV
            ws.cell(row=row_idx, column=11, value=(
                f'=IF(OR(A{row_idx}="",J{row_idx}="OK"),"",'
                f'MAX(0,D{row_idx}-(XLOOKUP("TargetLTV_BTC",tblSettings[Parameter],tblSettings[Waarde])*F{row_idx})))'
            ))

            # RequiredAddBTC - BTC to add to reach TargetLTV
            ws.cell(row=row_idx, column=12, value=(
                f'=IF(OR(A{row_idx}="",J{row_idx}="OK"),"",'
                f'MAX(0,(D{row_idx}/XLOOKUP("TargetLTV_BTC",tblSettings[Parameter],tblSettings[Waarde])-F{row_idx})/E{row_idx}))'
            ))

        # Create table
        table = create_excel_table(
            ws,
            table_name="tblBTCOnderpand",
            columns=onderpand_cols,
            start_row=3,
            start_col=1,
            num_empty_rows=10,
        )

        # Number formats
        for row in range(4, 14):
            ws.cell(row=row, column=3).number_format = '0.00000000'  # BTC
            ws.cell(row=row, column=4).number_format = '€ #,##0.00'  # EUR
            ws.cell(row=row, column=5).number_format = '€ #,##0.00'  # BTC prijs
            ws.cell(row=row, column=6).number_format = '€ #,##0.00'  # CollateralValue
            ws.cell(row=row, column=7).number_format = '0.00%'       # CurrentLTV
            ws.cell(row=row, column=8).number_format = '0.00%'       # LiquidationLTV
            ws.cell(row=row, column=9).number_format = '0.00%'       # DistanceToLiq
            ws.cell(row=row, column=11).number_format = '€ #,##0.00' # RepayEUR
            ws.cell(row=row, column=12).number_format = '0.00000000' # AddBTC

        # Conditional formatting for Status column
        add_status_conditional_formatting(ws, "J4:J13")

        # Stress Test Section
        add_section_header(ws, "Stress Test Scenarios", row=16, col=1, span=6)

        # Stress test headers
        stress_headers = ["Scenario", "BTC Drop", "New BTC Price", "LoanID", "New LTV", "New Status"]
        for col_idx, header in enumerate(stress_headers, start=1):
            cell = ws.cell(row=17, column=col_idx, value=header)
            cell.style = "header_style"

        # Stress Scenario 1
        ws.cell(row=18, column=1, value="Scenario 1")
        ws.cell(row=18, column=2, value='=XLOOKUP("StressBTCDown1",tblSettings[Parameter],tblSettings[Waarde])')
        ws.cell(row=18, column=2).number_format = '0%'
        ws.cell(row=18, column=3, value='=XLOOKUP("BTC",tblKoersen[Symbol],tblKoersen[EUR prijs])*(1+B18)')
        ws.cell(row=18, column=3).number_format = '€ #,##0.00'
        ws.cell(row=18, column=4, value='=A4')  # First BTC loan
        ws.cell(row=18, column=5, value='=IF(D18="","",D4/(C4*$C$18))')
        ws.cell(row=18, column=5).number_format = '0.00%'
        ws.cell(row=18, column=6, value=(
            '=IF(D18="","",'
            'IF(E18>=H4,"LIQUIDATION",'
            'IF((H4-E18)<=XLOOKUP("RiskBuffer_BTC",tblSettings[Parameter],tblSettings[Waarde]),"RISK",'
            'IF((H4-E18)<=XLOOKUP("AlertBuffer_BTC",tblSettings[Parameter],tblSettings[Waarde]),"ALERT","OK"))))'
        ))

        # Stress Scenario 2
        ws.cell(row=19, column=1, value="Scenario 2")
        ws.cell(row=19, column=2, value='=XLOOKUP("StressBTCDown2",tblSettings[Parameter],tblSettings[Waarde])')
        ws.cell(row=19, column=2).number_format = '0%'
        ws.cell(row=19, column=3, value='=XLOOKUP("BTC",tblKoersen[Symbol],tblKoersen[EUR prijs])*(1+B19)')
        ws.cell(row=19, column=3).number_format = '€ #,##0.00'
        ws.cell(row=19, column=4, value='=A4')
        ws.cell(row=19, column=5, value='=IF(D19="","",D4/(C4*$C$19))')
        ws.cell(row=19, column=5).number_format = '0.00%'
        ws.cell(row=19, column=6, value=(
            '=IF(D19="","",'
            'IF(E19>=H4,"LIQUIDATION",'
            'IF((H4-E19)<=XLOOKUP("RiskBuffer_BTC",tblSettings[Parameter],tblSettings[Waarde]),"RISK",'
            'IF((H4-E19)<=XLOOKUP("AlertBuffer_BTC",tblSettings[Parameter],tblSettings[Waarde]),"ALERT","OK"))))'
        ))

        # Conditional formatting for stress status
        add_status_conditional_formatting(ws, "F18:F19")

        # Column widths
        set_column_widths(ws, {
            1: 12,  # LoanID
            2: 15,  # Platform
            3: 14,  # BTC Onderpand
            4: 15,  # Openstaand
            5: 14,  # BTC prijs
            6: 16,  # CollateralValue
            7: 12,  # CurrentLTV
            8: 13,  # LiquidationLTV
            9: 13,  # DistanceToLiq
            10: 10, # Status
            11: 16, # RepayEUR
            12: 14, # AddBTC
        })

        freeze_panes(ws, row=4, col=1)

    # =========================================================================
    # CASH & OVERIG Sheet
    # =========================================================================

    def _build_cash_sheet(self) -> None:
        """Build the Cash & Other assets sheet."""
        ws = self.sheets["Cash & Overig"]

        add_sheet_title(ws, "Cash & Overig - Liquide Middelen", row=1, col=1)

        # Cash table
        cash_cols = ["Datum", "Rekening/Platform", "Bedrag(EUR)", "Type", "Opmerking"]
        cash_data = SAMPLE_DATA.get("tblCash", [])

        create_excel_table(
            ws,
            table_name="tblCash",
            columns=cash_cols,
            start_row=3,
            start_col=1,
            data=cash_data,
            num_empty_rows=15,
        )

        # Data validation for Type
        add_dropdown_validation(ws, "cash_type", "D4:D50")

        # Number formats
        for row in range(4, 50):
            ws.cell(row=row, column=1).number_format = 'YYYY-MM-DD'
            ws.cell(row=row, column=3).number_format = '€ #,##0.00'

        # Summary section
        summary_row = 2
        ws.cell(row=summary_row, column=7, value="Totaal Cash:")
        ws.cell(row=summary_row, column=7).font = Font(bold=True)
        ws.cell(row=summary_row, column=8, value="=SUM(tblCash[Bedrag(EUR)])")
        ws.cell(row=summary_row, column=8).number_format = '€ #,##0.00'
        ws.cell(row=summary_row, column=8).font = Font(bold=True)

        # Cash buffer calculation
        ws.cell(row=summary_row + 2, column=7, value="Maandlasten (inst.):")
        ws.cell(row=summary_row + 2, column=8, value='=XLOOKUP("MonthlyExpenses",tblSettings[Parameter],tblSettings[Waarde])')
        ws.cell(row=summary_row + 2, column=8).number_format = '€ #,##0.00'

        ws.cell(row=summary_row + 3, column=7, value="Buffer (maanden):")
        ws.cell(row=summary_row + 3, column=8, value='=IF(H4=0,0,H2/H4)')
        ws.cell(row=summary_row + 3, column=8).number_format = '0.0'

        ws.cell(row=summary_row + 4, column=7, value="Target (maanden):")
        ws.cell(row=summary_row + 4, column=8, value='=XLOOKUP("CashBufferMonthsTarget",tblSettings[Parameter],tblSettings[Waarde])')
        ws.cell(row=summary_row + 4, column=8).number_format = '0'

        # Column widths
        set_column_widths(ws, {
            1: 12,  # Datum
            2: 25,  # Rekening/Platform
            3: 14,  # Bedrag
            4: 15,  # Type
            5: 30,  # Opmerking
            6: 3,   # Spacer
            7: 18,  # Label
            8: 15,  # Value
        })

        freeze_panes(ws, row=4, col=1)

    # =========================================================================
    # ALERTS Sheet
    # =========================================================================

    def _build_alerts_sheet(self) -> None:
        """Build the Alerts sheet with rule-based insights."""
        ws = self.sheets["Alerts"]

        add_sheet_title(ws, "Alerts - Inzichten & Acties", row=1, col=1)

        # Alerts table columns
        alerts_cols = [
            "Date", "Severity", "Category", "Item",
            "Metric", "Value", "Threshold", "Insight", "SuggestedAction"
        ]

        # Create header
        for col_idx, header in enumerate(alerts_cols, start=1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.style = "header_style"

        # Alert formulas (each row checks a specific condition)
        alert_row = 4

        # Alert 1: BTC Loan LTV Status (for first BTC loan)
        ws.cell(row=alert_row, column=1, value="=TODAY()")
        ws.cell(row=alert_row, column=2, value='=IF(\'BTC als Onderpand\'!J4="RISK","CRITICAL",IF(\'BTC als Onderpand\'!J4="ALERT","ALERT",""))')
        ws.cell(row=alert_row, column=3, value='=IF(B4="","","BTC Lening")')
        ws.cell(row=alert_row, column=4, value='=IF(B4="","",\'BTC als Onderpand\'!A4)')
        ws.cell(row=alert_row, column=5, value='=IF(B4="","","LTV")')
        ws.cell(row=alert_row, column=6, value='=IF(B4="","",\'BTC als Onderpand\'!G4)')
        ws.cell(row=alert_row, column=7, value='=IF(B4="","",\'BTC als Onderpand\'!H4)')
        ws.cell(row=alert_row, column=8, value='=IF(B4="","",IF(B4="CRITICAL","LTV kritiek - liquidatie dreigt!","LTV te hoog - actie vereist"))')
        ws.cell(row=alert_row, column=9, value='=IF(B4="","",CONCAT("Los ",TEXT(\'BTC als Onderpand\'!K4,"€ #.##0")," af OF stort ",TEXT(\'BTC als Onderpand\'!L4,"0.0000")," BTC bij"))')
        alert_row += 1

        # Alert 2: Cash Buffer Alert
        ws.cell(row=alert_row, column=1, value="=TODAY()")
        ws.cell(row=alert_row, column=2, value='=IF(\'Cash & Overig\'!H5<XLOOKUP("CashBufferMonthsTarget",tblSettings[Parameter],tblSettings[Waarde]),"ALERT","")')
        ws.cell(row=alert_row, column=3, value='=IF(B5="","","Cash Buffer")')
        ws.cell(row=alert_row, column=4, value='=IF(B5="","","Totaal Cash")')
        ws.cell(row=alert_row, column=5, value='=IF(B5="","","Maanden buffer")')
        ws.cell(row=alert_row, column=6, value='=IF(B5="","",\'Cash & Overig\'!H5)')
        ws.cell(row=alert_row, column=7, value='=IF(B5="","",XLOOKUP("CashBufferMonthsTarget",tblSettings[Parameter],tblSettings[Waarde]))')
        ws.cell(row=alert_row, column=8, value='=IF(B5="","","Cashbuffer onder target - verhoog reserves")')
        ws.cell(row=alert_row, column=9, value='=IF(B5="","",CONCAT("Verhoog cash met €",TEXT((B7-\'Cash & Overig\'!H5)*\'Cash & Overig\'!H4,"#.##0")))')
        alert_row += 1

        # Alert 3: Concentration Risk
        ws.cell(row=alert_row, column=1, value="=TODAY()")
        # Calculate crypto concentration
        concentration_formula = (
            '=LET('
            'btc_val,SUM(tblBTC[BTC Aantal])*XLOOKUP("BTC",tblKoersen[Symbol],tblKoersen[EUR prijs],0),'
            'doge_val,SUM(tblDOGE[DOGE Aantal])*XLOOKUP("DOGE",tblKoersen[Symbol],tblKoersen[EUR prijs],0),'
            'crypto_total,btc_val+doge_val,'
            'vastgoed_val,SUM(tblVastgoed[Waarde(EUR)]),'
            'total,crypto_total+vastgoed_val+SUM(tblCash[Bedrag(EUR)]),'
            'concentration,IF(total=0,0,crypto_total/total),'
            'threshold,XLOOKUP("ConcentrationAlertPct",tblSettings[Parameter],tblSettings[Waarde]),'
            'IF(concentration>threshold,"ALERT",""))'
        )
        ws.cell(row=alert_row, column=2, value=concentration_formula)
        ws.cell(row=alert_row, column=3, value='=IF(B6="","","Concentratie")')
        ws.cell(row=alert_row, column=4, value='=IF(B6="","","Crypto Allocatie")')
        ws.cell(row=alert_row, column=5, value='=IF(B6="","","% van totaal")')
        # Value calculation
        ws.cell(row=alert_row, column=6, value=(
            '=IF(B6="","",LET('
            'btc_val,SUM(tblBTC[BTC Aantal])*XLOOKUP("BTC",tblKoersen[Symbol],tblKoersen[EUR prijs],0),'
            'doge_val,SUM(tblDOGE[DOGE Aantal])*XLOOKUP("DOGE",tblKoersen[Symbol],tblKoersen[EUR prijs],0),'
            'crypto_total,btc_val+doge_val,'
            'vastgoed_val,SUM(tblVastgoed[Waarde(EUR)]),'
            'total,crypto_total+vastgoed_val+SUM(tblCash[Bedrag(EUR)]),'
            'IF(total=0,0,crypto_total/total)))'
        ))
        ws.cell(row=alert_row, column=7, value='=IF(B6="","",XLOOKUP("ConcentrationAlertPct",tblSettings[Parameter],tblSettings[Waarde]))')
        ws.cell(row=alert_row, column=8, value='=IF(B6="","","Crypto-allocatie te hoog - overweeg diversificatie")')
        ws.cell(row=alert_row, column=9, value='=IF(B6="","","Overweeg winst te nemen of te diversifieren naar andere assets")')
        alert_row += 1

        # Alert 4: Stress Test Scenario 1
        ws.cell(row=alert_row, column=1, value="=TODAY()")
        ws.cell(row=alert_row, column=2, value='=IF(OR(\'BTC als Onderpand\'!F18="RISK",\'BTC als Onderpand\'!F18="LIQUIDATION"),"ALERT","")')
        ws.cell(row=alert_row, column=3, value='=IF(B7="","","Stress Test")')
        ws.cell(row=alert_row, column=4, value='=IF(B7="","","Scenario 1")')
        ws.cell(row=alert_row, column=5, value='=IF(B7="","","LTV bij daling")')
        ws.cell(row=alert_row, column=6, value='=IF(B7="","",\'BTC als Onderpand\'!E18)')
        ws.cell(row=alert_row, column=7, value='=IF(B7="","",\'BTC als Onderpand\'!H4)')
        ws.cell(row=alert_row, column=8, value='=IF(B7="","",CONCAT("Bij ",TEXT(\'BTC als Onderpand\'!B18,"0%")," BTC-daling: ",\'BTC als Onderpand\'!F18))')
        ws.cell(row=alert_row, column=9, value='=IF(B7="","","Overweeg preventief af te lossen of collateral te verhogen")')
        alert_row += 1

        # Alert 5: Stress Test Scenario 2
        ws.cell(row=alert_row, column=1, value="=TODAY()")
        ws.cell(row=alert_row, column=2, value='=IF(OR(\'BTC als Onderpand\'!F19="RISK",\'BTC als Onderpand\'!F19="LIQUIDATION"),"ALERT","")')
        ws.cell(row=alert_row, column=3, value='=IF(B8="","","Stress Test")')
        ws.cell(row=alert_row, column=4, value='=IF(B8="","","Scenario 2")')
        ws.cell(row=alert_row, column=5, value='=IF(B8="","","LTV bij daling")')
        ws.cell(row=alert_row, column=6, value='=IF(B8="","",\'BTC als Onderpand\'!E19)')
        ws.cell(row=alert_row, column=7, value='=IF(B8="","",\'BTC als Onderpand\'!H4)')
        ws.cell(row=alert_row, column=8, value='=IF(B8="","",CONCAT("Bij ",TEXT(\'BTC als Onderpand\'!B19,"0%")," BTC-daling: ",\'BTC als Onderpand\'!F19))')
        ws.cell(row=alert_row, column=9, value='=IF(B8="","","Actie vereist: liquidatierisico bij sterke daling")')

        # Create table (include all potential alert rows)
        create_excel_table(
            ws,
            table_name="tblAlerts",
            columns=alerts_cols,
            start_row=3,
            start_col=1,
            num_empty_rows=10,
        )

        # Number formats
        for row in range(4, 15):
            ws.cell(row=row, column=1).number_format = 'YYYY-MM-DD'
            ws.cell(row=row, column=6).number_format = '0.00%'
            ws.cell(row=row, column=7).number_format = '0.00%'

        # Conditional formatting for Severity
        add_severity_conditional_formatting(ws, "B4:B15")

        # Legend
        ws.cell(row=17, column=1, value="LEGENDA:")
        ws.cell(row=17, column=1).font = Font(bold=True)
        ws.cell(row=18, column=1, value="• Alleen actieve alerts worden getoond (lege rijen = geen alert)")
        ws.cell(row=19, column=1, value="• CRITICAL = Directe actie vereist")
        ws.cell(row=20, column=1, value="• ALERT = Aandacht vereist")
        ws.cell(row=21, column=1, value="• INFO = Ter informatie")

        # Column widths
        set_column_widths(ws, {
            1: 12,  # Date
            2: 10,  # Severity
            3: 14,  # Category
            4: 15,  # Item
            5: 15,  # Metric
            6: 12,  # Value
            7: 12,  # Threshold
            8: 45,  # Insight
            9: 50,  # SuggestedAction
        })

        freeze_panes(ws, row=4, col=1)

    # =========================================================================
    # DASHBOARD Sheet
    # =========================================================================

    def _build_dashboard_sheet(self) -> None:
        """Build the main Dashboard sheet."""
        ws = self.sheets["Dashboard"]

        # Title
        title_cell = ws.cell(row=1, column=1, value="FINANCIËLE COCKPIT")
        title_cell.font = Font(bold=True, size=18, color=COLORS["header_bg"])

        # Last update
        ws.cell(row=1, column=6, value="Laatste update:")
        ws.cell(row=1, column=7, value="=NOW()")
        ws.cell(row=1, column=7).number_format = 'DD-MM-YYYY HH:MM'

        # =====================================================================
        # SECTION 1: CRYPTO HOLDINGS
        # =====================================================================
        section_row = 3
        add_section_header(ws, "CRYPTO HOLDINGS", section_row, 1, span=4)

        row = section_row + 1
        # BTC Value
        ws.cell(row=row, column=1, value="Bitcoin (BTC)")
        ws.cell(row=row, column=2, value="=SUM(tblBTC[BTC Aantal])")
        ws.cell(row=row, column=2).number_format = '0.00000000 "BTC"'
        ws.cell(row=row, column=3, value='=SUM(tblBTC[BTC Aantal])*XLOOKUP("BTC",tblKoersen[Symbol],tblKoersen[EUR prijs],0)')
        ws.cell(row=row, column=3).number_format = '€ #,##0'
        row += 1

        # DOGE Value
        ws.cell(row=row, column=1, value="Dogecoin (DOGE)")
        ws.cell(row=row, column=2, value="=SUM(tblDOGE[DOGE Aantal])")
        ws.cell(row=row, column=2).number_format = '#,##0 "DOGE"'
        ws.cell(row=row, column=3, value='=SUM(tblDOGE[DOGE Aantal])*XLOOKUP("DOGE",tblKoersen[Symbol],tblKoersen[EUR prijs],0)')
        ws.cell(row=row, column=3).number_format = '€ #,##0'
        row += 1

        # Crypto Total
        ws.cell(row=row, column=1, value="Totaal Crypto")
        ws.cell(row=row, column=1).font = Font(bold=True)
        ws.cell(row=row, column=3, value="=C4+C5")
        ws.cell(row=row, column=3).number_format = '€ #,##0'
        ws.cell(row=row, column=3).font = Font(bold=True, size=12)
        row += 1

        # Spot vs Collateral breakdown
        ws.cell(row=row, column=1, value="  waarvan Spot:")
        ws.cell(row=row, column=3, value='=SUMIF(tblBTC[Type],"Spot",tblBTC[BTC Aantal])*XLOOKUP("BTC",tblKoersen[Symbol],tblKoersen[EUR prijs],0)+SUMIF(tblDOGE[Type],"Spot",tblDOGE[DOGE Aantal])*XLOOKUP("DOGE",tblKoersen[Symbol],tblKoersen[EUR prijs],0)')
        ws.cell(row=row, column=3).number_format = '€ #,##0'
        ws.cell(row=row, column=3).font = Font(italic=True)
        row += 1

        ws.cell(row=row, column=1, value="  waarvan Collateral:")
        ws.cell(row=row, column=3, value='=SUMIF(tblBTC[Type],"Collateral",tblBTC[BTC Aantal])*XLOOKUP("BTC",tblKoersen[Symbol],tblKoersen[EUR prijs],0)+SUMIF(tblDOGE[Type],"Collateral",tblDOGE[DOGE Aantal])*XLOOKUP("DOGE",tblKoersen[Symbol],tblKoersen[EUR prijs],0)')
        ws.cell(row=row, column=3).number_format = '€ #,##0'
        ws.cell(row=row, column=3).font = Font(italic=True)

        # =====================================================================
        # SECTION 2: VASTGOED
        # =====================================================================
        section_row = 11
        add_section_header(ws, "VASTGOED", section_row, 1, span=4)

        row = section_row + 1
        ws.cell(row=row, column=1, value="Totale Waarde")
        ws.cell(row=row, column=3, value="=SUM(tblVastgoed[Waarde(EUR)])")
        ws.cell(row=row, column=3).number_format = '€ #,##0'
        row += 1

        ws.cell(row=row, column=1, value="Hypotheken")
        ws.cell(row=row, column=3, value="=SUM(tblVastgoed[Hypotheek(EUR)])")
        ws.cell(row=row, column=3).number_format = '€ #,##0'
        ws.cell(row=row, column=3).font = Font(color="CC0000")
        row += 1

        ws.cell(row=row, column=1, value="Overige Leningen")
        ws.cell(row=row, column=3, value="=SUM(tblVastgoed[Overige Leningen(EUR)])")
        ws.cell(row=row, column=3).number_format = '€ #,##0'
        ws.cell(row=row, column=3).font = Font(color="CC0000")
        row += 1

        ws.cell(row=row, column=1, value="Netto Vastgoed")
        ws.cell(row=row, column=1).font = Font(bold=True)
        ws.cell(row=row, column=3, value="=SUM(tblVastgoed[Netto Waarde])")
        ws.cell(row=row, column=3).number_format = '€ #,##0'
        ws.cell(row=row, column=3).font = Font(bold=True, size=12)

        # =====================================================================
        # SECTION 3: CASH & LENINGEN
        # =====================================================================
        section_row = 18
        add_section_header(ws, "CASH & LENINGEN", section_row, 1, span=4)

        row = section_row + 1
        ws.cell(row=row, column=1, value="Cash & Overig")
        ws.cell(row=row, column=3, value="=SUM(tblCash[Bedrag(EUR)])")
        ws.cell(row=row, column=3).number_format = '€ #,##0'
        row += 1

        ws.cell(row=row, column=1, value="BTC-leningen")
        ws.cell(row=row, column=3, value='=SUMIF(tblLeningen[OnderpandType],"BTC",tblLeningen[Openstaand(EUR)])')
        ws.cell(row=row, column=3).number_format = '€ #,##0'
        ws.cell(row=row, column=3).font = Font(color="CC0000")
        row += 1

        ws.cell(row=row, column=1, value="Overige Leningen")
        ws.cell(row=row, column=3, value='=SUMIF(tblLeningen[OnderpandType],"Vastgoed",tblLeningen[Openstaand(EUR)])')
        ws.cell(row=row, column=3).number_format = '€ #,##0'
        ws.cell(row=row, column=3).font = Font(color="CC0000")
        row += 1

        ws.cell(row=row, column=1, value="Cash Buffer")
        ws.cell(row=row, column=3, value='=\'Cash & Overig\'!H5')
        ws.cell(row=row, column=3).number_format = '0.0 "maanden"'

        # =====================================================================
        # SECTION 4: NETTO VERMOGEN
        # =====================================================================
        section_row = 25
        add_section_header(ws, "NETTO VERMOGEN", section_row, 1, span=4)

        row = section_row + 1
        ws.cell(row=row, column=1, value="Totaal Activa")
        ws.cell(row=row, column=3, value="=C6+C12+C19")  # Crypto + Vastgoed + Cash
        ws.cell(row=row, column=3).number_format = '€ #,##0'
        row += 1

        ws.cell(row=row, column=1, value="Totaal Schulden")
        ws.cell(row=row, column=3, value="=C13+C14+C20+C21")  # All debts
        ws.cell(row=row, column=3).number_format = '€ #,##0'
        ws.cell(row=row, column=3).font = Font(color="CC0000")
        row += 1

        # Net Worth highlight
        nw_row = row
        ws.cell(row=nw_row, column=1, value="NETTO VERMOGEN")
        ws.cell(row=nw_row, column=1).font = Font(bold=True, size=14)
        ws.cell(row=nw_row, column=3, value="=C26-C27")
        ws.cell(row=nw_row, column=3).number_format = '€ #,##0'
        ws.cell(row=nw_row, column=3).font = Font(bold=True, size=14, color="006600")

        # Highlight net worth row
        for col in range(1, 5):
            cell = ws.cell(row=nw_row, column=col)
            cell.fill = PatternFill(start_color=COLORS["calc_bg"], end_color=COLORS["calc_bg"], fill_type="solid")

        # =====================================================================
        # SECTION 5: ALERTS SAMENVATTING (RIGHT SIDE)
        # =====================================================================
        section_row = 3
        add_section_header(ws, "TOP ALERTS", section_row, 6, span=3)

        row = section_row + 1
        # Show first 3 alerts with content
        for i in range(3):
            alert_row_num = 4 + i
            ws.cell(row=row + i, column=6, value=f'=IF(Alerts!B{alert_row_num}="","",Alerts!B{alert_row_num})')
            ws.cell(row=row + i, column=7, value=f'=IF(Alerts!B{alert_row_num}="","",Alerts!H{alert_row_num})')

        # Conditional formatting for alert severity
        add_severity_conditional_formatting(ws, "F4:F6")

        # Link to Alerts sheet
        ws.cell(row=row + 3, column=6, value="→ Zie Alerts sheet voor details")
        ws.cell(row=row + 3, column=6).font = Font(italic=True, color="0066CC")

        # =====================================================================
        # SECTION 6: STRESS TEST SAMENVATTING
        # =====================================================================
        section_row = 11
        add_section_header(ws, "STRESS TEST", section_row, 6, span=3)

        row = section_row + 1
        ws.cell(row=row, column=6, value="Scenario 1:")
        ws.cell(row=row, column=7, value='=XLOOKUP("StressBTCDown1",tblSettings[Parameter],tblSettings[Waarde])')
        ws.cell(row=row, column=7).number_format = '0%'
        ws.cell(row=row, column=8, value='=\'BTC als Onderpand\'!F18')
        row += 1

        ws.cell(row=row, column=6, value="Scenario 2:")
        ws.cell(row=row, column=7, value='=XLOOKUP("StressBTCDown2",tblSettings[Parameter],tblSettings[Waarde])')
        ws.cell(row=row, column=7).number_format = '0%'
        ws.cell(row=row, column=8, value='=\'BTC als Onderpand\'!F19')

        # Conditional formatting for stress status
        add_status_conditional_formatting(ws, "H12:H13")

        # =====================================================================
        # SECTION 7: KOERSEN SNAPSHOT
        # =====================================================================
        section_row = 16
        add_section_header(ws, "KOERSEN", section_row, 6, span=3)

        row = section_row + 1
        ws.cell(row=row, column=6, value="BTC")
        ws.cell(row=row, column=7, value='=XLOOKUP("BTC",tblKoersen[Symbol],tblKoersen[EUR prijs],0)')
        ws.cell(row=row, column=7).number_format = '€ #,##0.00'
        row += 1

        ws.cell(row=row, column=6, value="DOGE")
        ws.cell(row=row, column=7, value='=XLOOKUP("DOGE",tblKoersen[Symbol],tblKoersen[EUR prijs],0)')
        ws.cell(row=row, column=7).number_format = '€ 0.0000'

        # =====================================================================
        # SECTION 8: BTC LTV STATUS
        # =====================================================================
        section_row = 21
        add_section_header(ws, "BTC LENING STATUS", section_row, 6, span=3)

        row = section_row + 1
        ws.cell(row=row, column=6, value="Eerste lening:")
        ws.cell(row=row, column=6).font = Font(bold=True)
        row += 1

        ws.cell(row=row, column=6, value="LoanID")
        ws.cell(row=row, column=7, value='=\'BTC als Onderpand\'!A4')
        row += 1

        ws.cell(row=row, column=6, value="Current LTV")
        ws.cell(row=row, column=7, value='=\'BTC als Onderpand\'!G4')
        ws.cell(row=row, column=7).number_format = '0.00%'
        row += 1

        ws.cell(row=row, column=6, value="Distance to Liq")
        ws.cell(row=row, column=7, value='=\'BTC als Onderpand\'!I4')
        ws.cell(row=row, column=7).number_format = '0.00%'
        row += 1

        ws.cell(row=row, column=6, value="Status")
        ws.cell(row=row, column=7, value='=\'BTC als Onderpand\'!J4')

        # Conditional formatting for status
        add_status_conditional_formatting(ws, "G26:G26")

        # =====================================================================
        # COLUMN WIDTHS & FORMATTING
        # =====================================================================
        set_column_widths(ws, {
            1: 20,  # Label
            2: 18,  # Amount
            3: 15,  # Value EUR
            4: 3,   # Spacer
            5: 3,   # Spacer
            6: 15,  # Right section label
            7: 15,  # Right section value
            8: 12,  # Right section extra
        })

        freeze_panes(ws, row=2, col=1)


def main():
    """Main entry point."""
    # Determine output path
    if len(sys.argv) > 1:
        output_path = sys.argv[1]
    else:
        # Default to dist folder
        script_dir = Path(__file__).parent.parent
        dist_dir = script_dir / "dist"
        dist_dir.mkdir(exist_ok=True)
        output_path = str(dist_dir / "Financiele_Cockpit.xlsx")

    # Generate workbook
    generator = FinancialCockpitGenerator()
    generator.generate(output_path)

    print("\n✓ Financiële Cockpit gegenereerd!")
    print("\nVolgende stappen:")
    print("1. Open het bestand in Excel")
    print("2. Configureer Power Query voor live koersen (zie Koers-Setup)")
    print("3. Pas thresholds aan in Instellingen sheet")
    print("4. Vul je holdings in op de respectievelijke sheets")
    print("5. Dashboard en Alerts updaten automatisch!")


if __name__ == "__main__":
    main()
