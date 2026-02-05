"""
Model Configuration for Financiële Cockpit
Contains all asset definitions, sheet specifications, and default settings.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional

# =============================================================================
# ASSET DEFINITIONS
# =============================================================================

CRYPTO_ASSETS = [
    {
        "symbol": "BTC",
        "name": "Bitcoin",
        "coingecko_id": "bitcoin",
        "decimals": 8,
        "sheet_name": "Bitcoin Holdings",
        "table_name": "tblBTC",
    },
    {
        "symbol": "DOGE",
        "name": "Dogecoin",
        "coingecko_id": "dogecoin",
        "decimals": 4,
        "sheet_name": "Dogecoin Holdings",
        "table_name": "tblDOGE",
    },
]

# =============================================================================
# SETTINGS PARAMETERS (Default values)
# =============================================================================

SETTINGS_PARAMETERS = [
    # Refresh settings
    {
        "parameter": "RefreshMinutes",
        "value": 10,
        "unit": "minuten",
        "description": "Interval voor automatische koers-refresh in Excel",
    },
    # BTC Loan LTV Thresholds
    {
        "parameter": "TargetLTV_BTC",
        "value": 0.30,
        "unit": "ratio",
        "description": "Doel-LTV voor BTC-leningen (30%)",
    },
    {
        "parameter": "AlertBuffer_BTC",
        "value": 0.10,
        "unit": "ratio",
        "description": "Afstand tot liquidatie waarbij ALERT getriggerd wordt",
    },
    {
        "parameter": "RiskBuffer_BTC",
        "value": 0.05,
        "unit": "ratio",
        "description": "Afstand tot liquidatie waarbij RISK getriggerd wordt",
    },
    {
        "parameter": "DefaultLiquidationLTV_BTC",
        "value": 0.50,
        "unit": "ratio",
        "description": "Standaard liquidatie-LTV voor BTC-leningen (50%)",
    },
    # Cash & Debt thresholds
    {
        "parameter": "CashBufferMonthsTarget",
        "value": 6,
        "unit": "maanden",
        "description": "Minimaal aantal maanden cashbuffer",
    },
    {
        "parameter": "MaxDebtServiceRatio",
        "value": 0.35,
        "unit": "ratio",
        "description": "Maximale schuld-service ratio (35%)",
    },
    {
        "parameter": "MonthlyExpenses",
        "value": 5000,
        "unit": "EUR",
        "description": "Geschatte maandelijkse uitgaven (voor cashbuffer berekening)",
    },
    # Stress test scenarios
    {
        "parameter": "StressBTCDown1",
        "value": -0.20,
        "unit": "percentage",
        "description": "Stress scenario 1: BTC daalt met 20%",
    },
    {
        "parameter": "StressBTCDown2",
        "value": -0.40,
        "unit": "percentage",
        "description": "Stress scenario 2: BTC daalt met 40%",
    },
    # Concentration
    {
        "parameter": "ConcentrationAlertPct",
        "value": 0.60,
        "unit": "ratio",
        "description": "Alert als crypto > 60% van totaal vermogen",
    },
]

# =============================================================================
# SHEET SPECIFICATIONS
# =============================================================================

SHEETS_CONFIG = [
    {"name": "Dashboard", "order": 0, "type": "dashboard"},
    {"name": "Alerts", "order": 1, "type": "alerts"},
    {"name": "Instellingen", "order": 2, "type": "settings"},
    {"name": "Koersen", "order": 3, "type": "prices"},
    {"name": "Koers-Setup", "order": 4, "type": "instructions"},
    {"name": "Bitcoin Holdings", "order": 5, "type": "holdings"},
    {"name": "Dogecoin Holdings", "order": 6, "type": "holdings"},
    {"name": "Vastgoed", "order": 7, "type": "real_estate"},
    {"name": "Leningen", "order": 8, "type": "loans"},
    {"name": "BTC als Onderpand", "order": 9, "type": "collateral_view"},
    {"name": "Cash & Overig", "order": 10, "type": "cash"},
]

# =============================================================================
# TABLE DEFINITIONS
# =============================================================================

TABLE_DEFINITIONS = {
    "tblSettings": {
        "sheet": "Instellingen",
        "columns": ["Parameter", "Waarde", "Eenheid", "Toelichting"],
        "start_row": 2,
        "start_col": 1,
    },
    "tblLoanCollateralMap": {
        "sheet": "Instellingen",
        "columns": ["LoanID", "CollateralType", "CollateralRef", "Notes"],
        "start_row": 2,
        "start_col": 7,
    },
    "tblKoersen": {
        "sheet": "Koersen",
        "columns": ["Asset", "Symbol", "EUR prijs", "USD prijs", "Laatste update"],
        "start_row": 2,
        "start_col": 1,
    },
    "tblBTC": {
        "sheet": "Bitcoin Holdings",
        "columns": ["Datum", "Wallet/Platform", "BTC Aantal", "Type", "Opmerking"],
        "start_row": 2,
        "start_col": 1,
    },
    "tblDOGE": {
        "sheet": "Dogecoin Holdings",
        "columns": ["Datum", "Wallet/Platform", "DOGE Aantal", "Type", "Opmerking"],
        "start_row": 2,
        "start_col": 1,
    },
    "tblVastgoed": {
        "sheet": "Vastgoed",
        "columns": [
            "Pand",
            "Adres",
            "Type",
            "Waarde(EUR)",
            "Hypotheek(EUR)",
            "Overige Leningen(EUR)",
            "Netto Waarde",
        ],
        "start_row": 2,
        "start_col": 1,
    },
    "tblLeningen": {
        "sheet": "Leningen",
        "columns": [
            "LoanID",
            "Datum",
            "OnderpandType",
            "OnderpandRef",
            "Lener",
            "Hoofdsom(EUR)",
            "Openstaand(EUR)",
            "Rente%",
            "Looptijd",
            "LiquidatieLTV",
            "Opmerking",
        ],
        "start_row": 2,
        "start_col": 1,
    },
    "tblBTCOnderpand": {
        "sheet": "BTC als Onderpand",
        "columns": [
            "LoanID",
            "Platform/Lener",
            "BTC Onderpand",
            "Openstaand(EUR)",
            "BTC prijs(EUR)",
            "CollateralValue(EUR)",
            "CurrentLTV",
            "LiquidationLTV",
            "DistanceToLiq",
            "Status",
            "RequiredRepayEUR",
            "RequiredAddBTC",
        ],
        "start_row": 2,
        "start_col": 1,
    },
    "tblCash": {
        "sheet": "Cash & Overig",
        "columns": ["Datum", "Rekening/Platform", "Bedrag(EUR)", "Type", "Opmerking"],
        "start_row": 2,
        "start_col": 1,
    },
    "tblAlerts": {
        "sheet": "Alerts",
        "columns": [
            "Date",
            "Severity",
            "Category",
            "Item",
            "Metric",
            "Value",
            "Threshold",
            "Insight",
            "SuggestedAction",
        ],
        "start_row": 2,
        "start_col": 1,
    },
}

# =============================================================================
# SAMPLE DATA (for demonstration)
# =============================================================================

SAMPLE_DATA = {
    "tblBTC": [
        {
            "Datum": "2024-01-15",
            "Wallet/Platform": "Hardware Wallet",
            "BTC Aantal": 0.5,
            "Type": "Spot",
            "Opmerking": "Cold storage",
        },
        {
            "Datum": "2024-02-01",
            "Wallet/Platform": "BlockFi",
            "BTC Aantal": 0.25,
            "Type": "Collateral",
            "Opmerking": "Onderpand voor lening L001",
        },
        {
            "Datum": "2024-03-10",
            "Wallet/Platform": "Bitvavo",
            "BTC Aantal": 0.15,
            "Type": "Spot",
            "Opmerking": "Exchange",
        },
    ],
    "tblDOGE": [
        {
            "Datum": "2024-01-20",
            "Wallet/Platform": "Bitvavo",
            "DOGE Aantal": 10000,
            "Type": "Spot",
            "Opmerking": "",
        },
        {
            "Datum": "2024-02-15",
            "Wallet/Platform": "Kraken",
            "DOGE Aantal": 5000,
            "Type": "Spot",
            "Opmerking": "",
        },
    ],
    "tblVastgoed": [
        {
            "Pand": "Woning Amsterdam",
            "Adres": "Voorbeeldstraat 123, Amsterdam",
            "Type": "Eigen woning",
            "Waarde(EUR)": 650000,
            "Hypotheek(EUR)": 400000,
            "Overige Leningen(EUR)": 0,
        },
        {
            "Pand": "Appartement Utrecht",
            "Adres": "Voorbeeldlaan 45, Utrecht",
            "Type": "Verhuur",
            "Waarde(EUR)": 350000,
            "Hypotheek(EUR)": 200000,
            "Overige Leningen(EUR)": 25000,
        },
    ],
    "tblLeningen": [
        {
            "LoanID": "L001",
            "Datum": "2024-02-01",
            "OnderpandType": "BTC",
            "OnderpandRef": "BlockFi",
            "Lener": "BlockFi",
            "Hoofdsom(EUR)": 5000,
            "Openstaand(EUR)": 4800,
            "Rente%": 0.08,
            "Looptijd": "12 maanden",
            "LiquidatieLTV": 0.50,
            "Opmerking": "BTC-backed loan",
        },
        {
            "LoanID": "L002",
            "Datum": "2023-06-01",
            "OnderpandType": "Vastgoed",
            "OnderpandRef": "Appartement Utrecht",
            "Lener": "ING Bank",
            "Hoofdsom(EUR)": 30000,
            "Openstaand(EUR)": 25000,
            "Rente%": 0.045,
            "Looptijd": "60 maanden",
            "LiquidatieLTV": None,
            "Opmerking": "Verbouwingslening",
        },
    ],
    "tblCash": [
        {
            "Datum": "2024-03-01",
            "Rekening/Platform": "ING Betaalrekening",
            "Bedrag(EUR)": 15000,
            "Type": "Betaalrekening",
            "Opmerking": "",
        },
        {
            "Datum": "2024-03-01",
            "Rekening/Platform": "ING Spaarrekening",
            "Bedrag(EUR)": 35000,
            "Type": "Spaarrekening",
            "Opmerking": "Noodfonds",
        },
        {
            "Datum": "2024-03-01",
            "Rekening/Platform": "DEGIRO",
            "Bedrag(EUR)": 8000,
            "Type": "Belegging",
            "Opmerking": "ETF portfolio",
        },
    ],
    "tblKoersen": [
        {
            "Asset": "Bitcoin",
            "Symbol": "BTC",
            "EUR prijs": 45000,
            "USD prijs": 48000,
            "Laatste update": "=NOW()",
        },
        {
            "Asset": "Dogecoin",
            "Symbol": "DOGE",
            "EUR prijs": 0.08,
            "USD prijs": 0.085,
            "Laatste update": "=NOW()",
        },
    ],
    "tblLoanCollateralMap": [
        {
            "LoanID": "L001",
            "CollateralType": "BTC",
            "CollateralRef": "BlockFi",
            "Notes": "0.25 BTC als onderpand",
        },
    ],
}

# =============================================================================
# FORMULA TEMPLATES
# =============================================================================

FORMULA_TEMPLATES = {
    # Get setting value by parameter name
    "get_setting": '=XLOOKUP("{param}",tblSettings[Parameter],tblSettings[Waarde])',
    # Get price by symbol
    "get_price_eur": '=XLOOKUP("{symbol}",tblKoersen[Symbol],tblKoersen[EUR prijs])',
    # Sum holdings by type
    "sum_btc_spot": '=SUMIF(tblBTC[Type],"Spot",tblBTC[BTC Aantal])',
    "sum_btc_collateral": '=SUMIF(tblBTC[Type],"Collateral",tblBTC[BTC Aantal])',
    "sum_btc_total": "=SUM(tblBTC[BTC Aantal])",
    "sum_doge_total": "=SUM(tblDOGE[DOGE Aantal])",
    # Real estate totals
    "sum_vastgoed_waarde": "=SUM(tblVastgoed[Waarde(EUR)])",
    "sum_vastgoed_hypotheek": "=SUM(tblVastgoed[Hypotheek(EUR)])",
    "sum_vastgoed_netto": "=SUM(tblVastgoed[Netto Waarde])",
    # Loan totals
    "sum_leningen_openstaand": "=SUM(tblLeningen[Openstaand(EUR)])",
    # Cash totals
    "sum_cash": "=SUM(tblCash[Bedrag(EUR)])",
}

# =============================================================================
# STYLING CONSTANTS
# =============================================================================

COLORS = {
    "header_bg": "1F4E79",  # Dark blue
    "header_font": "FFFFFF",  # White
    "ok_bg": "C6EFCE",  # Light green
    "ok_font": "006100",  # Dark green
    "alert_bg": "FFEB9C",  # Light yellow/orange
    "alert_font": "9C5700",  # Dark orange
    "risk_bg": "FFC7CE",  # Light red
    "risk_font": "9C0006",  # Dark red
    "neutral_bg": "D9E1F2",  # Light blue
    "input_bg": "FFFFCC",  # Light yellow (for user input cells)
    "calc_bg": "E2EFDA",  # Light green (for calculated cells)
}

COLUMN_WIDTHS = {
    "date": 12,
    "currency": 15,
    "percentage": 12,
    "text_short": 15,
    "text_medium": 25,
    "text_long": 40,
    "number": 12,
    "status": 10,
}

# =============================================================================
# DATA VALIDATION OPTIONS
# =============================================================================

DATA_VALIDATIONS = {
    "holding_type": ["Spot", "Collateral"],
    "collateral_type": ["BTC", "Vastgoed"],
    "vastgoed_type": ["Eigen woning", "Verhuur", "Vakantiewoning", "Overig"],
    "cash_type": ["Betaalrekening", "Spaarrekening", "Belegging", "Crypto Exchange", "Overig"],
    "severity": ["INFO", "ALERT", "RISK", "CRITICAL"],
}
