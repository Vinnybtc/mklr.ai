# Financiële Cockpit

Een uitgebreide Excel-werkmap voor het beheren van je financiële overzicht, inclusief crypto holdings, vastgoed, leningen en automatische alerts.

## Features

- **Live Koersen**: Automatische prijsupdates via Excel Power Query
- **Holdings Tracking**: BTC, DOGE, vastgoed en cash
- **Lening Management**: LTV-berekeningen voor BTC-backed leningen
- **Automatische Alerts**: Rule-based inzichten op basis van jouw thresholds
- **Stress Testing**: Simuleer koersdalingen en bekijk de impact
- **Volledig Configureerbaar**: Alle drempels zijn aanpasbaar in de Instellingen sheet

## Installatie

### Vereisten

- Python 3.11+
- pip (Python package manager)

### Installatie stappen

```bash
# 1. Navigeer naar de project folder
cd financial-cockpit

# 2. Installeer dependencies
pip install openpyxl

# 3. Genereer de Excel werkmap
python src/generate_workbook.py
```

De werkmap wordt gegenereerd in `dist/Financiele_Cockpit.xlsx`

## Gebruik

### 1. Open de werkmap

Open `dist/Financiele_Cockpit.xlsx` in Microsoft Excel.

### 2. Configureer Power Query voor live koersen

Zie de **Koers-Setup** sheet voor gedetailleerde instructies:

1. Ga naar **Data > Get Data > From Web**
2. Voer de CoinGecko API URL in:
   ```
   https://api.coingecko.com/api/v3/simple/price?ids=bitcoin,dogecoin&vs_currencies=eur,usd
   ```
3. Transformeer de data en laad naar de Koersen sheet
4. Stel automatische refresh in (bijv. elke 10 minuten)

### 3. Pas thresholds aan

Open de **Instellingen** sheet en pas de parameters aan naar jouw situatie:

| Parameter | Default | Beschrijving |
|-----------|---------|--------------|
| RefreshMinutes | 10 | Koers refresh interval |
| TargetLTV_BTC | 0.30 | Doel-LTV voor BTC leningen (30%) |
| AlertBuffer_BTC | 0.10 | Alert trigger afstand tot liquidatie |
| RiskBuffer_BTC | 0.05 | Risk trigger afstand tot liquidatie |
| DefaultLiquidationLTV_BTC | 0.50 | Standaard liquidatie-LTV |
| CashBufferMonthsTarget | 6 | Minimale cashbuffer in maanden |
| MaxDebtServiceRatio | 0.35 | Maximale schuld-service ratio |
| StressBTCDown1 | -0.20 | Stress scenario 1: -20% |
| StressBTCDown2 | -0.40 | Stress scenario 2: -40% |
| ConcentrationAlertPct | 0.60 | Max crypto allocatie (60%) |

### 4. Vul je holdings in

#### Bitcoin Holdings sheet
- Datum van aankoop/transfer
- Wallet of Platform naam
- BTC Aantal (8 decimalen)
- Type: Spot of Collateral
- Eventuele opmerkingen

#### Dogecoin Holdings sheet
- Zelfde structuur als BTC

#### Vastgoed sheet
- Pand naam
- Adres
- Type (Eigen woning/Verhuur/etc.)
- Waarde, Hypotheek, Overige Leningen
- Netto Waarde wordt automatisch berekend

#### Leningen sheet
- LoanID (unieke identifier)
- OnderpandType: BTC of Vastgoed
- OnderpandRef: referentie naar wallet/platform of pand
- Rente, Looptijd, LiquidatieLTV

#### Cash & Overig sheet
- Rekeningen en beleggingen
- Type (Betaalrekening/Spaarrekening/Belegging/etc.)

### 5. Bekijk het Dashboard

Het **Dashboard** toont:
- Totale crypto waarde (BTC + DOGE)
- Vastgoed waarde en schulden
- Netto vermogen
- Top 3 Alerts
- Stress test resultaten
- Actuele koersen
- BTC lening LTV status

### 6. Monitor Alerts

De **Alerts** sheet toont actieve waarschuwingen:
- BTC LTV status (ALERT/RISK als te hoog)
- Cash buffer waarschuwingen
- Concentratie alerts (te veel crypto?)
- Stress test resultaten

## Sheet Overzicht

| Sheet | Doel |
|-------|------|
| Dashboard | Hoofdoverzicht met alle KPI's |
| Alerts | Actieve waarschuwingen en acties |
| Instellingen | Configuratie parameters |
| Koersen | Live prijzen (Power Query) |
| Koers-Setup | Instructies voor Power Query |
| Bitcoin Holdings | BTC portfolio |
| Dogecoin Holdings | DOGE portfolio |
| Vastgoed | Onroerend goed |
| Leningen | Alle leningen |
| BTC als Onderpand | LTV monitor voor BTC leningen |
| Cash & Overig | Liquide middelen |

## Alerts Logica

### BTC LTV Alerts
- **OK**: DistanceToLiq > AlertBuffer
- **ALERT**: AlertBuffer >= DistanceToLiq > RiskBuffer
- **RISK**: DistanceToLiq <= RiskBuffer

Bij ALERT/RISK wordt berekend:
- Hoeveel EUR aflossen om naar TargetLTV te komen
- Hoeveel BTC bijstorten om naar TargetLTV te komen

### Cash Buffer Alert
Getriggerd wanneer: `Cash / Maandlasten < CashBufferMonthsTarget`

### Concentratie Alert
Getriggerd wanneer: `Crypto Waarde / Totaal Vermogen > ConcentrationAlertPct`

### Stress Test Alerts
Waarschuwt als bij een gesimuleerde koersdaling:
- LTV boven liquidatiegrens komt (LIQUIDATION)
- LTV in RISK zone komt

## Nieuwe Asset Toevoegen

Om een nieuwe cryptocurrency toe te voegen:

### 1. Update model_config.py

```python
CRYPTO_ASSETS = [
    # Bestaande assets...
    {
        "symbol": "ETH",
        "name": "Ethereum",
        "coingecko_id": "ethereum",
        "decimals": 8,
        "sheet_name": "Ethereum Holdings",
        "table_name": "tblETH",
    },
]
```

### 2. Update SHEETS_CONFIG

```python
{"name": "Ethereum Holdings", "order": 7, "type": "holdings"},
```

### 3. Update TABLE_DEFINITIONS

```python
"tblETH": {
    "sheet": "Ethereum Holdings",
    "columns": ["Datum", "Wallet/Platform", "ETH Aantal", "Type", "Opmerking"],
    "start_row": 2,
    "start_col": 1,
},
```

### 4. Voeg sample data toe (optioneel)

### 5. Creëer de sheet builder functie in generate_workbook.py

### 6. Update Dashboard formules

### 7. Regenereer de werkmap

```bash
python src/generate_workbook.py
```

## Power Query URL's

### CoinGecko (gratis, rate-limited)
```
https://api.coingecko.com/api/v3/simple/price?ids=bitcoin,dogecoin&vs_currencies=eur,usd
```

### CoinCap (alternatief)
```
https://api.coincap.io/v2/assets?ids=bitcoin,dogecoin
```

## Troubleshooting

### Koersen updaten niet
1. Controleer internetverbinding
2. Controleer Power Query instellingen
3. CoinGecko heeft rate limits - verhoog RefreshMinutes
4. Voer handmatig koersen in als backup

### Formules tonen #REF! errors
1. Controleer of alle tabellen correct zijn aangemaakt
2. Controleer tabel namen (tblBTC, tblKoersen, etc.)
3. Zorg dat Instellingen sheet alle parameters bevat

### Alerts worden niet getoond
1. Controleer of data is ingevuld in holdings sheets
2. Controleer of koersen zijn ingevuld
3. Alerts tonen alleen wanneer thresholds overschreden worden

## Bestandsstructuur

```
financial-cockpit/
├── README.md
├── src/
│   ├── model_config.py      # Asset en sheet configuraties
│   ├── excel_builder.py     # Excel styling en utilities
│   └── generate_workbook.py # Hoofdscript
└── dist/
    └── Financiele_Cockpit.xlsx  # Gegenereerde werkmap
```

## Licentie

Dit project is bedoeld voor persoonlijk gebruik. Gebruik op eigen risico.

## Disclaimer

Deze tool is alleen bedoeld voor informatieve doeleinden en vormt geen financieel advies. Raadpleeg altijd een gekwalificeerde financieel adviseur voor belangrijke financiële beslissingen.
