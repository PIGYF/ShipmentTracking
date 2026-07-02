# Shipment Tracking

Excel-based shipment tracking prototype.

Supported now:

- DSV public tracking
- DGF tracking via DHL Shipment Tracking - Unified API
- Maersk provider template

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Put local credentials in `.env`. Do not commit `.env`.

## Run

DSV:

```powershell
python -m shipment_tracking.cli --dsv RTM0344100 --output exports\dsv_one.json
```

DGF:

```powershell
python -m shipment_tracking.cli --dgf FRAA47253 --output exports\dgf_one.json
```

DGF from Excel and write an updated copy:

```powershell
python -m shipment_tracking.cli --excel "D:\Downloads\2026-Import tracking list (1).xlsx" --sheet 2026 --carrier dgf --output exports\dgf_results.json --update-excel "exports\2026-Import tracking list.updated.xlsx"
```

Maersk template:

```powershell
python -m shipment_tracking.cli --maersk 261238481 --output exports\maersk_one.json
```

Current DGF key is only for rows where `货代 = DGF`, not `DHL EXPRESS`.
