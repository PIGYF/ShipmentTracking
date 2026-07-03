# Shipment Tracking

Excel shipment tracking refresh tool.

Currently supported:

- DSV
- DGF
- MAERSK

The current DGF key only applies to rows where `货代 = DGF`, not `DHL EXPRESS`.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Put API credentials in local `.env`. Do not commit `.env`.

## Run

If `python` is not available in PowerShell, use the included script:

```powershell
.\run_refresh.ps1 "D:\Downloads\2026-Import tracking list.xlsx"
```

Or use the batch file:

```bat
run_refresh.bat "D:\Downloads\2026-Import tracking list.xlsx"
```

Refresh pending supported shipments and save an updated Excel copy:

```powershell
python -m shipment_tracking.refresh_excel "D:\Downloads\2026-Import tracking list.xlsx"
```

Test without calling APIs or writing Excel:

```powershell
python -m shipment_tracking.refresh_excel "D:\Downloads\2026-Import tracking list.xlsx" --dry-run
```

Test a small batch:

```powershell
python -m shipment_tracking.refresh_excel "D:\Downloads\2026-Import tracking list.xlsx" --limit 10
```

By default it reads sheet `2026`, processes only `状态显示 = 未送货`, skips unsupported forwarders, and writes to:

```text
exports\2026-Import tracking list.updated.xlsx
```
