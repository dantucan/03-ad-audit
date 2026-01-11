# 03-ad-audit
## Innehåll
- `data/ad_export.json` – indata (AD-export)
- `src/ad_audit.ps1` – PowerShell-script för audit
- `output/ad_audit_report.txt` – genererad rapport
- `output/inactive_users.csv` – användare som inte loggat in på X dagar (om ingår)
- `output/computer_status.csv` – sammanfattning per site/OS (om ingår)

## Körning
Kör från repo-roten:

```bash
pwsh -File ./src/ad_audit.ps1
