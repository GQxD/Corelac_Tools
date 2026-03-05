# TOOL - CORELAC Utilities

Python repository for CORELAC plate management, plate plan generation/verification, and Excel/Nextcloud utilities.

## Requirements
- Recommended Python version: `3.12` (project baseline).

## Installation

```bash
python -m venv .venv
# Windows PowerShell
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Quick Start (what to run)
- Main application (plate management): `python src/app/gestionnaire_de_plaques_fast.py`
- Plate plan generation: `python src/plans/generer_plan_plaque_sequentielle_femelle.py`
- Plan verification and reports: `python src/plans/verifier_plan_sequentielle_femelle.py`
- Nextcloud QR generation (popup/env auth): `python src/nextcloud_qr/automat_link.py`

## Main Scripts
- `src/app/gestionnaire_de_plaques_fast.py`: main plate management app.
- `src/plans/generer_plan_plaque_sequentielle_femelle.py`: generates plate Excel files.
- `src/plans/verifier_plan_sequentielle_femelle.py`: verification and distribution reports.
- `src/traitement_excel/*.py`: batch plate-file edits and Word printing.
- `src/nextcloud_qr/*.py`: link/QR generation and Nextcloud automation.

## Configuration
Most scripts use absolute paths near the top of each file. Update `CONFIG` blocks before execution.

For `src/nextcloud_qr/automat_link.py`, set at least:

```powershell
$env:NEXTCLOUD_USERNAME="your_login"
$env:NEXTCLOUD_PASSWORD="your_token_or_password"
```

## Nextcloud QR Help (`automat_link.py`)
Recommended workflow to generate QR codes with temporary access:

1. In Nextcloud, go to `Settings` > `Security`.
2. Create an `app password` token, for example `qr_generation_temp`.
3. Run the script, then complete the popup fields:

```powershell
$env:NEXTCLOUD_USERNAME="your_login"
$env:NEXTCLOUD_PASSWORD="your_app_token"
python "src/nextcloud_qr/automat_link.py"
```

4. Verify that QR files and CSV are generated in the configured output folder.
5. Revoke the app token in `Settings` > `Security` once generation is complete.

Important: do not store tokens in source code.

## Development Support
This repository benefited from AI-assisted support for syntax optimization and documentation.

## License
This project is licensed under the GNU General Public License v3.0 (GPL-3.0).
See `LICENSE`.
