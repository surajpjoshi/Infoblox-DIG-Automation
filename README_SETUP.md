# Infoblox DIG Automation - New Computer Setup

1. Copy the complete project folder to the new Windows PC.
2. Open PowerShell in the project folder.
3. Run:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\setup.ps1
```

The script checks Python, creates `.venv`, installs `customtkinter`, `paramiko`, and `openpyxl`, checks PuTTYgen, verifies imports, and optionally launches `app.py`.

For PPK authentication, install PuTTY so `puttygen.exe` is available. Password and supported PEM authentication do not require PuTTYgen.

Manual launch after setup:

```powershell
.\.venv\Scripts\python.exe .\app.py
```
