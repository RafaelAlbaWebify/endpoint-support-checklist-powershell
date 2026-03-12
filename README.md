# Endpoint Support Checklist (PowerShell / WinForms)

A lightweight PowerShell WinForms utility for endpoint support workflows.

This tool inspects device status, records technical interventions, and maintains a local intervention history.

## Features

- Device status inspection
- BIOS version and BIOS date
- Secure Boot check
- TPM state detection
- BitLocker status
- Local intervention registration
- Local history log (CSV)
- JSON/TXT status export

## Why I built it

This project was created as a small internal-style support utility to standardize device checks and technical intervention logging.

The focus was on creating a simple, readable, and maintainable tool usable in real support workflows.

## Tech

- PowerShell
- WinForms
- CIM / WMI
- BitLocker / TPM / Secure Boot queries
- JSON / CSV persistence

## How to run

Open PowerShell and run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\EndpointSupportChecklist.ps1
