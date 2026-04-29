## Context

This project is part of my "Rafael Alba IT Lab", where I build practical tools and environments to improve real-world IT troubleshooting and infrastructure skills.

# Endpoint Support Checklist (PowerShell / WinForms)

A lightweight PowerShell WinForms utility for endpoint support workflows.

This tool inspects device status, records technical interventions, and maintains a local intervention history.

## Real-world usage

In endpoint support, many issues require quick validation of system state before deeper troubleshooting.

This tool helps standardize checks such as:
- BIOS and firmware state
- Secure Boot and TPM availability
- BitLocker configuration
- Device baseline information

This avoids:
- inconsistent manual checks
- missing critical information during incidents
- repeated diagnostic steps

## Troubleshooting mindset

The goal is not just to collect information, but to:

- quickly identify abnormal states
- compare against expected configurations
- document interventions consistently

In real environments, lack of structured checks often leads to:
- longer resolution times
- incomplete diagnostics
- repeated incidents

## Example scenario

A device fails to comply with security policies.

Initial symptoms:
- BitLocker not enforced
- TPM appears unavailable

Using this tool:
- TPM state is verified
- Secure Boot status is confirmed
- BIOS version is checked for compatibility issues

This allows faster identification of whether the issue is:
- configuration-related
- firmware-related
- or policy-related

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
