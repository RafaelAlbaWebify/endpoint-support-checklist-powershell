# Endpoint Support Checklist

A lightweight PowerShell WinForms utility for endpoint support workflows.

This tool inspects device status, records technical interventions and maintains a local intervention history so common endpoint checks become more consistent and repeatable.

## Problem It Solves

Endpoint support often starts with the same questions:

- Is TPM available and ready?
- Is Secure Boot enabled?
- Is BitLocker active?
- What BIOS version is installed?
- What was already checked or changed on this device?

When these checks are done manually, important details can be missed or recorded inconsistently. This tool creates a small structured workflow around those checks.

## Checks Included

- Device status inspection
- BIOS version and BIOS date
- Secure Boot status
- TPM state detection
- BitLocker status
- Local intervention registration
- Local history log in CSV format
- JSON/TXT status export

## Example Scenario

A device fails to comply with security policies.

Initial symptoms:

- BitLocker is not enforced.
- TPM appears unavailable.
- The user needs a fast answer before escalation.

Using this tool, the support technician can verify TPM state, Secure Boot status and BIOS version, then record the intervention in the local history. That makes it easier to decide whether the issue is configuration-related, firmware-related or policy-related.

## Tech Stack

- PowerShell
- WinForms
- CIM / WMI
- BitLocker, TPM and Secure Boot queries
- JSON, CSV and TXT persistence

## How To Run

Open PowerShell and run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\EndpointSupportChecklist.ps1
```

## Suggested Workflow

1. Launch the tool on the endpoint.
2. Review TPM, Secure Boot, BitLocker and BIOS information.
3. Export the status if evidence is needed for escalation.
4. Register the intervention with clear notes.
5. Attach the output to the ticket or support handover.

## Sample Output

See [examples/sample-output.md](examples/sample-output.md) for an example of how the exported status and intervention notes can be interpreted.

## Portfolio Value

This project demonstrates practical PowerShell GUI work, Windows endpoint support checks, evidence collection, local logging and support documentation habits.

## License

MIT

## Author

Rafael Alba
