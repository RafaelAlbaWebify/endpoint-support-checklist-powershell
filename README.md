# Endpoint Support Checklist

A lightweight PowerShell WinForms utility for repeatable Windows endpoint support checks, local intervention notes, and exportable evidence.

This project is part of my Rafael Alba IT Lab: practical IT Operations tools focused on endpoint troubleshooting, evidence collection, documentation, and repeatable support workflows.

## Current Status

Working public portfolio version.

The tool is designed for local endpoint support scenarios where a technician needs a quick, structured way to inspect common device-security and support-readiness signals before escalation or handover.

## Problem It Solves

Endpoint support often starts with the same questions:

- Is TPM available and ready?
- Is Secure Boot enabled?
- Is BitLocker active?
- What BIOS version is installed?
- What was already checked or changed on this device?
- Is there enough evidence to escalate the issue clearly?

When these checks are done manually, important details can be missed or recorded inconsistently. This tool creates a small structured workflow around those checks so endpoint support work becomes easier to repeat, document, and hand over.

## Checks Included

- Device status inspection
- BIOS version and BIOS date
- Secure Boot status
- TPM state detection
- BitLocker status
- Local intervention registration
- Local history log in CSV format
- JSON/TXT status export

## Example Support Scenario

A device fails to comply with security or access policies.

Initial symptoms:

- BitLocker is not enforced.
- TPM appears unavailable or not ready.
- Secure Boot status is unclear.
- The user needs a fast answer before escalation.

Using this tool, the support technician can verify TPM state, Secure Boot status, BitLocker status, and BIOS information, then record the intervention in the local history. That makes it easier to decide whether the issue is configuration-related, firmware-related, policy-related, or needs escalation to endpoint/security administration.

## Tech Stack

- PowerShell
- WinForms
- CIM / WMI
- BitLocker, TPM, and Secure Boot queries
- JSON, CSV, and TXT persistence

## How To Run

Open PowerShell and run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\EndpointSupportChecklist.ps1
```

## Suggested Workflow

1. Launch the tool on the endpoint.
2. Review TPM, Secure Boot, BitLocker, BIOS, and device information.
3. Export the status if evidence is needed for escalation.
4. Register the intervention with clear notes.
5. Attach the output to the ticket or support handover.
6. Update the support notes or runbook if the issue becomes recurring.

## Sample Output

See [examples/sample-output.md](examples/sample-output.md) for an example of how the exported status and intervention notes can be interpreted.

## Troubleshooting Notes

- A failed or unavailable check does not automatically confirm root cause.
- Some values depend on hardware, firmware, Windows edition, policy state, and execution permissions.
- BitLocker, TPM, and Secure Boot findings should be interpreted in the context of the device model and organization policy.
- This tool is intended to improve evidence collection and consistency, not to replace endpoint management platforms.

## Safety and Scope

This utility is intended for support evidence collection and local documentation.

It does not:

- enforce BitLocker
- change Secure Boot settings
- modify TPM configuration
- change BIOS settings
- remediate compliance issues automatically
- replace Intune, Group Policy, endpoint security platforms, or change-control procedures

## Homelab / FactoryOps Use

This project fits the planned FactoryOps homelab as an endpoint support and device-readiness scenario tool. It can support realistic practice cases such as:

- Windows client fails a security baseline check
- BitLocker status is unclear before escalation
- TPM or Secure Boot evidence is needed for Conditional Access / compliance investigation
- endpoint support handover requires structured evidence
- recurring endpoint issue needs a simple local intervention history

A future lab scenario should include the simulated device issue, exported evidence, support ticket notes, escalation summary, and a final runbook entry.

## Portfolio Value

This project demonstrates practical PowerShell GUI work, Windows endpoint support checks, evidence collection, local logging, ticket-ready exports, and support documentation habits.

It supports my public positioning around IT Operations, Systems Support, endpoint troubleshooting, Microsoft 365 / Entra ID support context, infrastructure support, and practical automation.

## Next Improvements

- Add screenshots of the WinForms interface.
- Add a short demo walkthrough using sanitized sample output.
- Add a FactoryOps endpoint compliance scenario.
- Add example ticket notes and escalation notes.
- Add a validation checklist for interpreting TPM, Secure Boot, and BitLocker findings safely.

## License

MIT

## Author

Rafael Alba
