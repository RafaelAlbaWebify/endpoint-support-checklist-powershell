# Endpoint Support Checklist

> Windows endpoint evidence and support-handover utility.

Endpoint Support Checklist is a lightweight PowerShell WinForms tool for repeatable endpoint checks, local intervention notes and ticket-ready evidence exports.

It is a supporting utility for IAM, security-hygiene and infrastructure investigations rather than a replacement for endpoint-management platforms.

## Problem it addresses

Endpoint support repeatedly begins with the same questions:

- Is TPM available and ready?
- Is Secure Boot enabled?
- Is BitLocker active?
- Which BIOS version is installed?
- What was already checked or changed?
- Is there enough evidence for a clear escalation?

The tool creates a consistent local workflow around those checks so important details are easier to review, record and hand over.

## Workflow

```text
run endpoint check
  -> review TPM / Secure Boot / BitLocker / BIOS signals
  -> add intervention notes
  -> export JSON / TXT / CSV evidence
  -> attach the result to a ticket or support handover
```

## Checks and outputs

- Device status inspection
- BIOS version and date
- Secure Boot status
- TPM state
- BitLocker status
- Local intervention registration
- CSV intervention history
- JSON and TXT status export

## Run

From the repository folder:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\EndpointSupportChecklist.ps1
```

The execution-policy change applies only to the current PowerShell process.

Suggested use:

1. Launch the tool on the endpoint.
2. Review TPM, Secure Boot, BitLocker, BIOS and device information.
3. Export the status when evidence is needed.
4. Record the intervention with clear notes.
5. Attach the output to the ticket or handover.
6. Update the runbook if the issue becomes recurring.

See [`examples/sample-output.md`](examples/sample-output.md) for a public-safe example.

## Interpretation notes

- An unavailable or failed check does not automatically confirm root cause.
- Results depend on hardware, firmware, Windows edition, policy state and permissions.
- TPM, Secure Boot and BitLocker findings must be interpreted against the device model and organization policy.
- The tool improves consistency and evidence quality; it does not replace Intune, Group Policy, endpoint security tools or change control.

## Safety boundary

The utility does not:

- enforce BitLocker;
- change Secure Boot settings;
- modify TPM configuration;
- change BIOS settings;
- remediate compliance issues automatically.

## Portfolio value

This project demonstrates practical PowerShell GUI work, Windows endpoint support, evidence collection, local logging and structured handover documentation.

## License

MIT

## Author

Rafael Alba
