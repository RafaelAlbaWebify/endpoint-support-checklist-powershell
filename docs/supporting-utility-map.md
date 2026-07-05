# Supporting Utility Map

Endpoint Support Checklist is not one of my six main flagships.

It is a supporting utility that provides endpoint-readiness evidence for several flagship areas.

## Related Flagships

| Related area | Related flagship | How this utility supports it |
|---|---|---|
| IAM | TRACE | Provides endpoint state evidence that may help explain access and device-compliance issues. |
| SOC | CustosOps | Provides defensive endpoint hygiene evidence such as TPM, Secure Boot and BitLocker state. |
| IPPO | OPSCORE | Provides local device state evidence that may be relevant in wider infrastructure or production-support investigations. |

## Evidence Produced

The tool can help capture:

- TPM readiness
- Secure Boot status
- BitLocker status
- BIOS version and date
- local support notes
- intervention history
- JSON / TXT / CSV exportable evidence

## Example Use

```text
User cannot access a Microsoft 365 resource
  -> access appears blocked by policy or device state
  -> endpoint state is checked locally
  -> TPM / Secure Boot / BitLocker evidence is exported
  -> evidence is attached to support ticket
  -> identity, endpoint or security team can validate next step
```

## Boundaries

This utility should remain evidence-first.

It should not:

- enforce BitLocker
- modify TPM
- modify Secure Boot
- change BIOS settings
- remediate compliance issues automatically
- bypass endpoint/security/change-control processes
