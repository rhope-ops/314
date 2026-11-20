# Code standard: what 33-332.Z adds

Decided by the owner, 2026-11-20 (docs/DECISIONS.md): DoD/M365 code, so the strictest settings the kit offers, the same as rhope-ops/RADE and rhope-ops/90-161.Z. A person must be able to read any script and know what it does before they run it on a gov device.

- Core rules and the check: [docs/code-standard.md](code-standard.md) (a kit file). Extras: [the RADE extras](https://github.com/rhope-ops/RADE/blob/main/docs/code-standard-rade.md) (X1 run target in the header, X2 Windows PowerShell 5.1, X4 decisions name who and when, X5 named values, X7 a discarded error says why).
- Settings: the `code-` lines of [.check](../.check): `code-c1-safety: on` (Network:, Writes:, Credentials: each start a header line with a value), `code-c2: full` (full help for each function longer than 10 lines), `code-max-line-chars: 200`, `code-generated: paste/*.ps1` (the paste bundles; their sources get the rules).
- Gate: [scripts/code_standard_gate.py](../scripts/code_standard_gate.py), a copy of 90-161.Z's (self-test `--test`, run by `scripts/check.sh`). The gate's rules come from RADE: a fix lands in RADE first, then in 90-161.Z and here.
- PowerShell runs on Windows PowerShell 5.1 and PnP.PowerShell 1.12.0 only: `tests/Lint.ps1` checks the syntax; `.github/workflows/ps51.yml` runs the tests on Windows PowerShell 5.1.
