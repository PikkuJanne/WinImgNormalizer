# Early Windows checks (M0-T03)

These development tests cover the import boundary and the two existing positional
invocations. They do not certify the later collision, metadata, frame, cancellation
or output-validation fixes. The characterization evidence in `legacy/` describes
those known defects separately.

Run `Initialize-TestDependencies.ps1` explicitly to download hash-pinned Pester
and portable ImageMagick into ignored `.scratch`. It verifies the archives and
extracts private copies without an installer, permanent PATH change, or persistent
execution-policy change. Normal image processing never calls this bootstrap.

```powershell
# Run separately in Windows PowerShell 5.1 and PowerShell 7 on Windows.
.\tests\Initialize-TestDependencies.ps1 -Download
.\tests\Invoke-Tests.ps1
```

The runner records actual host/tool versions, counts and NUnit results in ignored
scratch. It fails for a failed assertion, failed discovery, zero discovered tests
or any skipped test. `-DeliberateFailure` adds a real failing assertion; its expected
nonzero exit is checked by the Windows CI workflow after the normal suite passes.
No passing result should be inferred from merely importing Pester.

The application exposes `Invoke-WinImgNormalizer` and
`Invoke-WinImgNormalizerCommand` when dot-sourced. Their internal `OutputParent`
and `ProcessRunner` parameters let tests use owned synthetic scratch trees. The
public `.ps1 <sourceFolder> [maxBytes]` interface and `.bat` remain unchanged.

Invocation tests use a disposable copy with only the outer call given an internal
scratch output parent, then run its real `-File` entry in the current Windows shell.
They check ordinary JPEG decoding/dimensions, video hashes, mirrored directories,
timestamps, source preservation and logs. This is instrumented invocation evidence;
it does not exercise real Pictures resolution or manual drag-and-drop. Tests never
run normalization against the real Pictures tree, private media or drive roots.

The CI matrix uses Windows Server runners. Desktop results are recorded separately
in `docs/codex-winimg/evidence/M0-T03.json`; Server CI does not prove desktop launcher
or interruption behavior. PSScriptAnalyzer and the complete corpus belong to later
tasks.
