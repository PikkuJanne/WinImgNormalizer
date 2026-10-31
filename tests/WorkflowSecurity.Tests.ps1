BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $workflow = [IO.File]::ReadAllText((Join-Path $repository '.github/workflows/windows-tests.yml')).Replace("`r`n", "`n")

    function Get-WorkflowFieldBlock {
        param([string]$Text, [string]$Key, [int]$Indent)
        $lines = $Text -split '\r?\n'
        $start = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match ('^ {' + $Indent + '}' + [regex]::Escape($Key) + ':')) {
                $start = $i
                break
            }
        }
        if ($start -lt 0) { return '' }
        $block = @($lines[$start])
        for ($i = $start + 1; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^\s*(#.*)?$') { continue }
            $currentIndent = [regex]::Match($lines[$i], '^ *').Length
            if ($currentIndent -le $Indent) { break }
            $block += $lines[$i]
        }
        return $block -join "`n"
    }

    function Get-WorkflowPolicyViolations {
        param([string]$Text)
        $violations = @()
        # This sole build-time cache acquisition may read the repository's draft
        # dependency assets. Its ephemeral contents-read token is never inherited
        # by application tests, saved by checkout, printed or uploaded.
        $acquisitionPattern = '(?m)^      - name: Authenticate build-only runtime dependency cache\n        env:\n          GH_TOKEN: \$\{\{ github.token \}\}\n        run: \|\n          \./tools/release/Stage-Runtime\.ps1 -Download\n          if \(\$env:WINIMG_CANDIDATE_TAG\) \{ \./tests/Get-HotfixCandidate\.ps1 -Tag \$env:WINIMG_CANDIDATE_TAG \}\n'
        $acquisition = @([regex]::Matches($Text, $acquisitionPattern))
        if ($acquisition.Count -ne 1) { $violations += 'Only the exact build-only cache step may receive the contents-read token.' }
        $credentialPolicyText = [regex]::Replace($Text, $acquisitionPattern, '')
        # This is a deliberately narrow policy audit of this workflow's plain
        # YAML form, not a general YAML parser. Aliases/merges must be reviewed
        # explicitly because they can hide a permission or credential override.
        if ($Text -match '(?m)^\s*[^#\r\n]*:\s*[&*][A-Za-z0-9_-]+' -or $Text -match '(?m)^\s*<<:') {
            $violations += 'YAML aliases or merged security settings.'
        }
        if ($Text.Contains("`t")) { $violations += 'Nonstandard YAML indentation.' }
        $triggers = Get-WorkflowFieldBlock -Text $Text -Key 'on' -Indent 0
        $triggerLines = @($triggers -split '\n' | Select-Object -Skip 1 | Where-Object { $_ -match '^  \S' })
        if ($triggerLines.Count -ne 3 -or @($triggerLines | Where-Object { $_ -notmatch '^  (push|pull_request|workflow_dispatch):\s*$' }).Count -gt 0 -or
            @($triggerLines | Select-Object -Unique).Count -ne 3) {
            $violations += 'Only push, pull_request and workflow_dispatch are allowed.'
        }
        $dispatch = Get-WorkflowFieldBlock -Text $triggers -Key 'workflow_dispatch' -Indent 2
        $expectedDispatch = "  workflow_dispatch:`n    inputs:`n      candidate_tag:`n        description: Exact release bytes to verify before or after publication`n        required: false`n        default: ''`n        type: string"
        if ($dispatch -cne $expectedDispatch) { $violations += 'Dispatch may accept only the validated unpublished candidate tag.' }
        $inputPolicyText = $Text.Replace('      WINIMG_CANDIDATE_TAG: ${{ inputs.candidate_tag }}', '')
        if ($inputPolicyText -match '(?i)\$\{\{\s*inputs(?:\.|\[)') { $violations += 'Dispatch input must enter scripts only through the validated environment value.' }
        $lines = $Text -split '\r?\n'
        if ($Text -notmatch '(?m)^permissions:\s*$') { $violations += 'Missing explicit root permissions.' }
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -notmatch '^(?<indent> *)permissions:\s*(?<scalar>[^#]*)') { continue }
            $indent = $Matches.indent.Length
            if (-not [string]::IsNullOrWhiteSpace($Matches.scalar)) {
                $violations += 'Permissions scalar or flow mapping is prohibited.'
                continue
            }
            $entries = @()
            for ($j = $i + 1; $j -lt $lines.Count; $j++) {
                if ($lines[$j] -match '^\s*(#.*)?$') { continue }
                if ([regex]::Match($lines[$j], '^ *').Length -le $indent) { break }
                $entries += ($lines[$j] -replace '\s+#.*$', '').Trim()
            }
            if ($entries.Count -ne 1 -or $entries[0] -cne 'contents: read') {
                $violations += 'Every permissions block must grant contents read only.'
            }
        }
        if ($credentialPolicyText -match '(?i)\$\{\{\s*(?:secrets(?:\.|\[)|github\.token)' -or $credentialPolicyText -match '(?im)^\s*(token|GH_TOKEN|GITHUB_TOKEN|ssh-key|allow-unsafe-pr-checkout):' -or
            $credentialPolicyText -match '(?im)^\s*secrets:') { $violations += 'Explicit secrets or checkout credential override.' }
        if ($Text -match '(?i)\$\{\{\s*github\.event(?:\.|\[)') {
            $violations += 'Contributor-controlled event data interpolated into workflow.'
        }
        if ($Text -match '(?im)^\s*(?:- )?continue-on-error:') {
            $violations += 'Failed checks can be ignored.'
        }
        # Exact accepted pins bind the reviewed official source, so an arbitrary
        # forty-character hexadecimal string is never treated as reviewed.
        $pins = @{
            'actions/checkout' = '11d5960a326750d5838078e36cf38b85af677262'
            'actions/upload-artifact' = '043fb46d1a93c77aae656e7c1c64a875d1fc6a0a'
        }
        $uses = @([regex]::Matches($Text, '(?m)^\s*(?:- )?uses:\s*(?<reference>[^\r\n]*)'))
        if ($uses.Count -ne 2) { $violations += 'Unexpected external action inventory.' }
        foreach ($use in $uses) {
            $reference = [regex]::Match($use.Groups['reference'].Value, '^(?<action>[^@\s]+)@(?<sha>[^\s#]+)(?<comment>.*)$')
            $action = $reference.Groups['action'].Value
            if (-not $reference.Success -or -not $pins.ContainsKey($action) -or $reference.Groups['sha'].Value -cne $pins[$action] -or
                $reference.Groups['comment'].Value -notmatch '# v\d+\.\d+\.\d+;') {
                $violations += 'External action does not match a reviewed full SHA and version comment.'
            }
        }
        $steps = @([regex]::Matches($Text, '(?ms)^ {6}- (?:name|uses):[^\n]*\n(?:(?!^ {6}- ).)*'))
        $checkout = @($steps | Where-Object { $_.Value -match 'uses: actions/checkout@' })
        if ($checkout.Count -ne 1 -or $checkout[0].Value -notmatch '(?m)^ {10}persist-credentials: false\s*$' -or
            $checkout[0].Value -match '(?m)^ {10}(ref|repository):') {
            $violations += 'Checkout must use event commit and remove credentials before tests.'
        }
        $upload = @($steps | Where-Object { $_.Value -match 'uses: actions/upload-artifact@' })
        if ($upload.Count -ne 1) { $violations += 'Missing unique artifact upload.' }
        else {
            $pathBlock = Get-WorkflowFieldBlock -Text $upload[0].Value -Key 'path' -Indent 10
            $paths = @($pathBlock -split '\n' | Select-Object -Skip 1 | ForEach-Object { $_.Trim() })
            $allowlist = @('.scratch/ci-export/test-evidence.json', '.scratch/ci-export/static-analysis-evidence.json', '.scratch/ci-export/hotfix-evidence.json')
            if ($paths.Count -ne 3 -or @($paths | Where-Object { $_ -cnotin $allowlist }).Count -gt 0 -or
                @($paths | Select-Object -Unique).Count -ne 3) {
                $violations += 'Artifact paths must be the exact sanitized JSON allowlist.'
            }
            if ($upload[0].Value -notmatch '(?m)^ {10}if-no-files-found: error\s*$' -or
                $upload[0].Value -notmatch '(?m)^ {10}overwrite: false\s*$' -or
                $upload[0].Value -notmatch '(?m)^ {10}retention-days: 14\s*$') {
                $violations += 'Artifact absence, overwrite or retention policy was weakened.'
            }
        }
        if ($Text -notmatch '(?m)^ {4}runs-on: windows-2025\s*$' -or
            $Text -notmatch '(?m)^ {8}shell: \[powershell, pwsh\]\s*$' -or
            $Text -notmatch '(?m)^ {8}shell: \$\{\{ matrix\.shell \}\}\s*$' -or
            $Text -notmatch '(?m)^ {6}WINIMG_CI_RUNNER_LABEL: windows-2025\s*$') {
            $violations += 'Both Windows shells and identified Server image are required.'
        }
        if ($Text -notmatch 'Invoke-Tests\.ps1 -DownloadDependencies -ResultDirectory \.scratch/ci-tests' -or
            $Text -match 'Invoke-Tests\.ps1[^\r\n]* -Path(?:\s|:)' -or
            $Text -notmatch 'Invoke-StaticAnalysis\.ps1 -DownloadDependencies -ResultDirectory \.scratch/ci-analysis' -or
            $Text -notmatch 'Invoke-Tests\.ps1 -DeliberateFailure' -or
            $Text -notmatch 'Invoke-StaticAnalysis\.ps1 -DeliberateFailure' -or
            $Text -notmatch 'Export-TestEvidence\.ps1[^\r\n]*-OutputDirectory \.scratch/ci-export' -or
            $Text -notmatch 'Invoke-HotfixAcceptance\.ps1 -Build -CandidateTag \$env:WINIMG_CANDIDATE_TAG -DownloadDependencies -DisposableRunner') {
            $violations += 'Mandatory corpus, static analysis, controls and sanitized export are required.'
        }
        $cachePosition = $Text.IndexOf('      - name: Authenticate build-only runtime dependency cache', [StringComparison]::Ordinal)
        $acceptancePosition = $Text.IndexOf('      - name: Test exact offline portable bytes as an ordinary disposable account', [StringComparison]::Ordinal)
        $mandatoryPosition = $Text.IndexOf('      - name: Run mandatory tests with verified development dependencies', [StringComparison]::Ordinal)
        if ($cachePosition -lt 0 -or $acceptancePosition -le $cachePosition -or $mandatoryPosition -le $acceptancePosition -or
            $Text -notmatch 'if \(-not \$acceptance\.Passed -or \$acceptance\.PSObject\.Properties\[''Error''\]\)' -or
            $Text -notmatch 'Copy-Item -LiteralPath \.scratch/ci-hotfix/acceptance-summary\.json -Destination \.scratch/ci-export/hotfix-evidence\.json') {
            $violations += 'Packaged acceptance must run early, preserve mandatory gates and export only passing sanitized evidence.'
        }
        if (@([regex]::Matches($Text, '\$controlExit -ne 1')).Count -ne 2 -or
            $Text -notmatch '\$control\.total_count -ne \(\$normal\.total_count \+ 1\)' -or
            $Text -notmatch '\$control\.failed_count -ne 1' -or
            $Text -notmatch 'T007 runner credibility control\.deliberate assertion must fail' -or
            $Text -notmatch '\$control\.scoped_findings_count -ne 0' -or
            $Text -notmatch '\$control\.control_findings_count -ne 1' -or
            $Text -notmatch '\$control\.findings_count -ne 1' -or
            $Text -notmatch 'control/unsafe-expression\.ps1' -or
            $Text -notmatch 'PSAvoidUsingInvokeExpression') {
            $violations += 'Both controls must require exit one and the exact intended failure.'
        }
        return $violations
    }

    function Assert-WorkflowMutationRejected {
        param([string]$Mutated, [string]$ExpectedViolation)
        $Mutated | Should -Not -BeExactly $workflow
        @(Get-WorkflowPolicyViolations $Mutated) | Should -Contain $ExpectedViolation
    }
}

Describe 'T066 workflow least privilege, reviewed actions and synthetic evidence' {
    It 'accepts the actual workflow with both full Windows hosts and real failure controls' {
        @(Get-WorkflowPolicyViolations $workflow).Count | Should -Be 0
    }

    It 'rejects root contents write permission' {
        Assert-WorkflowMutationRejected ($workflow.Replace('contents: read', 'contents: write')) 'Every permissions block must grant contents read only.'
    }

    It 'rejects an injected job-level write override even with a safe root block' {
        $mutation = $workflow.Replace("  test:`n", "  test:`n    permissions:`n      contents: write`n")
        Assert-WorkflowMutationRejected $mutation 'Every permissions block must grant contents read only.'
    }

    It 'rejects implicit repository default permissions' {
        $mutation = [regex]::Replace($workflow, '(?m)^permissions:\s*\r?\n  contents: read\s*\r?\n', '')
        Assert-WorkflowMutationRejected $mutation 'Missing explicit root permissions.'
    }

    It 'rejects broad permission shorthand' {
        Assert-WorkflowMutationRejected ($workflow.Replace("permissions:`n  contents: read", 'permissions: write-all')) 'Permissions scalar or flow mapping is prohibited.'
    }

    It 'rejects a privileged pull request target trigger' {
        Assert-WorkflowMutationRejected ($workflow.Replace('  pull_request:', '  pull_request_target:')) 'Only push, pull_request and workflow_dispatch are allowed.'
    }

    It 'rejects credentials left in checked-out contributor code' {
        Assert-WorkflowMutationRejected ($workflow.Replace('persist-credentials: false', 'persist-credentials: true')) 'Checkout must use event commit and remove credentials before tests.'
    }

    It 'rejects removal of the explicit credential cleanup setting' {
        Assert-WorkflowMutationRejected ($workflow.Replace('persist-credentials: false', '# credential setting removed')) 'Checkout must use event commit and remove credentials before tests.'
    }

    It 'rejects secrets supplied to checkout' {
        $mutation = $workflow.Replace('persist-credentials: false', 'persist-credentials: false' + "`n          token: " + '${{ secrets.DEPLOY_TOKEN }}')
        Assert-WorkflowMutationRejected $mutation 'Explicit secrets or checkout credential override.'
    }

    It 'rejects exposing the workflow token as an environment override' {
        $mutation = $workflow.Replace('      WINIMG_CI_RUNNER_LABEL: windows-2025', '      WINIMG_CI_RUNNER_LABEL: windows-2025' + "`n      GH_TOKEN: " + '${{ github.token }}')
        Assert-WorkflowMutationRejected $mutation 'Explicit secrets or checkout credential override.'
    }

    It 'rejects broadening the build-only runtime cache token scope to application tests' {
        $mutation = $workflow.Replace('Stage-Runtime.ps1 -Download', 'tests/Invoke-Tests.ps1 -DownloadDependencies')
        Assert-WorkflowMutationRejected $mutation 'Only the exact build-only cache step may receive the contents-read token.'
        @(Get-WorkflowPolicyViolations $mutation) | Should -Contain 'Explicit secrets or checkout credential override.'
    }

    It 'rejects dispatch expressions interpolated into executable shell text' {
        $mutation = $workflow.Replace('          ./tests/Invoke-Tests.ps1 -DownloadDependencies', '          Write-Host "${{ inputs.candidate_tag }}"' + "`n          ./tests/Invoke-Tests.ps1 -DownloadDependencies")
        Assert-WorkflowMutationRejected $mutation 'Dispatch input must enter scripts only through the validated environment value.'
    }

    It 'rejects a permission YAML alias hiding the effective grant' {
        Assert-WorkflowMutationRejected ($workflow.Replace("permissions:`n  contents: read", 'permissions: *trusted')) 'YAML aliases or merged security settings.'
    }

    It 'rejects a floating action version despite its version comment' {
        Assert-WorkflowMutationRejected ($workflow.Replace('checkout@11d5960a326750d5838078e36cf38b85af677262', 'checkout@v4')) 'External action does not match a reviewed full SHA and version comment.'
    }

    It 'rejects an invented forty-character action pin' {
        $mutation = $workflow.Replace('upload-artifact@043fb46d1a93c77aae656e7c1c64a875d1fc6a0a', 'upload-artifact@0000000000000000000000000000000000000000')
        Assert-WorkflowMutationRejected $mutation 'External action does not match a reviewed full SHA and version comment.'
    }

    It 'rejects a Docker image action without a reviewed repository pin' {
        $mutation = $workflow.Replace('      - name: Run mandatory tests', "      - uses: docker://unreviewed-image:latest`n      - name: Run mandatory tests")
        Assert-WorkflowMutationRejected $mutation 'External action does not match a reviewed full SHA and version comment.'
    }

    It 'rejects contributor event expressions inserted into a script' {
        $mutation = $workflow.Replace("          ./tests/Invoke-Tests.ps1 -DownloadDependencies", '          Write-Host "${{ github.event.pull_request.title }}"' + "`n          ./tests/Invoke-Tests.ps1 -DownloadDependencies")
        Assert-WorkflowMutationRejected $mutation 'Contributor-controlled event data interpolated into workflow.'
    }

    It 'rejects suppressing a failing check' {
        $mutation = $workflow.Replace('      - name: Run mandatory tests', "      - continue-on-error: true`n        name: Run mandatory tests")
        Assert-WorkflowMutationRejected $mutation 'Failed checks can be ignored.'
    }

    It 'rejects uploading raw scratch through a recursive wildcard' {
        Assert-WorkflowMutationRejected ($workflow.Replace('.scratch/ci-export/test-evidence.json', '.scratch/**')) 'Artifact paths must be the exact sanitized JSON allowlist.'
    }

    It 'rejects an additional raw test XML artifact' {
        $mutation = $workflow.Replace('.scratch/ci-export/test-evidence.json', ".scratch/ci-export/test-evidence.json`n            .scratch/ci-tests/results.xml")
        Assert-WorkflowMutationRejected $mutation 'Artifact paths must be the exact sanitized JSON allowlist.'
    }

    It 'rejects accepting absent artifacts as a warning' {
        Assert-WorkflowMutationRejected ($workflow.Replace('if-no-files-found: error', 'if-no-files-found: warn')) 'Artifact absence, overwrite or retention policy was weakened.'
    }

    It 'rejects dropping Windows PowerShell 5.1 from the matrix' {
        Assert-WorkflowMutationRejected ($workflow.Replace('[powershell, pwsh]', '[pwsh]')) 'Both Windows shells and identified Server image are required.'
    }

    It 'rejects a moving latest runner label' {
        Assert-WorkflowMutationRejected ($workflow.Replace('runs-on: windows-2025', 'runs-on: windows-latest')) 'Both Windows shells and identified Server image are required.'
    }

    It 'rejects a narrowed CI test path replacing the mandatory corpus' {
        $mutation = $workflow.Replace('Invoke-Tests.ps1 -DownloadDependencies', 'Invoke-Tests.ps1 -Path tests/Normalizer.Tests.ps1 -DownloadDependencies')
        Assert-WorkflowMutationRejected $mutation 'Mandatory corpus, static analysis, controls and sanitized export are required.'
    }

    It 'rejects a broken assertion control that accepts native exit zero' {
        Assert-WorkflowMutationRejected ($workflow.Replace('$controlExit -ne 1', '$controlExit -ne 0')) 'Both controls must require exit one and the exact intended failure.'
    }

    It 'rejects an analyzer control that accepts an unrelated production finding' {
        Assert-WorkflowMutationRejected ($workflow.Replace('$control.scoped_findings_count -ne 0', '$control.scoped_findings_count -lt 0')) 'Both controls must require exit one and the exact intended failure.'
    }
}
