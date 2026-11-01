BeforeAll {
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $workflow = [IO.File]::ReadAllText((Join-Path $repository '.github/workflows/windows-tests.yml')).Replace("`r`n", "`n")
    $releaseWorkflow = [IO.File]::ReadAllText((Join-Path $repository '.github/workflows/release-verification.yml')).Replace("`r`n", "`n")

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
        $Text = $Text.Replace("`r`n", "`n")
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
            $allowlist = @('.scratch/ci-export/test-evidence.json', '.scratch/ci-export/static-analysis-evidence.json', '.scratch/ci-export/hotfix-evidence.json', '.scratch/ci-export/distribution-evidence.json')
            if ($paths.Count -ne 4 -or @($paths | Where-Object { $_ -cnotin $allowlist }).Count -gt 0 -or
                @($paths | Select-Object -Unique).Count -ne 4) {
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
            $Text -notmatch 'Invoke-HotfixAcceptance\.ps1 @acceptanceArguments') {
            $violations += 'Mandatory corpus, static analysis, controls and sanitized export are required.'
        }
        $expectedAcceptance = @'
      - name: Test exact offline portable bytes as an ordinary disposable account
        run: |
          $acceptanceArguments = @{ DownloadDependencies=$true; DisposableRunner=$true; ResultDirectory='.scratch/ci-hotfix' }
          if ($env:WINIMG_CANDIDATE_TAG) { $acceptanceArguments.CandidateTag=$env:WINIMG_CANDIDATE_TAG }
          else {
            $canonical = [IO.File]::ReadAllText((Join-Path (Get-Location) 'release-metadata.json')) | ConvertFrom-Json
            if ($canonical.release_state -eq 'published') { $acceptanceArguments.PublicTag=$canonical.tag }
            else { $acceptanceArguments.Build=$true }
          }
          ./tests/Invoke-HotfixAcceptance.ps1 @acceptanceArguments
'@
        $expectedAcceptance = $expectedAcceptance.Replace("`r`n", "`n")
        if (@([regex]::Matches($Text, [regex]::Escape($expectedAcceptance))).Count -ne 1) {
            $violations += 'Packaged acceptance must prefer explicit candidate, then published PublicTag, then unreleased Build.'
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

Describe 'Exact release verification workflow security' {
    BeforeAll {
        function Get-ExactReleaseViolations([string]$Text) {
            $Text=$Text.Replace("`r`n","`n");$violations=@()
            $triggers=Get-WorkflowFieldBlock -Text $Text -Key 'on' -Indent 0
            if(@($triggers -split '\n'|Where-Object {$_ -match '^  \S'}).Count -ne 1 -or $triggers -notmatch '(?m)^  workflow_dispatch:\s*$' -or $triggers -notmatch '(?m)^        required: true\s*$' -or $triggers -notmatch '(?m)^      candidate_tag:\s*$'){$violations+='Required exact-tag dispatch only.'}
            $permissions=Get-WorkflowFieldBlock -Text $Text -Key 'permissions' -Indent 0
            if($permissions -cne "permissions:`n  contents: read" -or [regex]::Matches($Text,'(?m)^\s*permissions:').Count -ne 1){$violations+='Contents-read permission only.'}
            if($Text -notmatch 'runs-on: windows-2025' -or $Text -notmatch 'shell: \[powershell, pwsh\]'){$violations+='Both identified Windows shells required.'}
            $uses=@([regex]::Matches($Text,'(?m)^\s*(?:- )?uses:\s*(?<reference>[^\r\n]*)'))
            if($uses.Count -ne 2 -or $Text -notmatch 'uses: actions/checkout@11d5960a326750d5838078e36cf38b85af677262 # v4\.4\.0;' -or $Text -notmatch 'uses: actions/upload-artifact@043fb46d1a93c77aae656e7c1c64a875d1fc6a0a # v7\.0\.1;'){$violations+='Reviewed action pins only.'}
            if($Text -notmatch '(?m)^          ref: \$\{\{ inputs\.candidate_tag \}\}\s*$' -or $Text -notmatch '(?m)^          persist-credentials: false\s*$' -or $Text -notmatch '(?m)^          fetch-depth: 0\s*$'){$violations+='Exact tag checkout without saved credentials required.'}
            if($Text -notmatch [regex]::Escape("if (`$env:WINIMG_CANDIDATE_TAG -notmatch '^v\d+\.\d+\.\d+$')")){$violations+='Stable tag validation required.'}
            $acquisition='      - name: Authenticate existing exact release artifacts'+"`n"+'        env:'+"`n"+'          GH_TOKEN: ${{ github.token }}'+"`n"+'        run: |'+"`n"+'          ./tests/Get-HotfixCandidate.ps1 -Tag $env:WINIMG_CANDIDATE_TAG'+"`n"
            $withoutAcquisition=$Text.Replace($acquisition,'')
            if($withoutAcquisition -ceq $Text -or $withoutAcquisition -match 'GH_TOKEN|GITHUB_TOKEN|github\.token|secrets\.|persist-credentials: true'){$violations+='Read token restricted to artifact acquisition.'}
            if($Text -match '(?i)-Build\b|Invoke-Tests\.ps1|Invoke-StaticAnalysis\.ps1|Stage-Runtime\.ps1' -or $Text -notmatch 'Invoke-HotfixAcceptance\.ps1 -CandidateTag \$env:WINIMG_CANDIDATE_TAG -DownloadDependencies -DisposableRunner' -or $Text -notmatch 'Invoke-DistributionAcceptance\.ps1 @arguments' -or $Text -notmatch 'ArchivePath = Join-Path \.scratch/ci-candidate \$candidate\.AssetFilename'){$violations+='Verify existing exact bytes without rebuilding.'}
            if($Text -notmatch 'if \(\$candidate\.Stable -and \$candidate\.Latest\) \{ \$arguments\.PublicTag = \$candidate\.Tag \}'){$violations+='Public source archive verification required.'}
            $pathBlock=Get-WorkflowFieldBlock -Text $Text -Key 'path' -Indent 10
            $paths=@($pathBlock -split '\n'|Select-Object -Skip 1|ForEach-Object {$_.Trim()})
            $allowed=@('.scratch/exact-release-export/hotfix-evidence.json','.scratch/exact-release-export/distribution-evidence.json')
            if($paths.Count -ne 2 -or @($paths|Where-Object {$_ -cnotin $allowed}).Count -or @($paths|Select-Object -Unique).Count -ne 2 -or $Text -notmatch "-not \`$hotfix\.Passed" -or $Text -notmatch "-not \`$distribution\.Passed" -or $Text -notmatch 'if-no-files-found: error' -or $Text -notmatch 'overwrite: false'){$violations+='Only passing sanitized evidence allowlist may be uploaded.'}
            if($Text -match '(?im)continue-on-error:|pull_request_target:|workflow_run:'){$violations+='No bypass triggers or ignored failures.'}
            $violations
        }
    }
    It 'keeps exact-byte publication verification separate from the full merge gate' {
        @(Get-ExactReleaseViolations $releaseWorkflow).Count | Should -Be 0
    }
    It 'requires an explicit tag rather than a default or optional dispatch' {
        @(Get-ExactReleaseViolations ($releaseWorkflow.Replace('required: true','required: false'))) | Should -Contain 'Required exact-tag dispatch only.'
    }
    It 'rejects write permission during release verification' {
        @(Get-ExactReleaseViolations ($releaseWorkflow.Replace('contents: read','contents: write'))) | Should -Contain 'Contents-read permission only.'
    }
    It 'rejects moving action references' {
        @(Get-ExactReleaseViolations ($releaseWorkflow.Replace('actions/checkout@11d5960a326750d5838078e36cf38b85af677262','actions/checkout@v4'))) | Should -Contain 'Reviewed action pins only.'
    }
    It 'requires both actual Windows hosts' {
        @(Get-ExactReleaseViolations ($releaseWorkflow.Replace('[powershell, pwsh]','[pwsh]'))) | Should -Contain 'Both identified Windows shells required.'
    }
    It 'rejects rebuilding bytes while claiming to verify a candidate' {
        @(Get-ExactReleaseViolations ($releaseWorkflow.Replace('-CandidateTag $env:WINIMG_CANDIDATE_TAG','-Build'))) | Should -Contain 'Verify existing exact bytes without rebuilding.'
    }
    It 'rejects raw scratch artifacts' {
        @(Get-ExactReleaseViolations ($releaseWorkflow.Replace('.scratch/exact-release-export/hotfix-evidence.json','.scratch/**'))) | Should -Contain 'Only passing sanitized evidence allowlist may be uploaded.'
    }
    It 'rejects persisting credentials in the exact tag checkout' {
        @(Get-ExactReleaseViolations ($releaseWorkflow.Replace('persist-credentials: false','persist-credentials: true'))) | Should -Contain 'Exact tag checkout without saved credentials required.'
    }
}

Describe 'T066 workflow least privilege, reviewed actions and synthetic evidence' {
    It 'accepts the actual workflow with both full Windows hosts and real failure controls' {
        @(Get-WorkflowPolicyViolations $workflow).Count | Should -Be 0
    }

    It 'accepts the same reviewed workflow with Windows line endings' {
        @(Get-WorkflowPolicyViolations ($workflow.Replace("`n", "`r`n"))).Count | Should -Be 0
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

    It 'rejects <Mode> routing that could rebuild published bytes or ignore an explicit candidate' -ForEach @(
        @{Mode='ignored explicit candidate';Before='if ($env:WINIMG_CANDIDATE_TAG) { $acceptanceArguments.CandidateTag=$env:WINIMG_CANDIDATE_TAG }';After='if ($false) { $acceptanceArguments.CandidateTag=$env:WINIMG_CANDIDATE_TAG }'},
        @{Mode='published rebuild';Before='if ($canonical.release_state -eq ''published'') { $acceptanceArguments.PublicTag=$canonical.tag }';After='if ($canonical.release_state -eq ''published'') { $acceptanceArguments.Build=$true }'},
        @{Mode='unreleased public fallback';Before='else { $acceptanceArguments.Build=$true }';After='else { $acceptanceArguments.PublicTag=''v1.0.0'' }'}
    ) {
        Assert-WorkflowMutationRejected ($workflow.Replace($Before,$After)) 'Packaged acceptance must prefer explicit candidate, then published PublicTag, then unreleased Build.'
    }

    It 'executes the actual workflow argument selection for <Mode> without processing or network access' -ForEach @(
        @{Mode='candidate while unreleased';State='unreleased';Candidate='v1.0.1';Expected='CandidateTag'},
        @{Mode='candidate after publication';State='published';Candidate='v1.0.1';Expected='CandidateTag'},
        @{Mode='canonical public download';State='published';Candidate='';Expected='PublicTag'},
        @{Mode='unreleased build';State='unreleased';Candidate='';Expected='Build'}
    ) {
        $route = [regex]::Match($workflow, '(?m)^      - name: Test exact offline portable bytes as an ordinary disposable account\n        run: \|\n(?<code>(?:^ {10}[^\r\n]*\n)+)')
        $route.Success | Should -BeTrue
        $code = $route.Groups['code'].Value -replace '(?m)^ {10}', ''
        $code = $code.Replace('./tests/Invoke-HotfixAcceptance.ps1','Invoke-HotfixArgumentFixture')
        function Invoke-HotfixArgumentFixture {
            param([switch]$Build,[string]$CandidateTag,[string]$PublicTag,[switch]$DownloadDependencies,[switch]$DisposableRunner,[string]$ResultDirectory)
            return [pscustomobject]@{Keys=@($PSBoundParameters.Keys);Build=[bool]$Build;CandidateTag=$CandidateTag;PublicTag=$PublicTag;DownloadDependencies=[bool]$DownloadDependencies;DisposableRunner=[bool]$DisposableRunner;ResultDirectory=$ResultDirectory}
        }
        $directory = Join-Path $repository ('.scratch/workflow-routing-' + [guid]::NewGuid().ToString('N'))
        [IO.Directory]::CreateDirectory($directory) | Out-Null
        [IO.File]::WriteAllText((Join-Path $directory 'release-metadata.json'), ('{"release_state":"'+$State+'","tag":"v1.0.1"}'), [Text.UTF8Encoding]::new($false))
        $oldCandidate = $env:WINIMG_CANDIDATE_TAG
        Push-Location -LiteralPath $directory
        try {
            $env:WINIMG_CANDIDATE_TAG = $Candidate
            $observed = & ([scriptblock]::Create($code))
        } finally { Pop-Location; $env:WINIMG_CANDIDATE_TAG=$oldCandidate }
        $observed.Keys | Should -Contain $Expected
        foreach ($other in @('CandidateTag','PublicTag','Build') | Where-Object {$_ -ne $Expected}) { $observed.Keys | Should -Not -Contain $other }
        if ($Expected -eq 'Build') { $observed.Build | Should -BeTrue } else { $observed.$Expected | Should -Be 'v1.0.1' }
        $observed.DownloadDependencies | Should -BeTrue
        $observed.DisposableRunner | Should -BeTrue
        $observed.ResultDirectory | Should -Be '.scratch/ci-hotfix'
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
