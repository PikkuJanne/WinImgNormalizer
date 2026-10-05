# Development acceptance helpers. Dot-sourcing only defines functions.
function Get-WinImgTestGateFailures {
    param([object]$Result, [string[]]$RequiredCases = @(), [string[]]$RequiredSuites = @())
    $failures = @()
    if ($Result.TotalCount -eq 0) { $failures += 'Zero tests discovered.' }
    if ($Result.FailedCount -gt 0) { $failures += 'One or more assertions failed.' }
    if ($Result.SkippedCount -gt 0) { $failures += 'Mandatory tests were skipped.' }
    if ($Result.NotRunCount -gt 0 -or $Result.InconclusiveCount -gt 0) { $failures += 'Mandatory tests were not completed.' }
    if ($Result.FailedBlocksCount -gt 0 -or $Result.FailedContainersCount -gt 0 -or $Result.Result -ne 'Passed') {
        $failures += 'Pester reported a failing run, block or container.'
    }
    $passedCases = @(Get-WinImgPassedTestCases $Result)
    foreach ($case in $RequiredCases) {
        if ($case -notin $passedCases) { $failures += ('Required case did not pass: ' + $case) }
    }
    foreach ($suite in $RequiredSuites) {
        $containers = @($Result.Containers | Where-Object { [IO.Path]::GetFileName([string]$_.Item) -eq $suite })
        if ($containers.Count -ne 1 -or $containers[0].TotalCount -eq 0 -or $containers[0].Result -ne 'Passed') {
            $failures += ('Required suite was absent, empty or failed: ' + $suite)
        }
    }
    return $failures
}

function Get-WinImgPassedTestCases {
    param([object]$Result)
    $cases = @(
        foreach ($test in $Result.Tests) {
            if ($test.Result -eq 'Passed') {
                foreach ($match in [regex]::Matches(($test.ExpandedPath -join '.'), '\bT[0-9]{3}\b')) { $match.Value }
            }
        }
    )
    return @($cases | Sort-Object -Unique)
}

function Get-WinImgExecutedCodecCoverage {
    param([object]$Result)
    foreach ($extension in @('jpg','jpeg','png','bmp','tif','tiff','gif','webp','heic','heif')) {
        $case = 'T065'
        $text = 'T065 fully decodes an actual .' + $extension + ' source and its JPEG derivative'
        if ($extension -in @('heic','heif')) { $case = 'T031'; $text = 'T031 decodes both actual HEVC collection images via .' + $extension + ',' }
        $count = @($Result.Tests | Where-Object { $_.Result -eq 'Passed' -and ($_.ExpandedPath -join '.').Contains($text) }).Count
        [pscustomobject]@{ extension=$extension; case_id=$case; passed_test_count=$count; executed=($count -eq 1) }
    }
}

function Assert-WinImgMandatoryTestManifest {
    param([string]$TestsRoot, [object]$Manifest)
    $actual = @(Get-ChildItem -LiteralPath $TestsRoot -Filter '*.Tests.ps1' -File | ForEach-Object { $_.Name } | Sort-Object)
    $expected = @($Manifest.suites | Sort-Object -Unique)
    if ($expected.Count -eq 0 -or $expected.Count -ne $Manifest.suites.Count -or
        (($actual -join '|') -ne ($expected -join '|'))) {
        throw 'Mandatory suite inventory differs from the reviewed manifest. Missing and unregistered suites fail acceptance.'
    }
    foreach ($suite in $expected) {
        if ($suite -notmatch '^[A-Za-z]+\.Tests\.ps1$') { throw 'Invalid mandatory suite name.' }
        $file = Get-Item -LiteralPath (Join-Path $TestsRoot $suite)
        if (($file.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Mandatory test suite must be a regular file.' }
    }
}
