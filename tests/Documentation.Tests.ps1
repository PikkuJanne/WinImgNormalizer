BeforeAll {
    if ($env:OS -ne 'Windows_NT') { throw 'Documentation examples require actual Windows.' }
    $repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $application = Join-Path $repository 'WinImgNormalizer.ps1'
    $publicPaths = @('README.md', 'docs/BEHAVIOR.md', 'SECURITY.md', 'docs/release/GETTING_STARTED.md')
    $documents = @{}
    foreach ($relative in $publicPaths) { $documents[$relative] = [IO.File]::ReadAllText((Join-Path $repository $relative)) }
    $scratch = Join-Path $repository '.scratch'
    function Assert-DocumentationAncestors([string]$Path) {
        $current = [IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($current) {
            if ($current.Exists -and ($current.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Documentation fixture crosses a reparse point.' }
            $current = $current.Parent
        }
    }
    Assert-DocumentationAncestors $scratch
    foreach ($pictures in @([Environment]::GetFolderPath('MyPictures'), (Join-Path $env:USERPROFILE 'Pictures'))) {
        if (-not $pictures) { continue }
        $protected = [IO.Path]::GetFullPath($pictures).TrimEnd('\', '/')
        if ($scratch.Equals($protected, [StringComparison]::OrdinalIgnoreCase) -or
            $scratch.StartsWith($protected + '\', [StringComparison]::OrdinalIgnoreCase) -or
            $protected.StartsWith($scratch + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Documentation fixtures overlap real Pictures.' }
    }
    & git -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'documentation-ignore-probe')
    if ($LASTEXITCODE -ne 0) { throw 'Documentation fixtures must already be ignored.' }
    $ownedRoot = Join-Path $scratch ('M4-T03-docs-' + [Guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $ownedRoot) { throw 'Documentation fixture ownership collision.' }
    $null = [IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'), 'M4-T03 owned synthetic documentation examples; retained for inspection; no real Pictures or recursive cleanup.', [Text.UTF8Encoding]::new($false))
    $observations = New-Object 'Collections.Generic.List[object]'
    $bindingsBefore = @($publicPaths + @('WinImgNormalizer.ps1', 'WinImgNormalizer.bat', 'release-metadata.json', 'tests/Documentation.Tests.ps1') | ForEach-Object {
        [pscustomobject]@{ Path = $_; Sha256 = (Get-FileHash -LiteralPath (Join-Path $repository $_)).Hash.ToLowerInvariant() }
    })
    if (-not $env:WINIMG_TEST_MAGICK -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)) { throw 'Explicit verified ImageMagick required.' }
    $magick = [IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK)
    Assert-DocumentationAncestors $magick
    $imagePin = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'legacy/toolchain.json') -Raw | ConvertFrom-Json
    (Get-FileHash -LiteralPath $magick).Hash.ToLowerInvariant() | Should -Be $imagePin.portable_imagemagick.executable_sha256
    . $application
    $hostExecutable = (Get-Process -Id $PID).Path
    $policiesBefore = @(Get-ExecutionPolicy -List | Where-Object Scope -ne Process | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join '|'
    function New-DocumentationDirectory([string]$Label) {
        $path = Join-Path $ownedRoot ($Label + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
        Assert-DocumentationAncestors $path
        if (Test-Path -LiteralPath $path) { throw 'Documentation fixture collision.' }
        $null = [IO.Directory]::CreateDirectory($path)
        return $path
    }
    function Get-DocumentationText([string]$Path) {
        return [regex]::Replace($documents[$Path].Replace('`', '').Replace('**', ''), '\s+', ' ')
    }
    $allPublicText = ($publicPaths | ForEach-Object { Get-DocumentationText $_ }) -join ' '
    function Get-DocumentationHeadings([string]$Text) {
        $counts = @{}
        return @([regex]::Matches($Text, '(?m)^#{1,6}\s+(.+?)\s*#*\s*$') | ForEach-Object {
            $title = $_.Groups[1].Value.Replace('`', '').Replace('*', '').ToLowerInvariant()
            $slug = [regex]::Replace($title, '[^\p{L}\p{N}_\- ]', '').Replace(' ', '-')
            if ($counts.ContainsKey($slug)) { $counts[$slug]++; $slug + '-' + $counts[$slug] } else { $counts[$slug] = 0; $slug }
        })
    }
    $packageLinks = @{
        'WinImgNormalizer.ps1' = 'WinImgNormalizer.ps1'; 'WinImgNormalizer.bat' = 'WinImgNormalizer.bat'
        'LICENSE' = 'LICENSE'; 'CHANGELOG.md' = 'CHANGELOG.md'
        'GETTING_STARTED.md' = 'docs/release/GETTING_STARTED.md'
        'THIRD_PARTY_NOTICES.md' = 'docs/release/THIRD_PARTY_NOTICES.md'
        'LICENSE-CC0.txt' = 'tests/fixtures/colour/LICENSE-CC0.txt'
    }
    function Assert-DocumentationLinks([string]$Relative, [string]$Text) {
        foreach ($link in [regex]::Matches($Text, '(?<!!)\[[^\]]*\]\((?<target><[^>]+>|[^)\s]+)(?:\s+"[^"]*")?\)')) {
            $target = $link.Groups['target'].Value.Trim('<', '>')
            if ($target -match '^https://') {
                if (-not [Uri]::IsWellFormedUriString($target, [UriKind]::Absolute)) { throw 'Malformed public HTTPS link.' }
                continue
            }
            if ($target -match '^[A-Za-z][A-Za-z0-9+.-]*:|^[\\/]' -or $target -match '[\x00-\x1f]') { throw 'Unsafe local documentation link.' }
            $parts = $target.Split('#', 2)
            $file = [Uri]::UnescapeDataString($parts[0]); $fragment = if ($parts.Count -eq 2) { [Uri]::UnescapeDataString($parts[1]) } else { '' }
            if (-not $file) { $source = $Relative }
            elseif ($Relative -eq 'docs/release/GETTING_STARTED.md') {
                if (-not $packageLinks.ContainsKey($file)) { throw ('Guide link is absent from the portable allowlist: ' + $file) }
                $source = $packageLinks[$file]
            } else {
                $full = [IO.Path]::GetFullPath((Join-Path (Split-Path -Parent (Join-Path $repository $Relative)) $file))
                if (-not $full.StartsWith($repository + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Local documentation link escapes the repository.' }
                $source = $full.Substring($repository.Length + 1)
            }
            $path = Join-Path $repository $source
            if (-not [IO.File]::Exists($path)) { throw ('Missing local documentation target: ' + $source) }
            if ($fragment -and $fragment -cnotin @(Get-DocumentationHeadings ([IO.File]::ReadAllText($path)))) { throw ('Missing documentation heading: ' + $target) }
        }
    }
    function Get-DocumentationExamples([string]$Text) {
        $examples = New-Object 'Collections.Generic.List[object]'
        foreach ($block in [regex]::Matches($Text, '(?ms)^```powershell[^\r\n]*\r?\n(?<body>.*?)^```\s*$')) {
            $tokens = $null; $errors = $null
            $ast = [Management.Automation.Language.Parser]::ParseInput($block.Groups['body'].Value, [ref]$tokens, [ref]$errors)
            if ($errors.Count) { throw 'Public PowerShell example contains a syntax error.' }
            foreach ($command in $ast.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] }, $true)) {
                if ($command.GetCommandName() -notmatch '^(?:\.\\|\./)?WinImgNormalizer\.ps1$') { continue }
                if ($command.InvocationOperator -eq [Management.Automation.Language.TokenKind]::Dot) { throw 'Documentation examples must execute the application, not dot-source definitions.' }
                $elements = @($command.CommandElements)
                if ($elements.Count -notin @(2, 3) -or $elements[1] -isnot [Management.Automation.Language.StringConstantExpressionAst]) { throw 'Application examples must use one literal positional source and optional byte cap.' }
                if ($command.Parent -isnot [Management.Automation.Language.PipelineAst] -or $command.Parent.PipelineElements.Count -ne 1 -or $command.Redirections.Count -or $command.Parent.Parent -ne $ast.EndBlock) { throw 'Application examples must be standalone commands.' }
                $cap = $null
                if ($elements.Count -eq 3) {
                    if ($elements[2] -isnot [Management.Automation.Language.ConstantExpressionAst] -or $elements[2].Extent.Text -notmatch '^[0-9]+$') { throw 'Documentation byte cap must be a plain positive integer literal.' }
                    $cap = ConvertTo-WinImgByteCap $elements[2].Extent.Text
                }
                $examples.Add([pscustomobject]@{ SourceLiteral = $elements[1].Value; Cap = $cap; Command = $command.Extent.Text })
            }
        }
        return @($examples.ToArray())
    }
    $examplesByDocument = @{}
    foreach ($relative in @('README.md', 'docs/release/GETTING_STARTED.md')) { $examplesByDocument[$relative] = @(Get-DocumentationExamples $documents[$relative]) }
    function Invoke-DocumentationCommand([object[]]$Arguments, [string]$OutputParent, [string]$MagickPath = $magick, [scriptblock]$PreflightRunner) {
        $parameters = @{ Arguments = $Arguments; OutputParent = $OutputParent; MagickPath = $MagickPath; NativeTemporaryRoot = $ownedRoot }
        if ($PreflightRunner) { $parameters.PreflightRunner = $PreflightRunner }
        $output = @(& Invoke-WinImgNormalizerCommand @parameters 6>&1 2>&1)
        $codes = @($output | Where-Object { $_ -is [int] -or $_ -is [long] })
        if ($codes.Count -ne 1) { throw 'Documentation command returned no unique application result.' }
        $result = [pscustomobject]@{ Code = $codes[0]; Text = ($output | Where-Object { $_ -isnot [int] -and $_ -isnot [long] }) -join "`n" }
        $observations.Add($result)
        return $result
    }
    function Get-DocumentationRun([string]$Parent) {
        $runs = @(Get-ChildItem -LiteralPath $Parent -Directory); $runs.Count | Should -Be 1
        $logs = @(Get-ChildItem -LiteralPath $runs[0].FullName -Recurse -File | Where-Object Extension -eq '.log'); $logs.Count | Should -Be 1
        return [pscustomobject]@{ Root = $runs[0].FullName; Log = [IO.File]::ReadAllText($logs[0].FullName) }
    }
    function Invoke-DocumentationMagick([string[]]$Arguments) {
        $text = @(& $magick @Arguments 2>&1); $code = $LASTEXITCODE
        if ($code -ne 0) { throw ('Synthetic documentation native probe failed: ' + ($text -join "`n")) }
        return $text -join "`n"
    }
    function Get-DocumentationFileState([string]$Path) {
        $item = Get-Item -LiteralPath $Path
        return ([ordered]@{ Hash = (Get-FileHash -LiteralPath $Path).Hash; Length = $item.Length; Created = $item.CreationTimeUtc.Ticks; Modified = $item.LastWriteTimeUtc.Ticks } | ConvertTo-Json -Compress)
    }
    function New-DocumentationVideo([string]$Path, [byte]$Seed = 17) {
        $bytes = New-Object byte[] 4096
        for ($i = 0; $i -lt $bytes.Length; $i++) { $bytes[$i] = [byte](($i * 29 + $Seed) % 256) }
        [IO.File]::WriteAllBytes($Path, $bytes)
        [IO.File]::SetLastWriteTimeUtc($Path, [DateTime]::Parse('2020-02-03T04:05:06Z').ToUniversalTime())
    }
    $versionText = Invoke-DocumentationMagick @('-version')
    $formatText = Invoke-DocumentationMagick @('-list', 'format')
    function New-DocumentationCapabilityControl([string]$Mode) {
        $version = $versionText; $formats = $formatText
        if ($Mode -eq 'version') { $version = $version.Replace('7.1.2-32', '7.1.1-47') }
        if ($Mode -eq 'jpeg') { $formats = [regex]::Replace($formats, '(?m)^(\s*JPEG\*?\s+(?:[A-Z0-9]+\s+)?)rw', '${1}r-') }
        return { param($Executable, $Arguments)
            $text = if ($Arguments[0] -eq '-version') { $version } elseif ($Arguments[0] -eq '-list') { $formats } else { throw 'Unexpected documentation control query.' }
            [pscustomobject]@{ ExitCode = 0; StdOut = $text; StdErr = ''; TimedOut = $false }
        }.GetNewClosure()
    }
}

Describe 'M4-T03 documentation parity with actual behavior (T072)' {
    It 'T072 resolves local public links and headings in repository and exact portable contexts' {
        foreach ($relative in $publicPaths) { { Assert-DocumentationLinks $relative $documents[$relative] } | Should -Not -Throw }
        { Assert-DocumentationLinks 'README.md' '[broken](missing-document.md)' } | Should -Throw '*Missing local*'
        { Assert-DocumentationLinks 'README.md' '[broken](README.md#absent-heading-control)' } | Should -Throw '*Missing documentation heading*'
        { Assert-DocumentationLinks 'docs/release/GETTING_STARTED.md' '[broken](../../SECURITY.md)' } | Should -Throw '*portable allowlist*'
    }

    It 'T072 accepts exactly both literal positional examples in README and portable guide and rejects unsafe drift' {
        foreach ($relative in $examplesByDocument.Keys) {
            $examples = @($examplesByDocument[$relative]); $examples.Count | Should -Be 2
            @($examples | Where-Object { $null -eq $_.Cap }).Count | Should -Be 1
            @($examples | Where-Object { $_.Cap -eq 2097152 }).Count | Should -Be 1
            foreach ($example in $examples) { $example.SourceLiteral | Should -BeExactly 'D:\Photos\2024' }
            $text = Get-DocumentationText $relative
            $text | Should -Match '1,048,576 bytes \(1 MiB\)'
            $text | Should -Match '(?i)best[- ]effort|Tiny targets may be impossible'
        }
        foreach ($bad in @('.\WinImgNormalizer.ps1 -Source ''D:\Photos\2024''', '.\WinImgNormalizer.ps1 ''D:\Photos\2024'' 1.5', '.\WinImgNormalizer.ps1 ''D:\Photos\2024'' 2097152 extra', '. .\WinImgNormalizer.ps1 ''D:\Photos\2024''', 'if ($false) { .\WinImgNormalizer.ps1 ''D:\Photos\2024'' }', '$saved = .\WinImgNormalizer.ps1 ''D:\Photos\2024''')) {
            $text = '```powershell' + "`n" + $bad + "`n" + '```'
            { Get-DocumentationExamples $text } | Should -Throw
        }
        Get-WinImgUsage | Should -Match ('WinImgNormalizer ' + [regex]::Escape((Get-WinImgVersion)))
        $allPublicText | Should -Match ([regex]::Escape('WinImgNormalizer.ps1 <sourceFolder> [maxBytes]'))
    }

    It 'T072 executes the documented <Label> positional byte target while preserving synthetic video bytes' -ForEach @(
        @{ Label = 'default 1 MiB'; Cap = $null; Expected = 1048576 }, @{ Label = 'custom 2 MiB'; Cap = 2097152; Expected = 2097152 }
    ) {
        $source = New-DocumentationDirectory 'example source with spaces'
        $parent = New-DocumentationDirectory 'example output'
        $video = Join-Path $source 'synthetic metadata.mp4'; New-DocumentationVideo $video
        $before = Get-DocumentationFileState $video
        $example = @($examplesByDocument['README.md'] | Where-Object { $_.Cap -eq $Cap })[0]
        $arguments = @($source)
        if ($null -ne $example.Cap) { $arguments += [string]$example.Cap }
        $result = Invoke-DocumentationCommand -Arguments $arguments -OutputParent $parent
        $result.Code | Should -Be 0 -Because $result.Text
        $run = Get-DocumentationRun $parent
        $run.Log | Should -Match ([regex]::Escape(('MaxBytes: {0} bytes ({1} MiB; best-effort target)' -f $Expected, ($Expected / 1MB))))
        (Get-FileHash -LiteralPath (Join-Path $run.Root 'synthetic metadata.mp4')).Hash | Should -BeExactly (Get-FileHash -LiteralPath $video).Hash
        Get-DocumentationFileState $video | Should -BeExactly $before
        $allPublicText | Should -Match '(?i)videos?.{0,90}(retain|keep).{0,50}metadata'
        $observations.Add([pscustomobject]@{ Kind = 'documented positional argument substitution into production command OutputParent seam'; DocumentCommand = $example.Command; CapBytes = $Expected; SourceUnchanged = $true; RealPictures = $false })
    }

    It 'T072 reports actual invalid byte target <Value> before creating destination output' -ForEach @(@{ Value = '0' }, @{ Value = '1.5' }) {
        $source = New-DocumentationDirectory 'bad byte source'; $parent = Join-Path $ownedRoot ('absent output-' + [Guid]::NewGuid().ToString('N'))
        $result = Invoke-DocumentationCommand -Arguments @($source, $Value) -OutputParent $parent
        $result.Code | Should -Be 1
        $message = 'maxBytes must be a positive whole number from 1 to 9223372036854775807.'
        $result.Text | Should -Match ([regex]::Escape($message))
        $allPublicText | Should -Match ([regex]::Escape($message))
        Test-Path -LiteralPath $parent | Should -BeFalse
    }

    It 'T072 reproduces unsafe nesting refusal with no destination creation or source mutation' {
        $source = New-DocumentationDirectory 'nesting source'
        $sentinel = Join-Path $source 'synthetic sentinel.txt'; [IO.File]::WriteAllText($sentinel, 'Owned synthetic nesting control.')
        $before = Get-DocumentationFileState $sentinel
        $parent = Join-Path $source 'absent Pictures substitute'
        $result = Invoke-DocumentationCommand -Arguments @($source) -OutputParent $parent
        $result.Code | Should -Be 1
        $message = 'The destination is inside or equal to the source, including a directory alias.'
        $result.Text | Should -Match ([regex]::Escape($message))
        $allPublicText | Should -Match ([regex]::Escape($message))
        Test-Path -LiteralPath $parent | Should -BeFalse
        Get-DocumentationFileState $sentinel | Should -BeExactly $before
    }

    It 'T072 reproduces the documented <Label> prerequisite failure before output allocation' -ForEach @(
        @{ Label = 'missing executable'; Mode = 'missing'; Message = 'ImageMagick magick.exe was not found. Install a supported build and put its executable on PATH.'; Fragment = 'ImageMagick magick.exe was not found.' },
        @{ Label = 'old version control'; Mode = 'version'; Message = 'ImageMagick 7.1.2-32 or newer supported 7.x build is required; update the dependency before running.'; Fragment = 'ImageMagick 7.1.2-32 or newer supported 7.x build is required' },
        @{ Label = 'missing JPEG writer control'; Mode = 'jpeg'; Message = 'This ImageMagick build has no JPEG encoder; install a build with JPEG write support.'; Fragment = 'This ImageMagick build has no JPEG encoder' }
    ) {
        $source = New-DocumentationDirectory 'dependency source'; $parent = Join-Path $ownedRoot ('absent prerequisite output-' + [Guid]::NewGuid().ToString('N'))
        if ($Mode -eq 'jpeg') { $null = Invoke-DocumentationMagick @('-size', '4x4', 'xc:red', ('PNG:' + (Join-Path $source 'synthetic readable.png'))) }
        $parameters = @{ Arguments = @($source); OutputParent = $parent }
        if ($Mode -eq 'missing') { $parameters.MagickPath = Join-Path $ownedRoot 'absent-magick.exe' } else { $parameters.PreflightRunner = New-DocumentationCapabilityControl $Mode }
        $result = Invoke-DocumentationCommand @parameters
        $result.Code | Should -Be 1 -Because $result.Text
        $result.Text | Should -Match ([regex]::Escape($Message))
        $Message.StartsWith($Fragment, [StringComparison]::Ordinal) | Should -BeTrue
        $allPublicText | Should -Match ([regex]::Escape($Fragment))
        Test-Path -LiteralPath $parent | Should -BeFalse
        if ($Mode -eq 'jpeg') {
            # JPEG writing is required for readable images. A genuine empty run
            # still requires ImageMagick but does not require its JPEG writer.
            $empty = New-DocumentationDirectory 'empty dependency source'
            $emptyParent = New-DocumentationDirectory 'empty dependency output'
            $emptyResult = Invoke-DocumentationCommand -Arguments @($empty) -OutputParent $emptyParent -PreflightRunner (New-DocumentationCapabilityControl 'jpeg')
            $emptyResult.Code | Should -Be 0 -Because $emptyResult.Text
            (Get-DocumentationRun $emptyParent).Log | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=0 Duplicates=0 Unsupported=0 Errors=0'
        }
    }

    It 'T072 retains an above-target lossy JPEG with white transparency and stripped source metadata' {
        $source = New-DocumentationDirectory 'alpha source'; $parent = New-DocumentationDirectory 'alpha output'
        $png = Join-Path $source 'private alpha.png'
        $null = Invoke-DocumentationMagick @('-size', '8x8', 'xc:rgba(180,30,30,0)', '-profile', (Join-Path $PSScriptRoot 'fixtures/colour/sRGB-v4.icc'), '-set', 'comment', 'synthetic-private-location', ('PNG:' + $png))
        $before = Get-DocumentationFileState $png
        $result = Invoke-DocumentationCommand -Arguments @($source, '1') -OutputParent $parent
        $result.Code | Should -Be 2 -Because $result.Text
        $run = Get-DocumentationRun $parent; $jpeg = Join-Path $run.Root 'private alpha.jpeg'
        (Get-Item -LiteralPath $jpeg).Length | Should -BeGreaterThan 1
        $native = '\\?\' + [IO.Path]::GetFullPath($jpeg)
        Invoke-DocumentationMagick @($native, '-format', '%m|%n|%[fx:round(255*p{0,0}.r)]|%[fx:round(255*p{0,0}.g)]|%[fx:round(255*p{0,0}.b)]', 'info:') | Should -BeExactly 'JPEG|1|255|255|255'
        $bytes = [IO.File]::ReadAllBytes($native); $offset = 2
        while ($offset -lt $bytes.Length - 1) {
            $bytes[$offset] | Should -Be 255
            $marker = $bytes[$offset + 1]; if ($marker -in @(218, 217)) { break }
            $marker | Should -Not -BeIn (@(225..239) + @(254))
            $length = [int]$bytes[$offset + 2] * 256 + $bytes[$offset + 3]; $length | Should -BeGreaterOrEqual 2
            $offset += $length + 2
        }
        Get-DocumentationFileState $png | Should -BeExactly $before
        $run.Log | Should -Match 'Alpha=WhiteAfterSrgb; OutputICC=None'
        $run.Log | Should -Match 'SizeWarnings=1'
        $allPublicText | Should -Match '(?i)lossy.{0,100}(JPEG|derivative)|JPEG.{0,100}lossy'
        $allPublicText | Should -Match '(?i)(strip|remove|lose).{0,90}metadata'
        $allPublicText | Should -Match '(?i)white.{0,60}(alpha|transparen)|(?:alpha|transparen).{0,60}white'
    }

    It 'T072 reports first displayed frame selection and a real omitted animation frame' {
        $source = New-DocumentationDirectory 'frame source'; $parent = New-DocumentationDirectory 'frame output'
        $gif = Join-Path $source 'two frames.gif'
        $null = Invoke-DocumentationMagick @('-size', '16x12', 'xc:red', 'xc:blue', '-delay', '10', '-loop', '0', ('GIF:' + $gif))
        $before = Get-DocumentationFileState $gif
        $result = Invoke-DocumentationCommand -Arguments @($source) -OutputParent $parent
        $result.Code | Should -Be 0 -Because $result.Text
        $run = Get-DocumentationRun $parent; $native = '\\?\' + (Join-Path $run.Root 'two frames.jpeg')
        Invoke-DocumentationMagick @('identify', '+ping', '-format', '%m|%w|%h|%n', $native) | Should -BeExactly 'JPEG|16|12|1'
        $run.Log | Should -Match 'SourceCount=2; Selected=1; Omitted=1; Unit=Frames; Policy=FirstDisplayedFrame; Decoder=GIF'
        Get-DocumentationFileState $gif | Should -BeExactly $before
        $allPublicText | Should -Match '(?i)GIF.{0,160}first displayed (?:logical )?frame'
        $allPublicText | Should -Match '(?i)TIFF.{0,100}first page'
        $allPublicText | Should -Match '(?i)HEIC.{0,140}(primary|first) image'
        $allPublicText | Should -Match '(?i)omitt|omission'
    }

    It 'T072 proves filename-time-length dedupe can skip different bytes and is not content verification' {
        $source = New-DocumentationDirectory 'dedupe source'; $parent = New-DocumentationDirectory 'dedupe output'
        foreach ($name in @('a', 'b')) { $null = [IO.Directory]::CreateDirectory((Join-Path $source $name)) }
        $first = Join-Path $source 'a/same.mp4'; $second = Join-Path $source 'b/same.mp4'
        New-DocumentationVideo $first 17; New-DocumentationVideo $second 29
        (Get-FileHash -LiteralPath $first).Hash | Should -Not -BeExactly (Get-FileHash -LiteralPath $second).Hash
        $before = @(Get-DocumentationFileState $first; Get-DocumentationFileState $second)
        $result = Invoke-DocumentationCommand -Arguments @($source) -OutputParent $parent
        $result.Code | Should -Be 0 -Because $result.Text
        $run = Get-DocumentationRun $parent
        $run.Log | Should -Match 'SUMMARY ConvertedImages=0 CopiedVideos=1 Duplicates=1 Unsupported=0 Errors=0'
        $run.Log | Should -Match 'Heuristic duplicate skipped: b[\\/]same.mp4'
        @(Get-DocumentationFileState $first; Get-DocumentationFileState $second) -join '|' | Should -BeExactly ($before -join '|')
        $allPublicText | Should -Match '(?i)heuristic'
        $allPublicText | Should -Match '(?i)(filename|file name).{0,170}(time|LastWriteTimeUtc).{0,170}(byte length|length|bytes)'
        $allPublicText | Should -Match '(?i)(not|no).{0,60}(content.hash|hash.based)|(?:does not|without).{0,60}(content hash|hash)'
    }

    It 'T072 documents cloud sync and private reports separately from local processing and recorded unsigned release status' {
        foreach ($relative in @('README.md', 'docs/release/GETTING_STARTED.md')) {
            $text = Get-DocumentationText $relative
            $text | Should -Match '(?i)does not upload'
            $text | Should -Match '(?i)(cloud[- ]sync|OneDrive).{0,180}(sync|upload)|(?:sync|upload).{0,180}(cloud[- ]sync|OneDrive)'
            $text | Should -Match '(?i)(log|report).{0,160}(private|sensitive).{0,100}(path|filename)|(?:private|sensitive).{0,100}(path|filename).{0,160}(log|report)'
        }
        $security = Get-DocumentationText 'SECURITY.md'
        $security | Should -Match '(?i)(redact|saniti[sz]|synthetic)'
        $security | Should -Match '(?i)(log|report).{0,140}(private|sensitive)|(?:private|sensitive).{0,140}(log|report)'
        $allPublicText | Should -Match '(?i)(Group Policy|organization|organisational|organizational)'
        $security | Should -Not -Match '(?i)[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}'
        $metadata = Get-Content -LiteralPath (Join-Path $repository 'release-metadata.json') -Raw | ConvertFrom-Json
        $metadata.version | Should -BeExactly (Get-WinImgVersion)
        if ($metadata.release_state -ceq 'unreleased') {
            foreach ($relative in @('README.md', 'docs/release/GETTING_STARTED.md')) { Get-DocumentationText $relative | Should -Match '(?i)unreleased' }
            foreach ($name in @('tag', 'release_date', 'release_url', 'asset_filename', 'asset_bytes', 'asset_sha256', 'download_url')) { $metadata.$name | Should -BeNullOrEmpty }
        } else {
            $metadata.release_state | Should -BeExactly 'published'
            $metadata.tag | Should -BeExactly 'v1.0.0'
            $metadata.asset_filename | Should -BeExactly 'WinImgNormalizer-1.0.0-portable.zip'
            $metadata.asset_bytes | Should -Be 184659
            $metadata.asset_sha256 | Should -BeExactly '251828028e144759c919645f423fde08641cabf42e04db7024d7ad17da4ba14d'
            $metadata.release_url | Should -BeExactly 'https://github.com/PikkuJanne/WinImgNormalizer/releases/tag/v1.0.0'
            $metadata.download_url | Should -BeExactly 'https://github.com/PikkuJanne/WinImgNormalizer/releases/download/v1.0.0/WinImgNormalizer-1.0.0-portable.zip'
            Get-DocumentationText 'README.md' | Should -Match '(?i)1\.0\.0 is published and unsigned'
            Get-DocumentationText 'README.md' | Should -Match '(?i)preparation.time unreleased wording'
            Get-DocumentationText 'docs/release/GETTING_STARTED.md' | Should -Match '(?i)unreleased'
            # The maintained generator verifies the dated witness and canonical
            # publication fields; this does not authenticate a fresh download.
            & (Join-Path $repository 'tools/release/Update-ReleaseMetadata.ps1') -Check
            if (-not $?) { throw 'Published documentation metadata check failed.' }
        }
        $allPublicText | Should -Match '(?i)unsigned'
        $allPublicText | Should -Match '(?i)checksum.{0,120}(not|does not).{0,120}(signature|authenticat)|(?:not|does not).{0,120}(signature|authenticat).{0,120}checksum'
    }
}

AfterAll {
    if ($ownedRoot -and $observations) { [IO.File]::WriteAllText((Join-Path $ownedRoot 'documentation-observations.json'), (@($observations.ToArray()) | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false)) }
    foreach ($binding in $bindingsBefore) { (Get-FileHash -LiteralPath (Join-Path $repository $binding.Path)).Hash.ToLowerInvariant() | Should -BeExactly $binding.Sha256 }
    if ($policiesBefore) { @(Get-ExecutionPolicy -List | Where-Object Scope -ne Process | ForEach-Object { $_.Scope.ToString() + '=' + $_.ExecutionPolicy.ToString() }) -join '|' | Should -BeExactly $policiesBefore }
}
