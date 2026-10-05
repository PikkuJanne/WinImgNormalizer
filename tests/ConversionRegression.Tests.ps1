BeforeAll {
    $repository=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    function Assert-CorpusAncestors([string]$Path) {
        $current=[IO.DirectoryInfo]::new([IO.Path]::GetFullPath($Path))
        while ($current) {
            if ($current.Exists -and ($current.Attributes -band [IO.FileAttributes]::ReparsePoint)) {throw 'Conversion corpus crosses a reparse point.'}
            $current=$current.Parent
        }
    }
    $scratch=Join-Path $repository '.scratch'; Assert-CorpusAncestors $scratch
    foreach($pictures in @([Environment]::GetFolderPath('MyPictures'),(Join-Path $env:USERPROFILE 'Pictures'))) {
        if(-not $pictures){continue}; $full=[IO.Path]::GetFullPath($pictures).TrimEnd('\','/')
        if($scratch.Equals($full,[StringComparison]::OrdinalIgnoreCase) -or $scratch.StartsWith($full+'\',[StringComparison]::OrdinalIgnoreCase) -or $full.StartsWith($scratch+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Conversion corpus overlaps real Pictures.'}
    }
    & git -C $repository check-ignore --quiet --no-index -- (Join-Path $scratch 'conversion-corpus-ignore-probe')
    if($LASTEXITCODE -ne 0){throw 'Conversion corpus must be ignored.'}
    $ownedRoot=Join-Path $scratch ('M2-T06-conversion-'+[Guid]::NewGuid().ToString('N'))
    if([IO.Directory]::Exists($ownedRoot)){throw 'Corpus ownership collision.'}
    $null=[IO.Directory]::CreateDirectory($ownedRoot)
    [IO.File]::WriteAllText((Join-Path $ownedRoot '.winimg-fixture-root'),'M2-T06 owned synthetic conversion corpus')
    function New-CorpusDirectory([string]$Relative) {
        $path=[IO.Path]::GetFullPath((Join-Path $ownedRoot $Relative))
        if(-not $path.StartsWith($ownedRoot+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Corpus escaped owned scratch.'}
        Assert-CorpusAncestors $path; $null=[IO.Directory]::CreateDirectory($path); return $path
    }
    if(-not $env:WINIMG_TEST_MAGICK -or -not [IO.Path]::IsPathRooted($env:WINIMG_TEST_MAGICK)){throw 'Explicit verified ImageMagick required.'}
    $magick=[IO.Path]::GetFullPath($env:WINIMG_TEST_MAGICK); Assert-CorpusAncestors $magick
    . (Join-Path $repository 'WinImgNormalizer.ps1')
    $observations=New-Object 'Collections.Generic.List[object]'
    $entries=New-Object 'Collections.Generic.List[object]'
    $comparisons=New-Object 'Collections.Generic.List[object]'
    $assets=Join-Path $PSScriptRoot 'fixtures/colour'
    $bindings=@((Join-Path $repository 'WinImgNormalizer.ps1'),(Join-Path $repository 'WinImgNormalizer.bat'),(Join-Path $PSScriptRoot 'ConversionRegression.Tests.ps1'),(Join-Path $PSScriptRoot 'Invoke-Tests.ps1'),(Join-Path $PSScriptRoot 'Initialize-TestDependencies.ps1'),(Join-Path $PSScriptRoot 'dependencies.json'),$magick)+@(Get-ChildItem -LiteralPath $assets -File | ForEach-Object{$_.FullName})
    $sourceBindings=@($bindings | ForEach-Object{[pscustomobject]@{Path=$_;Sha256=(Get-FileHash -LiteralPath $_).Hash.ToLowerInvariant()}})
    if(-not ('WinImgCorpusRgb8' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.IO;
public static class WinImgCorpusRgb8 {
    public static double[] Measure(string reference, string decoded) {
        byte[] a=File.ReadAllBytes(reference), b=File.ReadAllBytes(decoded);
        if(a.Length==0 || a.Length!=b.Length || a.Length%3!=0) throw new InvalidDataException("RGB8 grids differ.");
        double absolute=0, squared=0;
        for(int i=0;i<a.Length;i++){int d=(int)a[i]-(int)b[i];absolute+=Math.Abs(d);squared+=(double)d*d;}
        double mse=squared/a.Length;
        return new double[]{absolute/a.Length,Math.Sqrt(mse)/255.0,mse==0 ? Double.PositiveInfinity : 10*Math.Log10(255.0*255.0/mse),a.Length/3};
    }
}
'@
    }
    function Invoke-CorpusMagick([string[]]$Arguments,[int]$OutputLimit=16384) {
        # Fixture creation and independent reads require the same literal local
        # path boundary. A generator redirect would create a false regression.
        $literal=@('-define','registry:filename:literal=true')
        if($Arguments[0] -in @('-version','-list')){$native=$Arguments}
        elseif($Arguments[0] -eq 'identify'){$native=@('identify')+$literal+$Arguments[1..($Arguments.Length-1)]}
        else {$native=$literal+$Arguments}
        $result=Invoke-WinImgNativeProcess -Executable $magick -Arguments $native -OutputLimit $OutputLimit
        $observations.Add([pscustomobject]@{Kind='fixture/reference native process';Arguments=$native;Result=$result})
        if($result.ExitCode -ne 0 -or $result.StdErr -or -not $result.StreamsComplete -or $result.StdOutTruncated -or $result.StdErrTruncated){throw ('Corpus native probe failed: '+$result.StdErr)}
        return ($result.StdOut -replace "`r`n","`n").TrimEnd("`r","`n")
    }
    function Corpus-Native([string]$Path){return Get-WinImgNativeOutputPath $Path}
    $reference=[IO.File]::ReadAllText((Join-Path $assets 'REFERENCE.json')) | ConvertFrom-Json
    $manifest=[IO.File]::ReadAllText((Join-Path $assets 'MANIFEST.json')) | ConvertFrom-Json
    foreach($profile in $manifest.profiles){(Get-FileHash -LiteralPath (Join-Path $assets ([IO.Path]::GetFileName($profile.source_path)))).Hash.ToLowerInvariant() | Should -Be $profile.sha256}
    $wideProfile=Join-Path $assets 'AdobeCompat-v2.icc'; $srgbProfile=Join-Path $assets 'sRGB-v4.icc'; $cmykProfile=Join-Path $assets 'CGATS001Compat-v2-micro.icc'
    # Unicode is constructed from code points: the shared PS1 is UTF8 without
    # BOM, and Windows PowerShell 5.1 otherwise decodes source as ANSI.
    $unicode=[string][char]0x00C4+[char]0x00F6+[char]0x00DF+' '+[char]0xD83D+[char]0xDE00
    $literalName="p%d & O'Brien ("+$unicode+')! [0]'
    $source=New-CorpusDirectory ('source '+$unicode)
    $parent=New-CorpusDirectory ('output%03d [1] '+$unicode)
    $probeRoot=New-CorpusDirectory 'comparisons'
    $literalDir=Join-Path $source $literalName; $null=[IO.Directory]::CreateDirectory($literalDir)
    $neighborName=$literalName.Replace('%d','0'); $neighborDir=Join-Path $source $neighborName; $null=[IO.Directory]::CreateDirectory($neighborDir)
    function Add-CorpusEntry([string]$Relative,[string]$Kind,[int]$Width,[int]$Height,[object[]]$Expected,[object[]]$Points,[int]$Count=1,[string]$Decoder='', [string]$Unit='', [string]$Policy='') {
        $entry=[pscustomobject]@{Source=$Relative;Output=[IO.Path]::ChangeExtension($Relative,'.jpeg');Kind=$Kind;Width=$Width;Height=$Height;Expected=$Expected;Points=$Points;Count=$Count;Decoder=$Decoder;Unit=$Unit;Policy=$Policy}
        $entries.Add($entry); return Join-Path $source $Relative
    }
    function New-CorpusRgb([string]$Path,[object[]]$Values,[string]$Profile,[switch]$Alpha,[string]$Coder='PNG',[int]$Width=128,[int]$Height=96) {
        $arguments=@('-size',("${Width}x${Height}"),'xc:none')
        for($i=0;$i -lt 4;$i++){
            $v=$Values[$i]; $color='rgb({0},{1},{2})' -f $v[0],$v[1],$v[2]
            if($Alpha){$color='rgba({0},{1},{2},{3})' -f $v[0],$v[1],$v[2],([double]$v[3]/255).ToString('R',[Globalization.CultureInfo]::InvariantCulture)}
            $x=($i%2)*($Width/2); $y=[Math]::Floor($i/2)*($Height/2)
            $arguments+=@('-fill',$color,'-draw',('rectangle {0},{1} {2},{3}' -f $x,$y,($x+$Width/2-1),($y+$Height/2-1)))
        }
        if(-not $Alpha){$arguments+=@('-alpha','off')}; if($Profile){$arguments+=@('-profile',(Corpus-Native $Profile))}
        $null=Invoke-CorpusMagick ($arguments+@('-depth','8','-quality','100',($Coder+':'+(Corpus-Native $Path))))
    }
    $quad=@(@(32,24),@(96,24),@(32,72),@(96,72))
    $wide=Add-CorpusEntry ($literalName+'\wide[0]%.png') 'wide_rgb' 128 96 $reference.wide_rgb $quad
    New-CorpusRgb $wide $reference.rgb_source $wideProfile
    $alpha=Add-CorpusEntry ($literalName+'\alpha[1]%03d.png') 'tagged_alpha_white' 128 96 $reference.tagged_alpha_white $quad
    New-CorpusRgb $alpha $reference.alpha_source $wideProfile -Alpha
    Invoke-CorpusMagick @('identify','-format','%[profiles]|%[channels]',(Corpus-Native $alpha)) | Should -Match '^icc\|srgba'
    Invoke-CorpusMagick @((Corpus-Native $alpha),'-format','%[fx:p{32,24}.a]','info:') | Should -Be '0'
    $blue=@(@(32,32,224),@(32,32,224),@(32,32,224),@(32,32,224))
    $neighbor=Add-CorpusEntry ($neighborName+'\wide[0]%.png') 'literal_neighbor' 128 96 $blue $quad
    New-CorpusRgb $neighbor $blue
    $cmyk=Add-CorpusEntry ($literalName+'\cmyk%[1].tiff') 'tagged_cmyk' 128 96 $reference.tagged_cmyk $quad 1 'TIFF' 'Pages' 'FirstPage'
    $raw=Join-Path $probeRoot 'synthetic-cmyk.raw'; $bytes=New-Object byte[] (128*96*4)
    for($y=0;$y -lt 96;$y++){for($x=0;$x -lt 128;$x++){$q=[int]($x -ge 64)+2*[int]($y -ge 48); for($c=0;$c -lt 4;$c++){$bytes[4*($y*128+$x)+$c]=$reference.cmyk_source[$q][$c]}}}
    [IO.File]::WriteAllBytes($raw,$bytes)
    $null=Invoke-CorpusMagick @('-size','128x96','-depth','8',('CMYK:'+(Corpus-Native $raw)),'-profile',(Corpus-Native $cmykProfile),'-compress','None',('TIFF:'+(Corpus-Native $cmyk)))
    Invoke-CorpusMagick @('identify','-format','%[colorspace]|%[profiles]',(Corpus-Native $cmyk)) | Should -Be 'CMYK|icc'
    function Add-CorpusSegment([string]$Path,[byte]$Marker,[byte[]]$Payload){
        $bytes=[IO.File]::ReadAllBytes($Path); $length=$Payload.Length+2
        $segment=[byte[]]@(255,$Marker,[byte]($length -shr 8),[byte]($length -band 255))+$Payload
        [IO.File]::WriteAllBytes($Path,[byte[]](@($bytes[0..1])+@($segment)+@($bytes[2..($bytes.Length-1)])))
    }
    $order=$reference.orientation.expected_order_by_orientation[5]
    $orientExpected=@($order | ForEach-Object {,$reference.orientation.nominal_corners[$_]})
    $orientation=Add-CorpusEntry ($literalName+'\orient[0]%.jpg') 'orientation6' 64 96 $orientExpected @(@(16,24),@(48,24),@(16,72),@(48,72))
    New-CorpusRgb $orientation $reference.orientation.nominal_corners $srgbProfile -Coder JPEG -Width 96 -Height 64
    $exif=[Convert]::FromBase64String($reference.orientation.exif_template_base64)
    # Existing big-endian SHORT orientation entry is changed from1 to6 only.
    $ifd=6+([int]$exif[10]*16777216+[int]$exif[11]*65536+[int]$exif[12]*256+[int]$exif[13]); $count=[int]$exif[$ifd]*256+[int]$exif[$ifd+1]; $found=0
    for($i=0;$i -lt $count;$i++){$o=$ifd+2+12*$i; if([int]$exif[$o]*256+[int]$exif[$o+1] -eq 274){$exif[$o+8]=0;$exif[$o+9]=6;$found++}}
    $found | Should -Be 1; Add-CorpusSegment $orientation 225 $exif
    Add-CorpusSegment $orientation 225 ([Text.Encoding]::UTF8.GetBytes("http://ns.adobe.com/xap/1.0/`0<x:xmpmeta xmlns:x='adobe:ns:meta/'><rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#'><rdf:Description xmlns:exif='http://ns.adobe.com/exif/1.0/' exif:GPSLatitude='12,34.56N'/></rdf:RDF></x:xmpmeta>"))
    $title=[Text.Encoding]::ASCII.GetBytes('Owned corpus IPTC title'); $iptc=[byte[]]@(28,2,5,0,[byte]$title.Length)+$title
    $photoshop=[Text.Encoding]::ASCII.GetBytes("Photoshop 3.0`0"+'8BIM')+[byte[]]@(4,4,0,0,0,0,[byte]($iptc.Length -shr 8),[byte]($iptc.Length -band 255))+$iptc
    if($iptc.Length%2){$photoshop+=[byte]0}; Add-CorpusSegment $orientation 237 $photoshop
    Add-CorpusSegment $orientation 254 ([Text.Encoding]::ASCII.GetBytes('Owned synthetic private comment'))
    Invoke-CorpusMagick @('identify','-format','%[EXIF:Orientation]|%[EXIF:GPSLatitude]|%[profiles]',(Corpus-Native $orientation)) | Should -Match '^6\|.*\|.*exif.*icc.*iptc.*xmp'
    $framePoints=@(@(0,0),@(20,20),@(40,32)); $frameValues=@(@(255,255,255),@(224,32,32),@(255,255,255))
    $gif=Add-CorpusEntry ($literalName+'\first[1]%d.gif') 'gif_first_canvas' 64 48 $frameValues $framePoints 2 'GIF' 'Frames' 'FirstDisplayedFrame'
    # Multi-image writers can append a scene suffix even with literal mode.
    # Encode neutral owned containers, then byte-copy to declared source names.
    $generatedGif=Join-Path $probeRoot 'generated-first.gif'
    $null=Invoke-CorpusMagick @('-background','none','(','-size','24x20','xc:none','-fill','#E02020','-draw','rectangle 4,4 19,15','-set','page','64x48+8+10','-dispose','Background',')','(','-size','64x48','xc:#2020E0','-set','page','64x48+0+0',')','-delay','12','-loop','0','-adjoin',('GIF:'+(Corpus-Native $generatedGif)))
    [IO.File]::WriteAllBytes($gif,[IO.File]::ReadAllBytes($generatedGif))
    Invoke-CorpusMagick @('identify','+ping','-regard-warnings','-format','%m|%n|%w|%h|%g\n',(Corpus-Native $gif))|Should -Be "GIF|2|24|20|64x48+8+10`nGIF|2|64|48|64x48+0+0"
    $pages=Add-CorpusEntry ($literalName+'\pages%03d[1].tiff') 'tiff_first_page6' 48 80 @(@(32,208,64),@(224,208,32),@(224,32,32)) @(@(40,10),@(8,68),@(24,40)) 2 'TIFF' 'Pages' 'FirstPage'
    $generatedTiff=Join-Path $probeRoot 'generated-pages.tiff'
    $null=Invoke-CorpusMagick @('(','-size','80x48','xc:#E02020','-fill','#20D040','-draw','rectangle 0,0 23,15','-fill','#E0D020','-draw','rectangle 56,32 79,47',')','(','-size','56x32','xc:#2020E0',')','-orient','TopLeft','-compress','None','-adjoin',('TIFF:'+(Corpus-Native $generatedTiff)))
    [IO.File]::WriteAllBytes($pages,[IO.File]::ReadAllBytes($generatedTiff))
    $tiff=[IO.File]::ReadAllBytes($pages); $ifd=[BitConverter]::ToUInt32($tiff,4); $count=[BitConverter]::ToUInt16($tiff,$ifd); $found=0
    for($i=0;$i -lt $count;$i++){$o=[int]$ifd+2+12*$i;if([BitConverter]::ToUInt16($tiff,$o) -eq 274){[BitConverter]::ToUInt16($tiff,$o+2)|Should -Be 3;$tiff[$o+8]=6;$tiff[$o+9]=0;$found++}}
    $found | Should -Be 1; [IO.File]::WriteAllBytes($pages,$tiff)
    Invoke-CorpusMagick @('identify','+ping','-regard-warnings','-format','%m|%n|%w|%h|%[orientation]\n',(Corpus-Native $pages))|Should -Be "TIFF|2|80|48|RightTop`nTIFF|2|56|32|TopLeft"
    $webp=Add-CorpusEntry ($literalName+'\canvas%03d[1].webp') 'webp_first_alpha' 64 48 @(@(255,255,255),@(224,32,32),@(255,255,255)) @(@(0,0),@(24,20),@(40,32)) 2 'WEBP' 'Frames' 'FirstDisplayedFrame'
    $generatedWebp=Join-Path $probeRoot 'generated-canvas.webp'
    $null=Invoke-CorpusMagick @('-background','none','(','-size','64x48','xc:none','-fill','#E02020','-draw','rectangle 12,14 15,25 rectangle 18,14 27,25',')','(','-size','64x48','xc:#2020E0',')','-define','webp:lossless=true','-delay','12','-loop','0','-adjoin',('WEBP:'+(Corpus-Native $generatedWebp)))
    [IO.File]::WriteAllBytes($webp,[IO.File]::ReadAllBytes($generatedWebp))
    Invoke-CorpusMagick @('identify','+ping','-regard-warnings','-format','%m|%n|%w|%h\n',(Corpus-Native $webp))|Should -Be "WEBP|2|64|48`nWEBP|2|64|48"
    Invoke-CorpusMagick @((Corpus-Native $webp),'-format','%s|%[channels]|%[pixel:p{0,0}]\n','info:') | Should -Match '^0\|srgba[^|]*\|srgba\(0,0,0,0\)'
    # Same approved self-generated two-image HEVC collection as T031,1328 bytes,
    # SHA618ed4d0...; pillow-heif1.8.0/Pillow12.3.0, libheif1.23.4/x2654.3,
    # lossless444, primary_index0, no thumbnails or external source media.
    # Official generator provenance: https://pypi.org/pypi/pillow-heif/1.8.0/json
    # This fixture is a still collection, not verified timed HEIC animation.
    $heicBytes=[Convert]::FromBase64String('AAAAHGZ0eXBoZWl4AAAAAG1pZjFoZWl4bWlhZgAAAa1tZXRhAAAAAAAAACFoZGxyAAAAAAAAAABwaWN0AAAAAAAAAAAAAAAAAAAAADRpbG9jAAAAAERAAAIAAQAAAAAB0QABAAAAAAAAAvsAAgAAAAAEzAABAAAAAAAAAGQAAAA4aWluZgAAAAAAAgAAABVpbmZlAgAAAAABAABodmMxAAAAABVpbmZlAgAAAAACAABodmMxAAAAAA5waXRtAAAAAAABAAABBmlwcnAAAADeaXBjbwAAAHdodmNDAQQIAAAAAAAAAAAA//AA/P/4+AAADwNgAAEAF0ABDAH//wQIAAADAJ44AAADAAD/ugJAYQABACpCAQEECAAAAwCeOAAAAwAA/5AEECCy3VyU1zcBDQYEAAADAGQAAAMABCBiAAEACEQBwXAwYJEgAAAAE2NvbHJuY2x4AAEADQAGgAAAABRpc3BlAAAAAAAAAEAAAABAAAAAKGNsYXAAAABAAAAAAQAAADAAAAABAAAAAAAAAAL////wAAAAAgAAABBwaXhpAAAAAAMICAgAAAAgaXBtYQAAAAAAAAACAAEFgQIDBYQAAgWBAgMFhAAAA2dtZGF0AAAC9ygBrwW4H2xh+gG9DAhCEISWMYxjDyf//yJ2v//1PKhmJqy5cuXNFixYsWLFEz/9kHP//4f/DfwX2vK8ryvK8vy/L8vy/L8vy7wJmK2JYhZMCnSWPsP//XXrRppMtra2treHh4eHh4eG2m/tTEAV97rWta1xKUpSlB9/mgoA88yH5iVlq0luFnhZ4WeFnhdsXbF2xdsXbF2xdsWt8ZP//VXBbL3CBJO6vcYH5y2nrrciv//FVGPK5hbAiuinOitpOhP//i4Djne2zn3dFq9FuJAwEob//zKBcu7tycHBwcIDw8PDw8PDkitNBAIPOYQhCEKb3ve94eVX2AMLW3etWivT0Y4VeFXhV4VeGIRiEYhGIRiEYhGIRWMF//yVhr8glqY7IqcpPgWJKmcsf/9j2A/YGubTs+b+OMohcPf/8Nv0Y+pkZERiPuBRLVVzdkX67f/8vgrhwyfiut3Vx3R56eWCJf2Esm8JPmB9f/aggv6/loBYGKQOgAf9H/ov578t8N8N8N8N8P8P8P8P8P8P8P8OAOuJ//8QYFP5hD7u7u7vTMzMzMzMzHdfnf//k9El2nr1Xsuy7LsvBIEgSBIEgSBHaIAEFO3tkbHwP87ELa5wS2/a9SmKQACv+5qn6K5R4mcQgihAhqVkcT5a6Cn//sB1zc/6JkRLMlkbPiW96r+AC7cvpOxROaqhVNxtSTDZsjMNNzYAQs29+KsALZZ8st4COdqi3OG9tnvgBKx76uVcW2ZgDZQQC1RsiV8SVgADUYDxBCBLIklRUhWhpOapNrNqfJADLZvxpjSQEmpGrFzUgRaa2j+pAAMlofHKORErVSfc3m4HWiXiJfAwhVgf//MnlxJZ59khXss+d05Jft6AAjRXbtzOmlcV2KYShW538QUuO06AErHvbqofkn036XWIQ8XIUtQwUI43//498aOt8IxunGPCAlBSABHIAF+3/Rfh6VRso2Ohcf7V7S00Kp2ABAD84N+9zT1gqvWLegjddm65SKHigAAAAGAoAa8FuB9sY7///yeiS7T16r2XZdl2XgkCQJAkCQJAjr7//+RNXAF7a47UIyD2De973ve+B0HQdB0HQdBz05/+yDn//n8urWta173ve973gTMVsSxCyhy/TimpgRyrOKI=')
    $hash=[Security.Cryptography.SHA256]::Create(); try{[BitConverter]::ToString($hash.ComputeHash($heicBytes)).Replace('-','').ToLowerInvariant() | Should -Be '618ed4d0c36f19bfac97401f28d45d42df9e9125d093875ef794c8c4e675a69a'}finally{$hash.Dispose()}
    foreach($extension in @('heic','heif')){
        $path=Add-CorpusEntry ($literalName+'\primary-'+$extension+'[1]%.'+$extension) ('collection_'+$extension) 64 48 @(@(32,208,64),@(224,208,32),@(224,32,32)) @(@(10,8),@(52,40),@(32,24)) 2 $extension.ToUpperInvariant() 'Images' 'DecoderPrimaryOrFirstImage'
        [IO.File]::WriteAllBytes($path,$heicBytes)
        # HEIC/HEIF are aliases of this same HEVC collection decoder. A literal
        # bracket suffix can change the format label, not exposed image count.
        Invoke-CorpusMagick @('identify','+ping','-regard-warnings','-format','%n|%w|%h\n',(Corpus-Native $path))|Should -Be "2|64|48`n2|64|48"
        Invoke-CorpusMagick @((Corpus-Native $path),'-format','%s|%[fx:round(255*p{32,24}.r)]|%[fx:round(255*p{32,24}.b)]\n','info:') | Should -Be "0|224|32`n1|32|224"
    }
    $portrait=Add-CorpusEntry 'portrait[0]%.bmp' 'ordinary_bmp' 96 128 (,@(32,208,64)) (,@(48,64))
    $null=Invoke-CorpusMagick @('-size','96x128','xc:#20D040',('BMP:'+(Corpus-Native $portrait)))
    $noise=Add-CorpusEntry 'noise[%d].png' 'seeded_noise' 1280 960 @() @()
    $null=Invoke-CorpusMagick @('-seed','4319','-size','1280x960','xc:gray','+noise','Random','-depth','8',('PNG:'+(Corpus-Native $noise)))
    $duplicateDir=Join-Path $source 'z duplicate'; $null=[IO.Directory]::CreateDirectory($duplicateDir)
    $duplicate=Join-Path $duplicateDir 'wide[0]%.png'; [IO.File]::WriteAllBytes($duplicate,[IO.File]::ReadAllBytes($wide))
    $videoRel=$literalName+'\nested\opaque[0]%d.MP4'; $video=Join-Path $source $videoRel; $null=[IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($video))
    $videoBytes=New-Object byte[] 65537; for($i=0;$i -lt $videoBytes.Length;$i++){$videoBytes[$i]=[byte](($i*31+13)%256)}; [IO.File]::WriteAllBytes($video,$videoBytes)
    $null=[IO.Directory]::CreateDirectory((Join-Path $source 'empty[1]%d'))
    [IO.File]::WriteAllText((Join-Path $source 'ignored-owned.txt'),'Owned unsupported text')
    $hidden=Join-Path $source '.hidden-owned.txt'; [IO.File]::WriteAllText($hidden,'Owned hidden unsupported text'); [IO.File]::SetAttributes($hidden,[IO.FileAttributes]::Hidden)
    $sentinel=Join-Path $parent 'existing-output-sentinel.dat'; [IO.File]::WriteAllText($sentinel,'Existing output parent preserved')
    $fixed=[DateTime]::Parse('2020-02-03T04:05:06Z').ToUniversalTime()
    foreach($file in @(Get-ChildItem -LiteralPath $source -Recurse -Force -File)+@([IO.FileInfo]::new($sentinel))){[IO.File]::SetCreationTimeUtc($file.FullName,$fixed);[IO.File]::SetLastWriteTimeUtc($file.FullName,$fixed)}
    foreach($directory in @([IO.DirectoryInfo]::new($source))+@(Get-ChildItem -LiteralPath $source -Recurse -Force -Directory)){[IO.Directory]::SetCreationTimeUtc($directory.FullName,$fixed);[IO.Directory]::SetLastWriteTimeUtc($directory.FullName,$fixed)}
    function Get-CorpusState([string]$Root){
        [pscustomobject]@{Files=@(Get-ChildItem -LiteralPath $Root -Recurse -Force -File | Sort-Object FullName | ForEach-Object{[pscustomobject]@{Path=$_.FullName.Substring($Root.Length+1);Sha256=(Get-FileHash -LiteralPath $_.FullName).Hash.ToLowerInvariant();Length=$_.Length;Creation=$_.CreationTimeUtc.Ticks;Modified=$_.LastWriteTimeUtc.Ticks;Attributes=[int]$_.Attributes}});Directories=@(@([IO.DirectoryInfo]::new($Root))+@(Get-ChildItem -LiteralPath $Root -Recurse -Force -Directory) | Sort-Object FullName | ForEach-Object{[pscustomobject]@{Path=if($_.FullName -eq $Root){'.'}else{$_.FullName.Substring($Root.Length+1)};Creation=$_.CreationTimeUtc.Ticks;Modified=$_.LastWriteTimeUtc.Ticks;Attributes=[int]$_.Attributes}})} | ConvertTo-Json -Depth 6 -Compress
    }
    function Corpus-Ordinal([string[]]$Values){$list=New-Object 'Collections.Generic.List[string]';foreach($v in $Values){$list.Add($v)};$list.Sort([StringComparer]::Ordinal);return $list.ToArray() -join '|'}
    function Assert-CorpusPrivacy([string]$Path,[switch]$AllowGrayscale){
        $bytes=[IO.File]::ReadAllBytes((Corpus-Native $Path)); $bytes[0]|Should -Be 255; $bytes[1]|Should -Be 216; $o=2;$progressive=0
        while($o+3 -lt $bytes.Length){
            $bytes[$o]|Should -Be 255;$marker=$bytes[$o+1];if($marker -in @(218,217)){break}
            $marker|Should -Not -BeIn @(225,226,237,254);$length=[int]$bytes[$o+2]*256+[int]$bytes[$o+3];$length|Should -BeGreaterOrEqual 2;($o+2+$length)|Should -BeLessOrEqual $bytes.Length
            if($marker -eq 194){
                $progressive++;$bytes[$o+11]|Should -Be 34
                if($bytes[$o+9] -eq 1){$AllowGrayscale|Should -BeTrue}
                else{$bytes[$o+9]|Should -Be 3;$bytes[$o+14]|Should -Be 17;$bytes[$o+17]|Should -Be 17}
            }
            $o+=2+$length
        }
        $progressive|Should -Be 1
    }
    function Get-CorpusPixels([string]$Path,[object[]]$Points){
        foreach($point in $Points){$format='%[fx:round(255*p{{{0},{1}}}.r)]|%[fx:round(255*p{{{0},{1}}}.g)]|%[fx:round(255*p{{{0},{1}}}.b)]' -f $point[0],$point[1]; ,@((Invoke-CorpusMagick @((Corpus-Native $Path),'-format',$format,'info:')).Split('|') | ForEach-Object{[int]$_})}
    }
    $before=Get-CorpusState $source
    [IO.File]::WriteAllText((Join-Path $probeRoot 'source-before.json'),$before)
    $entries.Count|Should -Be 12
    foreach($entry in $entries){
        [IO.File]::Exists((Join-Path $source $entry.Source))|Should -BeTrue;[IO.Path]::GetExtension($entry.Source)|Should -BeIn @('.png','.tiff','.jpg','.gif','.webp','.heic','.heif','.bmp')
        $entry.Expected.Count|Should -Be $entry.Points.Count
        foreach($tuple in $entry.Expected){$tuple.Count|Should -Be 3};foreach($point in $entry.Points){$point.Count|Should -Be 2}
    }
    $sentinelBefore=(Get-FileHash -LiteralPath $sentinel).Hash
    # Reference.json calibrates3/255 before JPEG and12/255 for uniform patch
    # centers. Frame/canvas source samples use the same pinned12/255 coding
    # tolerance; no such bound is asserted over edges or the noisy whole image.
    $jpegTolerance=[int]$reference.tolerance.jpeg_channel_units
    $preJpegTolerance=[int]$reference.tolerance.pre_jpeg_channel_units
    foreach($entry in @($entries.ToArray() | Where-Object Kind -in @('wide_rgb','tagged_alpha_white','tagged_cmyk'))){
        $pre=Join-Path $probeRoot ('native-lossless-'+$entry.Kind+'.png')
        $null=Invoke-CorpusMagick @('-define','image:frames=0',(Corpus-Native (Join-Path $source $entry.Source)),'-auto-orient','+black-point-compensation','-intent','Relative','-profile',(Corpus-Native $srgbProfile),'-background','white','-alpha','remove','-alpha','off','-strip','-depth','8',('PNG:'+(Corpus-Native $pre)))
        $samples=@(Get-CorpusPixels $pre $entry.Points)
        for($q=0;$q -lt $entry.Expected.Count;$q++){for($c=0;$c -lt 3;$c++){[Math]::Abs($samples[$q][$c]-$entry.Expected[$q][$c]) | Should -BeLessOrEqual $preJpegTolerance}}
        $known=Join-Path $probeRoot ('independent-tuples-'+$entry.Kind+'.png')
        New-CorpusRgb -Path $known -Values $entry.Expected
        $observations.Add([pscustomobject]@{Kind='independent Pillow LCMS reference';Source=$entry.Source;NativeLossless=$pre;KnownTuplePng=$known;ActualSamples=$samples;ExpectedSamples=$entry.Expected;Tolerance=$preJpegTolerance})
    }
    $version=Invoke-CorpusMagick @('-version')
    $formats=Invoke-CorpusMagick @('-list','format') -OutputLimit 65536
    $codecVersions=@($formats -split "`n"|Where-Object{$_ -match '^\s*(BMP|GIF|PNG|TIFF|JPEG|HEIC|HEIF|WEBP)\*?\s'})
    $noiseInitial=Join-Path $probeRoot 'noise-initial100-65537.jpg'
    $null=Invoke-CorpusMagick @('-quiet','-regard-warnings',(Corpus-Native $noise),'-auto-orient','-colorspace','sRGB','-background','white','-alpha','remove','-alpha','off','-strip','-sampling-factor','4:2:0','-interlace','Line','-resize','100%','-define','jpeg:extent=65537B',('JPEG:'+(Corpus-Native $noiseInitial)))
    ([IO.FileInfo]::new($noiseInitial)).Length | Should -BeGreaterThan 65537
    Invoke-CorpusMagick @('identify','+ping','-regard-warnings','-format','%m|%w|%h|%n',(Corpus-Native $noiseInitial)) | Should -Be 'JPEG|1280|960|1'
    $firstRun=$null; $firstState=$null
}

AfterAll {
    if($ownedRoot){
        $endBindings=@($bindings | ForEach-Object{[pscustomobject]@{Path=$_;Sha256=(Get-FileHash -LiteralPath $_).Hash.ToLowerInvariant()}})
        $artifactFiles=@(Get-ChildItem -LiteralPath $ownedRoot -Recurse -Force -File | ForEach-Object{[pscustomobject]@{Path=$_.FullName;Sha256=(Get-FileHash -LiteralPath $_.FullName).Hash.ToLowerInvariant();Length=$_.Length}})
        $record=[pscustomobject]@{Task='M2-T06';Recipe='Owned mixed tree: CC0 compact ICC/Pillow tuples, real EXIF6/privacy, GIF/WebP first alpha canvases, TIFF first oriented page, approved HEVC primary collections, literal paths, seed4319 noise, opaque video';NativeVersion=$version;NativeCodecFormats=$codecVersions;SourceBindings=$sourceBindings;EndSourceBindings=$endBindings;FixtureStateBefore=$before;FixtureStateFinal=if($source -and (Get-Command Get-CorpusState -ErrorAction SilentlyContinue)){Get-CorpusState $source}else{$null};ArtifactFiles=$artifactFiles;Entries=if($entries){$entries.ToArray()}else{@()};Comparisons=if($comparisons){$comparisons.ToArray()}else{@()};NativeProcesses=if($observations){$observations.ToArray()}else{@()}}
        [IO.File]::WriteAllText((Join-Path $ownedRoot 'comparison-artifacts.json'),($record | ConvertTo-Json -Depth 20),[Text.UTF8Encoding]::new($false))
        Write-Host ('Conversion corpus owned artifacts: '+$ownedRoot)
    }
}

Describe 'M2-T06 integrated conversion corpus (T046)' {
    It 'T046 combines actual conversion features at <Label> with source preservation and decoded comparison artifacts' -ForEach @(@{Label='default';Cap=1048576},@{Label='exact-small';Cap=65537}) {
        $parameters=@{Source=$source;OutputParent=$parent;MagickPath=$magick}
        if($Label -ne 'default'){$parameters.MaxBytes=$Cap}
        # No ProcessRunner, preflight or filesystem mocks: every application
        # probe, transform, validation and finalization uses the actual runtime.
        $output=@(& Invoke-WinImgNormalizer @parameters 6>&1 3>&1 2>&1)
        $codes=@($output | Where-Object{$_ -is [int] -or $_ -is [long]});$codes.Count|Should -Be 1;$codes[0]|Should -Be 0 -Because ($output -join "`n")
        $runs=@(Get-ChildItem -LiteralPath $parent -Directory | Sort-Object CreationTimeUtc)
        $runs.Count|Should -Be $(if($Label -eq 'default'){1}else{2}); $run=$runs[-1].FullName
        $logs=@(Get-ChildItem -LiteralPath (Join-Path $run '.WinImgNormalizer/reports') -File);$logs.Count|Should -Be 1;$log=[IO.File]::ReadAllText($logs[0].FullName)
        # Unsupported files are outside the runtime media inventory; their
        # unchanged bytes/times are part of the complete source-state check.
        $log|Should -Match ('MaxBytes: '+$Cap+' bytes');$log|Should -Match 'SUMMARY ConvertedImages=12 CopiedVideos=1 Duplicates=1 Unsupported=0 Errors=0 SizeWarnings=0 NativeWarnings=0'
        $log|Should -Not -Match 'ERR IMG:|WARN IMG:|RETRY IMG:'
        Get-CorpusState $source|Should -Be $before
        (Get-FileHash -LiteralPath $sentinel).Hash|Should -Be $sentinelBefore
        ([IO.FileInfo]::new($sentinel)).CreationTimeUtc.Ticks|Should -Be $fixed.Ticks;([IO.FileInfo]::new($sentinel)).LastWriteTimeUtc.Ticks|Should -Be $fixed.Ticks
        $wanted=@($entries.ToArray() | ForEach-Object{$_.Output})+@($videoRel)
        $actual=@(Get-ChildItem -LiteralPath $run -Recurse -Force -File | Where-Object Extension -ne '.log' | ForEach-Object{$_.FullName.Substring($run.Length+1)})
        Corpus-Ordinal $actual|Should -Be (Corpus-Ordinal $wanted)
        foreach($dir in @((ConvertFrom-Json $before).Directories | Where-Object Path -ne '.')){[IO.Directory]::Exists((Join-Path $run $dir.Path))|Should -BeTrue}
        @(Get-ChildItem -LiteralPath (Join-Path $run '.WinImgNormalizer/work') -Force).Count|Should -Be 0
        (Get-FileHash -LiteralPath (Join-Path $run $videoRel)).Hash|Should -Be (Get-FileHash -LiteralPath $video).Hash
        ([IO.FileInfo]::new((Join-Path $run $videoRel))).CreationTimeUtc.Ticks|Should -Be $fixed.Ticks
        ([IO.FileInfo]::new((Join-Path $run $videoRel))).LastWriteTimeUtc.Ticks|Should -Be $fixed.Ticks
        $rows=New-Object 'Collections.Generic.List[object]'
        foreach($entry in $entries){
            $target=Join-Path $run $entry.Output
            $match=[regex]::Match($log,('OK IMG: '+[regex]::Escape($entry.Source)+' -> '+[regex]::Escape($entry.Output)+' \[(\d+) bytes, MaxBytes=(\d+), Width=(\d+), Height=(\d+), Scale=(\d+)%\]'))
            $match.Success|Should -BeTrue -Because $entry.Source
            $bytes=[long]$match.Groups[1].Value;$width=[int]$match.Groups[3].Value;$height=[int]$match.Groups[4].Value;$scale=[int]$match.Groups[5].Value
            [long]$match.Groups[2].Value|Should -Be $Cap
            $bytes|Should -Be ([IO.FileInfo]::new($target)).Length;$bytes|Should -BeLessOrEqual $Cap
            ([IO.FileInfo]::new($target)).CreationTimeUtc.Ticks|Should -Be $fixed.Ticks;([IO.FileInfo]::new($target)).LastWriteTimeUtc.Ticks|Should -Be $fixed.Ticks
            $scale|Should -BeIn @(100,90,80,70,60,50);$width|Should -BeLessOrEqual $entry.Width;$height|Should -BeLessOrEqual $entry.Height
            $width|Should -Be ([int][Math]::Floor($entry.Width*$scale/100+0.5));$height|Should -Be ([int][Math]::Floor($entry.Height*$scale/100+0.5))
            if($entry.Kind -ne 'seeded_noise'){$scale|Should -Be 100}
            if($entry.Kind -eq 'seeded_noise' -and $Label -eq 'exact-small'){$scale|Should -BeLessThan 100}
            Invoke-CorpusMagick @('identify','+ping','-regard-warnings','-format','%m|%w|%h|%n',(Corpus-Native $target))|Should -Be ("JPEG|$width|$height|1")
            Assert-CorpusPrivacy $target -AllowGrayscale:($entry.Kind -eq 'seeded_noise')
            $actualPixels=@(Get-CorpusPixels $target $entry.Points);$maximum=0;$absolute=0;$samples=0
            for($q=0;$q -lt $entry.Expected.Count;$q++){for($c=0;$c -lt 3;$c++){$difference=[Math]::Abs($actualPixels[$q][$c]-$entry.Expected[$q][$c]);$difference|Should -BeLessOrEqual $jpegTolerance;$maximum=[Math]::Max($maximum,$difference);$absolute+=$difference;$samples++}}
            if($entry.Decoder){$log|Should -Match ([regex]::Escape(('SOURCE IMG: {0} (SourceCount={1}; Selected=1; Omitted={2}; Unit={3}; Policy={4}; Decoder={5})' -f $entry.Source,$entry.Count,($entry.Count-1),$entry.Unit,$entry.Policy,$entry.Decoder)))}
            $decoded=Join-Path $probeRoot ($Label+'-'+$rows.Count+'-decoded.png')
            $null=Invoke-CorpusMagick @((Corpus-Native $target),('PNG:'+(Corpus-Native $decoded)))
            $rmse=$null;$aligned=$null;$metricResult=$null;$rgbMetrics=$null
            $block=[regex]::Match($log,('COLOUR IMG: '+[regex]::Escape($entry.Source)+'.*?OK IMG: '+[regex]::Escape($entry.Source)),'Singleline').Value
            $attempts=@([regex]::Matches($block,'NATIVE IMG: Attempt=(\d+); Scale=(\d+)%; Category=(\w+); Exit=(\w+); TransientRetries=(\d+)/2'))
            $scaleTrace=@($attempts | ForEach-Object{[int]$_.Groups[2].Value})
            ($scaleTrace -join '|')|Should -Be ((@(100,90,80,70,60,50)|Where-Object{$_ -ge $scale}) -join '|')
            for($a=0;$a -lt $attempts.Count;$a++){[int]$attempts[$a].Groups[1].Value|Should -Be ($a+1);$attempts[$a].Groups[3].Value|Should -Be 'Success';$attempts[$a].Groups[4].Value|Should -Be '0';$attempts[$a].Groups[5].Value|Should -Be '0'}
            if($entry.Kind -eq 'seeded_noise'){
                $aligned=Join-Path $probeRoot ($Label+'-noise-lossless-aligned.png')
                $null=Invoke-CorpusMagick @((Corpus-Native $noise),'-colorspace','sRGB','-resize',("${width}x${height}!"),'-depth','8',('PNG:'+(Corpus-Native $aligned)))
                # This shared-decoder full-grid normalized RMSE is descriptive.
                # Native compare exit1 means differing pixels, not app failure.
                $metricResult=Invoke-WinImgNativeProcess -Executable $magick -Arguments @('compare','-define','registry:filename:literal=true','-metric','RMSE',(Corpus-Native $aligned),(Corpus-Native $target),'null:')
                $metricResult.ExitCode | Should -BeIn @(0,1)
                $metricResult.StreamsComplete | Should -BeTrue
                $metric=[regex]::Match($metricResult.StdErr.Trim(),'^[0-9.eE+\-]+ \(([0-9.eE+\-]+)\)$');$metric.Success | Should -BeTrue
                $rmse=[double]::Parse($metric.Groups[1].Value,[Globalization.CultureInfo]::InvariantCulture);$rmse | Should -BeGreaterOrEqual 0;$rmse | Should -BeLessOrEqual 1
                $rawReference=Join-Path $probeRoot ($Label+'-noise-reference.rgb');$rawDecoded=Join-Path $probeRoot ($Label+'-noise-decoded.rgb')
                $null=Invoke-CorpusMagick @((Corpus-Native $aligned),'-depth','8',('RGB:'+(Corpus-Native $rawReference)))
                $null=Invoke-CorpusMagick @((Corpus-Native $decoded),'-depth','8',('RGB:'+(Corpus-Native $rawDecoded)))
                $values=[WinImgCorpusRgb8]::Measure($rawReference,$rawDecoded)
                $rgbMetrics=[pscustomobject]@{MaeRgb8=$values[0];NormalizedRmse=$values[1];PsnrDb=$values[2];Pixels=$values[3];ReferenceRgb=$rawReference;DecodedRgb=$rawDecoded;Qualification='Descriptive full-grid differences at independently selected output grid; no perceptual pass threshold.'}
                [Math]::Abs($values[1]-$rmse)|Should -BeLessThan 0.00001
            }
            $rows.Add([pscustomobject]@{Source=$entry.Source;Output=$entry.Output;Kind=$entry.Kind;Bytes=$bytes;Cap=$Cap;Width=$width;Height=$height;Scale=$scale;ScaleTrace=$scaleTrace;Status='Converted';MaximumSampleError=if($samples){$maximum}else{$null};SampleMae=if($samples){[double]$absolute/$samples}else{$null};ActualSamples=$actualPixels;ExpectedSamples=$entry.Expected;NormalizedNoiseRmse=$rmse;Rgb8NoiseMetrics=$rgbMetrics;MetricNativeResult=$metricResult;AlignedLossless=$aligned;OutputSha256=(Get-FileHash -LiteralPath $target).Hash.ToLowerInvariant();DecodedPng=$decoded;DecodedPngSha256=(Get-FileHash -LiteralPath $decoded).Hash.ToLowerInvariant()})
        }
        $log|Should -Match ([regex]::Escape(('Heuristic duplicate skipped: z duplicate\wide[0]%.png (retained source: {0}; retained output: {1}; retained status: Converted)' -f $entries[0].Source,$entries[0].Output)))
        $comparisons.Add([pscustomobject]@{Label=$Label;Cap=$Cap;ApplicationExit=$codes[0];Run=$run;Log=$logs[0].FullName;LogSha256=(Get-FileHash -LiteralPath $logs[0].FullName).Hash.ToLowerInvariant();SourceStateAfter=(Get-CorpusState $source);Rows=$rows.ToArray()})
        [IO.File]::WriteAllText((Join-Path $probeRoot ($Label+'-source-after.json')),(Get-CorpusState $source))
        if($Label -eq 'default'){$script:firstRun=$run;$script:firstState=Get-CorpusState $run}
        else{$run|Should -Not -Be $script:firstRun;Get-CorpusState $script:firstRun|Should -Be $script:firstState}
    }
}
