# Development inventory only; does not import or run the application.
$ErrorActionPreference = 'Stop'
$pester = @(Get-Module -ListAvailable Pester | Select-Object -ExpandProperty Version | ForEach-Object { $_.ToString() })
$analyzer = @(Get-Module -ListAvailable PSScriptAnalyzer | Select-Object -ExpandProperty Version | ForEach-Object { $_.ToString() })
$os = Get-CimInstance Win32_OperatingSystem
[ordered]@{
  schema_version = 1
  os_caption = $os.Caption
  os_version = $os.Version
  os_build = $os.BuildNumber
  os_architecture = $os.OSArchitecture
  process_bits = [IntPtr]::Size * 8
  powershell_version = $PSVersionTable.PSVersion.ToString()
  powershell_edition = $PSVersionTable.PSEdition
  culture = [Globalization.CultureInfo]::CurrentCulture.Name
  ui_culture = [Globalization.CultureInfo]::CurrentUICulture.Name
  pester_versions = $pester
  analyzer_versions = $analyzer
  execution_policy = @(Get-ExecutionPolicy -List | ForEach-Object { [ordered]@{ scope = $_.Scope.ToString(); policy = $_.ExecutionPolicy.ToString() } })
} | ConvertTo-Json -Depth 5
