<#
WOW Inventory Scanner — PowerShell WPF GUI (v2.3.20)
Author: ChatGPT (for Digby)

Run:
powershell -ExecutionPolicy Bypass -File ".\WOW-Inventory-Scanner-v2.3.20.ps1"

Folder layout (same folder as this script):
.\Data\    -> Carts.csv, Scanners.csv, Mics.csv, Tangents.csv (optional), "Model Map.csv"
.\Output\  -> auto-created; saves Asset_Submission_*.csv, *_unknown_peripherals.csv, *_links.csv

Key changes vs v2.3.19:
- Removed "Choose Output" button. Output is fixed to .\Output next to the script.
- CSVs are auto-loaded from .\Data next to the script (with safe fallback to script folder).
- Paths are displayed read-only in the header/status.
#>

[CmdletBinding()]param()

Add-Type -AssemblyName PresentationCore,PresentationFramework | Out-Null

# ================= Helpers =================
function Normalize([object]$s){ if($null -eq $s){return ''}; return ($s -as [string]).Trim() }
function Upper([object]$s){ return (Normalize $s).ToUpperInvariant() }
function Normalize-HeaderName([string]$s){ if($null -eq $s){return ''}; return ([regex]::Replace($s.Trim(),'[^A-Za-z0-9]','').ToUpperInvariant()) }
function Set-Label([System.Windows.Controls.TextBlock]$Lbl,[string]$Text,[System.Windows.Media.Brush]$Brush){ if($Lbl){ $Lbl.Text=$Text; $Lbl.Foreground=$Brush } }
function Show-Toast([string]$Message,[System.Windows.Media.Brush]$Brush=[System.Windows.Media.Brushes]::DarkGreen){ if($script:lblStatus){ $script:lblStatus.Text=$Message; $script:lblStatus.Foreground=$Brush } else { Write-Host $Message } }
function Show-Error([string]$Message){ [void][System.Windows.MessageBox]::Show($Message,'Validation', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) }
function Import-Table([Parameter(Mandatory)][string]$Path){ if(-not (Test-Path $Path)){ throw "File not found: $Path" }; $ext=[IO.Path]::GetExtension($Path).ToLowerInvariant(); switch($ext){ '.csv'{Import-Csv -Path $Path} default{ throw "Unsupported file type '$ext' (use CSV)" } } }

function Build-Lookup{
  param($Rows,[string[]]$KeySynonyms,[switch]$CaseSensitive)
  $lookup=@{}; if(-not $Rows){ return $lookup }
  $headers = if($Rows.Count -gt 0){ $Rows[0].PSObject.Properties.Name } else { @() }
  $hmap=@{}; foreach($h in $headers){ $hmap[(Normalize-HeaderName $h)]=$h }
  $keyHeader=$null
  foreach($try in $KeySynonyms){ $norm = Normalize-HeaderName $try; if($hmap.ContainsKey($norm)){ $keyHeader = $hmap[$norm]; break } }
  if(-not $keyHeader){ return $lookup }
  foreach($r in $Rows){
    $key = $r.($keyHeader)
    if($null -ne $key){
      $normVal = if($CaseSensitive){ [string]$key } else { ([string]$key).Trim().ToUpperInvariant() }
      if(-not $lookup.ContainsKey($normVal)){ $lookup[$normVal]=$r }
    }
  }
  return $lookup
}

function Find-RowByKeys{ param([object[]]$Rows,[string[]]$Keys,[string]$Needle)
  if(-not $Rows -or -not $Keys -or [string]::IsNullOrWhiteSpace($Needle)){ return $null }
  $needleN = $Needle.Trim().ToUpperInvariant()
  foreach($k in $Keys){ foreach($r in $Rows){ try{ $v=$r.($k); if($null -ne $v){ if(([string]$v).Trim().ToUpperInvariant() -eq $needleN){ return $r } } } catch{} } }
  return $null
}

function Get-FirstProp{ param($Row,[string[]]$Candidates)
  if(-not $Row){ return $null }
  foreach($c in $Candidates){ if($Row.PSObject.Properties.Name -icontains $c){ return $Row.($c) } }
  return $null
}

function Describe-Row{ param($Row,[string[]]$Preferred)
  if(-not $Row){ return '<not found>' }
  $props = $Row.PSObject.Properties.Name; $parts=@()
  foreach($c in $Preferred){ $hit = $props | Where-Object { $_ -ieq $c }; if($hit){ $parts += ("{0}: {1}" -f $c, ($Row.($hit))) } }
  if(-not $parts){ $first = $props | Select-Object -First 3; foreach($c in $first){ $parts += ("{0}: {1}" -f $c, ($Row.($c))) } }
  return ($parts -join ' | ')
}

function Trim-AssetTag($s){
  $s = Upper $s
  if([string]::IsNullOrWhiteSpace($s)){ return '' }
  if($s.StartsWith('C') -and $s.Length -gt 7){ return $s.Substring(0,7) }
  return $s
}

function Get-Ordinal($d){
  if($d -in 11,12,13){ return 'th' }
  switch($d % 10){
    1 {'st'} 2 {'nd'} 3 {'rd'} Default {'th'}
  }
}
function Format-LongDate([datetime]$dt){
  $suf = Get-Ordinal $dt.Day
  return ("{0} {1}{2} {3}" -f $dt.ToString('MMMM'), $dt.Day, $suf, $dt.Year)
}

function Build-PeripheralAssetName([string]$Parent,[string]$Suffix){
  if([string]::IsNullOrWhiteSpace($Parent)){ return '' } else { return ($Parent + $Suffix) }
}

# ============== App State ==============
$global:State = [ordered]@{
  Auditor=$env:USERNAME; Company='ISLH'; Manufacturer='Ergotron';
  TangentRows=@(); TangentLookup=@{};
  CartRows=@(); CartLookup=@{};
  ScannerRows=@(); ScannerLookup=@{};
  MicRows=@(); MicLookup=@{};
  ModelRows=@(); ModelLookup=@{};
  ModelMpnList=@();
  OutputPath = $null
  UnknownLogPath=$null; LinkPath=$null;
  RecentItems = New-Object System.Collections.ObjectModel.ObservableCollection[psobject]
}

# ============== Base dir & path config ==============
function global:Get-BaseDir {
  if($script:BaseDir){ return $script:BaseDir }
  if($MyInvocation -and $MyInvocation.MyCommand -and $MyInvocation.MyCommand.Path){
    $script:BaseDir = Split-Path -Path $MyInvocation.MyCommand.Path -Parent; return $script:BaseDir
  }
  if($PSScriptRoot){ $script:BaseDir = $PSScriptRoot; return $script:BaseDir }
  if($PSCommandPath){ $script:BaseDir = Split-Path -Path $PSCommandPath -Parent; return $script:BaseDir }
  $script:BaseDir = (Get-Location).Path; return $script:BaseDir
}

# New: centralize Data/Output under the script folder
function global:Configure-Paths {
  $base = Get-BaseDir
  $script:DataDir   = Join-Path $base 'Data'
  $script:OutputDir = Join-Path $base 'Output'
  if(-not (Test-Path $script:OutputDir)){ New-Item -ItemType Directory -Path $script:OutputDir | Out-Null }
  if(-not (Test-Path $script:DataDir)){ New-Item -ItemType Directory -Path $script:DataDir | Out-Null }

  $global:State.OutputPath = Join-Path $script:OutputDir ("Asset_Submission_" + (Get-Date -Format 'yyyyMMdd_HHmm') + ".csv")
  $global:State.UnknownLogPath = $global:State.OutputPath -replace '\.csv$', '_unknown_peripherals.csv'
  $global:State.LinkPath       = $global:State.OutputPath -replace '\.csv$', '_links.csv'
}

function global:Resolve-File($baseDir, [string[]]$candidates){
  foreach($name in $candidates){
    $p = Join-Path -Path $baseDir -ChildPath $name
    if(Test-Path $p){ return $p }
  }
  return $null
}

# ============== UI (XAML) ==============
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="WOW Inventory Scanner v2.3.20" Height="980" Width="1220" WindowStartupLocation="CenterScreen">
  <ScrollViewer VerticalScrollBarVisibility="Auto">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="2.6*"/>
      <ColumnDefinition Width="3.4*"/>
    </Grid.ColumnDefinitions>

    <DockPanel Grid.ColumnSpan="2" LastChildFill="False" Margin="0,0,0,10">
      <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
        <TextBlock Text="Name:" VerticalAlignment="Center"/>
        <TextBox x:Name="txtAuditor" Width="200" Margin="6,0,8,0"/>
        <!-- Removed Choose Output button -->
      </StackPanel>
      <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
        <TextBlock Text="Carts:" VerticalAlignment="Center"/>
        <TextBlock x:Name="lblCarts" Margin="4,0,12,0" VerticalAlignment="Center"/>
        <TextBlock Text="Scanners:" VerticalAlignment="Center"/>
        <TextBlock x:Name="lblScanners" Margin="4,0,12,0" VerticalAlignment="Center"/>
        <TextBlock Text="Mics:" VerticalAlignment="Center"/>
        <TextBlock x:Name="lblMics" Margin="4,0,12,0" VerticalAlignment="Center"/>
        <TextBlock Text="Tangents:" VerticalAlignment="Center"/>
        <TextBlock x:Name="lblTangents" Margin="4,0,12,0" VerticalAlignment="Center"/>
        <TextBlock Text="Models:" VerticalAlignment="Center"/>
        <TextBlock x:Name="lblModels" Margin="4,0,12,0" VerticalAlignment="Center"/>
        <TextBlock Text="Output:" VerticalAlignment="Center" Margin="12,0,0,0"/>
        <TextBlock x:Name="lblOutputPath" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Width="300"/>
      </StackPanel>
    </DockPanel>

    <!-- Left column -->
    <StackPanel Grid.Row="1" Grid.Column="0" Margin="0,0,12,0">
      <TextBlock Text="Parent Tangent Hostname (AO...)" />
      <TextBox x:Name="txtParentHostname" FontSize="18"/>

      <TextBlock Text="Cart Serial Number (WOW)" Margin="0,8,0,0"/>
      <TextBox x:Name="txtCartSerial" FontSize="18"/>

      <TextBlock Text="Cart Asset Tag (new sticker)" Margin="0,8,0,0"/>
      <TextBox x:Name="txtCartAssetTag" FontSize="18"/>

      <TextBlock Text="Manufacturer Part Number (MPN)" Margin="0,8,0,0"/>
      <TextBox x:Name="txtMpn" FontSize="16"/>

      <TextBlock Text="Model (auto from MPN)" Margin="0,8,0,0"/>
      <TextBox x:Name="txtModel" FontSize="16" IsReadOnly="True" Background="#FFF6F6F6"/>

      <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
        <CheckBox x:Name="chkSharps" Content="Sharps bin on cart" Margin="0,0,16,0"/>
        <CheckBox x:Name="chkSanitizer" Content="Sanitizer on cart"/>
      </StackPanel>

      <TextBlock Text="Hand Scanner Asset Tag (optional)" Margin="0,12,0,0"/>
      <TextBox x:Name="txtScannerAssetTag" FontSize="18" Height="34" Width="340"/>

      <TextBlock Text="Dragon Mic Asset Tag (optional)" Margin="0,12,0,0"/>
      <TextBox x:Name="txtMicAssetTag" FontSize="18" Height="34" Width="340"/>

      <TextBlock Text="Notes (appended to Additional Comments)" Margin="0,12,0,0"/>
      <TextBox x:Name="txtNotes" Height="70" TextWrapping="Wrap" AcceptsReturn="True"/>

      <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
        <Button x:Name="btnClear" Content="Clear" Width="100"/>
        <Button x:Name="btnSave" Content="Save" Margin="8,0,0,0" Width="120"/>
      </StackPanel>

      <Border Margin="0,10,0,0" Padding="8" BorderBrush="Gray" BorderThickness="1" CornerRadius="4">
        <StackPanel>
          <TextBlock Text="Derived Cart Asset Name (AO-CRT)" FontWeight="Bold"/>
          <TextBlock x:Name="lblCartAssetName" FontSize="16"/>
        </StackPanel>
      </Border>
    </StackPanel>

    <!-- Right column -->
    <StackPanel Grid.Row="1" Grid.Column="1">
      <GroupBox Header="Cart Lookup (from Carts CSV)" Margin="0,0,0,8">
        <StackPanel>
          <TextBlock x:Name="lblCartInfo" Padding="8" TextWrapping="Wrap"/>
        </StackPanel>
      </GroupBox>
      <GroupBox Header="Hand Scanner Lookup (by Asset Tag)" Margin="0,0,0,8">
        <TextBlock x:Name="lblScannerInfo" Padding="8" TextWrapping="Wrap"/>
      </GroupBox>
      <GroupBox Header="Dragon Mic Lookup (by Asset Tag)" Margin="0,0,0,8">
        <TextBlock x:Name="lblMicInfo" Padding="8" TextWrapping="Wrap"/>
      </GroupBox>
      <GroupBox Header="Recent Saves (last 5)" Margin="0,0,0,8">
        <DataGrid x:Name="dgRecent" Height="80" IsReadOnly="True" AutoGenerateColumns="False"/>
      </GroupBox>

      <!-- Missing Peripherals -->
      <GroupBox Header="Add Missing Peripherals" Margin="0,0,0,8">
        <StackPanel>
          <GroupBox Header="Add Hand Scanner" Margin="0,4,0,8">
            <Grid Margin="6">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>

              <StackPanel Grid.Row="0" Grid.Column="0" Margin="0,0,8,4">
                <TextBlock Text="Peripheral Asset Tag"/>
                <TextBox x:Name="txtMissScanAssetTag" FontSize="16" Height="30"/>
              </StackPanel>
              <StackPanel Grid.Row="0" Grid.Column="1" Margin="0,0,8,4">
                <TextBlock Text="Peripheral Serial Number"/>
                <TextBox x:Name="txtMissScanSerial" FontSize="16" Height="30"/>
              </StackPanel>
              <StackPanel Grid.Row="0" Grid.Column="2" Margin="0,0,0,4">
                <TextBlock Text="MPN"/>
                <ComboBox x:Name="cboMissScanMpn" Height="30" IsEditable="True"/>
              </StackPanel>

              <StackPanel Grid.Row="1" Grid.Column="0" Margin="0,0,8,0">
                <TextBlock Text="Manufacturer"/>
                <TextBox x:Name="txtMissScanManu" FontSize="16" Height="30"/>
              </StackPanel>
              <StackPanel Grid.Row="1" Grid.Column="1" Margin="0,0,8,0">
                <TextBlock Text="Model"/>
                <TextBox x:Name="txtMissScanModel" FontSize="16" Height="30"/>
              </StackPanel>
              <StackPanel Grid.Row="1" Grid.Column="2" Margin="0,0,0,0">
                <TextBlock Text="Peripheral Asset Name (auto AO-xxx-SCN)" />
                <TextBox x:Name="txtMissScanAssetName" FontSize="16" Height="30" IsReadOnly="True" Background="#FFF6F6F6"/>
              </StackPanel>
            </Grid>
          </GroupBox>

          <GroupBox Header="Add Dragon Mic" Margin="0,0,0,8">
            <Grid Margin="6">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
              </Grid.RowDefinitions>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>

              <StackPanel Grid.Row="0" Grid.Column="0" Margin="0,0,8,4">
                <TextBlock Text="Peripheral Asset Tag"/>
                <TextBox x:Name="txtMissMicAssetTag" FontSize="16" Height="30"/>
              </StackPanel>
              <StackPanel Grid.Row="0" Grid.Column="1" Margin="0,0,8,4">
                <TextBlock Text="Peripheral Serial Number"/>
                <TextBox x:Name="txtMissMicSerial" FontSize="16" Height="30"/>
              </StackPanel>
              <StackPanel Grid.Row="0" Grid.Column="2" Margin="0,0,0,4">
                <TextBlock Text="MPN"/>
                <ComboBox x:Name="cboMissMicMpn" Height="30" IsEditable="True"/>
              </StackPanel>

              <StackPanel Grid.Row="1" Grid.Column="0" Margin="0,0,8,0">
                <TextBlock Text="Manufacturer"/>
                <TextBox x:Name="txtMissMicManu" FontSize="16" Height="30"/>
              </StackPanel>
              <StackPanel Grid.Row="1" Grid.Column="1" Margin="0,0,8,0">
                <TextBlock Text="Model"/>
                <TextBox x:Name="txtMissMicModel" FontSize="16" Height="30"/>
              </StackPanel>
              <StackPanel Grid.Row="1" Grid.Column="2" Margin="0,0,0,0">
                <TextBlock Text="Peripheral Asset Name (auto AO-xxx-MIC)" />
                <TextBox x:Name="txtMissMicAssetName" FontSize="16" Height="30" IsReadOnly="True" Background="#FFF6F6F6"/>
              </StackPanel>
            </Grid>
          </GroupBox>
        </StackPanel>
      </GroupBox>
    </StackPanel>

    <DockPanel Grid.Row="2" Grid.ColumnSpan="2" Margin="0,10,0,0">
      <TextBlock Text="Status:" FontWeight="Bold"/>
      <TextBlock x:Name="lblStatus" Margin="6,0,0,0"/>
      <DockPanel DockPanel.Dock="Right">
        <TextBlock Text="Output:" FontWeight="Bold"/>
        <TextBlock x:Name="lblOutput" Margin="6,0,0,0"/>
        <TextBlock Text="  Unknown log:" FontWeight="Bold" Margin="12,0,0,0"/>
        <TextBlock x:Name="lblUnknown" Margin="6,0,0,0"/>
      </DockPanel>
    </DockPanel>

  </Grid>
  </ScrollViewer>
</Window>
"@

$window = [Windows.Markup.XamlReader]::Parse($xaml)
function B([string]$n){ $c=$window.FindName($n); if(-not $c){ Write-Host "[WARN] missing XAML element: $n" -ForegroundColor Yellow }; return $c }

# Controls
$script:txtAuditor=B 'txtAuditor'
$script:lblCarts=B 'lblCarts'; $script:lblScanners=B 'lblScanners'; $script:lblMics=B 'lblMics'; $script:lblTangents=B 'lblTangents'; $script:lblModels=B 'lblModels'; $script:lblOutputPath=B 'lblOutputPath'; $script:lblOutput=B 'lblOutput'; $script:lblUnknown=B 'lblUnknown'; $script:lblStatus=B 'lblStatus'

$script:txtParentHostname=B 'txtParentHostname'; $script:txtCartSerial=B 'txtCartSerial'; $script:txtCartAssetTag=B 'txtCartAssetTag'; $script:txtMpn=B 'txtMpn'; $script:txtModel=B 'txtModel'
$script:chkSharps=B 'chkSharps'; $script:chkSanitizer=B 'chkSanitizer'; $script:txtScannerAssetTag=B 'txtScannerAssetTag'; $script:txtMicAssetTag=B 'txtMicAssetTag'; $script:txtNotes=B 'txtNotes'
$script:btnSave=B 'btnSave'; $script:btnClear=B 'btnClear'; $script:lblCartAssetName=B 'lblCartAssetName'

$script:lblCartInfo=B 'lblCartInfo'; $script:lblScannerInfo=B 'lblScannerInfo'; $script:lblMicInfo=B 'lblMicInfo'; $script:dgRecent=B 'dgRecent'

# Missing peripherals controls
$script:txtMissScanAssetTag=B 'txtMissScanAssetTag'; $script:txtMissScanSerial=B 'txtMissScanSerial'; $script:cboMissScanMpn=B 'cboMissScanMpn'; $script:txtMissScanManu=B 'txtMissScanManu'; $script:txtMissScanModel=B 'txtMissScanModel'; $script:txtMissScanAssetName=B 'txtMissScanAssetName'
$script:txtMissMicAssetTag=B 'txtMissMicAssetTag'; $script:txtMissMicSerial=B 'txtMissMicSerial'; $script:cboMissMicMpn=B 'cboMissMicMpn'; $script:txtMissMicManu=B 'txtMissMicManu'; $script:txtMissMicModel=B 'txtMissMicModel'; $script:txtMissMicAssetName=B 'txtMissMicAssetName'

# ============== Paths, DataGrid & initial labels ==============
Configure-Paths

if($dgRecent){
  $dgRecent.ItemsSource = $global:State.RecentItems
  foreach($p in @('Date','Name','CartSerial','CartName','CartTag','Model','MPN','Parent','MicAT','ScannerAT','Comments')){
    $col = New-Object System.Windows.Controls.DataGridTextColumn
    $col.Header=$p; $col.Binding=New-Object System.Windows.Data.Binding($p)
    [void]$dgRecent.Columns.Add($col)
  }
}
Set-Label $lblCartInfo '' ([System.Windows.Media.Brushes]::Gray)
Set-Label $lblScannerInfo '' ([System.Windows.Media.Brushes]::Gray)
Set-Label $lblMicInfo '' ([System.Windows.Media.Brushes]::Gray)
if($txtAuditor){ $txtAuditor.Text=$global:State.Auditor }
if($lblOutputPath){ $lblOutputPath.Text=$global:State.OutputPath }
if($lblOutput){ $lblOutput.Text=$global:State.OutputPath }
if($lblUnknown){ $lblUnknown.Text=$global:State.UnknownLogPath }
if($lblModels){ $lblModels.Text='' }

# ============== Output/Unknown file helpers ==============
function Ensure-OutputPaths{
  $outDir = (Split-Path -Path $global:State.OutputPath -Parent)
  if($outDir -and -not (Test-Path $outDir)){ New-Item -ItemType Directory -Path $outDir | Out-Null }
}
Ensure-OutputPaths

function Ensure-UnknownHeader{
  if(-not (Test-Path $global:State.UnknownLogPath)){
    $hdr = 'Submitted to Asset Management,Date,Entered by Name,Company,Peripheral Serial Number,Peripheral Asset Type,Peripheral Asset Name,Peripheral Asset Tag,Mobile Cart Serial Name (Parent),Manufacturer,Model,Manufacturer Part Number,Order Configuration,Comments'
    $hdr | Out-File -FilePath $global:State.UnknownLogPath -Encoding UTF8
  }
}
Ensure-UnknownHeader

# ============== Bulk load CSVs from .\Data (with fallback) ==============
function global:Load-CartsFrom([string]$p){
  if(-not $p){ return $false } try{
    $rows=Import-Table $p; $global:State.CartRows=$rows
    $global:State.CartLookup=Build-Lookup -Rows $rows -KeySynonyms @('Cart Serial Number','CartSerial','SerialNumber','Serial_Number','serial_number','WOWSerial','Serial')
    Set-Label $lblCarts ("Loaded: " + [IO.Path]::GetFileName($p) + " (" + $rows.Count + ")") ([System.Windows.Media.Brushes]::DarkSlateBlue)
    Show-Toast "Carts loaded: $($rows.Count)" ([System.Windows.Media.Brushes]::DarkSlateBlue)
    return $true
  } catch{ Show-Toast ("Carts load failed: " + $_.Exception.Message) ([System.Windows.Media.Brushes]::Firebrick); return $false }
}
function global:Load-ScannersFrom([string]$p){
  if(-not $p){ return $false } try{
    $rows=Import-Table $p; $global:State.ScannerRows=$rows
    $global:State.ScannerLookup=Build-Lookup -Rows $rows -KeySynonyms @('asset_tag','AssetTag','Asset Tag','Asset','Asset_Number','AssetNumber','Asset_Tag')
    Set-Label $lblScanners ("Loaded: " + [IO.Path]::GetFileName($p) + " (" + $rows.Count + ")") ([System.Windows.Media.Brushes]::DarkSlateBlue)
    Show-Toast "Scanners loaded: $($rows.Count)" ([System.Windows.Media.Brushes]::DarkSlateBlue)
    return $true
  } catch{ Show-Toast ("Scanners load failed: " + $_.Exception.Message) ([System.Windows.Media.Brushes]::Firebrick); return $false }
}
function global:Load-MicsFrom([string]$p){
  if(-not $p){ return $false } try{
    $rows=Import-Table $p; $global:State.MicRows=$rows
    $global:State.MicLookup=Build-Lookup -Rows $rows -KeySynonyms @('asset_tag','AssetTag','Asset Tag','Asset','Asset_Number','AssetNumber','Asset_Tag')
    Set-Label $lblMics ("Loaded: " + [IO.Path]::GetFileName($p) + " (" + $rows.Count + ")") ([System.Windows.Media.Brushes]::DarkSlateBlue)
    Show-Toast "Mics loaded: $($rows.Count)" ([System.Windows.Media.Brushes]::DarkSlateBlue)
    return $true
  } catch{ Show-Toast ("Mics load failed: " + $_.Exception.Message) ([System.Windows.Media.Brushes]::Firebrick); return $false }
}
function global:Load-TangentsFrom([string]$p){
  if(-not $p){ return $false } try{
    $rows=Import-Table $p; $global:State.TangentRows=$rows
    $global:State.TangentLookup=Build-Lookup -Rows $rows -KeySynonyms @('Hostname','ComputerName','Name','HostName')
    Set-Label $lblTangents ("Loaded: " + [IO.Path]::GetFileName($p) + " (" + $rows.Count + ")") ([System.Windows.Media.Brushes]::DarkSlateBlue)
    Show-Toast "Tangents loaded: $($rows.Count)" ([System.Windows.Media.Brushes]::DarkSlateBlue)
    return $true
  } catch{ Show-Toast ("Tangents load failed: " + $_.Exception.Message) ([System.Windows.Media.Brushes]::DarkGoldenrod); return $false }
}
function global:Load-ModelMapFrom([string]$p){
  if(-not $p){ return $false } try{
    $rows=Import-Table $p; $global:State.ModelRows=$rows
    $global:State.ModelLookup=Build-Lookup -Rows $rows -KeySynonyms @('Manufacturer Part Number','u_manufacturer_part_number','manufacturer_part_number','Mfr Part Number','PartNumber','MPN')
    $mpns = @()
    foreach($r in $rows){
      $mpn = Get-FirstProp -Row $r -Candidates @('Manufacturer Part Number','u_manufacturer_part_number','manufacturer_part_number','Mfr Part Number','PartNumber','MPN')
      if($mpn){ $mpns += ([string]$mpn).Trim() }
    }
    $mpns = $mpns | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique
    $global:State.ModelMpnList = $mpns
    if($cboMissScanMpn){ $cboMissScanMpn.ItemsSource = $mpns }
    if($cboMissMicMpn){ $cboMissMicMpn.ItemsSource = $mpns }
    if($lblModels){ Set-Label $lblModels ("Loaded: " + [IO.Path]::GetFileName($p) + " (" + $rows.Count + ")") ([System.Windows.Media.Brushes]::DarkSlateBlue) }
    Show-Toast "Model map loaded: $($rows.Count)" ([System.Windows.Media.Brushes]::DarkSlateBlue)
    Update-ModelFromMpn
    return $true
  } catch{ Show-Toast ("Model map load failed: " + $_.Exception.Message) ([System.Windows.Media.Brushes]::Firebrick); return $false }
}

function global:LoadAllCSVs{
  $base = Get-BaseDir
  $data = Join-Path $base 'Data'

  $carts = Resolve-File $data @('Carts.csv'); if(-not $carts){ $carts = Resolve-File $base @('Carts.csv') }
  if($carts){ if(Load-CartsFrom $carts){ $ok++ } } else { Set-Label $lblCarts 'Missing: Data\Carts.csv' ([System.Windows.Media.Brushes]::Firebrick); $miss++ }

  $scanners = Resolve-File $data @('Scanners.csv'); if(-not $scanners){ $scanners = Resolve-File $base @('Scanners.csv') }
  if($scanners){ if(Load-ScannersFrom $scanners){ $ok++ } } else { Set-Label $lblScanners 'Missing: Data\Scanners.csv' ([System.Windows.Media.Brushes]::Firebrick); $miss++ }

  $mics = Resolve-File $data @('Mics.csv'); if(-not $mics){ $mics = Resolve-File $base @('Mics.csv') }
  if($mics){ if(Load-MicsFrom $mics){ $ok++ } } else { Set-Label $lblMics 'Missing: Data\Mics.csv' ([System.Windows.Media.Brushes]::Firebrick); $miss++ }

  $tangents = Resolve-File $data @('Tangents.csv'); if(-not $tangents){ $tangents = Resolve-File $base @('Tangents.csv') }
  if($tangents){ if(Load-TangentsFrom $tangents){ $ok++ } } else { Set-Label $lblTangents 'Missing: Data\Tangents.csv' ([System.Windows.Media.Brushes]::DarkGoldenrod) }

  $model = Resolve-File $data @('Model Map.csv'); if(-not $model){ $model = Resolve-File $base @('Model Map.csv') }
  if($model){ if(Load-ModelMapFrom $model){ $ok++ } } else { Set-Label $lblModels 'Missing: Data\Model Map.csv' ([System.Windows.Media.Brushes]::Firebrick); $miss++ }

  if($miss -eq 0){ Show-Toast "Reloaded all CSVs from Data ($ok loaded)" ([System.Windows.Media.Brushes]::DarkGreen) }
  else { Show-Toast "Reload: $ok file(s) loaded, $miss missing. Base: $data" ([System.Windows.Media.Brushes]::DarkOrange) }
}

# ============== Derived labels & focus ==============
function Update-CartAssetName{
  if(-not $lblCartAssetName){return}
  $ph=Normalize $txtParentHostname.Text
  if([string]::IsNullOrWhiteSpace($ph)){ $lblCartAssetName.Text='' } else { $lblCartAssetName.Text = ($ph + '-CRT') }
  if($txtMissScanAssetName){ $txtMissScanAssetName.Text = (Build-PeripheralAssetName $ph '-SCN') }
  if($txtMissMicAssetName){  $txtMissMicAssetName.Text  = (Build-PeripheralAssetName $ph '-MIC') }
}
if($txtParentHostname){ $txtParentHostname.Add_TextChanged({ Update-CartAssetName }) }
Update-CartAssetName

foreach($tb in @($txtAuditor,$txtParentHostname,$txtCartSerial,$txtCartAssetTag,$txtMpn,$txtModel,$txtScannerAssetTag,$txtMicAssetTag,$txtNotes,$txtMissScanAssetTag,$txtMissScanSerial,$txtMissScanManu,$txtMissScanModel,$txtMissMicAssetTag,$txtMissMicSerial,$txtMissMicManu,$txtMissMicModel)){ if($tb){ $tb.Add_GotKeyboardFocus({ param($s,$e) $s.SelectAll() }) } }

function Move-Focus($to){ if($to){ $to.Focus() } }
if($txtParentHostname){   $txtParentHostname.Add_KeyDown({ if($_.Key -eq 'Enter'){ Move-Focus $txtCartSerial } }) }
if($txtCartSerial){       $txtCartSerial.Add_KeyDown({     if($_.Key -eq 'Enter'){ Move-Focus $txtCartAssetTag } }) }
if($txtCartAssetTag){     $txtCartAssetTag.Add_KeyDown({   if($_.Key -eq 'Enter'){ Move-Focus $txtMpn } }) }
if($txtMpn){              $txtMpn.Add_KeyDown({            if($_.Key -eq 'Enter'){ Move-Focus $txtModel } }) }
if($txtModel){            $txtModel.Add_KeyDown({          if($_.Key -eq 'Enter'){ Move-Focus $txtScannerAssetTag } }) }
if($txtScannerAssetTag){  $txtScannerAssetTag.Add_KeyDown({if($_.Key -eq 'Enter'){ Move-Focus $txtMicAssetTag } }) }
if($txtMicAssetTag){      $txtMicAssetTag.Add_KeyDown({    if($_.Key -eq 'Enter'){ Move-Focus $txtNotes } }) }
if($txtNotes){            $txtNotes.Add_KeyDown({          if($_.Key -eq 'Enter'){ Do-Save } }) }

# ============== Model autoload for main MPN ==============
function Update-ModelFromMpn{
  if(-not $txtMpn -or -not $txtModel){ return }
  $mpn = Upper $txtMpn.Text
  if([string]::IsNullOrWhiteSpace($mpn)){ $txtModel.Text=''; return }
  $row=$null
  if($global:State.ModelLookup.ContainsKey($mpn)){
    $row = $global:State.ModelLookup[$mpn]
  } else {
    $row = Find-RowByKeys -Rows $global:State.ModelRows -Keys @('Manufacturer Part Number','u_manufacturer_part_number','manufacturer_part_number','Mfr Part Number','PartNumber','MPN') -Needle $mpn
    if($row){ $global:State.ModelLookup[$mpn]=$row }
  }
  if($row){
    $model = Get-FirstProp -Row $row -Candidates @('Model','Model Name','model_name','Asset Model')
    if($model){ $txtModel.Text = ([string]$model).Trim() } else { $txtModel.Text='' }
  } else { $txtModel.Text = '' }
}
if($txtMpn){ $txtMpn.Add_TextChanged({ Update-ModelFromMpn }) }

# ============== Lookup renderers ==============
function Render-CartInfo{
  $serial = Upper $txtCartSerial.Text
  if([string]::IsNullOrWhiteSpace($serial)){ Set-Label $lblCartInfo '' ([System.Windows.Media.Brushes]::Gray); return }
  $row=$null
  if($global:State.CartLookup.ContainsKey($serial)){ $row=$global:State.CartLookup[$serial] } else { $row = Find-RowByKeys -Rows $global:State.CartRows -Keys @('Cart Serial Number','CartSerial','SerialNumber','Serial_Number','serial_number','WOWSerial','Serial') -Needle $serial; if($row){ $global:State.CartLookup[$serial]=$row } }
  if($row){
    $mpn = Get-FirstProp -Row $row -Candidates @('u_manufacturer_part_number','Manufacturer Part Number','manufacturer_part_number','Mfr Part Number','PartNumber','MPN')
    if($mpn){ $txtMpn.Text = ([string]$mpn).Trim(); Update-ModelFromMpn }
    $desc = Describe-Row -Row $row -Preferred 'serial_number','Cart Serial Number','Model','u_manufacturer_part_number','Manufacturer Part Number','asset','name'
    Set-Label $lblCartInfo ('FOUND - ' + $desc) ([System.Windows.Media.Brushes]::DarkGreen)
  } else {
    Set-Label $lblCartInfo ('NOT FOUND: ' + $serial) ([System.Windows.Media.Brushes]::Firebrick)
  }
}
if($txtCartSerial){ $txtCartSerial.Add_TextChanged({ Render-CartInfo }) }

function Render-ScannerInfo{
  $asset = Trim-AssetTag $txtScannerAssetTag.Text
  if([string]::IsNullOrWhiteSpace($asset)){ Set-Label $lblScannerInfo '' ([System.Windows.Media.Brushes]::Gray); return $false }
  $row=$null; if($global:State.ScannerLookup.ContainsKey($asset)){ $row=$global:State.ScannerLookup[$asset] } else { $row = Find-RowByKeys -Rows $global:State.ScannerRows -Keys @('asset_tag','AssetTag','Asset Tag','Asset','Asset_Number','AssetNumber','Asset_Tag') -Needle $asset; if($row){ $global:State.ScannerLookup[$asset]=$row } }
  if($row){ Set-Label $lblScannerInfo ('FOUND - [interpreted: ' + $asset + '] - ' + (Describe-Row -Row $row -Preferred 'AssetTag','asset_tag','Model','SerialNumber','Owner')) ([System.Windows.Media.Brushes]::DarkGreen); return $true } else { Set-Label $lblScannerInfo ('NOT FOUND [interpreted: ' + $asset + ']') ([System.Windows.Media.Brushes]::Firebrick); return $false }
}
if($txtScannerAssetTag){ $txtScannerAssetTag.Add_TextChanged({ Render-ScannerInfo | Out-Null }) }

function Render-MicInfo{
  $asset = Trim-AssetTag $txtMicAssetTag.Text
  if([string]::IsNullOrWhiteSpace($asset)){ Set-Label $lblMicInfo '' ([System.Windows.Media.Brushes]::Gray); return $false }
  $row=$null; if($global:State.MicLookup.ContainsKey($asset)){ $row=$global:State.MicLookup[$asset] } else { $row = Find-RowByKeys -Rows $global:State.MicRows -Keys @('asset_tag','AssetTag','Asset Tag','Asset','Asset_Number','AssetNumber','Asset_Tag') -Needle $asset; if($row){ $global:State.MicLookup[$asset]=$row } }
  if($row){ Set-Label $lblMicInfo ('FOUND - [interpreted: ' + $asset + '] - ' + (Describe-Row -Row $row -Preferred 'AssetTag','asset_tag','Model','SerialNumber','Owner')) ([System.Windows.Media.Brushes]::DarkGreen); return $true } else { Set-Label $lblMicInfo ('NOT FOUND [interpreted: ' + $asset + ']') ([System.Windows.Media.Brushes]::Firebrick); return $false }
}
if($txtMicAssetTag){ $txtMicAssetTag.Add_TextChanged({ Render-MicInfo | Out-Null }) }

# ============== Save helpers ==============
function Ensure-CsvAppend($path, $obj){
  if($null -eq $obj){ return }
  $obj | Export-Csv -Path $path -NoTypeInformation -Append -Encoding UTF8
}

function Add-UnknownRow {
  param(
    [Parameter(Mandatory)][string]$Type,            # 'Scanner' or 'Mic'
    [string]$AssetTag,
    [string]$SerialNumber,
    [string]$MPN,
    [string]$Manufacturer,
    [string]$Model
  )
  Ensure-UnknownHeader
  $parent = Normalize $txtParentHostname.Text
  $assetName = ''
  if(-not [string]::IsNullOrWhiteSpace($parent)){
    if($Type -ieq 'Scanner'){ $assetName = $parent + '-SCN' } else { $assetName = $parent + '-MIC' }
  }

  $ptype = 'Microphone'
  if($Type -ieq 'Scanner'){ $ptype = 'Barcode Scanner' }

  $manuOut = ''; if($Manufacturer){ $manuOut = ($Manufacturer.Trim()) }
  $modelOut = ''; if($Model){ $modelOut = ($Model.Trim()) }
  $mpnOut   = ''; if($MPN){ $mpnOut   = ($MPN.Trim()) }
  $serialOut= ''; if($SerialNumber){ $serialOut = ($SerialNumber.Trim()) }
  $atagOut  = Trim-AssetTag $AssetTag

  $row = [pscustomobject][ordered]@{
    'Submitted to Asset Management'   = ''
    'Date'                            = (Format-LongDate (Get-Date))
    'Entered by Name'                 = $global:State.Auditor
    'Company'                         = $global:State.Company
    'Peripheral Serial Number'        = $serialOut
    'Peripheral Asset Type'           = $ptype
    'Peripheral Asset Name'           = $assetName
    'Peripheral Asset Tag'            = $atagOut
    'Mobile Cart Serial Name (Parent)'= ($parent + '-CRT')
    'Manufacturer'                    = $manuOut
    'Model'                           = $modelOut
    'Manufacturer Part Number'        = $mpnOut
    'Order Configuration'             = ''
    'Comments'                        = ''
  }
  Ensure-CsvAppend $global:State.UnknownLogPath $row
}

# ============== Main Save ==============
function Do-Save{
  Ensure-OutputPaths
  $global:State.Auditor = Normalize $txtAuditor.Text
  $parent = Normalize $txtParentHostname.Text
  $cartSN = Upper $txtCartSerial.Text
  $cartAT = Upper $txtCartAssetTag.Text
  if([string]::IsNullOrWhiteSpace($parent)){ Show-Error 'Parent hostname (AO...) is required.'; return }
  if([string]::IsNullOrWhiteSpace($cartSN)){ Show-Error 'Cart Serial Number is required.'; return }

  $model = Normalize $txtModel.Text
  $mpn   = Normalize $txtMpn.Text
  $notes = Normalize $txtNotes.Text

  $scannerAT = Trim-AssetTag $txtScannerAssetTag.Text
  $micAT     = Trim-AssetTag $txtMicAssetTag.Text

  # Validate unknown peripherals
  $needScannerMissing = $false
  $needMicMissing = $false

  if($scannerAT){
    $scannerKnown = $global:State.ScannerLookup.ContainsKey($scannerAT) -or (Find-RowByKeys -Rows $global:State.ScannerRows -Keys @('asset_tag','AssetTag','Asset Tag','Asset','Asset_Number','AssetNumber','Asset_Tag') -Needle $scannerAT)
    if(-not $scannerKnown){ $needScannerMissing = $true }
  }
  if($micAT){
    $micKnown = $global:State.MicLookup.ContainsKey($micAT) -or (Find-RowByKeys -Rows $global:State.MicRows -Keys @('asset_tag','AssetTag','Asset Tag','Asset','Asset_Number','AssetNumber','Asset_Tag') -Needle $micAT)
    if(-not $micKnown){ $needMicMissing = $true }
  }

  if($needScannerMissing){
    if([string]::IsNullOrWhiteSpace($txtMissScanAssetTag.Text)){ $txtMissScanAssetTag.Text = $scannerAT }
    if([string]::IsNullOrWhiteSpace($txtMissScanSerial.Text) -or [string]::IsNullOrWhiteSpace($cboMissScanMpn.Text)){
      Show-Error "Scanner asset tag '$scannerAT' not found in the Scanners list.`nPlease fill the 'Add Hand Scanner' section: Asset Tag, Serial Number, and MPN."
      return
    }
  }
  if($needMicMissing){
    if([string]::IsNullOrWhiteSpace($txtMissMicAssetTag.Text)){ $txtMissMicAssetTag.Text = $micAT }
    if([string]::IsNullOrWhiteSpace($txtMissMicSerial.Text) -or [string]::IsNullOrWhiteSpace($cboMissMicMpn.Text)){
      Show-Error "Mic asset tag '$micAT' not found in the Mics list.`nPlease fill the 'Add Dragon Mic' section: Asset Tag, Serial Number, and MPN."
      return
    }
  }

  # Comments
  $comments=@()
  if($chkSharps.IsChecked){$comments+='Sharps Bin'}
  if($chkSanitizer.IsChecked){$comments+='Sanitizer'}
  if($scannerAT){$comments+='Hand Scanner'}
  if($micAT){$comments+='Dragon Mic'}
  if($notes){$comments+=$notes}
  $commentsJoined = ($comments -join ', ')

  # Unknown rows write-through if needed
  if($needScannerMissing){
    Add-UnknownRow -Type 'Scanner' -AssetTag $txtMissScanAssetTag.Text -SerialNumber $txtMissScanSerial.Text -MPN $cboMissScanMpn.Text -Manufacturer $txtMissScanManu.Text -Model $txtMissScanModel.Text
  }
  if($needMicMissing){
    Add-UnknownRow -Type 'Mic' -AssetTag $txtMissMicAssetTag.Text -SerialNumber $txtMissMicSerial.Text -MPN $cboMissMicMpn.Text -Manufacturer $txtMissMicManu.Text -Model $txtMissMicModel.Text
  }

  # Main output
  $row = [pscustomobject][ordered]@{
    'Submitted to Asset Management' = ''
    'Date Added to List'            = (Get-Date -Format 'M/d/yy')
    'Entered by Name'               = $global:State.Auditor
    'Company'                       = $global:State.Company
    'Cart Serial Number'            = $cartSN
    'Cart Asset Name'               = ($parent + '-CRT')
    'Cart Asset Tag'                = $cartAT
    'Manufacturer'                  = $global:State.Manufacturer
    'Model'                         = $model
    'Manufacturer Part Number'      = $mpn
    'Order Configuration'           = ''
    'Parent Name (All-In-One)'      = $parent
    'Additional Comments'           = $commentsJoined
  }
  $row | Export-Csv -Path $global:State.OutputPath -NoTypeInformation -Append -Encoding UTF8

  $lnk = [pscustomobject][ordered]@{
    ParentHostname  = $parent
    CartAssetName   = ($parent + '-CRT')
    MicAssetTag     = $micAT
    ScannerAssetTag = $scannerAT
  }
  $lnk | Export-Csv -Path $global:State.LinkPath -NoTypeInformation -Append -Encoding UTF8

  $global:State.RecentItems.Add([pscustomobject]@{
    Date      = $row.'Date Added to List'
    Name      = $row.'Entered by Name'
    CartSerial= $row.'Cart Serial Number'
    CartName  = $row.'Cart Asset Name'
    CartTag   = $row.'Cart Asset Tag'
    Model     = $row.'Model'
    MPN       = $row.'Manufacturer Part Number'
    Parent    = $row.'Parent Name (All-In-One)'
    MicAT     = $lnk.MicAssetTag
    ScannerAT = $lnk.ScannerAssetTag
    Comments  = $row.'Additional Comments'
  })
  if($global:State.RecentItems.Count -gt 5){ $global:State.RecentItems.RemoveAt(0) }

  Show-Toast ('Saved: ' + $row.'Cart Asset Name') ([System.Windows.Media.Brushes]::DarkGreen)
  Clear-Form -keepAuditor
}

function Clear-Form([switch]$keepAuditor){
  if(-not $keepAuditor -and $txtAuditor){ $txtAuditor.Text=$global:State.Auditor }
  $txtParentHostname.Text=''; $txtCartSerial.Text=''; $txtCartAssetTag.Text=''
  $txtModel.Text=''; $txtMpn.Text=''
  $chkSharps.IsChecked=$false; $chkSanitizer.IsChecked=$false
  $txtScannerAssetTag.Text=''; $txtMicAssetTag.Text=''; $txtNotes.Text=''
  $txtMissScanAssetTag.Text=''; $txtMissScanSerial.Text=''; $cboMissScanMpn.Text=''; $txtMissScanManu.Text=''; $txtMissScanModel.Text=''
  $txtMissMicAssetTag.Text='';  $txtMissMicSerial.Text='';  $cboMissMicMpn.Text='';  $txtMissMicManu.Text='';  $txtMissMicModel.Text=''
  Set-Label $lblCartInfo '' ([System.Windows.Media.Brushes]::Gray)
  Set-Label $lblScannerInfo '' ([System.Windows.Media.Brushes]::Gray)
  Set-Label $lblMicInfo '' ([System.Windows.Media.Brushes]::Gray)
  Update-CartAssetName
  $txtParentHostname.Focus()
}

# ============== Missing peripherals MPN handlers ==============
function Apply-ModelMap-ToFields {
  param(
    [string]$mpn,
    [System.Windows.Controls.TextBox]$txtManu,
    [System.Windows.Controls.TextBox]$txtModelBox
  )
  if([string]::IsNullOrWhiteSpace($mpn)){
    if($txtManu){ $txtManu.Text = '' }
    if($txtModelBox){ $txtModelBox.Text = '' }
    return
  }
  $key = (Upper $mpn)
  $row = $null
  if($global:State.ModelLookup.ContainsKey($key)){
    $row = $global:State.ModelLookup[$key]
  } else {
    $row = Find-RowByKeys -Rows $global:State.ModelRows -Keys @('Manufacturer Part Number','u_manufacturer_part_number','manufacturer_part_number','Mfr Part Number','PartNumber','MPN') -Needle $key
  }
  $manu = Get-FirstProp -Row $row -Candidates @('Manufacturer')
  $model= Get-FirstProp -Row $row -Candidates @('Model','Model Name','model_name','Asset Model')

  if($txtManu){
    if($manu){ $txtManu.Text = [string]$manu } else { $txtManu.Text = '' }
  }
  if($txtModelBox){
    if($model){ $txtModelBox.Text = [string]$model } else { $txtModelBox.Text = '' }
  }
}

if($cboMissScanMpn){
  $cboMissScanMpn.Add_SelectionChanged({ Apply-ModelMap-ToFields -mpn $cboMissScanMpn.SelectedItem -txtManu $txtMissScanManu -txtModelBox $txtMissScanModel })
  $cboMissScanMpn.Add_DropDownClosed({ Apply-ModelMap-ToFields -mpn $cboMissScanMpn.Text -txtManu $txtMissScanManu -txtModelBox $txtMissScanModel })
  $cboMissScanMpn.Add_LostFocus({ Apply-ModelMap-ToFields -mpn $cboMissScanMpn.Text -txtManu $txtMissScanManu -txtModelBox $txtMissScanModel })
  $cboMissScanMpn.Add_KeyUp({ Apply-ModelMap-ToFields -mpn $cboMissScanMpn.Text -txtManu $txtMissScanManu -txtModelBox $txtMissScanModel })
}
if($cboMissMicMpn){
  $cboMissMicMpn.Add_SelectionChanged({ Apply-ModelMap-ToFields -mpn $cboMissMicMpn.SelectedItem -txtManu $txtMissMicManu -txtModelBox $txtMissMicModel })
  $cboMissMicMpn.Add_DropDownClosed({ Apply-ModelMap-ToFields -mpn $cboMissMicMpn.Text -txtManu $txtMissMicManu -txtModelBox $txtMissMicModel })
  $cboMissMicMpn.Add_LostFocus({ Apply-ModelMap-ToFields -mpn $cboMissMicMpn.Text -txtManu $txtMissMicManu -txtModelBox $txtMissMicModel })
  $cboMissMicMpn.Add_KeyUp({ Apply-ModelMap-ToFields -mpn $cboMissMicMpn.Text -txtManu $txtMissMicManu -txtModelBox $txtMissMicModel })
}

# ============== Buttons / keys ==============
if($btnSave){ $btnSave.Add_Click({ Do-Save }) }
if($btnClear){ $btnClear.Add_Click({ Clear-Form }) }
if($txtAuditor){ $txtAuditor.Add_TextChanged({ $global:State.Auditor = Normalize $txtAuditor.Text }) }
$window.Add_KeyDown({ if($_.Key -eq 'F5'){ Do-Save } elseif($_.Key -eq 'F6'){ Clear-Form } })

# ============== Auto-load CSVs on launch ==============
$window.Add_Loaded({ 
  try{ 
    # Counters used in LoadAllCSVs
    $script:ok = 0; $script:miss = 0
    LoadAllCSVs 
  } catch { 
    Show-Toast ("Auto-load failed: " + $_.Exception.Message) ([System.Windows.Media.Brushes]::DarkOrange) 
  } 
})

[void]$window.ShowDialog()
