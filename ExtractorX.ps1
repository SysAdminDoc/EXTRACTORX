<#
.SYNOPSIS
    ExtractorX v2.1.0 - Open Source Bulk Archive Extraction Tool
.DESCRIPTION
    A comprehensive recreation of ExtractNow built in PowerShell WPF.
    Features: System tray, context menu integration, full settings panel,
    directory monitoring, password cycling, nested extraction, verbose
    file-by-file output, progress bar, column sorting, row coloring,
    virtualized ListView, and more.
.AUTHOR
    SysAdminDoc
.VERSION
    2.1.0
.LICENSE
    MIT
#>

#Requires -Version 5.1
param(
    [string[]]$FilesToExtract,
    [Alias('target')][string]$TargetPath,
    [switch]$minimize,
    [switch]$minimizetotray
)

# =====================================================================
# STA Check & Assembly Loading
# =====================================================================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
try { Add-Type -AssemblyName Microsoft.VisualBasic } catch {}

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $sp = $MyInvocation.MyCommand.Definition
    $al = @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',"`"$sp`"")
    if ($TargetPath) { $al += '-target'; $al += "`"$TargetPath`"" }
    if ($minimize) { $al += '-minimize' }
    if ($minimizetotray) { $al += '-minimizetotray' }
    if ($FilesToExtract) { $al += $FilesToExtract }
    Start-Process powershell.exe -ArgumentList $al
    exit
}

# =====================================================================
# Globals & Constants
# =====================================================================
$script:AppName    = "ExtractorX"
$script:AppVersion = "2.1.0"
$script:AppDataDir = Join-Path $env:APPDATA $script:AppName
$script:ConfigPath = Join-Path $script:AppDataDir "config.json"
$script:PasswordFile = Join-Path $script:AppDataDir "passwords.dat"
$script:LogDir     = Join-Path $script:AppDataDir "logs"
$script:ScriptPath = $PSCommandPath
if (-not $script:ScriptPath) { $script:ScriptPath = $MyInvocation.MyCommand.Definition }
$script:7zPath     = $null

$script:ArchiveExtensions = @(
    '.zip','.7z','.rar',
    '.tar','.gz','.gzip','.tgz','.bz2','.bzip2','.tbz2','.tbz',
    '.xz','.txz','.lzma','.tlz','.lz','.zst','.zstd','.z',
    '.iso','.cab','.arj','.lzh','.lha','.wim','.001',
    '.cpio','.rpm','.deb'
)

$script:MagicBytes = @{
    '504B0304'='zip';'504B0506'='zip';'504B0708'='zip'
    '377ABCAF271C'='7z';'526172211A0700'='rar';'526172211A07'='rar'
    '1F8B'='gz';'425A68'='bz2';'FD377A585A00'='xz';'4D534346'='cab'
    '7573746172'='tar';'4344303031'='iso'
}

# =====================================================================
# Configuration
# =====================================================================
$script:DefaultConfig = @{
    # Destination
    OutputPath             = '{ArchiveFolder}\{ArchiveName}'
    OverwriteMode          = 'Always'
    # Process - Post Actions
    PostAction             = 'None'
    PostActionFolder       = ''
    DeleteAfterExtract     = $false
    OpenDestAfterExtract   = $false
    # Process - Archive Recursion
    NestedExtraction       = $true
    NestedMaxDepth         = 5
    NestedApplyPostAction  = $false
    # Process - Cleanup
    RemoveDuplicateFolder  = $true
    RenameSingleFile       = $false
    DeleteBrokenFiles      = $false
    # Process - Batch Complete
    CloseOnComplete        = $false
    CloseOnCompleteAlways  = $false
    ClearListOnComplete    = $false
    # General - Window
    AlwaysOnTop            = $false
    MinimizeToTray         = $true
    LogHistory             = $true
    AutoSwitchToHistory    = $true
    DeepArchiveDetection   = $true
    # Drag & Drop
    AutoExtractOnDrop      = $false
    DragDropFilterType     = 'None'
    DragDropFilterMask     = ''
    # Passwords
    UsePasswordList        = $true
    PromptOnExhaustion     = $false
    AssumeOnePassword      = $true
    PasswordTimeout        = 45
    # Files / Exclusions
    FileExclusions         = 'Thumbs.db;desktop.ini;.DS_Store'
    # Monitor
    WatchFolders           = @()
    WatchAutoExtract       = $true
    # Explorer / Context Menu
    CtxEnabled             = $false
    CtxGrouped             = $true
    CtxExtractHere         = $true
    CtxExtractToFolder     = $true
    CtxEnqueue             = $true
    CtxSearchArchives      = $true
    # File Associations
    FileAssociations       = @()
    # Advanced
    ThreadPriority         = 'Normal'
    SoundsEnabled          = $true
    # External Processors
    ExternalProcessors     = @()
    # Window state
    WindowWidth            = 1100
    WindowHeight           = 750
    WindowLeft             = -1
    WindowTop              = -1
}

$script:UIQueue = [System.Collections.Concurrent.ConcurrentQueue[hashtable]]::new()

function Send-UIMessage {
    param([string]$Type, [hashtable]$Data = @{})
    $script:UIQueue.Enqueue((@{ Type = $Type } + $Data))
}

# =====================================================================
# Helper Functions
# =====================================================================
function Initialize-AppDirectories {
    @($script:AppDataDir, $script:LogDir) | ForEach-Object {
        if (-not (Test-Path $_)) { New-Item -Path $_ -ItemType Directory -Force | Out-Null }
    }
}

function Load-Config {
    if (Test-Path $script:ConfigPath) {
        try {
            $json = Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
            $config = @{}
            $script:DefaultConfig.Keys | ForEach-Object {
                if ($null -ne $json.$_) { $config[$_] = $json.$_ } else { $config[$_] = $script:DefaultConfig[$_] }
            }
            return $config
        } catch { return $script:DefaultConfig.Clone() }
    }
    return $script:DefaultConfig.Clone()
}

function Save-Config { param([hashtable]$Config); try { $Config | ConvertTo-Json -Depth 5 | Set-Content $script:ConfigPath -Force } catch {} }

function Find-7Zip {
    $locs = @(
        (Join-Path $script:AppDataDir "7z.exe"),
        "C:\Program Files\7-Zip\7z.exe","C:\Program Files (x86)\7-Zip\7z.exe",
        (Join-Path $env:LOCALAPPDATA "Programs\7-Zip\7z.exe")
    )
    $pe = Get-Command 7z.exe -ErrorAction SilentlyContinue
    if ($pe) { $locs = @($pe.Source) + $locs }
    foreach ($l in $locs) { if (Test-Path $l) { return $l } }
    return $null
}

function Get-FileSizeString {
    param([long]$Size)
    if ($Size -ge 1GB) { return "{0:N2} GB" -f ($Size/1GB) }
    if ($Size -ge 1MB) { return "{0:N1} MB" -f ($Size/1MB) }
    if ($Size -ge 1KB) { return "{0:N0} KB" -f ($Size/1KB) }
    return "$Size B"
}

function Test-IsArchive { param([string]$Path); return ([IO.Path]::GetExtension($Path).ToLower() -in $script:ArchiveExtensions) }

function Test-MagicBytes {
    param([string]$FilePath)
    try {
        $s = [IO.File]::OpenRead($FilePath)
        try { $buf = New-Object byte[] 16; $r = $s.Read($buf,0,16)
            if ($r -lt 2) { return $false }
            $hex = ($buf[0..($r-1)] | ForEach-Object { $_.ToString('X2') }) -join ''
            foreach ($sig in $script:MagicBytes.Keys) { if ($hex.StartsWith($sig)) { return $true } }
        } finally { $s.Dispose() }
    } catch {}
    return $false
}

# Multi-volume archive detection - returns $true if file is a NON-FIRST volume (should be skipped)
function Test-IsMultiVolumePart {
    param([string]$FileName)
    # Modern RAR: file.part2.rar, file.part03.rar, etc. (skip anything > part1)
    if ($FileName -match '\.part(\d+)\.rar$') { return ([int]$Matches[1] -gt 1) }
    # Old RAR continuation: .r00, .r01, .r99, .s00, .s01 (always non-first)
    if ($FileName -match '\.[rs]\d{2}$') { return $true }
    # Split volumes: .7z.002, .zip.003, .tar.002, etc. (non-.001 already excluded by ext list, but guard for magic bytes)
    if ($FileName -match '\.\w+\.(\d{3})$' -and $Matches[1] -ne '001') { return $true }
    return $false
}

function Resolve-OutputPath {
    param([string]$Template, [string]$ArchivePath)
    $dir = [IO.Path]::GetDirectoryName($ArchivePath)
    $name = [IO.Path]::GetFileNameWithoutExtension($ArchivePath)
    if ($name -match '\.tar$') { $name = [IO.Path]::GetFileNameWithoutExtension($name) }
    $ext = [IO.Path]::GetExtension($ArchivePath).TrimStart('.')
    $fname = [IO.Path]::GetFileName($ArchivePath)
    $pname = Split-Path $dir -Leaf
    $r = $Template
    $r = $r -replace '\{ArchiveFolder\}', $dir
    $r = $r -replace '\{ArchiveName\}', $name
    $r = $r -replace '\{ArchiveExtension\}', $ext
    $r = $r -replace '\{ArchiveFileName\}', $fname
    $r = $r -replace '\{ArchiveFolderName\}', $pname
    $r = $r -replace '\{ArchivePath\}', $ArchivePath
    $r = $r -replace '\{Guid\}', ([Guid]::NewGuid().ToString('N').Substring(0,8))
    $r = $r -replace '\{Desktop\}', [Environment]::GetFolderPath('Desktop')
    $r = $r -replace '\{UserProfile\}', $env:USERPROFILE
    $r = $r -replace '\{Program Files\}', $env:ProgramFiles
    $r = $r -replace '\{Windows\}', $env:windir
    $r = $r -replace '\{Date\}', (Get-Date -Format 'yyyyMMdd')
    $r = $r -replace '\{Time\}', (Get-Date -Format 'HHmmss')
    if ($r -match '\{Env:([^}]+)\}') {
        $r = [regex]::Replace($r, '\{Env:([^}]+)\}', { param($m) [Environment]::GetEnvironmentVariable($m.Groups[1].Value) })
    }
    if ($r -match '\{ArchiveNameUnique\}') {
        $b = $r -replace '\{ArchiveNameUnique\}', $name
        if (Test-Path $b) { $c=1; do { $t=$r -replace '\{ArchiveNameUnique\}',"$name ($c)"; $c++ } while (Test-Path $t); $r=$t } else { $r=$b }
    }
    return $r
}

# =====================================================================
# Password Management (DPAPI encrypted)
# =====================================================================
function Load-Passwords {
    if (Test-Path $script:PasswordFile) {
        try {
            $lines = Get-Content $script:PasswordFile -Force
            $pws = @()
            foreach ($ln in $lines) {
                if ([string]::IsNullOrWhiteSpace($ln)) { continue }
                try {
                    $sec = $ln | ConvertTo-SecureString -ErrorAction Stop
                    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
                    $plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
                    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
                    $pws += $plain
                } catch { $pws += $ln }
            }
            return $pws
        } catch { return @() }
    }
    return @()
}

function Save-Passwords {
    param([string[]]$Passwords)
    try {
        $enc = $Passwords | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
            $_ | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString
        }
        $enc | Set-Content $script:PasswordFile -Force
    } catch {}
}

# =====================================================================
# Context Menu Management (multiple entries)
# =====================================================================
function Install-ContextMenuEntries {
    param([hashtable]$Cfg)
    $sp = $script:ScriptPath; if (-not $sp) { return }
    $menuName = "ExtractorX"
    $grouped = $Cfg.CtxGrouped

    # Remove old entries first
    Uninstall-ContextMenuEntries

    foreach ($ext in $script:ArchiveExtensions) {
        $base = "HKCU:\Software\Classes\SystemFileAssociations\$ext\shell"
        if ($grouped) {
            $gp = "$base\$menuName"
            try {
                New-Item -Path $gp -Force | Out-Null
                Set-ItemProperty -Path $gp -Name "(Default)" -Value "ExtractorX" -Force
                Set-ItemProperty -Path $gp -Name "SubCommands" -Value "" -Force
                Set-ItemProperty -Path $gp -Name "MUIVerb" -Value "ExtractorX" -Force
                $shell = "$gp\shell"
                if ($Cfg.CtxExtractHere) {
                    New-Item "$shell\01_here\command" -Force -Value "powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File `"$sp`" -target `"{ArchiveFolder}`" `"%1`"" | Out-Null
                    Set-ItemProperty "$shell\01_here" -Name "(Default)" -Value "Extract here" -Force
                }
                if ($Cfg.CtxExtractToFolder) {
                    New-Item "$shell\02_folder\command" -Force -Value "powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File `"$sp`" `"%1`"" | Out-Null
                    Set-ItemProperty "$shell\02_folder" -Name "(Default)" -Value "Extract to folder" -Force
                }
                if ($Cfg.CtxEnqueue) {
                    New-Item "$shell\03_enqueue\command" -Force -Value "powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File `"$sp`" `"%1`"" | Out-Null
                    Set-ItemProperty "$shell\03_enqueue" -Name "(Default)" -Value "Add to ExtractorX" -Force
                }
            } catch {}
        } else {
            if ($Cfg.CtxExtractHere) {
                try {
                    New-Item "$base\${menuName}_Here\command" -Force -Value "powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File `"$sp`" -target `"{ArchiveFolder}`" `"%1`"" | Out-Null
                    Set-ItemProperty "$base\${menuName}_Here" -Name "(Default)" -Value "ExtractorX: Extract here" -Force
                } catch {}
            }
            if ($Cfg.CtxExtractToFolder) {
                try {
                    New-Item "$base\${menuName}_Folder\command" -Force -Value "powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File `"$sp`" `"%1`"" | Out-Null
                    Set-ItemProperty "$base\${menuName}_Folder" -Name "(Default)" -Value "ExtractorX: Extract to folder" -Force
                } catch {}
            }
        }
    }
    # Directory entry: Search for archives
    if ($Cfg.CtxSearchArchives) {
        try {
            $dp = "HKCU:\Software\Classes\Directory\shell\$menuName"
            New-Item "$dp\command" -Force -Value "powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File `"$sp`" `"%1`"" | Out-Null
            Set-ItemProperty $dp -Name "(Default)" -Value "Search for archives with ExtractorX" -Force
        } catch {}
    }
}

function Uninstall-ContextMenuEntries {
    $menuName = "ExtractorX"
    foreach ($ext in $script:ArchiveExtensions) {
        $base = "HKCU:\Software\Classes\SystemFileAssociations\$ext\shell"
        @("$base\$menuName","$base\${menuName}_Here","$base\${menuName}_Folder","$base\${menuName}_Enqueue") | ForEach-Object {
            if (Test-Path $_) { Remove-Item $_ -Recurse -Force -ErrorAction SilentlyContinue }
        }
    }
    $dp = "HKCU:\Software\Classes\Directory\shell\$menuName"
    if (Test-Path $dp) { Remove-Item $dp -Recurse -Force -ErrorAction SilentlyContinue }
    # Background entry
    $bp = "HKCU:\Software\Classes\Directory\Background\shell\$menuName"
    if (Test-Path $bp) { Remove-Item $bp -Recurse -Force -ErrorAction SilentlyContinue }
}

# =====================================================================
# Initialize
# =====================================================================
Initialize-AppDirectories
$script:Config = Load-Config
$script:7zPath = Find-7Zip
$script:Passwords = Load-Passwords

# Override output path if /target specified
if ($TargetPath) { $script:Config.OutputPath = $TargetPath }


# =====================================================================
# WPF XAML - Main Window
# =====================================================================
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="ExtractorX v2.1.0" Width="1100" Height="750" MinWidth="850" MinHeight="500"
    WindowStartupLocation="CenterScreen" Background="#FF0A0A12" Foreground="#FFD4D4E0" AllowDrop="True"
    WindowStyle="None" AllowsTransparency="False"
    FontFamily="Segoe UI" TextOptions.TextRenderingMode="ClearType" TextOptions.TextFormattingMode="Display">
    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="40" ResizeBorderThickness="6" GlassFrameThickness="0" CornerRadius="0" UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>
    <Window.Resources>
        <!-- Color System -->
        <SolidColorBrush x:Key="Acc" Color="#FF00B4D8"/>
        <SolidColorBrush x:Key="AccH" Color="#FF33C5E3"/>
        <SolidColorBrush x:Key="AccP" Color="#FF0090AA"/>
        <SolidColorBrush x:Key="Ok" Color="#FF6BCB77"/>
        <SolidColorBrush x:Key="Err" Color="#FFFF6B6B"/>
        <SolidColorBrush x:Key="Wrn" Color="#FFFFD93D"/>
        <SolidColorBrush x:Key="Bg" Color="#FF0A0A12"/>
        <SolidColorBrush x:Key="Sf" Color="#FF101018"/>
        <SolidColorBrush x:Key="Sf2" Color="#FF16161F"/>
        <SolidColorBrush x:Key="Sf3" Color="#FF1E1E2A"/>
        <SolidColorBrush x:Key="Sf4" Color="#FF262636"/>
        <SolidColorBrush x:Key="Bd" Color="#FF2A2A3C"/>
        <SolidColorBrush x:Key="BdH" Color="#FF3A3A50"/>
        <SolidColorBrush x:Key="Tx" Color="#FFD4D4E0"/>
        <SolidColorBrush x:Key="Td" Color="#FF6E6E84"/>
        <SolidColorBrush x:Key="TxB" Color="#FFEEEEF4"/>

        <!-- Dark ScrollBar -->
        <Style x:Key="ScrollBarThumb" TargetType="Thumb">
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Thumb">
                <Border x:Name="tb" Background="#FF3A3A4E" CornerRadius="4" Margin="2"/>
                <ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="tb" Property="Background" Value="#FF50506A"/></Trigger>
                <Trigger Property="IsDragging" Value="True"><Setter TargetName="tb" Property="Background" Value="{StaticResource Acc}"/></Trigger></ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="Transparent"/><Setter Property="Width" Value="10"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ScrollBar">
                <Border Background="#08FFFFFF" CornerRadius="5"><Track x:Name="PART_Track" IsDirectionReversed="True">
                    <Track.Thumb><Thumb Style="{StaticResource ScrollBarThumb}"/></Track.Thumb></Track></Border>
            </ControlTemplate></Setter.Value></Setter>
            <Style.Triggers><Trigger Property="Orientation" Value="Horizontal">
                <Setter Property="Height" Value="10"/><Setter Property="Width" Value="Auto"/>
            </Trigger></Style.Triggers>
        </Style>

        <!-- Tooltip -->
        <Style TargetType="ToolTip">
            <Setter Property="Background" Value="#FF1A1A26"/><Setter Property="Foreground" Value="{StaticResource Tx}"/>
            <Setter Property="BorderBrush" Value="{StaticResource Bd}"/><Setter Property="Padding" Value="10,6"/>
            <Setter Property="FontSize" Value="11.5"/>
        </Style>

        <!-- Buttons -->
        <Style x:Key="Btn" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource Sf3}"/><Setter Property="Foreground" Value="{StaticResource Tx}"/>
            <Setter Property="BorderThickness" Value="1"/><Setter Property="BorderBrush" Value="{StaticResource Bd}"/>
            <Setter Property="Padding" Value="14,7"/><Setter Property="FontSize" Value="12"/><Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
                <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
                <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="{StaticResource Sf4}"/><Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource BdH}"/></Trigger>
                    <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Background" Value="#FF2E2E42"/></Trigger>
                    <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.35"/></Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>
        <Style x:Key="AccBtn" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource Acc}"/><Setter Property="Foreground" Value="#FF0A0A12"/><Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="20,8"/><Setter Property="FontSize" Value="13"/><Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
                <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
                <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="{StaticResource AccH}"/></Trigger>
                    <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Background" Value="{StaticResource AccP}"/></Trigger>
                    <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.35"/></Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>
        <Style x:Key="DanBtn" TargetType="Button">
            <Setter Property="Background" Value="#FF8B2020"/><Setter Property="Foreground" Value="White"/><Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="14,7"/><Setter Property="FontSize" Value="12"/><Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
                <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
                <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#FFA52828"/></Trigger>
                    <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Background" Value="#FF6E1818"/></Trigger>
                    <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.35"/></Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>

        <!-- TextBox -->
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#FF0E0E16"/><Setter Property="Foreground" Value="{StaticResource Tx}"/>
            <Setter Property="BorderBrush" Value="{StaticResource Bd}"/><Setter Property="Padding" Value="10,7"/>
            <Setter Property="CaretBrush" Value="{StaticResource Acc}"/><Setter Property="FontSize" Value="12"/>
            <Setter Property="SelectionBrush" Value="{StaticResource AccP}"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="TextBox">
                <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="6" Padding="{TemplateBinding Padding}">
                    <ScrollViewer x:Name="PART_ContentHost" Margin="0"/></Border>
                <ControlTemplate.Triggers><Trigger Property="IsFocused" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource Acc}"/></Trigger></ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>

        <!-- CheckBox -->
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource Tx}"/><Setter Property="FontSize" Value="12"/><Setter Property="Margin" Value="0,5"/><Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="CheckBox">
                <StackPanel Orientation="Horizontal">
                    <Border x:Name="box" Width="18" Height="18" CornerRadius="4" BorderThickness="1.5" BorderBrush="{StaticResource Bd}" Background="#FF0E0E16" VerticalAlignment="Center" Margin="0,0,8,0">
                        <TextBlock x:Name="check" Text="&#x2713;" FontSize="12" FontWeight="Bold" Foreground="{StaticResource Acc}" HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed"/></Border>
                    <ContentPresenter VerticalAlignment="Center"/></StackPanel>
                <ControlTemplate.Triggers>
                    <Trigger Property="IsChecked" Value="True"><Setter TargetName="check" Property="Visibility" Value="Visible"/><Setter TargetName="box" Property="BorderBrush" Value="{StaticResource Acc}"/><Setter TargetName="box" Property="Background" Value="#FF0D2A33"/></Trigger>
                    <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="box" Property="BorderBrush" Value="{StaticResource BdH}"/></Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>

        <!-- ListBox -->
        <Style TargetType="ListBox">
            <Setter Property="Background" Value="#FF0E0E16"/><Setter Property="BorderBrush" Value="{StaticResource Bd}"/><Setter Property="Foreground" Value="{StaticResource Tx}"/>
            <Setter Property="FontSize" Value="12"/><Setter Property="Padding" Value="4"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ListBox">
                <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="6" Padding="{TemplateBinding Padding}">
                    <ScrollViewer Focusable="False"><ItemsPresenter/></ScrollViewer></Border>
            </ControlTemplate></Setter.Value></Setter>
        </Style>

        <!-- ComboBox -->
        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="#FF0E0E16"/><Setter Property="Foreground" Value="{StaticResource Tx}"/>
            <Setter Property="BorderBrush" Value="{StaticResource Bd}"/><Setter Property="FontSize" Value="12"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ComboBox">
                <Grid><ToggleButton x:Name="ToggleButton" Focusable="False" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press">
                    <ToggleButton.Template><ControlTemplate TargetType="ToggleButton">
                        <Border Background="#FF0E0E16" BorderBrush="{StaticResource Bd}" BorderThickness="1" CornerRadius="6" Padding="8,7">
                            <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="20"/></Grid.ColumnDefinitions>
                                <TextBlock Grid.Column="1" Text="&#x25BC;" FontSize="9" Foreground="{StaticResource Td}" HorizontalAlignment="Center" VerticalAlignment="Center"/></Grid></Border>
                    </ControlTemplate></ToggleButton.Template></ToggleButton>
                <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" Margin="10,7,28,6" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                <Popup x:Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                    <Border Background="{StaticResource Sf3}" BorderBrush="{StaticResource BdH}" BorderThickness="1" CornerRadius="6" Margin="0,2,0,0" MaxHeight="300" MinWidth="{TemplateBinding ActualWidth}">
                        <ScrollViewer><ItemsPresenter/></ScrollViewer></Border></Popup></Grid>
            </ControlTemplate></Setter.Value></Setter>
        </Style>
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="Transparent"/><Setter Property="Foreground" Value="{StaticResource Tx}"/>
            <Setter Property="Padding" Value="10,7"/><Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ComboBoxItem">
                <Border x:Name="bd" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                    <ContentPresenter/></Border>
                <ControlTemplate.Triggers><Trigger Property="IsHighlighted" Value="True"><Setter TargetName="bd" Property="Background" Value="{StaticResource Sf4}"/></Trigger></ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>

        <!-- ListView -->
        <Style TargetType="ListView">
            <Setter Property="Background" Value="#FF0D0D15"/><Setter Property="BorderBrush" Value="{StaticResource Bd}"/><Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Foreground" Value="{StaticResource Tx}"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ListView">
                <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" SnapsToDevicePixels="True">
                    <ScrollViewer Style="{DynamicResource {x:Static GridView.GridViewScrollViewerStyleKey}}" Padding="0">
                        <ItemsPresenter/></ScrollViewer></Border>
            </ControlTemplate></Setter.Value></Setter>
        </Style>
        <Style TargetType="ListViewItem">
            <Setter Property="Foreground" Value="{StaticResource Tx}"/><Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="4,5"/><Setter Property="HorizontalContentAlignment" Value="Stretch"/><Setter Property="FontSize" Value="12"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ListViewItem">
                <Border x:Name="bd" Background="Transparent" Padding="{TemplateBinding Padding}" BorderThickness="0,0,0,1" BorderBrush="#08FFFFFF" SnapsToDevicePixels="True">
                    <GridViewRowPresenter VerticalAlignment="Center"/></Border>
                <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#FF151522"/></Trigger>
                    <Trigger Property="IsSelected" Value="True"><Setter TargetName="bd" Property="Background" Value="#FF141430"/><Setter TargetName="bd" Property="BorderBrush" Value="#20FFFFFF"/></Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>
        <Style TargetType="GridViewColumnHeader">
            <Setter Property="Background" Value="{StaticResource Sf2}"/><Setter Property="Foreground" Value="{StaticResource Td}"/>
            <Setter Property="BorderThickness" Value="0,0,1,0"/><Setter Property="BorderBrush" Value="#15FFFFFF"/>
            <Setter Property="Padding" Value="10,8"/><Setter Property="FontSize" Value="11"/><Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="GridViewColumnHeader">
                <Border x:Name="hbd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
                    <TextBlock Text="{TemplateBinding Content}" Foreground="{TemplateBinding Foreground}" FontSize="{TemplateBinding FontSize}" FontWeight="{TemplateBinding FontWeight}" TextAlignment="Left"/>
                </Border>
                <ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="hbd" Property="Background" Value="{StaticResource Sf3}"/><Setter Property="Foreground" Value="{StaticResource Tx}"/></Trigger></ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>

        <!-- TabControl / TabItem -->
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/><Setter Property="BorderThickness" Value="0"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="{StaticResource Td}"/><Setter Property="FontSize" Value="12"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="TabItem">
                <Border x:Name="bd" Padding="14,9" Margin="0,0,2,0" CornerRadius="6,6,0,0" Background="Transparent" Cursor="Hand">
                    <ContentPresenter ContentSource="Header"/></Border>
                <ControlTemplate.Triggers><Trigger Property="IsSelected" Value="True"><Setter TargetName="bd" Property="Background" Value="{StaticResource Sf3}"/><Setter Property="Foreground" Value="{StaticResource Acc}"/><Setter Property="FontWeight" Value="SemiBold"/></Trigger>
                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#FF181825"/></Trigger></ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>

        <!-- Context Menu -->
        <Style TargetType="ContextMenu">
            <Setter Property="Background" Value="#FF181822"/><Setter Property="BorderBrush" Value="{StaticResource BdH}"/>
            <Setter Property="Foreground" Value="{StaticResource Tx}"/><Setter Property="Padding" Value="4,6"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ContextMenu">
                <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="8" Padding="{TemplateBinding Padding}">
                    <StackPanel IsItemsHost="True"/></Border>
            </ControlTemplate></Setter.Value></Setter>
        </Style>
        <Style TargetType="MenuItem">
            <Setter Property="Foreground" Value="{StaticResource Tx}"/><Setter Property="FontSize" Value="12"/><Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="MenuItem">
                <Border x:Name="mbd" Background="Transparent" Padding="12,7" CornerRadius="4">
                    <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                        <TextBlock x:Name="chk" Width="16" Text="&#x2713;" Foreground="{StaticResource Acc}" Visibility="Collapsed" VerticalAlignment="Center"/>
                        <ContentPresenter Grid.Column="1" ContentSource="Header" VerticalAlignment="Center"/></Grid></Border>
                <ControlTemplate.Triggers>
                    <Trigger Property="IsHighlighted" Value="True"><Setter TargetName="mbd" Property="Background" Value="{StaticResource Sf4}"/></Trigger>
                    <Trigger Property="IsChecked" Value="True"><Setter TargetName="chk" Property="Visibility" Value="Visible"/></Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>
        <Style TargetType="Separator">
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Separator">
                <Border Height="1" Background="{StaticResource Bd}" Margin="8,4"/>
            </ControlTemplate></Setter.Value></Setter>
        </Style>
        <!-- Window Control Buttons -->
        <Style x:Key="WinBtn" TargetType="Button">
            <Setter Property="Width" Value="46"/><Setter Property="Height" Value="32"/><Setter Property="FontSize" Value="10"/>
            <Setter Property="Foreground" Value="{StaticResource Td}"/><Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/><Setter Property="Cursor" Value="Hand"/>
            <Setter Property="WindowChrome.IsHitTestVisibleInChrome" Value="True"/>
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
                <Border x:Name="bd" Background="{TemplateBinding Background}" SnapsToDevicePixels="True">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
                <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#FF2A2A3C"/></Trigger>
                    <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Background" Value="#FF3A3A50"/></Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>
        <Style x:Key="WinCloseBtn" TargetType="Button" BasedOn="{StaticResource WinBtn}">
            <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
                <Border x:Name="bd" Background="Transparent" SnapsToDevicePixels="True">
                    <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
                <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#FFE81123"/><Setter Property="Foreground" Value="White"/></Trigger>
                    <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Background" Value="#FFBF0F1D"/></Trigger>
                </ControlTemplate.Triggers>
            </ControlTemplate></Setter.Value></Setter>
        </Style>
    </Window.Resources>
    <!-- Outer border for window edge -->
    <Border BorderBrush="#FF1A1A28" BorderThickness="1">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Custom Title Bar -->
        <Border Background="#FF08080E" Padding="0" BorderBrush="{StaticResource Bd}" BorderThickness="0,0,0,1">
            <Grid Height="40">
                <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <!-- App Branding -->
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="14,0,0,0">
                    <TextBlock Text="EXTRACTOR" FontSize="12" FontWeight="Bold" Foreground="#FF8888A0" VerticalAlignment="Center"/>
                    <TextBlock Text="X" FontSize="12" FontWeight="Bold" Foreground="{StaticResource Acc}" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <TextBlock Text="&#x2014;" Foreground="#FF2A2A3C" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <TextBlock x:Name="TitleBarText" Text="Ready" FontSize="11" Foreground="#FF505068" VerticalAlignment="Center"/>
                </StackPanel>
                <!-- Window Controls -->
                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Top">
                    <Button x:Name="BtnMinimize" Style="{StaticResource WinBtn}">
                        <TextBlock Text="&#x2015;" FontSize="10" Foreground="{StaticResource Td}"/></Button>
                    <Button x:Name="BtnMaximize" Style="{StaticResource WinBtn}">
                        <TextBlock x:Name="MaximizeIcon" Text="&#x25A1;" FontSize="12" Foreground="{StaticResource Td}"/></Button>
                    <Button x:Name="BtnClose" Style="{StaticResource WinCloseBtn}">
                        <TextBlock Text="&#x2715;" FontSize="11" Foreground="{StaticResource Td}"/></Button>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Toolbar -->
        <Border Grid.Row="1" Background="{StaticResource Sf}" Padding="14,7" BorderBrush="{StaticResource Bd}" BorderThickness="0,0,0,1">
            <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                <WrapPanel VerticalAlignment="Center">
                    <Button x:Name="BtnAddFiles" Content="+ Files" Style="{StaticResource Btn}" Margin="0,0,5,0" ToolTip="Add archive files"/>
                    <Button x:Name="BtnAddFolder" Content="+ Folder" Style="{StaticResource Btn}" Margin="0,0,5,0" ToolTip="Scan folder for archives"/>
                    <Border Width="1" Height="22" Background="{StaticResource Bd}" Margin="8,0"/>
                    <Button x:Name="BtnExtractAll" Content="Extract All" Style="{StaticResource AccBtn}" Margin="0,0,5,0"/>
                    <Button x:Name="BtnStopAll" Content="Stop" Style="{StaticResource DanBtn}" Margin="0,0,5,0" IsEnabled="False"/>
                    <Border Width="1" Height="22" Background="{StaticResource Bd}" Margin="8,0"/>
                    <Button x:Name="BtnRemoveSelected" Content="Remove" Style="{StaticResource Btn}" Margin="0,0,5,0"/>
                    <Button x:Name="BtnClearQueue" Content="Clear" Style="{StaticResource Btn}" Margin="0,0,5,0"/>
                </WrapPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <CheckBox x:Name="ChkDeleteAfter" Content="Delete after extract" Foreground="{StaticResource Err}" Margin="0,0,12,0" VerticalAlignment="Center" ToolTip="Permanently delete archives after successful extraction"/>
                    <Border Width="1" Height="22" Background="{StaticResource Bd}" Margin="0,0,10,0"/>
                    <Button x:Name="BtnSettings" Style="{StaticResource Btn}" Padding="10,7" Margin="0,0,4,0" ToolTip="Settings" WindowChrome.IsHitTestVisibleInChrome="True">
                        <TextBlock Text="&#x2699;" FontSize="14" Foreground="{StaticResource Td}"/></Button>
                    <Button x:Name="BtnAbout" Style="{StaticResource Btn}" Padding="10,7" ToolTip="About ExtractorX" WindowChrome.IsHitTestVisibleInChrome="True">
                        <TextBlock Text="&#x2139;" FontSize="13" Foreground="{StaticResource Td}"/></Button>
                </StackPanel>
            </Grid>
        </Border>

        <!-- Main Content: Queue + Log -->
        <Grid Grid.Row="2" Margin="10,8,10,0">
            <Grid.RowDefinitions>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition x:Name="LogRowDef" Height="200"/>
            </Grid.RowDefinitions>

            <!-- Drop overlay -->
            <Border x:Name="DropOverlay" Grid.RowSpan="3" Background="#CC0A0A12" Visibility="Collapsed" Panel.ZIndex="99" CornerRadius="12">
                <Border Background="{StaticResource Sf2}" CornerRadius="16" Padding="50" HorizontalAlignment="Center" VerticalAlignment="Center" BorderBrush="{StaticResource Acc}" BorderThickness="2">
                    <StackPanel>
                        <TextBlock Text="&#x2B73;" FontSize="40" Foreground="{StaticResource Acc}" HorizontalAlignment="Center" Margin="0,0,0,10"/>
                        <TextBlock Text="Drop Archives Here" FontSize="20" FontWeight="SemiBold" Foreground="{StaticResource TxB}" HorizontalAlignment="Center"/>
                        <TextBlock Text="Files or folders containing archives" FontSize="12" Foreground="{StaticResource Td}" HorizontalAlignment="Center" Margin="0,6,0,0"/>
                    </StackPanel>
                </Border>
            </Border>

            <!-- Queue -->
            <ListView x:Name="QueueList" SelectionMode="Extended" FontSize="12"
                VirtualizingStackPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling"
                VirtualizingStackPanel.ScrollUnit="Pixel" ScrollViewer.IsDeferredScrollingEnabled="True">
                <ListView.View>
                    <GridView>
                        <GridViewColumn Header="File" Width="280" DisplayMemberBinding="{Binding Filename}"/>
                        <GridViewColumn Header="Directory" Width="260" DisplayMemberBinding="{Binding Directory}"/>
                        <GridViewColumn Header="Size" Width="80" DisplayMemberBinding="{Binding Size}"/>
                        <GridViewColumn Header="Type" Width="60" DisplayMemberBinding="{Binding Extension}"/>
                        <GridViewColumn Header="Status" Width="150" DisplayMemberBinding="{Binding Status}"/>
                        <GridViewColumn Header="Time" Width="70" DisplayMemberBinding="{Binding Elapsed}"/>
                    </GridView>
                </ListView.View>
                <ListView.ContextMenu>
                    <ContextMenu>
                        <MenuItem x:Name="CtxAddFiles" Header="Add Archives..."/>
                        <MenuItem x:Name="CtxSearchFolder" Header="Search for Archives..."/>
                        <Separator/>
                        <MenuItem x:Name="CtxExtractUnextracted" Header="Extract Only Unextracted"/>
                        <MenuItem x:Name="CtxTestOnly" Header="Test Decompression Only" IsCheckable="True"/>
                        <Separator/>
                        <MenuItem x:Name="CtxOpenFolder" Header="Open Containing Folder"/>
                        <MenuItem x:Name="CtxOpenOutput" Header="Open Output Folder"/>
                        <MenuItem x:Name="CtxRemoveItem" Header="Remove Selected"/>
                    </ContextMenu>
                </ListView.ContextMenu>
            </ListView>

            <!-- Splitter -->
            <GridSplitter Grid.Row="1" Height="6" HorizontalAlignment="Stretch" Background="Transparent" ResizeBehavior="PreviousAndNext" Cursor="SizeNS">
                <GridSplitter.Template><ControlTemplate TargetType="GridSplitter">
                    <Border Background="Transparent"><Border Height="2" CornerRadius="1" Background="{StaticResource Bd}" VerticalAlignment="Center" HorizontalAlignment="Center" Width="60"/></Border>
                </ControlTemplate></GridSplitter.Template>
            </GridSplitter>

            <!-- Log/History Panel -->
            <Border Grid.Row="2" Background="{StaticResource Sf}" BorderBrush="{StaticResource Bd}" BorderThickness="1" CornerRadius="8">
                <Grid>
                    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                    <Border Background="{StaticResource Sf2}" Padding="10,6" CornerRadius="8,8,0,0">
                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                <TextBlock Text="LOG" FontSize="10" FontWeight="Bold" Foreground="{StaticResource Td}" VerticalAlignment="Center" Margin="0,0,8,0"/>
                                <Border Width="1" Height="14" Background="{StaticResource Bd}"/>
                            </StackPanel>
                            <WrapPanel Grid.Column="1">
                                <Button x:Name="BtnToggleLog" Content="Hide" Style="{StaticResource Btn}" Padding="8,3" FontSize="10" Margin="0,0,3,0"/>
                                <Button x:Name="BtnClearLog" Content="Clear" Style="{StaticResource Btn}" Padding="8,3" FontSize="10" Margin="0,0,3,0"/>
                                <Button x:Name="BtnExportLog" Content="Export" Style="{StaticResource Btn}" Padding="8,3" FontSize="10"/>
                            </WrapPanel>
                        </Grid>
                    </Border>
                    <ScrollViewer x:Name="LogScroll" Grid.Row="1" VerticalScrollBarVisibility="Auto" Padding="10,6">
                        <TextBlock x:Name="LogBox" TextWrapping="Wrap" FontFamily="Cascadia Mono, Consolas, Courier New" FontSize="11"/>
                    </ScrollViewer>
                </Grid>
            </Border>
        </Grid>

        <!-- Output Path Bar -->
        <Border Grid.Row="3" Background="{StaticResource Sf}" Padding="12,7" BorderBrush="{StaticResource Bd}" BorderThickness="0,1,0,0">
            <Grid><Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
                <TextBlock Text="OUTPUT" VerticalAlignment="Center" Foreground="{StaticResource Td}" FontSize="10" FontWeight="Bold" Margin="0,0,10,0"/>
                <TextBox x:Name="TxtOutputPath" Grid.Column="1"/>
                <Button x:Name="BtnMacroInsert" Grid.Column="2" Style="{StaticResource Btn}" Padding="10,5" FontSize="11" Margin="6,0,0,0" ToolTip="Insert output path macro">
                    <TextBlock Text="{}{ }" FontFamily="Consolas" Foreground="{StaticResource Td}"/>
                    <Button.ContextMenu>
                        <ContextMenu x:Name="MacroMenu">
                            <MenuItem Header="{}{ArchiveFolder}" Tag="{}{ArchiveFolder}"/>
                            <MenuItem Header="{}{ArchiveName}" Tag="{}{ArchiveName}"/>
                            <MenuItem Header="{}{ArchiveNameUnique}" Tag="{}{ArchiveNameUnique}"/>
                            <MenuItem Header="{}{ArchiveExtension}" Tag="{}{ArchiveExtension}"/>
                            <MenuItem Header="{}{ArchiveFileName}" Tag="{}{ArchiveFileName}"/>
                            <MenuItem Header="{}{ArchiveFolderName}" Tag="{}{ArchiveFolderName}"/>
                            <Separator/>
                            <MenuItem Header="{}{Desktop}" Tag="{}{Desktop}"/>
                            <MenuItem Header="{}{UserProfile}" Tag="{}{UserProfile}"/>
                            <MenuItem Header="{}{Guid}" Tag="{}{Guid}"/>
                            <MenuItem Header="{}{Date}" Tag="{}{Date}"/>
                            <MenuItem Header="{}{Time}" Tag="{}{Time}"/>
                            <MenuItem Header="{}{Env:TEMP}" Tag="{}{Env:TEMP}"/>
                        </ContextMenu>
                    </Button.ContextMenu>
                </Button>
                <Button x:Name="BtnBrowseOutput" Grid.Column="3" Content="..." Style="{StaticResource Btn}" Padding="10,5" FontSize="11" Margin="4,0,0,0"/>
            </Grid>
        </Border>

        <!-- Status Bar -->
        <Border Grid.Row="4" Background="#FF08080E" BorderBrush="{StaticResource Bd}" BorderThickness="0,1,0,0">
            <Grid>
                <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                <!-- Progress bar -->
                <Border Height="2" Background="#FF151520">
                    <Border x:Name="ProgressBar" Background="{StaticResource Acc}" HorizontalAlignment="Left" Width="0"/>
                </Border>
                <Grid Grid.Row="1" Margin="14,6,14,6"><Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                    <TextBlock x:Name="QueueCount" Text="Queue: 0 items" Foreground="{StaticResource Td}" FontSize="11" VerticalAlignment="Center" Margin="0,0,14,0"/>
                    <TextBlock x:Name="SelectionInfo" Grid.Column="1" Foreground="{StaticResource Td}" FontSize="11" VerticalAlignment="Center" Margin="0,0,14,0"/>
                    <TextBlock x:Name="TotalSizeText" Grid.Column="2" Foreground="{StaticResource Td}" FontSize="11" VerticalAlignment="Center" Margin="0,0,14,0"/>
                    <TextBlock x:Name="ProgressText" Grid.Column="3" Foreground="{StaticResource Acc}" FontSize="11" VerticalAlignment="Center" Margin="0,0,14,0"/>
                    <TextBlock x:Name="StatusText" Grid.Column="4" Foreground="{StaticResource Td}" FontSize="11" VerticalAlignment="Center" Margin="0,0,14,0"/>
                    <TextBlock x:Name="SevenZipStatus" Grid.Column="5" FontSize="10" VerticalAlignment="Center" Margin="0,0,14,0"/>
                    <TextBlock x:Name="WatchStatus" Grid.Column="6" Foreground="{StaticResource Wrn}" FontSize="10" VerticalAlignment="Center" Margin="0,0,14,0"/>
                    <TextBlock Grid.Column="7" Text="v2.1.0" Foreground="#FF404058" FontSize="10" VerticalAlignment="Center"/>
                </Grid>
            </Grid>
        </Border>
    </Grid>
    </Border>
</Window>
"@


# =====================================================================
# Build Window & Map Controls
# =====================================================================
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$ui = @{}
@(
    'QueueList','DropOverlay','TxtOutputPath','BtnBrowseOutput','BtnMacroInsert','MacroMenu',
    'BtnAddFiles','BtnAddFolder','BtnRemoveSelected','BtnClearQueue','BtnExtractAll','BtnStopAll',
    'BtnSettings','BtnAbout','ChkDeleteAfter',
    'BtnMinimize','BtnMaximize','BtnClose','MaximizeIcon','TitleBarText',
    'StatusText','SevenZipStatus','QueueCount','ProgressText','WatchStatus',
    'ProgressBar','SelectionInfo','TotalSizeText',
    'LogBox','LogScroll','BtnClearLog','BtnExportLog','BtnToggleLog',
    'CtxAddFiles','CtxSearchFolder','CtxExtractUnextracted','CtxTestOnly','CtxOpenFolder','CtxOpenOutput','CtxRemoveItem'
) | ForEach-Object { $el = $window.FindName($_); if ($el) { $ui[$_] = $el } }

# Thread-safe state
$script:State = [hashtable]::Synchronized(@{
    IsExtracting         = $false
    StopRequested        = $false
    TestOnly             = $false
    WatchersActive       = $false
    Watchers             = [System.Collections.ArrayList]::new()
    SevenZipPath         = $script:7zPath
    Downloading7z        = $false
    IsScanning           = $false
    ScanCancelRequested  = $false
    AutoExtractAfterScan = $false
})

# =====================================================================
# System Tray Icon
# =====================================================================
$script:TrayIcon = New-Object System.Windows.Forms.NotifyIcon
$script:TrayIcon.Text = "ExtractorX v$($script:AppVersion)"
# Create a simple icon programmatically
try {
    $bmp = New-Object System.Drawing.Bitmap(16,16)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::FromArgb(0,180,216))
    $g.FillRectangle([System.Drawing.Brushes]::White, 3, 3, 10, 10)
    $g.FillRectangle([System.Drawing.Brushes]::FromKnownColor('SteelBlue'), 5, 5, 6, 6)
    $g.Dispose()
    $script:TrayIcon.Icon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
    $bmp.Dispose()
} catch {
    $script:TrayIcon.Icon = [System.Drawing.SystemIcons]::Application
}
$script:TrayIcon.Visible = $false

# Tray context menu
$trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
$trayShow = $trayMenu.Items.Add("Show ExtractorX")
$trayShow.Font = New-Object System.Drawing.Font($trayShow.Font, [System.Drawing.FontStyle]::Bold)
$traySep1 = $trayMenu.Items.Add("-")
$trayExtract = $trayMenu.Items.Add("Extract All")
$trayStop = $trayMenu.Items.Add("Stop")
$trayStop.Enabled = $false
$traySep2 = $trayMenu.Items.Add("-")
$trayExit = $trayMenu.Items.Add("Exit")
$script:TrayIcon.ContextMenuStrip = $trayMenu

$trayShow.Add_Click({ $window.Show(); $window.WindowState = 'Normal'; $window.Activate(); $script:TrayIcon.Visible = $false })
$script:TrayIcon.Add_DoubleClick({ $window.Show(); $window.WindowState = 'Normal'; $window.Activate(); $script:TrayIcon.Visible = $false })
$trayExtract.Add_Click({ Start-QueueExtraction -Silent })
$trayStop.Add_Click({ $script:State.StopRequested = $true })
$trayExit.Add_Click({ $script:ForceClose = $true; $window.Close() })
$script:ForceClose = $false

# =====================================================================
# Window Chrome Controls (Custom Title Bar)
# =====================================================================
$ui['BtnMinimize'].Add_Click({ $window.WindowState = 'Minimized' })
$ui['BtnMaximize'].Add_Click({
    if ($window.WindowState -eq 'Maximized') { $window.WindowState = 'Normal' }
    else { $window.WindowState = 'Maximized' }
})
$ui['BtnClose'].Add_Click({ $script:ForceClose = $true; $window.Close() })
$window.Add_StateChanged({
    if ($window.WindowState -eq 'Maximized') { $ui['MaximizeIcon'].Text = [char]0x29C9 }
    else { $ui['MaximizeIcon'].Text = [char]0x25A1 }
})

# =====================================================================
# UI Log Helper
# =====================================================================
function Write-UILog {
    param([string]$Message, [string]$Color = '#FF8888A0')
    $ts = Get-Date -Format "HH:mm:ss"
    $run = New-Object System.Windows.Documents.Run("[$ts] $Message`n")
    $run.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Color)
    $ui['LogBox'].Inlines.Add($run)
    while ($ui['LogBox'].Inlines.Count -gt 3000) { $ui['LogBox'].Inlines.Remove($ui['LogBox'].Inlines.FirstInline) }
    $ui['LogScroll'].ScrollToEnd()
    if ($script:Config.LogHistory) {
        $logFile = Join-Path $script:LogDir "ExtractorX_$(Get-Date -Format 'yyyyMMdd').log"
        try { "[$ts] $Message" | Add-Content -Path $logFile -Force -ErrorAction SilentlyContinue } catch {}
    }
}

# =====================================================================
# Load Settings Into UI
# =====================================================================
if ($script:Config.WindowWidth -gt 0 -and $script:Config.WindowLeft -ge 0) {
    $window.Width = $script:Config.WindowWidth; $window.Height = $script:Config.WindowHeight
    $window.Left = $script:Config.WindowLeft; $window.Top = $script:Config.WindowTop
    $window.WindowStartupLocation = 'Manual'
}
$ui['TxtOutputPath'].Text = $script:Config.OutputPath
$ui['ChkDeleteAfter'].IsChecked = [bool]$script:Config.DeleteAfterExtract
if ($script:Config.AlwaysOnTop) { $window.Topmost = $true }

if ($script:7zPath) {
    $ui['SevenZipStatus'].Text = "7-Zip: Ready"
    $ui['SevenZipStatus'].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF6BCB77')
} else {
    $ui['SevenZipStatus'].Text = "7-Zip: Not found"
    $ui['SevenZipStatus'].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FFFFD93D')
}

# =====================================================================
# UI Timer - Polls ConcurrentQueue
# =====================================================================
$script:UITimer = New-Object System.Windows.Threading.DispatcherTimer
$script:UITimer.Interval = [TimeSpan]::FromMilliseconds(100)
$script:UITimer.Add_Tick({
    $msg = $null; $n = 0
    while ($script:UIQueue.TryDequeue([ref]$msg) -and $n -lt 60) {
        $n++
        switch ($msg.Type) {
            'Log'    { Write-UILog -Message $msg.Message -Color $msg.Color }
            'Status' { $ui['StatusText'].Text = $msg.Text; $ui['StatusText'].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($msg.Color) }
            'Progress' { $ui['ProgressText'].Text = $msg.Text }
            'ProgressBarUpdate' {
                try {
                    $pct = [double]$msg.Percent
                    $barMaxW = $ui['ProgressBar'].Parent.ActualWidth
                    if ($barMaxW -gt 0) { $ui['ProgressBar'].Width = [Math]::Min($barMaxW, $barMaxW * $pct / 100.0) }
                } catch {}
            }
            'TrayTextUpdate' { try { $script:TrayIcon.Text = $msg.Text.Substring(0, [Math]::Min(63, $msg.Text.Length)); $ui['TitleBarText'].Text = $msg.Text -replace '^ExtractorX - ',''; $window.Title = $msg.Text } catch {} }
            'TrayTip' { if ($script:TrayIcon.Visible -and $script:Config.MinimizeToTray) { $script:TrayIcon.ShowBalloonTip(3000, "ExtractorX", $msg.Text, [System.Windows.Forms.ToolTipIcon]::Info) } }
            'QueueUpdate' {
                $idx = $msg.Index
                if ($idx -ge 0 -and $idx -lt $ui['QueueList'].Items.Count) {
                    $old = $ui['QueueList'].Items[$idx]
                    $color = switch -Wildcard ($msg.Status) {
                        '*Success*'    { '#FF6BCB77' }
                        'Test OK*'     { '#FF6BCB77' }
                        '*Fail*'       { '#FFFF6B6B' }
                        '*Error*'      { '#FFFF6B6B' }
                        'Extracting*'  { '#FF00B4D8' }
                        'Trying*'      { '#FFFFD93D' }
                        default        { '#FFD4D4E0' }
                    }
                    $ui['QueueList'].Items[$idx] = [PSCustomObject]@{
                        Filename=$old.Filename; Directory=$old.Directory; Size=$old.Size
                        Extension=$old.Extension; Status=$msg.Status
                        Elapsed = if ($msg.Elapsed) { $msg.Elapsed } else { $old.Elapsed }
                        FullPath=$old.FullPath; SizeBytes=$old.SizeBytes; StatusColor=$color
                    }
                    try {
                        $lvi = $ui['QueueList'].ItemContainerGenerator.ContainerFromIndex($idx)
                        if ($lvi) { $lvi.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($color) }
                    } catch {}
                }
            }
            'QueueCountRefresh' {
                $t=$ui['QueueList'].Items.Count; $q=0; $d=0; $f=0; $totalBytes=[long]0
                foreach ($i in $ui['QueueList'].Items) {
                    if ($i.Status -eq 'Queued') { $q++ }
                    elseif ($i.Status -like '*Success*' -or $i.Status -eq 'Test OK') { $d++ }
                    elseif ($i.Status -like '*Fail*' -or $i.Status -like '*Error*') { $f++ }
                    if ($i.SizeBytes) { $totalBytes += [long]$i.SizeBytes }
                }
                $ui['QueueCount'].Text = "Queue: $t | Pending: $q | Done: $d | Failed: $f"
                if ($totalBytes -gt 0) { $ui['TotalSizeText'].Text = "Total: $(Get-FileSizeString $totalBytes)" }
                else { $ui['TotalSizeText'].Text = '' }
            }
            'ExtractionDone' {
                $script:State.IsExtracting = $false
                $ui['BtnExtractAll'].IsEnabled = $true; $ui['BtnStopAll'].IsEnabled = $false
                $ui['BtnRemoveSelected'].IsEnabled = $true; $ui['BtnClearQueue'].IsEnabled = $true
                $trayStop.Enabled = $false; $trayExtract.Enabled = $true
                # Reset progress bar and tray text
                $ui['ProgressBar'].Width = 0
                $script:TrayIcon.Text = "ExtractorX v$($script:AppVersion)"
                $ui['TitleBarText'].Text = "Ready"
                $window.Title = "ExtractorX v$($script:AppVersion)"
                # Sounds
                if ($script:Config.SoundsEnabled) {
                    if ($msg.Failed -gt 0) { [System.Media.SystemSounds]::Hand.Play() }
                    else { [System.Media.SystemSounds]::Asterisk.Play() }
                }
                Send-UIMessage -Type 'QueueCountRefresh'
                # Batch complete actions
                if ($msg.Completed -gt 0 -and $msg.Failed -eq 0 -and $script:Config.CloseOnComplete) {
                    $script:ForceClose = $true; $window.Close()
                } elseif ($script:Config.CloseOnCompleteAlways -and ($msg.Completed + $msg.Failed) -gt 0) {
                    $script:ForceClose = $true; $window.Close()
                }
                if ($script:Config.ClearListOnComplete) { $ui['QueueList'].Items.Clear(); $script:QueuedPaths.Clear(); Send-UIMessage -Type 'QueueCountRefresh' }
                # Re-trigger for watch auto-extract
                if ($script:State.WatchersActive -and $script:Config.WatchAutoExtract) {
                    $hq = $false; foreach ($item in $ui['QueueList'].Items) { if ($item.Status -eq 'Queued') { $hq=$true; break } }
                    if ($hq) { Start-QueueExtraction -Silent }
                }
            }
            'SevenZipFound' {
                $script:State.Downloading7z = $false
                $ui['SevenZipStatus'].Text = "7-Zip: Ready"
                $ui['SevenZipStatus'].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF6BCB77')
                $ui['BtnExtractAll'].IsEnabled = $true
                if ($script:State.WatchersActive -and $script:Config.WatchAutoExtract) { Start-QueueExtraction -Silent }
            }
            'SevenZipFailed' {
                $script:State.Downloading7z = $false
                $ui['SevenZipStatus'].Text = "7-Zip: Failed"
                $ui['SevenZipStatus'].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FFFF6B6B')
                $ui['BtnExtractAll'].IsEnabled = $true
            }
            'WatchFileDetected' {
                $fp = $msg.FilePath
                if ((Test-Path $fp) -and -not $script:QueuedPaths.Contains($fp)) {
                    Add-FileDirect $fp
                    Send-UIMessage -Type 'QueueCountRefresh'
                    if ($script:Config.WatchAutoExtract) { Start-QueueExtraction -Silent }
                }
            }
            'ScanBatch' {
                # Batch of archive files from background scanner
                foreach ($item in $msg.Items) {
                    if ($script:QueuedPaths.Contains($item.FullPath)) { continue }
                    [void]$script:QueuedPaths.Add($item.FullPath)
                    $ui['QueueList'].Items.Add([PSCustomObject]@{
                        Filename=$item.Filename; Directory=$item.Directory
                        Size=(Get-FileSizeString $item.Size); Extension=$item.Extension
                        Status='Queued'; Elapsed=''; FullPath=$item.FullPath; SizeBytes=[long]$item.Size; StatusColor='#FFD4D4E0'
                    })
                }
                Send-UIMessage -Type 'QueueCountRefresh'
            }
            'ScanProgress' {
                $ui['StatusText'].Text = $msg.Text
                $ui['StatusText'].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FFFFD93D')
            }
            'ScanDone' {
                $script:State.IsScanning = $false
                if (-not $script:State.IsExtracting) { $ui['BtnStopAll'].IsEnabled = $false }
                $ui['StatusText'].Text = "Ready"
                $ui['StatusText'].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF6BCB77')
                Send-UIMessage -Type 'QueueCountRefresh'
                # Auto-extract if requested (from drop-auto-extract)
                if ($script:State.AutoExtractAfterScan -and -not $msg.Cancelled) {
                    $script:State.AutoExtractAfterScan = $false
                    Start-QueueExtraction -Silent
                }
                # Check for pending scan paths
                $pendingPath = $null
                if ($script:PendingScanPaths.TryDequeue([ref]$pendingPath)) {
                    $batch = @($pendingPath)
                    while ($script:PendingScanPaths.TryDequeue([ref]$pendingPath)) { $batch += $pendingPath }
                    Start-ScanJob -Paths $batch
                }
            }
        }
    }
})
$script:UITimer.Start()

# Re-apply row colors when virtualized containers are recycled
$ui['QueueList'].ItemContainerGenerator.Add_StatusChanged({
    if ($ui['QueueList'].ItemContainerGenerator.Status -eq 'ContainersGenerated') {
        for ($i = 0; $i -lt $ui['QueueList'].Items.Count; $i++) {
            $item = $ui['QueueList'].Items[$i]
            if ($item.StatusColor -and $item.StatusColor -ne '#FFD4D4E0') {
                try {
                    $lvi = $ui['QueueList'].ItemContainerGenerator.ContainerFromIndex($i)
                    if ($lvi) { $lvi.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($item.StatusColor) }
                } catch {}
            }
        }
    }
})

# =====================================================================
# Queue Management
# =====================================================================
# Fast duplicate tracking (O(1) vs iterating ListView)
$script:QueuedPaths = New-Object 'System.Collections.Generic.HashSet[string]'([System.StringComparer]::OrdinalIgnoreCase)

# Lightweight single-file add for watch events and known files (no scanning needed)
function Add-FileDirect {
    param([string]$FilePath)
    if ($script:QueuedPaths.Contains($FilePath)) { return }
    $fi = Get-Item $FilePath -Force -ErrorAction SilentlyContinue
    if (-not $fi -or $fi.PSIsContainer) { return }
    [void]$script:QueuedPaths.Add($FilePath)
    $ui['QueueList'].Items.Add([PSCustomObject]@{
        Filename=$fi.Name; Directory=$fi.DirectoryName
        Size=(Get-FileSizeString $fi.Length); Extension=$fi.Extension.TrimStart('.').ToUpper()
        Status='Queued'; Elapsed=''; FullPath=$FilePath; SizeBytes=[long]$fi.Length; StatusColor='#FFD4D4E0'
    })
}

# Background directory scanner - never blocks the UI
function Start-ScanJob {
    param([string[]]$Paths, [switch]$AutoExtractAfter)
    if ($script:State.IsScanning) {
        Write-UILog "Scan already in progress - queuing paths" '#FFFFD93D'
        # Queue these for after current scan finishes
        foreach ($p in $Paths) { $script:PendingScanPaths.Enqueue($p) }
        return
    }

    # Separate files vs directories
    $directFiles = @()
    $scanDirs = @()
    foreach ($p in $Paths) {
        if (-not (Test-Path $p)) { continue }
        $item = Get-Item $p -Force -ErrorAction SilentlyContinue
        if (-not $item) { continue }
        if ($item.PSIsContainer) { $scanDirs += $p }
        else { $directFiles += $p }
    }

    # Add individual files immediately (fast, no recursion)
    $addedDirect = 0
    $deep = [bool]$script:Config.DeepArchiveDetection
    foreach ($f in $directFiles) {
        if ($script:QueuedPaths.Contains($f)) { continue }
        $fn = [IO.Path]::GetFileName($f)
        if (Test-IsMultiVolumePart $fn) { Write-UILog "  Skipped volume part: $fn" '#FF505068'; continue }
        $isArch = (Test-IsArchive $f) -or ($deep -and (Test-MagicBytes $f))
        if ($isArch) { Add-FileDirect $f; $addedDirect++ }
    }
    if ($addedDirect -gt 0) { Send-UIMessage -Type 'QueueCountRefresh' }

    # If no directories to scan, we're done
    if ($scanDirs.Count -eq 0) {
        if ($AutoExtractAfter -and $script:Config.AutoExtractOnDrop) { Start-QueueExtraction -Silent }
        return
    }

    # Launch background scan
    $script:State.IsScanning = $true
    $script:State.ScanCancelRequested = $false
    if ($AutoExtractAfter) { $script:State.AutoExtractAfterScan = $true }

    # Snapshot existing queue paths for duplicate filtering in runspace
    $existingPaths = [System.Collections.Generic.HashSet[string]]::new($script:QueuedPaths, [System.StringComparer]::OrdinalIgnoreCase)

    Write-UILog "Scanning $($scanDirs.Count) directory(ies)..." '#FF00B4D8'
    $ui['StatusText'].Text = "Scanning..."
    $ui['StatusText'].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FFFFD93D')
    $ui['BtnStopAll'].IsEnabled = $true

    $sRS = [runspacefactory]::CreateRunspace(); $sRS.Open()
    $sRS.SessionStateProxy.SetVariable('state', $script:State)
    $sRS.SessionStateProxy.SetVariable('uiQueue', $script:UIQueue)
    $sRS.SessionStateProxy.SetVariable('scanDirs', $scanDirs)
    $sRS.SessionStateProxy.SetVariable('archiveExts', $script:ArchiveExtensions)
    $sRS.SessionStateProxy.SetVariable('deepDetect', [bool]$script:Config.DeepArchiveDetection)
    $sRS.SessionStateProxy.SetVariable('filterType', $script:Config.DragDropFilterType)
    $sRS.SessionStateProxy.SetVariable('filterMask', $script:Config.DragDropFilterMask)
    $sRS.SessionStateProxy.SetVariable('existingPaths', $existingPaths)

    $sPS = [powershell]::Create().AddScript({
        function Send-Msg($T,$D) { $uiQueue.Enqueue((@{Type=$T}+$D)) }

        # Pre-build extension lookup HashSet for speed
        $extSet = New-Object 'System.Collections.Generic.HashSet[string]'([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($e in $archiveExts) { [void]$extSet.Add($e) }

        # Pre-parse filter masks
        $inclMasks = @(); $exclMasks = @()
        if ($filterType -eq 'Inclusion' -and $filterMask) { $inclMasks = @(($filterMask -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
        elseif ($filterType -eq 'Exclusion' -and $filterMask) { $exclMasks = @(($filterMask -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }

        # Magic bytes signatures
        $magicSigs = @{
            '504B0304'='zip';'504B0506'='zip';'504B0708'='zip'
            '377ABCAF271C'='7z';'526172211A0700'='rar';'526172211A07'='rar'
            '1F8B'='gz';'425A68'='bz2';'FD377A585A00'='xz';'4D534346'='cab'
            '7573746172'='tar';'4344303031'='iso'
        }

        function Test-Magic([string]$FilePath) {
            try {
                $s = [IO.File]::OpenRead($FilePath)
                try {
                    $buf = New-Object byte[] 16; $r = $s.Read($buf,0,16)
                    if ($r -lt 2) { return $false }
                    $hex = ($buf[0..($r-1)] | ForEach-Object { $_.ToString('X2') }) -join ''
                    foreach ($sig in $magicSigs.Keys) { if ($hex.StartsWith($sig)) { return $true } }
                } finally { $s.Dispose() }
            } catch {}
            return $false
        }

        $dupes = New-Object 'System.Collections.Generic.HashSet[string]'($existingPaths, [System.StringComparer]::OrdinalIgnoreCase)
        $batch = [System.Collections.ArrayList]::new()
        $scanned = 0; $found = 0; $dirsScanned = 0; $lastProgressTime = [datetime]::UtcNow

        foreach ($rootDir in $scanDirs) {
            if ($state.ScanCancelRequested) { break }
            Send-Msg 'ScanProgress' @{Text="Scanning: $rootDir"; Scanned=$scanned; Found=$found; Dir=$rootDir}

            # Manual recursive walk - handles access errors per-directory without terminating
            try {
                $dirStack = New-Object System.Collections.Stack
                $dirStack.Push($rootDir)
                while ($dirStack.Count -gt 0 -and -not $state.ScanCancelRequested) {
                    $curDir = $dirStack.Pop()
                    # Enumerate subdirectories
                    try { foreach ($sub in [IO.Directory]::EnumerateDirectories($curDir)) { $dirStack.Push($sub) } } catch {}
                    # Enumerate files in current directory
                    $fileEnum = $null
                    try { $fileEnum = [IO.Directory]::EnumerateFiles($curDir) } catch { continue }
                    foreach ($filePath in $fileEnum) {
                    if ($state.ScanCancelRequested) { break }
                    $scanned++

                    # Progress every 200ms or 500 files
                    if ($scanned % 500 -eq 0 -or ([datetime]::UtcNow - $lastProgressTime).TotalMilliseconds -gt 200) {
                        $dispCurDir = [IO.Path]::GetDirectoryName($filePath)
                        # Truncate long paths for display
                        $dispDir = $dispCurDir
                        if ($dispDir.Length -gt 70) { $dispDir = $dispDir.Substring(0,30) + "..." + $dispDir.Substring($dispDir.Length-37) }
                        Send-Msg 'ScanProgress' @{Text="Scanning: $scanned files | $found archives | $dispDir"; Scanned=$scanned; Found=$found; Dir=$dispCurDir}
                        $lastProgressTime = [datetime]::UtcNow
                    }

                    # Duplicate check
                    if ($dupes.Contains($filePath)) { continue }

                    # Filter masks
                    $fileName = [IO.Path]::GetFileName($filePath)
                    if ($inclMasks.Count -gt 0) {
                        $hit=$false; foreach ($m in $inclMasks) { if ($fileName -like $m) { $hit=$true; break } }
                        if (-not $hit) { continue }
                    }
                    if ($exclMasks.Count -gt 0) {
                        $skip=$false; foreach ($m in $exclMasks) { if ($fileName -like $m) { $skip=$true; break } }
                        if ($skip) { continue }
                    }

                    # Extension check
                    $ext = [IO.Path]::GetExtension($filePath)
                    $isArchive = $extSet.Contains($ext)

                    # Magic bytes fallback
                    if (-not $isArchive -and $deepDetect) { $isArchive = Test-Magic $filePath }

                    if ($isArchive) {
                        # Multi-volume detection: skip non-first volumes
                        $skipVol = $false
                        if ($fileName -match '\.part(\d+)\.rar$') { if ([int]$Matches[1] -gt 1) { $skipVol=$true } }
                        elseif ($fileName -match '\.[rs]\d{2}$') { $skipVol=$true }
                        elseif ($fileName -match '\.\w+\.(\d{3})$' -and $Matches[1] -ne '001') { $skipVol=$true }
                        if ($skipVol) { continue }
                        [void]$dupes.Add($filePath)
                        $found++
                        try { $fi = [IO.FileInfo]::new($filePath); $fSize = $fi.Length } catch { $fSize = 0 }
                        [void]$batch.Add(@{
                            FullPath=$filePath; Filename=$fileName
                            Directory=[IO.Path]::GetDirectoryName($filePath)
                            Size=$fSize; Extension=$ext.TrimStart('.').ToUpper()
                        })

                        # Flush batch every 50 items
                        if ($batch.Count -ge 50) {
                            Send-Msg 'ScanBatch' @{Items=@($batch.ToArray())}
                            $batch.Clear()
                        }
                    }
                }
                }
            } catch {
                Send-Msg 'Log' @{Message="Scan error in $rootDir : $_"; Color='#FFFF6B6B'}
            }
            $dirsScanned++
        }

        # Flush remaining
        if ($batch.Count -gt 0) { Send-Msg 'ScanBatch' @{Items=@($batch.ToArray())} }

        $statusMsg = if ($state.ScanCancelRequested) { "Scan cancelled: $found archives in $scanned files" } else { "Scan complete: $found archives in $("{0:N0}" -f $scanned) files" }
        Send-Msg 'Log' @{Message=$statusMsg; Color=if($found -gt 0){'#FF6BCB77'}else{'#FFFFD93D'}}
        Send-Msg 'ScanDone' @{Found=$found; Scanned=$scanned; Cancelled=[bool]$state.ScanCancelRequested}
    })
    $sPS.Runspace = $sRS; $sPS.BeginInvoke() | Out-Null
    $script:scanPS = $sPS; $script:scanRS = $sRS
}

# Pending scan paths queue for sequential scanning
$script:PendingScanPaths = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()


# =====================================================================
# EXTRACTION ENGINE - Reusable from button, watch, tray, context menu
# =====================================================================
function Start-QueueExtraction {
    param([switch]$Silent, [switch]$UnextractedOnly)
    if ($script:State.IsExtracting) { return }
    if ($script:State.Downloading7z) { if (-not $Silent) { [System.Windows.MessageBox]::Show("7-Zip downloading...","ExtractorX",'OK','Information') }; return }

    $queuedItems = @()
    for ($i=0; $i -lt $ui['QueueList'].Items.Count; $i++) {
        $st = $ui['QueueList'].Items[$i].Status
        if ($UnextractedOnly) { if ($st -eq 'Queued') { $queuedItems += @{ Index=$i; FullPath=$ui['QueueList'].Items[$i].FullPath; Filename=$ui['QueueList'].Items[$i].Filename } } }
        else { if ($st -eq 'Queued') { $queuedItems += @{ Index=$i; FullPath=$ui['QueueList'].Items[$i].FullPath; Filename=$ui['QueueList'].Items[$i].Filename } } }
    }
    if ($queuedItems.Count -eq 0) { if (-not $Silent) { [System.Windows.MessageBox]::Show("No items to extract.","ExtractorX",'OK','Information') }; return }

    $outTpl = $ui['TxtOutputPath'].Text.Trim()
    if (-not $outTpl) { if (-not $Silent) { [System.Windows.MessageBox]::Show("Output path is empty.","ExtractorX",'OK','Warning') }; return }

    # 7-Zip download if needed
    if (-not $script:State.SevenZipPath) {
        $script:State.Downloading7z = $true; $ui['BtnExtractAll'].IsEnabled = $false
        Send-UIMessage -Type 'Status' -Data @{ Text="Downloading 7-Zip..."; Color='#FFFFD93D' }
        Send-UIMessage -Type 'Log' -Data @{ Message="Downloading 7-Zip..."; Color='#FFFFD93D' }
        $dlRS = [runspacefactory]::CreateRunspace(); $dlRS.Open()
        $dlRS.SessionStateProxy.SetVariable('state', $script:State)
        $dlRS.SessionStateProxy.SetVariable('uiQueue', $script:UIQueue)
        $dlRS.SessionStateProxy.SetVariable('appDir', $script:AppDataDir)
        $dlPS = [powershell]::Create().AddScript({
            function Send-Msg($T,$D) { $uiQueue.Enqueue((@{Type=$T}+$D)) }
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                $t = Join-Path $appDir "7zr.exe"
                Invoke-WebRequest -Uri "https://www.7-zip.org/a/7zr.exe" -OutFile $t -UseBasicParsing
                $t7 = Join-Path $appDir "7z.exe"; Copy-Item $t $t7 -Force
                try { $ex = Join-Path $appDir "7z-extra.7z"; Invoke-WebRequest -Uri "https://www.7-zip.org/a/7z2408-extra.7z" -OutFile $ex -UseBasicParsing; & $t x $ex -o"$appDir" -y 2>$null; Remove-Item $ex -Force -ErrorAction SilentlyContinue } catch {}
                if (Test-Path $t7) { $state.SevenZipPath=$t7; Send-Msg 'SevenZipFound' @{Path=$t7} } else { $state.SevenZipPath=$t; Send-Msg 'SevenZipFound' @{Path=$t} }
                Send-Msg 'Log' @{Message="7-Zip ready";Color='#FF6BCB77'}; Send-Msg 'Status' @{Text="Ready";Color='#FF6BCB77'}
            } catch { Send-Msg 'SevenZipFailed' @{}; Send-Msg 'Log' @{Message="7-Zip download failed: $_";Color='#FFFF6B6B'} }
        })
        $dlPS.Runspace = $dlRS; $dlPS.BeginInvoke() | Out-Null
        $script:dl7zPS=$dlPS; $script:dl7zRS=$dlRS
        return
    }

    # Lock UI
    $script:State.IsExtracting = $true; $script:State.StopRequested = $false
    $ui['BtnExtractAll'].IsEnabled=$false; $ui['BtnStopAll'].IsEnabled=$true
    $ui['BtnRemoveSelected'].IsEnabled=$false; $ui['BtnClearQueue'].IsEnabled=$false
    $trayStop.Enabled=$true; $trayExtract.Enabled=$false
    Send-UIMessage -Type 'Status' -Data @{ Text="Extracting..."; Color='#FFFFD93D' }

    # Snapshot all settings
    $settings = @{
        OutputTemplate     = $ui['TxtOutputPath'].Text
        Overwrite          = $script:Config.OverwriteMode
        PostAction         = $script:Config.PostAction
        PostFolder         = $script:Config.PostActionFolder
        DeleteAfterExtract = [bool]$ui['ChkDeleteAfter'].IsChecked
        OpenDestAfter      = [bool]$script:Config.OpenDestAfterExtract
        Exclusions         = @(($script:Config.FileExclusions -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        NestedEnabled      = [bool]$script:Config.NestedExtraction
        NestedDepth        = [int]$script:Config.NestedMaxDepth
        NestedApplyPost    = [bool]$script:Config.NestedApplyPostAction
        RemoveDupe         = [bool]$script:Config.RemoveDuplicateFolder
        RenameSingle       = [bool]$script:Config.RenameSingleFile
        DeleteBroken       = [bool]$script:Config.DeleteBrokenFiles
        Passwords          = @($script:Passwords)
        SevenZip           = $script:State.SevenZipPath
        ArchiveExts        = @($script:ArchiveExtensions)
        TestOnly           = [bool]$script:State.TestOnly
        ThreadPriority     = $script:Config.ThreadPriority
        ExternalProcs      = @($script:Config.ExternalProcessors)
    }

    $rs = [runspacefactory]::CreateRunspace(); $rs.Open()
    $rs.SessionStateProxy.SetVariable('state', $script:State)
    $rs.SessionStateProxy.SetVariable('uiQueue', $script:UIQueue)
    $rs.SessionStateProxy.SetVariable('queuedItems', $queuedItems)
    $rs.SessionStateProxy.SetVariable('settings', $settings)

    $ps = [powershell]::Create().AddScript({
        function Send-Msg($T,$D) { $uiQueue.Enqueue((@{Type=$T}+$D)) }

        # Set thread priority
        try {
            switch ($settings.ThreadPriority) {
                'Low'         { [System.Threading.Thread]::CurrentThread.Priority = [System.Threading.ThreadPriority]::BelowNormal }
                'BelowNormal' { [System.Threading.Thread]::CurrentThread.Priority = [System.Threading.ThreadPriority]::BelowNormal }
                'AboveNormal' { [System.Threading.Thread]::CurrentThread.Priority = [System.Threading.ThreadPriority]::AboveNormal }
                'High'        { [System.Threading.Thread]::CurrentThread.Priority = [System.Threading.ThreadPriority]::Highest }
                default       { [System.Threading.Thread]::CurrentThread.Priority = [System.Threading.ThreadPriority]::Normal }
            }
        } catch {}
        try { Add-Type -AssemblyName Microsoft.VisualBasic } catch {}

        function Run-7z {
            param([string]$Archive, [string]$OutDir, [string]$Pw='', [switch]$Test, [switch]$Silent)
            $a = @($(if($Test){'t'}else{'x'}), "`"$Archive`"")
            if (-not $Test) { $a += "-o`"$OutDir`""; $a += '-y' }
            # -bb1 = show extracted filenames, -bsp1 = show progress to stdout
            $a += '-bb1'; $a += '-bsp1'
            switch ($settings.Overwrite) { 'Always'{$a+='-aoa'} 'Never'{$a+='-aos'} 'Rename'{$a+='-aou'} }
            if ($Pw) { $a += "-p`"$Pw`"" }
            foreach ($e in $settings.Exclusions) { if ($e) { $a += "-xr!$e" } }

            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = $settings.SevenZip; $psi.Arguments = ($a -join ' ')
            $psi.UseShellExecute=$false; $psi.RedirectStandardOutput=$true; $psi.RedirectStandardError=$true; $psi.CreateNoWindow=$true

            $proc = New-Object System.Diagnostics.Process; $proc.StartInfo = $psi
            $errSB = [System.Text.StringBuilder]::new()
            # Shared queue for real-time output lines
            $outLines = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

            $outEvt = Register-ObjectEvent $proc OutputDataReceived -Action {
                if ($null -ne $EventArgs.Data -and $EventArgs.Data.Length -gt 0) {
                    $Event.MessageData.Enqueue($EventArgs.Data)
                }
            } -MessageData $outLines
            $errEvt = Register-ObjectEvent $proc ErrorDataReceived -Action {
                if ($null -ne $EventArgs.Data) { $Event.MessageData.AppendLine($EventArgs.Data) }
            } -MessageData $errSB

            try { $proc.Start() | Out-Null } catch {
                Unregister-Event $outEvt.Name -EA SilentlyContinue; Unregister-Event $errEvt.Name -EA SilentlyContinue
                return @{ Success=$false; ExitCode=-1; NeedsPassword=$false; Error="Cannot start 7z: $_"; FileCount=0 }
            }
            $proc.BeginOutputReadLine(); $proc.BeginErrorReadLine()

            $killed=$false; $fileCount=0; $lastStatusLine=''
            while (-not $proc.WaitForExit(250)) {
                if ($state.StopRequested) { try { $proc.Kill(); $killed=$true } catch {}; break }
                # Flush output lines to UI in real-time
                $line = $null
                while ($outLines.TryDequeue([ref]$line)) {
                    if (-not $Silent) {
                        # Parse 7z output: "- filename" = extracted file, "T filename" = tested file
                        if ($line -match '^- (.+)$' -or $line -match '^Extracting\s+(.+)$' -or $line -match '^T (.+)$') {
                            $fn = $Matches[1].Trim()
                            if ($fn.Length -gt 80) { $fn = '...' + $fn.Substring($fn.Length - 77) }
                            Send-Msg 'Log' @{Message="    $fn";Color='#FF6E6E84'}
                            $fileCount++
                        } elseif ($line -match '^\s*(\d+)%') {
                            $lastStatusLine = $line.Trim()
                        }
                    }
                }
            }
            try { $proc.WaitForExit() } catch {}
            # Final flush
            $line = $null
            while ($outLines.TryDequeue([ref]$line)) {
                if (-not $Silent -and ($line -match '^- (.+)$' -or $line -match '^Extracting\s+(.+)$' -or $line -match '^T (.+)$')) {
                    $fn = $Matches[1].Trim()
                    if ($fn.Length -gt 80) { $fn = '...' + $fn.Substring($fn.Length - 77) }
                    Send-Msg 'Log' @{Message="    $fn";Color='#FF6E6E84'}
                    $fileCount++
                }
            }
            Unregister-Event $outEvt.Name -EA SilentlyContinue; Unregister-Event $errEvt.Name -EA SilentlyContinue

            $ec = if ($killed) { -1 } else { try { $proc.ExitCode } catch { -1 } }
            $stderr = $errSB.ToString(); $proc.Dispose()
            return @{
                Success = ($ec -eq 0 -or $ec -eq 1); ExitCode=$ec; FileCount=$fileCount
                NeedsPassword = ($ec -eq 2 -and ($stderr -match 'password|Wrong password|encrypted|Can not open encrypted'))
                Error = if ($stderr) { $stderr } else { '' }
            }
        }

        function Resolve-Out($Template, $ArchivePath) {
            $d=[IO.Path]::GetDirectoryName($ArchivePath); $n=[IO.Path]::GetFileNameWithoutExtension($ArchivePath)
            if ($n -match '\.tar$') { $n=[IO.Path]::GetFileNameWithoutExtension($n) }
            $e=[IO.Path]::GetExtension($ArchivePath).TrimStart('.'); $fn=[IO.Path]::GetFileName($ArchivePath); $pn=Split-Path $d -Leaf
            $r=$Template
            $r=$r -replace '\{ArchiveFolder\}',$d; $r=$r -replace '\{ArchiveName\}',$n; $r=$r -replace '\{ArchiveExtension\}',$e
            $r=$r -replace '\{ArchiveFileName\}',$fn; $r=$r -replace '\{ArchiveFolderName\}',$pn; $r=$r -replace '\{ArchivePath\}',$ArchivePath
            $r=$r -replace '\{Guid\}',([Guid]::NewGuid().ToString('N').Substring(0,8))
            $r=$r -replace '\{Desktop\}',[Environment]::GetFolderPath('Desktop')
            $r=$r -replace '\{UserProfile\}',$env:USERPROFILE; $r=$r -replace '\{Windows\}',$env:windir
            $r=$r -replace '\{Date\}',(Get-Date -Format 'yyyyMMdd'); $r=$r -replace '\{Time\}',(Get-Date -Format 'HHmmss')
            if ($r -match '\{Env:([^}]+)\}') { $r=[regex]::Replace($r,'\{Env:([^}]+)\}',{param($m) [Environment]::GetEnvironmentVariable($m.Groups[1].Value)}) }
            if ($r -match '\{ArchiveNameUnique\}') {
                $b=$r -replace '\{ArchiveNameUnique\}',$n
                if (Test-Path $b) { $c=1; do { $t=$r -replace '\{ArchiveNameUnique\}',"$n ($c)"; $c++ } while (Test-Path $t); $r=$t } else { $r=$b }
            }
            return $r
        }

        function Apply-PostAction($ArchivePath, $ArchiveName) {
            if ($settings.DeleteAfterExtract) {
                try { Remove-Item $ArchivePath -Force -ErrorAction Stop; Send-Msg 'Log' @{Message="  Deleted: $ArchiveName";Color='#FF8888A0'} }
                catch { Send-Msg 'Log' @{Message="  Delete failed: $ArchiveName";Color='#FFFF6B6B'} }
            } elseif (Test-Path $ArchivePath) {
                switch ($settings.PostAction) {
                    'Recycle' { try { [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($ArchivePath,'OnlyErrorDialogs','SendToRecycleBin'); Send-Msg 'Log' @{Message="  Recycled: $ArchiveName";Color='#FF8888A0'} } catch {} }
                    'MoveToFolder' { if ($settings.PostFolder -and (Test-Path $settings.PostFolder)) { try { Move-Item $ArchivePath $settings.PostFolder -Force; Send-Msg 'Log' @{Message="  Moved: $ArchiveName";Color='#FF8888A0'} } catch {} } }
                    'Delete' { try { Remove-Item $ArchivePath -Force; Send-Msg 'Log' @{Message="  Deleted: $ArchiveName";Color='#FF8888A0'} } catch {} }
                }
            }
        }

        # --- Main loop ---
        $completed=0; $failed=0; $total=$queuedItems.Count; $totalFiles=0
        $batchSW = [System.Diagnostics.Stopwatch]::StartNew()
        Send-Msg 'Log' @{Message="--- Batch: $total archive(s) ---";Color='#FF00B4D8'}
        foreach ($qi in $queuedItems) {
            if ($state.StopRequested) { break }
            $idx=$qi.Index; $archPath=$qi.FullPath; $archName=$qi.Filename
            $sw=[System.Diagnostics.Stopwatch]::StartNew()

            Send-Msg 'QueueUpdate' @{Index=$idx;Status='Extracting...'}
            Send-Msg 'Progress' @{Text="$($completed+$failed+1)/$total : $archName"}
            Send-Msg 'ProgressBarUpdate' @{Percent=(($completed+$failed)/$total*100)}
            Send-Msg 'TrayTextUpdate' @{Text="ExtractorX - $($completed+$failed+1)/$total $archName"}
            Send-Msg 'Log' @{Message="Extracting: $archName";Color='#FF00B4D8'}
            Send-Msg 'Log' @{Message="  Source: $archPath";Color='#FF505068'}

            $outDir = Resolve-Out $settings.OutputTemplate $archPath
            Send-Msg 'Log' @{Message="  Output: $outDir";Color='#FF505068'}

            if ($settings.TestOnly) {
                $result = Run-7z -Archive $archPath -OutDir $outDir -Test
            } else {
                if (-not (Test-Path $outDir)) { try { New-Item -Path $outDir -ItemType Directory -Force | Out-Null } catch { $failed++; Send-Msg 'QueueUpdate' @{Index=$idx;Status='Error: output dir'}; continue } }
                $result = Run-7z -Archive $archPath -OutDir $outDir
            }
            $pwUsed = ''

            # Password cycling
            if ($result.NeedsPassword -and $settings.Passwords.Count -gt 0) {
                Send-Msg 'Log' @{Message="  Encrypted - cycling $($settings.Passwords.Count) passwords";Color='#FFFFD93D'}
                Send-Msg 'QueueUpdate' @{Index=$idx;Status='Trying passwords...'}
                foreach ($pw in $settings.Passwords) {
                    if ($state.StopRequested) { break }
                    if ($settings.TestOnly) { $result = Run-7z -Archive $archPath -OutDir $outDir -Pw $pw -Test -Silent }
                    else { $result = Run-7z -Archive $archPath -OutDir $outDir -Pw $pw -Silent }
                    if ($result.Success) {
                        $pwUsed=$pw
                        # Re-extract with verbose output now that we have the right password
                        if (-not $settings.TestOnly) {
                            $result = Run-7z -Archive $archPath -OutDir $outDir -Pw $pw
                        }
                        break
                    }
                }
            }

            $sw.Stop(); $elapsed = "{0:mm\:ss}" -f $sw.Elapsed

            if ($result.Success) {
                $completed++
                $fc = if ($result.FileCount -gt 0) { " ($($result.FileCount) files)" } else { '' }
                $totalFiles += $result.FileCount
                $st = if ($settings.TestOnly) { "Test OK$fc" } elseif ($pwUsed) { "Success (pw)$fc" } else { "Success$fc" }
                Send-Msg 'QueueUpdate' @{Index=$idx;Status=$st;Elapsed=$elapsed}
                Send-Msg 'Log' @{Message="  Done: $archName ($elapsed) $fc";Color='#FF6BCB77'}

                if (-not $settings.TestOnly) {
                    # Duplicate folder removal
                    if ($settings.RemoveDupe) {
                        $aName = [IO.Path]::GetFileNameWithoutExtension($archPath)
                        if ($aName -match '\.tar$') { $aName = [IO.Path]::GetFileNameWithoutExtension($aName) }
                        $inner = Join-Path $outDir $aName
                        if (Test-Path $inner) {
                            $items = @(Get-ChildItem $outDir -Force)
                            if ($items.Count -eq 1 -and $items[0].PSIsContainer -and $items[0].Name -eq $aName) {
                                try { $tmp="${outDir}_flat$(Get-Random)"; Rename-Item $outDir $tmp -Force; Move-Item (Join-Path $tmp $aName) $outDir -Force; Remove-Item $tmp -Recurse -Force -EA SilentlyContinue } catch {}
                            }
                        }
                    }

                    # Single file rename
                    if ($settings.RenameSingle -and (Test-Path $outDir)) {
                        $files = @(Get-ChildItem $outDir -File -Force)
                        $dirs = @(Get-ChildItem $outDir -Directory -Force)
                        if ($files.Count -eq 1 -and $dirs.Count -eq 0) {
                            $aName = [IO.Path]::GetFileNameWithoutExtension($archPath)
                            $newName = $aName + $files[0].Extension
                            try { Rename-Item $files[0].FullName $newName -Force; Send-Msg 'Log' @{Message="  Renamed: $($files[0].Name) -> $newName";Color='#FF8888A0'} } catch {}
                        }
                    }

                    # Nested extraction
                    if ($settings.NestedEnabled) {
                        $seen = New-Object 'System.Collections.Generic.HashSet[string]'
                        $nq = @(Get-ChildItem -Path $outDir -Recurse -File -EA SilentlyContinue |
                            Where-Object {
                                $_.Extension.ToLower() -in $settings.ArchiveExts -and
                                -not ($_.Name -match '\.part(\d+)\.rar$' -and [int]$Matches[1] -gt 1) -and
                                -not ($_.Name -match '\.[rs]\d{2}$') -and
                                -not ($_.Name -match '\.\w+\.(\d{3})$' -and $Matches[1] -ne '001')
                            } | ForEach-Object {
                                try { $h=(Get-FileHash $_.FullName -Algorithm MD5 -EA Stop).Hash; if ($seen.Add($h)){$_.FullName} } catch { $_.FullName }
                            })
                        $depth=1
                        while ($nq.Count -gt 0 -and $depth -le $settings.NestedDepth -and -not $state.StopRequested) {
                            Send-Msg 'Log' @{Message="  Nested depth $depth : $($nq.Count) archive(s)";Color='#FFFFD93D'}
                            $next=@()
                            foreach ($na in $nq) {
                                if ($state.StopRequested) { break }
                                $nn=[IO.Path]::GetFileNameWithoutExtension($na)
                                $no=Join-Path ([IO.Path]::GetDirectoryName($na)) $nn
                                Send-Msg 'Log' @{Message="    Nested: $([IO.Path]::GetFileName($na))";Color='#FFCBA6F7'}
                                if (-not (Test-Path $no)) { New-Item $no -ItemType Directory -Force | Out-Null }
                                $nr = Run-7z -Archive $na -OutDir $no
                                if ($nr.NeedsPassword -and $settings.Passwords.Count -gt 0) { foreach ($pw in $settings.Passwords) { $nr = Run-7z -Archive $na -OutDir $no -Pw $pw -Silent; if ($nr.Success) { $nr = Run-7z -Archive $na -OutDir $no -Pw $pw; break } } }
                                if ($nr.Success) {
                                    if ($settings.NestedApplyPost) { Apply-PostAction $na ([IO.Path]::GetFileName($na)) }
                                    else { Remove-Item $na -Force -EA SilentlyContinue }
                                    Get-ChildItem -Path $no -Recurse -File -EA SilentlyContinue |
                                        Where-Object {
                                            $_.Extension.ToLower() -in $settings.ArchiveExts -and
                                            -not ($_.Name -match '\.part(\d+)\.rar$' -and [int]$Matches[1] -gt 1) -and
                                            -not ($_.Name -match '\.[rs]\d{2}$') -and
                                            -not ($_.Name -match '\.\w+\.(\d{3})$' -and $Matches[1] -ne '001')
                                        } | ForEach-Object {
                                            try { $h=(Get-FileHash $_.FullName -Algorithm MD5 -EA Stop).Hash; if ($seen.Add($h)){$next+=$_.FullName} } catch { $next+=$_.FullName }
                                        }
                                }
                            }
                            $nq=$next; $depth++
                        }
                    }

                    # Post actions
                    Apply-PostAction $archPath $archName

                    # Open destination
                    if ($settings.OpenDestAfter -and (Test-Path $outDir)) {
                        try { Start-Process explorer.exe $outDir } catch {}
                    }

                    # External processors
                    foreach ($ep in $settings.ExternalProcs) {
                        $epExt = $ep.Extension; $epCmd = $ep.Command
                        if ($epExt -and $epCmd) {
                            $aExt = [IO.Path]::GetExtension($archPath).TrimStart('.')
                            if ($aExt -ieq $epExt) {
                                $cmd = $epCmd -replace '\{ArchivePath\}',$archPath -replace '\{Destination\}',$outDir
                                try { Start-Process cmd.exe -ArgumentList "/c $cmd" -WindowStyle Hidden -Wait } catch {}
                            }
                        }
                    }
                }
            } else {
                $failed++
                $em = if ($result.NeedsPassword) { "Failed - Password" } else { "Failed" }
                Send-Msg 'QueueUpdate' @{Index=$idx;Status=$em;Elapsed=$elapsed}
                $errText = $result.Error; if ($errText -and $errText.Length -gt 100) { $errText=$errText.Substring(0,100) }
                Send-Msg 'Log' @{Message="  FAIL: $archName - $errText";Color='#FFFF6B6B'}
                # Delete broken files
                if ($settings.DeleteBroken -and -not $settings.TestOnly -and (Test-Path $outDir)) {
                    try { Remove-Item $outDir -Recurse -Force -EA Stop; Send-Msg 'Log' @{Message="  Cleaned broken output: $outDir";Color='#FF8888A0'} } catch {}
                }
            }
        }

        $batchSW.Stop(); $totalElapsed = "{0:mm\:ss}" -f $batchSW.Elapsed
        $fm = if ($state.StopRequested) { "Stopped ($completed done, $failed failed) in $totalElapsed" } else { "Complete: $completed extracted, $failed failed ($totalFiles files) in $totalElapsed" }
        $fc = if ($failed -gt 0) { '#FFFFD93D' } elseif ($state.StopRequested) { '#FFFF6B6B' } else { '#FF6BCB77' }
        Send-Msg 'Log' @{Message="--- $fm ---";Color=$fc}
        Send-Msg 'Status' @{Text=$fm;Color=$fc}; Send-Msg 'Progress' @{Text=''}; Send-Msg 'Log' @{Message=$fm;Color=$fc}
        Send-Msg 'ProgressBarUpdate' @{Percent=100}
        Send-Msg 'TrayTextUpdate' @{Text="ExtractorX - $fm"}
        Send-Msg 'TrayTip' @{Text=$fm}
        Send-Msg 'ExtractionDone' @{Completed=$completed;Failed=$failed}
    })
    $ps.Runspace = $rs; $script:extractHandle = $ps.BeginInvoke()
    $script:extractPS=$ps; $script:extractRS=$rs
}


# =====================================================================
# Event Handlers
# =====================================================================

# Drag & Drop
$window.Add_DragEnter({ if ($_.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) { $_.Effects=[Windows.DragDropEffects]::Copy; $ui['DropOverlay'].Visibility='Visible' } })
$window.Add_DragLeave({ $ui['DropOverlay'].Visibility='Collapsed' })
$window.Add_Drop({
    $ui['DropOverlay'].Visibility='Collapsed'
    if ($_.Data.GetDataPresent([Windows.DataFormats]::FileDrop)) {
        $files = @($_.Data.GetData([Windows.DataFormats]::FileDrop))
        Start-ScanJob -Paths $files -AutoExtractAfter:($script:Config.AutoExtractOnDrop)
    }
})

# Toolbar buttons
$ui['BtnAddFiles'].Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog; $dlg.Title="Select Archives"
    $dlg.Filter="Archives|*.zip;*.7z;*.rar;*.tar;*.gz;*.tgz;*.bz2;*.xz;*.iso;*.cab;*.wim|All|*.*"; $dlg.Multiselect=$true
    if ($dlg.ShowDialog()) { Start-ScanJob -Paths $dlg.FileNames }
})
$ui['BtnAddFolder'].Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog; $dlg.Description="Scan for archives"
    if ($dlg.ShowDialog() -eq 'OK') { Start-ScanJob -Paths @($dlg.SelectedPath) }
})
$ui['BtnExtractAll'].Add_Click({ Start-QueueExtraction })
$ui['BtnStopAll'].Add_Click({
    $script:State.StopRequested=$true
    $script:State.ScanCancelRequested=$true
    Send-UIMessage -Type 'Status' -Data @{Text="Stopping...";Color='#FFFF6B6B'}
    Send-UIMessage -Type 'Log' -Data @{Message="Stop requested";Color='#FFFF6B6B'}
})
$ui['BtnRemoveSelected'].Add_Click({
    $sel=@($ui['QueueList'].SelectedItems); foreach ($s in $sel) { [void]$script:QueuedPaths.Remove($s.FullPath); $ui['QueueList'].Items.Remove($s) }
    Send-UIMessage -Type 'QueueCountRefresh'
})
$ui['BtnClearQueue'].Add_Click({ $ui['QueueList'].Items.Clear(); $script:QueuedPaths.Clear(); Send-UIMessage -Type 'QueueCountRefresh' })
$ui['BtnBrowseOutput'].Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq 'OK') { $ui['TxtOutputPath'].Text = $dlg.SelectedPath }
})

# Macro insert button
$ui['BtnMacroInsert'].Add_Click({ $ui['MacroMenu'].IsOpen = $true })
foreach ($mi in $ui['MacroMenu'].Items) {
    if ($mi -is [System.Windows.Controls.MenuItem] -and $mi.Tag) {
        $mi.Add_Click({ $ui['TxtOutputPath'].Text += $_.Source.Tag }.GetNewClosure())
    }
}

# Queue context menu
$ui['CtxAddFiles'].Add_Click({
    $dlg = New-Object Microsoft.Win32.OpenFileDialog; $dlg.Title="Add Archives"; $dlg.Multiselect=$true
    $dlg.Filter="Archives|*.zip;*.7z;*.rar;*.tar;*.gz;*.tgz;*.bz2;*.xz;*.iso;*.cab;*.wim|All|*.*"
    if ($dlg.ShowDialog()) { Start-ScanJob -Paths $dlg.FileNames }
})
$ui['CtxSearchFolder'].Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog; $dlg.Description="Search for archives"
    if ($dlg.ShowDialog() -eq 'OK') { Start-ScanJob -Paths @($dlg.SelectedPath) }
})
$ui['CtxExtractUnextracted'].Add_Click({ Start-QueueExtraction -UnextractedOnly })
$ui['CtxTestOnly'].Add_Click({ $script:State.TestOnly = [bool]$ui['CtxTestOnly'].IsChecked })
$ui['CtxOpenFolder'].Add_Click({
    $sel = $ui['QueueList'].SelectedItem
    if ($sel -and $sel.Directory -and (Test-Path $sel.Directory)) { Start-Process explorer.exe $sel.Directory }
})
$ui['CtxRemoveItem'].Add_Click({
    $sel=@($ui['QueueList'].SelectedItems); foreach ($s in $sel) { [void]$script:QueuedPaths.Remove($s.FullPath); $ui['QueueList'].Items.Remove($s) }
    Send-UIMessage -Type 'QueueCountRefresh'
})

# Log buttons
$ui['BtnClearLog'].Add_Click({ $ui['LogBox'].Inlines.Clear() })
$ui['BtnExportLog'].Add_Click({
    $dlg = New-Object Microsoft.Win32.SaveFileDialog; $dlg.Filter="Text|*.txt"; $dlg.FileName="ExtractorX_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    if ($dlg.ShowDialog()) {
        $sb=[System.Text.StringBuilder]::new()
        foreach ($il in $ui['LogBox'].Inlines) { if ($il -is [System.Windows.Documents.Run]) { [void]$sb.Append($il.Text) } }
        $sb.ToString() | Set-Content $dlg.FileName -Force
    }
})

# Keyboard
$window.Add_KeyDown({
    if ($_.Key -eq 'V' -and [System.Windows.Input.Keyboard]::Modifiers -eq 'Control') {
        if ([System.Windows.Clipboard]::ContainsFileDropList()) { Start-ScanJob -Paths @([System.Windows.Clipboard]::GetFileDropList()) }
    }
    if ($_.Key -eq 'Delete' -and -not $script:State.IsExtracting) {
        $sel=@($ui['QueueList'].SelectedItems); foreach ($s in $sel) { [void]$script:QueuedPaths.Remove($s.FullPath); $ui['QueueList'].Items.Remove($s) }
        Send-UIMessage -Type 'QueueCountRefresh'
    }
})

# =====================================================================
# Double-click queue item to open output folder
# =====================================================================
$ui['QueueList'].Add_MouseDoubleClick({
    $sel = $ui['QueueList'].SelectedItem
    if ($sel -and $sel.FullPath -and ($sel.Status -like '*Success*' -or $sel.Status -eq 'Test OK')) {
        $outDir = Resolve-OutputPath -Template $ui['TxtOutputPath'].Text -ArchivePath $sel.FullPath
        if (Test-Path $outDir) { Start-Process explorer.exe $outDir }
        elseif (Test-Path ([IO.Path]::GetDirectoryName($outDir))) { Start-Process explorer.exe ([IO.Path]::GetDirectoryName($outDir)) }
    } elseif ($sel -and $sel.Directory -and (Test-Path $sel.Directory)) {
        Start-Process explorer.exe $sel.Directory
    }
})

# =====================================================================
# Column Sorting
# =====================================================================
$script:SortColumn = ''
$script:SortAscending = $true
$ui['QueueList'].AddHandler(
    [System.Windows.Controls.GridViewColumnHeader]::ClickEvent,
    [System.Windows.RoutedEventHandler]{
        param($sender, $e)
        $header = $e.OriginalSource
        if ($header -isnot [System.Windows.Controls.GridViewColumnHeader]) { return }
        if (-not $header.Column) { return }
        $binding = $header.Column.DisplayMemberBinding
        if (-not $binding) { return }
        $prop = $binding.Path.Path
        if ($prop -eq $script:SortColumn) { $script:SortAscending = -not $script:SortAscending }
        else { $script:SortColumn = $prop; $script:SortAscending = $true }

        $items = @($ui['QueueList'].Items | ForEach-Object { $_ })
        if ($script:SortAscending) {
            $sorted = $items | Sort-Object -Property $prop
        } else {
            $sorted = $items | Sort-Object -Property $prop -Descending
        }
        $ui['QueueList'].Items.Clear()
        foreach ($item in $sorted) { $ui['QueueList'].Items.Add($item) }
    }
)

# =====================================================================
# Selection Changed - show count + size
# =====================================================================
$ui['QueueList'].Add_SelectionChanged({
    $cnt = $ui['QueueList'].SelectedItems.Count
    if ($cnt -gt 0) {
        $selBytes = [long]0
        foreach ($item in $ui['QueueList'].SelectedItems) { if ($item.SizeBytes) { $selBytes += [long]$item.SizeBytes } }
        $sizeStr = if ($selBytes -gt 0) { " ($(Get-FileSizeString $selBytes))" } else { '' }
        $ui['SelectionInfo'].Text = "$cnt selected$sizeStr"
    } else {
        $ui['SelectionInfo'].Text = ''
    }
})

# =====================================================================
# Log Panel Toggle
# =====================================================================
$script:LogVisible = $true
$script:LogRowDef = $window.FindName('LogRowDef')
$ui['BtnToggleLog'].Add_Click({
    if ($script:LogVisible) {
        $script:LogRowDef.Height = [System.Windows.GridLength]::new(0)
        $ui['BtnToggleLog'].Content = "Show Log"
        $script:LogVisible = $false
    } else {
        $script:LogRowDef.Height = [System.Windows.GridLength]::new(200)
        $ui['BtnToggleLog'].Content = "Hide"
        $script:LogVisible = $true
    }
})

# =====================================================================
# About Dialog
# =====================================================================
$ui['BtnAbout'].Add_Click({
    $aw = New-Object System.Windows.Window
    $aw.Title = "About ExtractorX"; $aw.Width = 420; $aw.Height = 320; $aw.ResizeMode = 'NoResize'
    $aw.WindowStartupLocation = 'CenterOwner'; $aw.Owner = $window
    $aw.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF0A0A12')
    Apply-DarkTheme $aw
    $sp = New-Object System.Windows.Controls.StackPanel; $sp.Margin = "32,28"
    $hdr = New-Object System.Windows.Controls.StackPanel; $hdr.Orientation = 'Horizontal'; $hdr.Margin = "0,0,0,4"
    $t1a = New-Object System.Windows.Controls.TextBlock; $t1a.Text = "EXTRACTOR"; $t1a.FontSize = 24; $t1a.FontWeight = 'Bold'
    $t1a.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FFEEEEF4')
    $t1b = New-Object System.Windows.Controls.TextBlock; $t1b.Text = "X"; $t1b.FontSize = 24; $t1b.FontWeight = 'Bold'
    $t1b.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF00B4D8')
    $hdr.Children.Add($t1a); $hdr.Children.Add($t1b); $sp.Children.Add($hdr)
    $t2 = New-Object System.Windows.Controls.TextBlock; $t2.Text = "v$($script:AppVersion)"; $t2.FontSize = 13
    $t2.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF6E6E84'); $t2.Margin = "0,0,0,18"
    $sp.Children.Add($t2)
    $sep = New-Object System.Windows.Controls.Border; $sep.Height = 1; $sep.Margin = "0,0,0,18"
    $sep.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF2A2A3C')
    $sp.Children.Add($sep)
    $t3 = New-Object System.Windows.Controls.TextBlock; $t3.TextWrapping = 'Wrap'; $t3.FontSize = 12; $t3.LineHeight = 20
    $t3.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FFD4D4E0')
    $t3.Text = "Open-source bulk archive extraction tool.`nInspired by ExtractNow by Nathan Moinvaziri.`n`nPowered by 7-Zip (Igor Pavlov, LGPL).`nBuilt with PowerShell WPF."
    $t3.Margin = "0,0,0,18"
    $sp.Children.Add($t3)
    $t4 = New-Object System.Windows.Controls.TextBlock; $t4.Text = "github.com/SysAdminDoc"; $t4.FontSize = 12
    $t4.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF00B4D8'); $t4.Margin = "0,0,0,6"
    $t5 = New-Object System.Windows.Controls.TextBlock; $t5.Text = "MIT License"; $t5.FontSize = 11
    $t5.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF6E6E84')
    @($t4,$t5) | ForEach-Object { $sp.Children.Add($_) }
    $aw.Content = $sp; [void]$aw.ShowDialog()
})

# =====================================================================
# Open Output Folder (context menu)
# =====================================================================
$ui['CtxOpenOutput'].Add_Click({
    $sel = $ui['QueueList'].SelectedItem
    if ($sel -and $sel.FullPath) {
        $outDir = Resolve-OutputPath -Template $ui['TxtOutputPath'].Text -ArchivePath $sel.FullPath
        if (Test-Path $outDir) { Start-Process explorer.exe $outDir }
        elseif (Test-Path ([IO.Path]::GetDirectoryName($outDir))) { Start-Process explorer.exe ([IO.Path]::GetDirectoryName($outDir)) }
    }
})

# =====================================================================
# Dark Theme ResourceDictionary for child windows
# =====================================================================
function Apply-DarkTheme {
    param([System.Windows.Window]$Win)
    $Win.FontFamily = [System.Windows.Media.FontFamily]::new('Segoe UI')
    [xml]$rdXaml = @"
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <SolidColorBrush x:Key="Acc" Color="#FF00B4D8"/>
    <SolidColorBrush x:Key="Sf3" Color="#FF1E1E2A"/>
    <SolidColorBrush x:Key="Sf4" Color="#FF262636"/>
    <SolidColorBrush x:Key="Bd" Color="#FF2A2A3C"/>
    <SolidColorBrush x:Key="BdH" Color="#FF3A3A50"/>
    <SolidColorBrush x:Key="Tx" Color="#FFD4D4E0"/>
    <SolidColorBrush x:Key="Td" Color="#FF6E6E84"/>
    <Style x:Key="ScrollBarThumb" TargetType="Thumb">
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Thumb">
            <Border x:Name="tb" Background="#FF3A3A4E" CornerRadius="4" Margin="2"/>
            <ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="tb" Property="Background" Value="#FF50506A"/></Trigger></ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style TargetType="ScrollBar">
        <Setter Property="Background" Value="Transparent"/><Setter Property="Width" Value="10"/>
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ScrollBar">
            <Border Background="#08FFFFFF" CornerRadius="5"><Track x:Name="PART_Track" IsDirectionReversed="True">
                <Track.Thumb><Thumb Style="{StaticResource ScrollBarThumb}"/></Track.Thumb></Track></Border>
        </ControlTemplate></Setter.Value></Setter>
        <Style.Triggers><Trigger Property="Orientation" Value="Horizontal">
            <Setter Property="Height" Value="10"/><Setter Property="Width" Value="Auto"/>
        </Trigger></Style.Triggers>
    </Style>
    <Style TargetType="Button">
        <Setter Property="Background" Value="{StaticResource Sf3}"/><Setter Property="Foreground" Value="{StaticResource Tx}"/>
        <Setter Property="BorderThickness" Value="1"/><Setter Property="BorderBrush" Value="{StaticResource Bd}"/>
        <Setter Property="Padding" Value="14,7"/><Setter Property="FontSize" Value="12"/><Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/></Border>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="{StaticResource Sf4}"/><Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource BdH}"/></Trigger>
                <Trigger Property="IsPressed" Value="True"><Setter TargetName="bd" Property="Background" Value="#FF2E2E42"/></Trigger>
                <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.35"/></Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style TargetType="TextBox">
        <Setter Property="Background" Value="#FF0E0E16"/><Setter Property="Foreground" Value="{StaticResource Tx}"/>
        <Setter Property="BorderBrush" Value="{StaticResource Bd}"/><Setter Property="Padding" Value="10,7"/>
        <Setter Property="CaretBrush" Value="{StaticResource Acc}"/><Setter Property="FontSize" Value="12"/><Setter Property="SelectionBrush" Value="#FF0090AA"/>
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="TextBox">
            <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="6" Padding="{TemplateBinding Padding}">
                <ScrollViewer x:Name="PART_ContentHost" Margin="0"/></Border>
            <ControlTemplate.Triggers><Trigger Property="IsFocused" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="{StaticResource Acc}"/></Trigger></ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style TargetType="CheckBox">
        <Setter Property="Foreground" Value="{StaticResource Tx}"/><Setter Property="FontSize" Value="12"/><Setter Property="Margin" Value="0,5"/><Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="CheckBox">
            <StackPanel Orientation="Horizontal">
                <Border x:Name="box" Width="18" Height="18" CornerRadius="4" BorderThickness="1.5" BorderBrush="{StaticResource Bd}" Background="#FF0E0E16" VerticalAlignment="Center" Margin="0,0,8,0">
                    <TextBlock x:Name="check" Text="&#x2713;" FontSize="12" FontWeight="Bold" Foreground="{StaticResource Acc}" HorizontalAlignment="Center" VerticalAlignment="Center" Visibility="Collapsed"/></Border>
                <ContentPresenter VerticalAlignment="Center"/></StackPanel>
            <ControlTemplate.Triggers>
                <Trigger Property="IsChecked" Value="True"><Setter TargetName="check" Property="Visibility" Value="Visible"/><Setter TargetName="box" Property="BorderBrush" Value="{StaticResource Acc}"/><Setter TargetName="box" Property="Background" Value="#FF0D2A33"/></Trigger>
                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="box" Property="BorderBrush" Value="{StaticResource BdH}"/></Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style TargetType="ComboBox">
        <Setter Property="Background" Value="#FF0E0E16"/><Setter Property="Foreground" Value="{StaticResource Tx}"/>
        <Setter Property="BorderBrush" Value="{StaticResource Bd}"/><Setter Property="FontSize" Value="12"/>
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ComboBox">
            <Grid><ToggleButton x:Name="ToggleButton" Focusable="False" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}" ClickMode="Press">
                <ToggleButton.Template><ControlTemplate TargetType="ToggleButton">
                    <Border Background="#FF0E0E16" BorderBrush="{StaticResource Bd}" BorderThickness="1" CornerRadius="6" Padding="8,7">
                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="20"/></Grid.ColumnDefinitions>
                            <TextBlock Grid.Column="1" Text="&#x25BC;" FontSize="9" Foreground="{StaticResource Td}" HorizontalAlignment="Center" VerticalAlignment="Center"/></Grid></Border>
                </ControlTemplate></ToggleButton.Template></ToggleButton>
            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" Margin="10,7,28,6" VerticalAlignment="Center" HorizontalAlignment="Left"/>
            <Popup x:Name="Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                <Border Background="#FF1E1E2A" BorderBrush="{StaticResource BdH}" BorderThickness="1" CornerRadius="6" Margin="0,2,0,0" MaxHeight="300" MinWidth="{TemplateBinding ActualWidth}">
                    <ScrollViewer><ItemsPresenter/></ScrollViewer></Border></Popup></Grid>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style TargetType="ComboBoxItem">
        <Setter Property="Background" Value="Transparent"/><Setter Property="Foreground" Value="{StaticResource Tx}"/>
        <Setter Property="Padding" Value="10,7"/><Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ComboBoxItem">
            <Border x:Name="bd" Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
                <ContentPresenter/></Border>
            <ControlTemplate.Triggers><Trigger Property="IsHighlighted" Value="True"><Setter TargetName="bd" Property="Background" Value="{StaticResource Sf4}"/></Trigger></ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style TargetType="ListBox">
        <Setter Property="Background" Value="#FF0E0E16"/><Setter Property="BorderBrush" Value="{StaticResource Bd}"/><Setter Property="Foreground" Value="{StaticResource Tx}"/><Setter Property="FontSize" Value="12"/>
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ListBox">
            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="6" Padding="4">
                <ScrollViewer Focusable="False"><ItemsPresenter/></ScrollViewer></Border>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style TargetType="ListBoxItem">
        <Setter Property="Foreground" Value="{StaticResource Tx}"/><Setter Property="Padding" Value="8,5"/><Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ListBoxItem">
            <Border x:Name="bd" Background="Transparent" Padding="{TemplateBinding Padding}" CornerRadius="4">
                <ContentPresenter/></Border>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#FF1A1A28"/></Trigger>
                <Trigger Property="IsSelected" Value="True"><Setter TargetName="bd" Property="Background" Value="#FF141430"/></Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style TargetType="TabControl">
        <Setter Property="Background" Value="Transparent"/><Setter Property="BorderThickness" Value="0"/>
    </Style>
    <Style TargetType="TabItem">
        <Setter Property="Foreground" Value="{StaticResource Td}"/><Setter Property="FontSize" Value="12"/>
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="TabItem">
            <Border x:Name="bd" Padding="14,9" Margin="0,0,2,0" CornerRadius="6,6,0,0" Background="Transparent" Cursor="Hand">
                <ContentPresenter ContentSource="Header"/></Border>
            <ControlTemplate.Triggers><Trigger Property="IsSelected" Value="True"><Setter TargetName="bd" Property="Background" Value="{StaticResource Sf3}"/><Setter Property="Foreground" Value="{StaticResource Acc}"/><Setter Property="FontWeight" Value="SemiBold"/></Trigger>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#FF181825"/></Trigger></ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
    </Style>
    <Style TargetType="ToolTip">
        <Setter Property="Background" Value="#FF1A1A26"/><Setter Property="Foreground" Value="{StaticResource Tx}"/>
        <Setter Property="BorderBrush" Value="{StaticResource Bd}"/><Setter Property="Padding" Value="10,6"/>
    </Style>
</ResourceDictionary>
"@
    $reader = New-Object System.Xml.XmlNodeReader $rdXaml
    $rd = [System.Windows.Markup.XamlReader]::Load($reader)
    $Win.Resources.MergedDictionaries.Add($rd)
}

# =====================================================================
# Settings Window
# =====================================================================
$ui['BtnSettings'].Add_Click({
    $sw = New-Object System.Windows.Window
    $sw.Title = "ExtractorX Settings"; $sw.Width = 680; $sw.Height = 600
    $sw.WindowStartupLocation = 'CenterOwner'; $sw.Owner = $window
    $sw.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF0A0A12')
    $sw.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FFD4D4E0')
    $sw.ResizeMode = 'NoResize'
    Apply-DarkTheme $sw

    $cfg = $script:Config  # reference

    # Build settings UI in code (avoids massive XAML)
    $mainGrid = New-Object System.Windows.Controls.Grid
    $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height='*'}))
    $mainGrid.RowDefinitions.Add((New-Object System.Windows.Controls.RowDefinition -Property @{Height='Auto'}))

    $tabs = New-Object System.Windows.Controls.TabControl
    $tabs.Background = [System.Windows.Media.Brushes]::Transparent; $tabs.BorderThickness = "0"

    # Helper: create a settings tab
    function New-SettingsTab($Header) {
        $ti = New-Object System.Windows.Controls.TabItem; $ti.Header = "  $Header  "
        $sv = New-Object System.Windows.Controls.ScrollViewer; $sv.VerticalScrollBarVisibility = 'Auto'; $sv.Padding = "16"
        $sp = New-Object System.Windows.Controls.StackPanel; $sp.MaxWidth = 600; $sp.HorizontalAlignment = 'Left'
        $sv.Content = $sp; $ti.Content = $sv
        return @{Tab=$ti; Panel=$sp}
    }
    function New-Label($Text, [switch]$Header) {
        $tb = New-Object System.Windows.Controls.TextBlock; $tb.Text = $Text; $tb.Margin = "0,14,0,6"
        if ($Header) { $tb.FontSize=10; $tb.FontWeight='Bold'; $tb.Foreground=[System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF00B4D8'); $tb.Margin="0,18,0,8" }
        else { $tb.Foreground=[System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF6E6E84'); $tb.FontSize=12 }
        return $tb
    }
    function New-Check($Text, [bool]$Checked) {
        $cb = New-Object System.Windows.Controls.CheckBox; $cb.Content=$Text; $cb.IsChecked=$Checked
        $cb.Foreground=[System.Windows.Media.BrushConverter]::new().ConvertFromString('#FFD4D4E0'); $cb.Margin="0,5"; $cb.FontSize=12
        return $cb
    }

    # --- General Tab ---
    $gen = New-SettingsTab "General"
    $gen.Panel.Children.Add((New-Label "W I N D O W" -Header))
    $chkAlwaysTop = New-Check "Always show on top" $cfg.AlwaysOnTop
    $chkMinToTray = New-Check "Minimize to system tray" $cfg.MinimizeToTray
    $chkLogHist   = New-Check "Log extraction history" $cfg.LogHistory
    $chkAutoSwitch = New-Check "Auto-switch to history view on extract" $cfg.AutoSwitchToHistory
    $chkDeepDetect = New-Check "Deep archive detection (magic bytes)" $cfg.DeepArchiveDetection
    @($chkAlwaysTop,$chkMinToTray,$chkLogHist,$chkAutoSwitch,$chkDeepDetect) | ForEach-Object { $gen.Panel.Children.Add($_) }
    $tabs.Items.Add($gen.Tab)

    # --- Destination Tab ---
    $dest = New-SettingsTab "Destination"
    $dest.Panel.Children.Add((New-Label "O U T P U T   P A T H" -Header))
    $dest.Panel.Children.Add((New-Label "Uses macros. See main window for macro insertion."))
    $dest.Panel.Children.Add((New-Label "O V E R W R I T E" -Header))
    $cmbOver = New-Object System.Windows.Controls.ComboBox; $cmbOver.Width=250; $cmbOver.HorizontalAlignment='Left'; $cmbOver.Margin="0,4"
    @('Always overwrite','Never overwrite','Keep and rename') | ForEach-Object { $cmbOver.Items.Add($_) | Out-Null }
    switch ($cfg.OverwriteMode) { 'Always'{$cmbOver.SelectedIndex=0} 'Never'{$cmbOver.SelectedIndex=1} 'Rename'{$cmbOver.SelectedIndex=2} default{$cmbOver.SelectedIndex=0} }
    $dest.Panel.Children.Add($cmbOver)
    $tabs.Items.Add($dest.Tab)

    # --- Process Tab ---
    $proc = New-SettingsTab "Process"
    $proc.Panel.Children.Add((New-Label "A R C H I V E   R E C U R S I O N" -Header))
    $chkNested = New-Check "Extract archives within archives" $cfg.NestedExtraction
    $chkNestedPost = New-Check "Apply post-action to nested archives" $cfg.NestedApplyPostAction
    $proc.Panel.Children.Add($chkNested); $proc.Panel.Children.Add($chkNestedPost)
    $proc.Panel.Children.Add((New-Label "Nested max depth:"))
    $cmbDepth = New-Object System.Windows.Controls.ComboBox; $cmbDepth.Width=100; $cmbDepth.HorizontalAlignment='Left'; $cmbDepth.Margin="0,4"
    @('3','5','10','20') | ForEach-Object { $cmbDepth.Items.Add($_) | Out-Null }
    switch ($cfg.NestedMaxDepth) { 3{$cmbDepth.SelectedIndex=0} 10{$cmbDepth.SelectedIndex=2} 20{$cmbDepth.SelectedIndex=3} default{$cmbDepth.SelectedIndex=1} }
    $proc.Panel.Children.Add($cmbDepth)

    $proc.Panel.Children.Add((New-Label "S U C C E S S F U L   E X T R A C T I O N" -Header))
    $cmbPost = New-Object System.Windows.Controls.ComboBox; $cmbPost.Width=250; $cmbPost.HorizontalAlignment='Left'; $cmbPost.Margin="0,4"
    @('Do nothing','Move to Recycle Bin','Move to folder','Delete permanently') | ForEach-Object { $cmbPost.Items.Add($_) | Out-Null }
    switch ($cfg.PostAction) { 'None'{$cmbPost.SelectedIndex=0} 'Recycle'{$cmbPost.SelectedIndex=1} 'MoveToFolder'{$cmbPost.SelectedIndex=2} 'Delete'{$cmbPost.SelectedIndex=3} default{$cmbPost.SelectedIndex=0} }
    $proc.Panel.Children.Add($cmbPost)
    $proc.Panel.Children.Add((New-Label "Move-to folder path:"))
    $txtPostFolder = New-Object System.Windows.Controls.TextBox; $txtPostFolder.Text=$cfg.PostActionFolder; $txtPostFolder.Margin="0,4"
    $proc.Panel.Children.Add($txtPostFolder)
    $chkOpenDest = New-Check "Open destination folder after extraction" $cfg.OpenDestAfterExtract
    $proc.Panel.Children.Add($chkOpenDest)

    $proc.Panel.Children.Add((New-Label "C L E A N U P" -Header))
    $chkRemoveDupe = New-Check "Remove duplicate archive name folder" $cfg.RemoveDuplicateFolder
    $chkRenameSingle = New-Check "Rename single file after archive name" $cfg.RenameSingleFile
    $chkDelBroken = New-Check "Delete extracted files that are broken" $cfg.DeleteBrokenFiles
    @($chkRemoveDupe,$chkRenameSingle,$chkDelBroken) | ForEach-Object { $proc.Panel.Children.Add($_) }

    $proc.Panel.Children.Add((New-Label "B A T C H   C O M P L E T E" -Header))
    $chkCloseOk = New-Check "Close program on success" $cfg.CloseOnComplete
    $chkCloseAlways = New-Check "Close program even if not successful" $cfg.CloseOnCompleteAlways
    $chkClearList = New-Check "Clear archive list on completion" $cfg.ClearListOnComplete
    @($chkCloseOk,$chkCloseAlways,$chkClearList) | ForEach-Object { $proc.Panel.Children.Add($_) }
    $tabs.Items.Add($proc.Tab)

    # --- Explorer Tab ---
    $expl = New-SettingsTab "Explorer"
    $expl.Panel.Children.Add((New-Label "C O N T E X T   M E N U" -Header))
    $chkCtxEnabled = New-Check "Enable context menu" $cfg.CtxEnabled
    $chkCtxGrouped = New-Check "Group entries in submenu" $cfg.CtxGrouped
    $chkCtxHere    = New-Check "Show: Extract here" $cfg.CtxExtractHere
    $chkCtxFolder  = New-Check "Show: Extract to folder" $cfg.CtxExtractToFolder
    $chkCtxEnqueue = New-Check "Show: Add to ExtractorX" $cfg.CtxEnqueue
    $chkCtxSearch  = New-Check "Show: Search for archives (directories)" $cfg.CtxSearchArchives
    @($chkCtxEnabled,$chkCtxGrouped,$chkCtxHere,$chkCtxFolder,$chkCtxEnqueue,$chkCtxSearch) | ForEach-Object { $expl.Panel.Children.Add($_) }
    $tabs.Items.Add($expl.Tab)

    # --- Drag & Drop Tab ---
    $dd = New-SettingsTab "Drag & Drop"
    $dd.Panel.Children.Add((New-Label "D R A G   &   D R O P" -Header))
    $chkAutoExtDrop = New-Check "Automatically extract dropped archives" $cfg.AutoExtractOnDrop
    $dd.Panel.Children.Add($chkAutoExtDrop)
    $dd.Panel.Children.Add((New-Label "Filter type:"))
    $cmbFilter = New-Object System.Windows.Controls.ComboBox; $cmbFilter.Width=200; $cmbFilter.HorizontalAlignment='Left'; $cmbFilter.Margin="0,4"
    @('None','Inclusion','Exclusion') | ForEach-Object { $cmbFilter.Items.Add($_) | Out-Null }
    switch ($cfg.DragDropFilterType) { 'Inclusion'{$cmbFilter.SelectedIndex=1} 'Exclusion'{$cmbFilter.SelectedIndex=2} default{$cmbFilter.SelectedIndex=0} }
    $dd.Panel.Children.Add($cmbFilter)
    $dd.Panel.Children.Add((New-Label "Filter mask (semicolon-separated):"))
    $txtFilterMask = New-Object System.Windows.Controls.TextBox; $txtFilterMask.Text=$cfg.DragDropFilterMask; $txtFilterMask.Margin="0,4"
    $dd.Panel.Children.Add($txtFilterMask)
    $tabs.Items.Add($dd.Tab)

    # --- Passwords Tab ---
    $pwTab = New-SettingsTab "Passwords"
    $pwTab.Panel.Children.Add((New-Label "P A S S W O R D   L I S T" -Header))
    $chkUsePw = New-Check "Use password list for encrypted archives" $cfg.UsePasswordList
    $chkPromptExh = New-Check "Prompt for password when list exhausted" $cfg.PromptOnExhaustion
    $chkOnePerArc = New-Check "Assume one password per archive (faster)" $cfg.AssumeOnePassword
    @($chkUsePw,$chkPromptExh,$chkOnePerArc) | ForEach-Object { $pwTab.Panel.Children.Add($_) }
    $pwTab.Panel.Children.Add((New-Label "Passwords ($($script:Passwords.Count) stored):"))

    $pwList = New-Object System.Windows.Controls.ListBox; $pwList.Height=150; $pwList.Margin="0,4"
    foreach ($pw in $script:Passwords) {
        $m = if ($pw.Length -le 4) { '*'*$pw.Length } else { $pw.Substring(0,2)+('*'*($pw.Length-4))+$pw.Substring($pw.Length-2) }
        $pwList.Items.Add($m) | Out-Null
    }
    $pwTab.Panel.Children.Add($pwList)

    $pwBtnPanel = New-Object System.Windows.Controls.WrapPanel; $pwBtnPanel.Margin="0,6,0,0"
    $btnPwAdd = New-Object System.Windows.Controls.Button; $btnPwAdd.Content="Add Password"; $btnPwAdd.Padding="12,6"; $btnPwAdd.Margin="0,0,6,0"
    $btnPwRemove = New-Object System.Windows.Controls.Button; $btnPwRemove.Content="Remove"; $btnPwRemove.Padding="12,6"; $btnPwRemove.Margin="0,0,6,0"
    $btnPwImport = New-Object System.Windows.Controls.Button; $btnPwImport.Content="Import File"; $btnPwImport.Padding="12,6"
    $pwBtnPanel.Children.Add($btnPwAdd); $pwBtnPanel.Children.Add($btnPwRemove); $pwBtnPanel.Children.Add($btnPwImport)
    $pwTab.Panel.Children.Add($pwBtnPanel)

    $btnPwAdd.Add_Click({
        $input = [Microsoft.VisualBasic.Interaction]::InputBox("Enter password:", "Add Password", "")
        if ($input) {
            $script:Passwords += $input
            $m = if ($input.Length -le 4) { '*'*$input.Length } else { $input.Substring(0,2)+('*'*($input.Length-4))+$input.Substring($input.Length-2) }
            $pwList.Items.Add($m); Save-Passwords -Passwords $script:Passwords
        }
    })
    $btnPwRemove.Add_Click({
        $idx = $pwList.SelectedIndex
        if ($idx -ge 0 -and $idx -lt $script:Passwords.Count) {
            $newList=@(); for ($i=0; $i -lt $script:Passwords.Count; $i++) { if ($i -ne $idx) { $newList+=$script:Passwords[$i] } }
            $script:Passwords=$newList; $pwList.Items.RemoveAt($idx); Save-Passwords -Passwords $script:Passwords
        }
    })
    $btnPwImport.Add_Click({
        $dlg = New-Object Microsoft.Win32.OpenFileDialog; $dlg.Filter="Text|*.txt|All|*.*"
        if ($dlg.ShowDialog()) {
            $c=0; Get-Content $dlg.FileName | Where-Object { $_.Trim() } | ForEach-Object {
                $pw=$_.Trim(); if ($pw -notin $script:Passwords) {
                    $script:Passwords+=$pw; $c++
                    $m = if ($pw.Length -le 4) { '*'*$pw.Length } else { $pw.Substring(0,2)+('*'*($pw.Length-4))+$pw.Substring($pw.Length-2) }
                    $pwList.Items.Add($m)
                }
            }
            Save-Passwords -Passwords $script:Passwords
        }
    })
    $tabs.Items.Add($pwTab.Tab)

    # --- Files Tab ---
    $files = New-SettingsTab "Files"
    $files.Panel.Children.Add((New-Label "E X C L U S I O N S" -Header))
    $files.Panel.Children.Add((New-Label "Files matching these masks will not be extracted (semicolon-separated):"))
    $txtExcl = New-Object System.Windows.Controls.TextBox; $txtExcl.Text=$cfg.FileExclusions; $txtExcl.Margin="0,4"
    $files.Panel.Children.Add($txtExcl)
    $tabs.Items.Add($files.Tab)

    # --- Monitor Tab ---
    $mon = New-SettingsTab "Monitor"
    $mon.Panel.Children.Add((New-Label "W A T C H   F O L D E R S" -Header))
    $mon.Panel.Children.Add((New-Label "Directories to monitor for new archives:"))
    $watchList = New-Object System.Windows.Controls.ListBox; $watchList.Height=150; $watchList.Margin="0,4"
    foreach ($wf in $cfg.WatchFolders) { $watchList.Items.Add($wf) | Out-Null }
    $mon.Panel.Children.Add($watchList)

    $wBtnPanel = New-Object System.Windows.Controls.WrapPanel; $wBtnPanel.Margin="0,6,0,0"
    $btnWAdd = New-Object System.Windows.Controls.Button; $btnWAdd.Content="+ Add"; $btnWAdd.Padding="12,6"; $btnWAdd.Margin="0,0,6,0"
    $btnWRemove = New-Object System.Windows.Controls.Button; $btnWRemove.Content="Remove"; $btnWRemove.Padding="12,6"
    $wBtnPanel.Children.Add($btnWAdd); $wBtnPanel.Children.Add($btnWRemove)
    $mon.Panel.Children.Add($wBtnPanel)

    $chkWatchAuto = New-Check "Automatically extract new archives" $cfg.WatchAutoExtract
    $mon.Panel.Children.Add($chkWatchAuto)

    $btnWAdd.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog; $dlg.Description="Select folder to monitor"
        if ($dlg.ShowDialog() -eq 'OK') {
            $dup=$false; foreach ($i in $watchList.Items) { if ($i -eq $dlg.SelectedPath) { $dup=$true; break } }
            if (-not $dup) { $watchList.Items.Add($dlg.SelectedPath) }
        }
    })
    $btnWRemove.Add_Click({ $s=$watchList.SelectedItem; if ($s) { $watchList.Items.Remove($s) } })
    $tabs.Items.Add($mon.Tab)

    # --- Advanced Tab ---
    $adv = New-SettingsTab "Advanced"
    $adv.Panel.Children.Add((New-Label "T H R E A D   P R I O R I T Y" -Header))
    $cmbPriority = New-Object System.Windows.Controls.ComboBox; $cmbPriority.Width=200; $cmbPriority.HorizontalAlignment='Left'; $cmbPriority.Margin="0,4"
    @('Low','BelowNormal','Normal','AboveNormal','High') | ForEach-Object { $cmbPriority.Items.Add($_) | Out-Null }
    switch ($cfg.ThreadPriority) { 'Low'{$cmbPriority.SelectedIndex=0} 'BelowNormal'{$cmbPriority.SelectedIndex=1} 'AboveNormal'{$cmbPriority.SelectedIndex=3} 'High'{$cmbPriority.SelectedIndex=4} default{$cmbPriority.SelectedIndex=2} }
    $adv.Panel.Children.Add($cmbPriority)

    $adv.Panel.Children.Add((New-Label "N O T I F I C A T I O N S" -Header))
    $chkSounds = New-Check "Play sounds on completion" $cfg.SoundsEnabled
    $adv.Panel.Children.Add($chkSounds)

    $adv.Panel.Children.Add((New-Label "E X T E R N A L   P R O C E S S O R S" -Header))
    $adv.Panel.Children.Add((New-Label "Extension|Command pairs (one per line, pipe-separated):"))
    $txtExtProcs = New-Object System.Windows.Controls.TextBox; $txtExtProcs.AcceptsReturn=$true; $txtExtProcs.Height=100
    $txtExtProcs.VerticalScrollBarVisibility='Auto'; $txtExtProcs.Margin="0,4"; $txtExtProcs.TextWrapping='Wrap'
    $epLines = @(); foreach ($ep in $cfg.ExternalProcessors) { if ($ep.Extension -and $ep.Command) { $epLines += "$($ep.Extension)|$($ep.Command)" } }
    $txtExtProcs.Text = $epLines -join "`r`n"
    $adv.Panel.Children.Add($txtExtProcs)

    $adv.Panel.Children.Add((New-Label "C O N F I G" -Header))
    $btnOpenCfg = New-Object System.Windows.Controls.Button; $btnOpenCfg.Content="Open Config Directory"; $btnOpenCfg.Padding="12,6"; $btnOpenCfg.HorizontalAlignment='Left'; $btnOpenCfg.Margin="0,6"
    $btnOpenCfg.Add_Click({ Start-Process explorer.exe $script:AppDataDir })
    $adv.Panel.Children.Add($btnOpenCfg)
    $btnResetCfg = New-Object System.Windows.Controls.Button; $btnResetCfg.Content="Reset All Settings"; $btnResetCfg.Padding="12,6"; $btnResetCfg.HorizontalAlignment='Left'; $btnResetCfg.Margin="0,6"
    $btnResetCfg.Add_Click({
        $script:Config = $script:DefaultConfig.Clone(); Save-Config -Config $script:Config
        [System.Windows.MessageBox]::Show("Settings reset to defaults.","ExtractorX",'OK','Information')
    })
    $adv.Panel.Children.Add($btnResetCfg)
    $tabs.Items.Add($adv.Tab)

    [System.Windows.Controls.Grid]::SetRow($tabs, 0)
    $mainGrid.Children.Add($tabs)

    # Save/Cancel buttons
    $btnPanel = New-Object System.Windows.Controls.StackPanel; $btnPanel.Orientation='Horizontal'; $btnPanel.HorizontalAlignment='Right'; $btnPanel.Margin="16,10"
    [System.Windows.Controls.Grid]::SetRow($btnPanel, 1)
    $btnSave = New-Object System.Windows.Controls.Button; $btnSave.Content="Save Settings"; $btnSave.Padding="20,9"; $btnSave.Margin="0,0,8,0"
    $btnSave.Background=[System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF00B4D8')
    $btnSave.Foreground=[System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF0A0A12')
    $btnSave.FontWeight='SemiBold'
    $btnCancel = New-Object System.Windows.Controls.Button; $btnCancel.Content="Cancel"; $btnCancel.Padding="20,9"
    $btnPanel.Children.Add($btnSave); $btnPanel.Children.Add($btnCancel)
    $mainGrid.Children.Add($btnPanel)

    $sw.Content = $mainGrid

    $btnCancel.Add_Click({ $sw.Close() })
    $btnSave.Add_Click({
        # General
        $cfg.AlwaysOnTop = [bool]$chkAlwaysTop.IsChecked; $window.Topmost = $cfg.AlwaysOnTop
        $cfg.MinimizeToTray = [bool]$chkMinToTray.IsChecked
        $cfg.LogHistory = [bool]$chkLogHist.IsChecked
        $cfg.AutoSwitchToHistory = [bool]$chkAutoSwitch.IsChecked
        $cfg.DeepArchiveDetection = [bool]$chkDeepDetect.IsChecked
        # Destination
        $cfg.OverwriteMode = @('Always','Never','Rename')[$cmbOver.SelectedIndex]
        # Process
        $cfg.NestedExtraction = [bool]$chkNested.IsChecked
        $cfg.NestedApplyPostAction = [bool]$chkNestedPost.IsChecked
        $cfg.NestedMaxDepth = @(3,5,10,20)[$cmbDepth.SelectedIndex]
        $cfg.PostAction = @('None','Recycle','MoveToFolder','Delete')[$cmbPost.SelectedIndex]
        $cfg.PostActionFolder = $txtPostFolder.Text
        $cfg.OpenDestAfterExtract = [bool]$chkOpenDest.IsChecked
        $cfg.RemoveDuplicateFolder = [bool]$chkRemoveDupe.IsChecked
        $cfg.RenameSingleFile = [bool]$chkRenameSingle.IsChecked
        $cfg.DeleteBrokenFiles = [bool]$chkDelBroken.IsChecked
        $cfg.CloseOnComplete = [bool]$chkCloseOk.IsChecked
        $cfg.CloseOnCompleteAlways = [bool]$chkCloseAlways.IsChecked
        $cfg.ClearListOnComplete = [bool]$chkClearList.IsChecked
        # Explorer
        $cfg.CtxEnabled = [bool]$chkCtxEnabled.IsChecked; $cfg.CtxGrouped = [bool]$chkCtxGrouped.IsChecked
        $cfg.CtxExtractHere = [bool]$chkCtxHere.IsChecked; $cfg.CtxExtractToFolder = [bool]$chkCtxFolder.IsChecked
        $cfg.CtxEnqueue = [bool]$chkCtxEnqueue.IsChecked; $cfg.CtxSearchArchives = [bool]$chkCtxSearch.IsChecked
        # Apply context menu
        if ($cfg.CtxEnabled) { Install-ContextMenuEntries -Cfg $cfg } else { Uninstall-ContextMenuEntries }
        # Drag & Drop
        $cfg.AutoExtractOnDrop = [bool]$chkAutoExtDrop.IsChecked
        $cfg.DragDropFilterType = @('None','Inclusion','Exclusion')[$cmbFilter.SelectedIndex]
        $cfg.DragDropFilterMask = $txtFilterMask.Text
        # Passwords
        $cfg.UsePasswordList = [bool]$chkUsePw.IsChecked
        $cfg.PromptOnExhaustion = [bool]$chkPromptExh.IsChecked
        $cfg.AssumeOnePassword = [bool]$chkOnePerArc.IsChecked
        # Files
        $cfg.FileExclusions = $txtExcl.Text
        # Monitor
        $cfg.WatchFolders = @($watchList.Items | ForEach-Object { $_ })
        $cfg.WatchAutoExtract = [bool]$chkWatchAuto.IsChecked
        # Advanced
        $cfg.ThreadPriority = @('Low','BelowNormal','Normal','AboveNormal','High')[$cmbPriority.SelectedIndex]
        $cfg.SoundsEnabled = [bool]$chkSounds.IsChecked
        # External processors
        $cfg.ExternalProcessors = @()
        foreach ($ln in ($txtExtProcs.Text -split "`r`n|`n")) {
            $parts = $ln -split '\|',2
            if ($parts.Count -eq 2 -and $parts[0].Trim() -and $parts[1].Trim()) {
                $cfg.ExternalProcessors += @{Extension=$parts[0].Trim(); Command=$parts[1].Trim()}
            }
        }

        Save-Config -Config $cfg

        # Restart watchers if config changed
        Start-Watchers
        Send-UIMessage -Type 'Log' -Data @{Message="Settings saved";Color='#FF6BCB77'}
        Send-UIMessage -Type 'Status' -Data @{Text="Settings saved";Color='#FF6BCB77'}
        $sw.Close()
    })

    [void]$sw.ShowDialog()
})


# =====================================================================
# Watch Folder System
# =====================================================================
function Stop-Watchers {
    foreach ($w in $script:State.Watchers) {
        try { $w.Watcher.EnableRaisingEvents = $false; $w.Watcher.Dispose() } catch {}
        try { if ($w.RunspacePS) { $w.RunspacePS.Stop(); $w.RunspacePS.Dispose() } } catch {}
        try { if ($w.RunspaceRS) { $w.RunspaceRS.Close(); $w.RunspaceRS.Dispose() } } catch {}
    }
    $script:State.Watchers.Clear()
    $script:State.WatchersActive = $false
    $ui['WatchStatus'].Text = ""
}

function Start-Watchers {
    Stop-Watchers
    $folders = @($script:Config.WatchFolders)
    if ($folders.Count -eq 0) { return }

    foreach ($folder in $folders) {
        if (-not (Test-Path $folder)) { continue }

        $wRS = [runspacefactory]::CreateRunspace(); $wRS.Open()
        $wRS.SessionStateProxy.SetVariable('uiQueue', $script:UIQueue)
        $wRS.SessionStateProxy.SetVariable('watchFolder', $folder)
        $wRS.SessionStateProxy.SetVariable('archiveExts', $script:ArchiveExtensions)
        $wRS.SessionStateProxy.SetVariable('recursive', $true)

        $wPS = [powershell]::Create().AddScript({
            $q = $uiQueue
            $fsw = New-Object IO.FileSystemWatcher
            $fsw.Path = $watchFolder
            $fsw.IncludeSubdirectories = $recursive
            $fsw.NotifyFilter = [IO.NotifyFilters]::FileName -bor [IO.NotifyFilters]::LastWrite
            $fsw.EnableRaisingEvents = $true

            $debounce = [System.Collections.Concurrent.ConcurrentDictionary[string,datetime]]::new()

            $action = {
                $p = $Event.SourceEventArgs.FullPath
                $ext = [IO.Path]::GetExtension($p).ToLower()
                if ($ext -notin $archiveExts) { return }
                # Skip non-first multi-volume parts
                $fn = [IO.Path]::GetFileName($p)
                if ($fn -match '\.part(\d+)\.rar$' -and [int]$Matches[1] -gt 1) { return }
                if ($fn -match '\.[rs]\d{2}$') { return }
                if ($fn -match '\.\w+\.(\d{3})$' -and $Matches[1] -ne '001') { return }
                $now = [datetime]::UtcNow
                if ($debounce.ContainsKey($p)) {
                    $last = $debounce[$p]
                    if (($now - $last).TotalSeconds -lt 3) { return }
                }
                $debounce[$p] = $now
                # Wait for file to be ready (not locked by download)
                $ready = $false
                for ($i = 0; $i -lt 20; $i++) {
                    Start-Sleep -Milliseconds 500
                    try {
                        $s = [IO.File]::Open($p, 'Open', 'Read', 'None')
                        $s.Dispose(); $ready = $true; break
                    } catch {}
                }
                if ($ready) { $q.Enqueue(@{ Type = 'WatchFileDetected'; FilePath = $p }) }
            }

            Register-ObjectEvent $fsw Created -Action $action | Out-Null
            Register-ObjectEvent $fsw Renamed -Action $action | Out-Null

            # Keep runspace alive
            while ($true) { Start-Sleep -Seconds 5 }
        })
        $wPS.Runspace = $wRS
        $handle = $wPS.BeginInvoke()

        $script:State.Watchers.Add(@{
            Folder = $folder
            RunspacePS = $wPS; RunspaceRS = $wRS; Handle = $handle
            Watcher = $null
        }) | Out-Null
    }

    $script:State.WatchersActive = $true
    $ui['WatchStatus'].Text = "Watching $($folders.Count) folder(s)"
    Write-UILog "Watchers started for $($folders.Count) folder(s)" '#FF6BCB77'
}

# Auto-start watchers if config has watch folders
if ($script:Config.WatchFolders.Count -gt 0) { Start-Watchers }

# =====================================================================
# Window Minimize-to-Tray Behavior
# =====================================================================
$window.Add_StateChanged({
    if ($window.WindowState -eq 'Minimized' -and $script:Config.MinimizeToTray) {
        $window.Hide()
        $script:TrayIcon.Visible = $true
    }
})

# =====================================================================
# Window Closing - Save Config & Cleanup
# =====================================================================
$window.Add_Closing({
    param($sender, $e)
    # If tray-minimize is on and not force-close, minimize instead
    if ($script:Config.MinimizeToTray -and -not $script:ForceClose -and -not $script:State.IsExtracting) {
        # Actually close on X click — tray behavior is only on minimize
    }

    # Save window position and all config
    $script:Config.WindowWidth  = [int]$window.ActualWidth
    $script:Config.WindowHeight = [int]$window.ActualHeight
    $script:Config.WindowLeft   = [int]$window.Left
    $script:Config.WindowTop    = [int]$window.Top
    $script:Config.OutputPath   = $ui['TxtOutputPath'].Text
    $script:Config.DeleteAfterExtract = [bool]$ui['ChkDeleteAfter'].IsChecked
    Save-Config -Config $script:Config

    # Stop watchers
    Stop-Watchers

    # Stop scanning if running
    if ($script:State.IsScanning) { $script:State.ScanCancelRequested = $true }
    try { if ($script:scanPS) { $script:scanPS.Stop(); $script:scanPS.Dispose() } } catch {}
    try { if ($script:scanRS) { $script:scanRS.Close(); $script:scanRS.Dispose() } } catch {}

    # Stop extraction if running
    if ($script:State.IsExtracting) { $script:State.StopRequested = $true }

    # Cleanup runspaces
    try { if ($script:extractPS) { $script:extractPS.Stop(); $script:extractPS.Dispose() } } catch {}
    try { if ($script:extractRS) { $script:extractRS.Close(); $script:extractRS.Dispose() } } catch {}
    try { if ($script:dl7zPS) { $script:dl7zPS.Stop(); $script:dl7zPS.Dispose() } } catch {}
    try { if ($script:dl7zRS) { $script:dl7zRS.Close(); $script:dl7zRS.Dispose() } } catch {}

    # Dispose tray icon
    $script:TrayIcon.Visible = $false
    $script:TrayIcon.Dispose()

    # Stop timer
    $script:UITimer.Stop()
})

# =====================================================================
# Command-Line File Loading
# =====================================================================
if ($FilesToExtract -and $FilesToExtract.Count -gt 0) {
    $cliFiles = @()
    foreach ($f in $FilesToExtract) {
        $resolved = if ([IO.Path]::IsPathRooted($f)) { $f } else { Join-Path $PWD $f }
        if (Test-Path $resolved) { $cliFiles += $resolved }
    }
    if ($cliFiles.Count -gt 0) {
        Start-ScanJob -Paths $cliFiles
        Write-UILog "Scanning $($cliFiles.Count) path(s) from command line" '#FF00B4D8'
    }
}

# =====================================================================
# Startup Minimization
# =====================================================================
if ($minimize) {
    $window.WindowState = 'Minimized'
} elseif ($minimizetotray) {
    $window.Hide()
    $script:TrayIcon.Visible = $true
}

# =====================================================================
# Initial Status
# =====================================================================
Write-UILog "ExtractorX v$($script:AppVersion) ready" '#FF00B4D8'
if ($script:7zPath) {
    Write-UILog "7-Zip: $($script:7zPath)" '#FF6BCB77'
} else {
    Write-UILog "7-Zip not found - will download on first extract" '#FFFFD93D'
}
$ui['StatusText'].Text = "Ready"
$ui['StatusText'].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#FF6BCB77')
Send-UIMessage -Type 'QueueCountRefresh'

# =====================================================================
# Run!
# =====================================================================
[void]$window.ShowDialog()
