# =============================================================================
#  GP Launcher - GhostPractice Database/Server Selector
# =============================================================================
#  A WPF UI built in PowerShell that lets the user pick a (server, database)
#  combination and launch GhostPractice against it. Each "connection" is a
#  named pair of SQL Server\Instance + database. When the user clicks one, we:
#    1. Locate GP's per-version user.config folder under %LOCALAPPDATA%\Korbitec
#    2. Write a user.config XML pointing at the chosen server/database
#    3. Launch GhostPractice.exe (which then reads that user.config)
#    4. Bump the usage counter so the most-used connection floats to the top
#
#  Files:
#    config.json      - persisted settings (gpPath + connections + usage + lastLaunched)
#    GPLauncher.ps1   - this file
#    gplauncher.log   - rolling diagnostic log, rewritten on each launch
#    icon.ico         - app icon
#    gp-logo.png      - optional: drop a 300x60 transparent PNG to override the monogram
#    gp-logo.svg      - the original GhostPractice SVG (kept for reference)
# =============================================================================

# WPF and Windows Forms assemblies needed for the UI and any MessageBox calls.
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# -----------------------------------------------------------------------------
# Path setup - everything lives next to the script so the folder is portable.
# -----------------------------------------------------------------------------
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ConfigPath = Join-Path $ScriptDir "config.json"
$IconPath   = Join-Path $ScriptDir "icon.ico"
$LogoPath   = Join-Path $ScriptDir "gp-logo.png"
$LogPath    = Join-Path $ScriptDir "gplauncher.log"

$AppVersion = "Beta 0.0.8"

# -----------------------------------------------------------------------------
# Brand colors (from the official GhostPractice logo).
# Used throughout the UI for accents so the launcher feels visually consistent
# with the GP product itself.
# -----------------------------------------------------------------------------
$BrandPurple = '#6d1652'
$BrandPink   = '#cc3369'

# -----------------------------------------------------------------------------
# Easter egg: "Created by AZ" snark.
# Each click on the credits link increments $script:AZClickCount and (if the
# count maps to a scripted line) shows that line as a small fade-in pill above
# the credits. After the scripted list runs out we cycle randomly through the
# wildcard pool. The counter resets to 0 after 60 seconds of no clicks so the
# joke restarts fresh if someone walks away and comes back.
# -----------------------------------------------------------------------------
$AZScriptedSnark = @(
    $null                                                          # 1: just the spin
    "Whoa, easy there"                                            # 2
    "That's enough now"                                           # 3
    "Ouch, that hurts"                                            # 4
    "Seriously, click somewhere else"                             # 5
    "If you do that one more time, I'm gonna snitch to Aubrey"    # 6
    "Right, that's it. Calling Aubrey now..."                     # 7
    "Aubrey says: please stop"                                    # 8
    "I'm getting dizzy"                                           # 9
    "Why are you like this"                                       # 10
    "Fine. Have it your way."                                     # 11
)
$AZWildcardSnark = @(
    "Still here?"
    "What did AZ ever do to you"
    "Touch grass"
    "AZ is filing an HR complaint"
    "This is being logged"
    "Bored at work?"
    "Have you tried the Settings button?"
    "Get back to billable hours"
    "AZ is taking a coffee break"
    "Productivity meter: -42%"
)
$script:AZClickCount = 0
$script:AZResetTimer = $null

# -----------------------------------------------------------------------------
# Tiny diagnostic logger. Click handlers use this so we can see what happened
# after the fact (useful when WPF event handlers fail silently).
# -----------------------------------------------------------------------------
function Write-Log($msg) {
    try {
        $line = "{0:yyyy-MM-dd HH:mm:ss}  {1}" -f (Get-Date), $msg
        Add-Content -Path $LogPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch { }
}

# Reset log each launch so it stays small.
try { Set-Content -Path $LogPath -Value "--- GP Launcher $AppVersion start ---" -Encoding UTF8 } catch { }

# =============================================================================
#  CONFIG LAYER - load, save, migrate, and shape the connections data
# =============================================================================

# Persist the config to disk as pretty-printed JSON.
function Save-Config($cfg) {
    $cfg | ConvertTo-Json -Depth 5 | Set-Content $ConfigPath -Encoding UTF8
}

# Factory for a connection record. Always use this rather than building hashtables
# inline so every connection has the exact same property names/order.
function New-Connection($name, $server, $database, $regenerateReport = $true) {
    return [PSCustomObject]@{
        name             = $name              # Display label shown on the launcher button
        server           = $server            # SQL Server\Instance (e.g. "10.10.10.21\SQLEXPRESS2022")
        database         = $database          # Database name on that server
        regenerateReport = [bool]$regenerateReport  # Per-connection toggle for RegenerateReport in user.config
    }
}

# Read config.json from disk. If it doesn't exist, write a sensible default.
# This function also migrates old single-server configs (`{ server, databases[] }`)
# to the new multi-connection schema, defensively cleans null/malformed entries,
# and ensures the `lastLaunched` field exists.
function Load-Config {
    # First-run case: no file on disk yet. Seed a default config and persist it.
    if (-not (Test-Path $ConfigPath)) {
        $default = [PSCustomObject]@{
            gpPath        = "C:\Program Files (x86)\Korbicom\GhostPractice\GhostPractice.exe"
            connections   = @()
            usage         = [PSCustomObject]@{}
            lastLaunched  = ""
        }
        Save-Config $default
        return $default
    }

    $cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    $needsSave = $false  # Tracks whether any of the cleanups below mutated the config.

    # --- Schema migration: { server, databases[] }  ->  { connections[] } ---
    if (-not $cfg.PSObject.Properties['connections']) {
        $oldServer = if ($cfg.PSObject.Properties['server']) { $cfg.server } else { "10.10.10.21" }
        $oldDbs    = if ($cfg.PSObject.Properties['databases']) { @($cfg.databases) } else { @() }

        $migrated = New-Object System.Collections.ArrayList
        foreach ($db in $oldDbs) {
            if ($db) { [void]$migrated.Add((New-Connection $db $oldServer $db)) }
        }
        $cfg | Add-Member -NotePropertyName 'connections' -NotePropertyValue ([object[]]$migrated.ToArray()) -Force
        $needsSave = $true
    }

    # Strip legacy fields once they've been migrated.
    if ($cfg.PSObject.Properties['server'])    { $cfg.PSObject.Properties.Remove('server');    $needsSave = $true }
    if ($cfg.PSObject.Properties['databases']) { $cfg.PSObject.Properties.Remove('databases'); $needsSave = $true }

    # --- Self-healing: drop any null/malformed connection entries ---
    # Also back-fills the per-connection regenerateReport flag (default true,
    # matching the pre-0.0.8 behavior of always-regenerate) for any record
    # that doesn't have it yet.
    $clean = New-Object System.Collections.ArrayList
    foreach ($c in @($cfg.connections)) {
        if ($c -and $c.PSObject.Properties['name'] -and $c.name) {
            $regen = $true
            if ($c.PSObject.Properties['regenerateReport']) {
                $regen = [bool]$c.regenerateReport
            } else {
                $needsSave = $true   # We're about to add the missing field.
            }
            [void]$clean.Add((New-Connection $c.name $c.server $c.database $regen))
        }
    }
    if ($clean.Count -ne (@($cfg.connections)).Count) { $needsSave = $true }
    $cfg.connections = [object[]]$clean.ToArray()

    # --- Make sure usage{} exists and has an entry per connection ---
    if (-not $cfg.PSObject.Properties['usage']) {
        $cfg | Add-Member -NotePropertyName 'usage' -NotePropertyValue ([PSCustomObject]@{})
        $needsSave = $true
    }
    foreach ($c in $cfg.connections) {
        if ($c.name -and -not $cfg.usage.PSObject.Properties[$c.name]) {
            $cfg.usage | Add-Member -NotePropertyName $c.name -NotePropertyValue 0
            $needsSave = $true
        }
    }

    # --- Make sure lastLaunched{} exists (used to mark the most recent button) ---
    if (-not $cfg.PSObject.Properties['lastLaunched']) {
        $cfg | Add-Member -NotePropertyName 'lastLaunched' -NotePropertyValue ""
        $needsSave = $true
    }

    if ($needsSave) { Save-Config $cfg }
    return $cfg
}

$Config = Load-Config

# =============================================================================
#  GP INTEGRATION - locate GhostPractice's per-version config folder
# =============================================================================

# GhostPractice stores its user.config at a path like:
#   %LOCALAPPDATA%\Korbitec\GhostPractice.exe_StrongName_<hash>\<version>\user.config
function Get-GPConfigTargetPath {
    $korbitecBase = Join-Path $env:LOCALAPPDATA "Korbitec"
    if (-not (Test-Path $korbitecBase)) { return $null }

    $strongNameDirs = Get-ChildItem $korbitecBase -Directory -Filter "GhostPractice.exe_StrongName_*" -ErrorAction SilentlyContinue
    if (-not $strongNameDirs -or $strongNameDirs.Count -eq 0) { return $null }
    $strongNameDir = $strongNameDirs | Select-Object -First 1

    $versionDirs = Get-ChildItem $strongNameDir.FullName -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+$' } |
        Sort-Object { [version]$_.Name } -Descending

    if (-not $versionDirs -or $versionDirs.Count -eq 0) { return $null }
    return $versionDirs[0].FullName
}

# -----------------------------------------------------------------------------
# Update an existing user.config in place, changing only the four settings we
# care about and leaving everything else untouched.
#
# Why this exists:
#   GhostPractice's user.config holds ~45 settings - column widths, form
#   geometry, the ApplicationDeviceID, HasUpgraded migration flags, etc.
#   Earlier versions of this launcher overwrote the whole file from a hard-
#   coded template every launch, silently wiping all of those preferences.
#   This function fixes that by editing only what we own.
#
# Settings we touch:
#   Server             - chosen connection's server\instance
#   Database           - chosen connection's database
#   ApplicationVersion - the auto-detected installed GP version
#   RegenerateReport   - per-connection: True or False from $regenerateReport
#
# Anything else in user.config is left exactly as GP wrote it.
# -----------------------------------------------------------------------------
function Update-UserConfig($configFile, $server, $database, $version, $regenerateReport) {
    [xml]$doc = Get-Content $configFile -Raw

    $regenStr = if ($regenerateReport) { 'True' } else { 'False' }
    $updates = [ordered]@{
        'Server'             = $server
        'Database'           = $database
        'ApplicationVersion' = $version
        'RegenerateReport'   = $regenStr
    }

    # We need this if we have to insert a missing <setting> node.
    $userSettings = $doc.SelectSingleNode("//*[local-name()='Korbicom.Mustang.ApplicationUserSettings']")

    foreach ($name in $updates.Keys) {
        $newValue = $updates[$name]
        $node = $doc.SelectSingleNode("//*[local-name()='setting' and @name='$name']")

        if ($node) {
            # Existing setting - update its <value> child only.
            $valueNode = $node.SelectSingleNode("*[local-name()='value']")
            if ($valueNode) {
                $valueNode.InnerText = [string]$newValue
            } else {
                # Setting tag exists but no <value> child somehow; add one.
                $valueElem = $doc.CreateElement('value')
                $valueElem.InnerText = [string]$newValue
                [void]$node.AppendChild($valueElem)
            }
        } elseif ($userSettings) {
            # Setting doesn't exist yet - append a fresh <setting> node.
            $settingElem = $doc.CreateElement('setting')
            $settingElem.SetAttribute('name', $name)
            $settingElem.SetAttribute('serializeAs', 'String')
            $valueElem = $doc.CreateElement('value')
            $valueElem.InnerText = [string]$newValue
            [void]$settingElem.AppendChild($valueElem)
            [void]$userSettings.AppendChild($settingElem)
        }
    }

    $doc.Save($configFile)
}

# Tracks whether we've already taken the per-session backup.
$script:UserConfigBackupTaken = $false

# -----------------------------------------------------------------------------
# Take a one-time-per-session backup of user.config before the first edit.
# Subsequent launches in the same session reuse the same .bak (we only want a
# pre-launcher snapshot, not a snapshot after every click).
# -----------------------------------------------------------------------------
function Backup-UserConfigOnce($configFile) {
    if ($script:UserConfigBackupTaken) { return }
    if (-not (Test-Path $configFile)) { return }
    $backupFile = "$configFile.bak"
    try {
        Copy-Item $configFile $backupFile -Force -ErrorAction Stop
        $script:UserConfigBackupTaken = $true
        Write-Log ("Backed up user.config to {0}" -f $backupFile)
    } catch {
        Write-Log ("user.config backup failed: " + $_)
    }
}

# -----------------------------------------------------------------------------
# Build the user.config XML that GhostPractice reads on startup.
# -----------------------------------------------------------------------------
function Build-UserConfig($server, $database, $version, $regenerateReport = $true) {
    $regenStr = if ($regenerateReport) { 'True' } else { 'False' }
    return @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
    <configSections>
        <sectionGroup name="userSettings" type="System.Configuration.UserSettingsGroup, System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" >
            <section name="Korbicom.Mustang.ApplicationUserSettings" type="System.Configuration.ClientSettingsSection, System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" allowExeDefinition="MachineToLocalUser" requirePermission="false" />
        </sectionGroup>
    </configSections>
    <userSettings>
        <Korbicom.Mustang.ApplicationUserSettings>
            <setting name="FormLocation" serializeAs="String"><value>0, 0</value></setting>
            <setting name="HasUpgraded" serializeAs="String"><value>True</value></setting>
            <setting name="Server" serializeAs="String"><value>$server</value></setting>
            <setting name="Database" serializeAs="String"><value>$database</value></setting>
            <setting name="ApplicationVersion" serializeAs="String"><value>$version</value></setting>
            <setting name="MessagingProgram" serializeAs="String"><value>1</value></setting>
            <setting name="ShowDBConnectionDetails" serializeAs="String"><value>False</value></setting>
            <setting name="FormSize" serializeAs="String"><value>88, 28</value></setting>
            <setting name="ShowConversionTools" serializeAs="String"><value>False</value></setting>
            <setting name="ApplicationDeviceID" serializeAs="String"><value>00000000-0000-0000-0000-000000000000</value></setting>
            <setting name="InvoiceWizardOpenMaximized" serializeAs="String"><value>False</value></setting>
            <setting name="AllowAdvancedSupport" serializeAs="String"><value>False</value></setting>
            <setting name="EnableGenerateStatementDocuments" serializeAs="String"><value>False</value></setting>
            <setting name="AllowAllocationOfOldMatters" serializeAs="String"><value>False</value></setting>
            <setting name="ShowHMRCTestOptions" serializeAs="String"><value>False</value></setting>
            <setting name="ShowMonthEndGB" serializeAs="String"><value>False</value></setting>
            <setting name="ShowDocumentDataMigrationVersion10" serializeAs="String"><value>False</value></setting>
            <setting name="ShowDocumentMetaDataMigration" serializeAs="String"><value>False</value></setting>
            <setting name="ShowTransactionID" serializeAs="String"><value>False</value></setting>
            <setting name="ResyncDocumentNames" serializeAs="String"><value>False</value></setting>
            <setting name="GenerateDataset" serializeAs="String"><value>False</value></setting>
            <setting name="LogProgress" serializeAs="String"><value>False</value></setting>
            <setting name="EnableEditingE4XML" serializeAs="String"><value>False</value></setting>
            <setting name="RegenerateReport" serializeAs="String"><value>$regenStr</value></setting>
            <setting name="AllowAgingRecalculate" serializeAs="String"><value>False</value></setting>
            <setting name="ShowWarningForUncomittedTransactions" serializeAs="String"><value>False</value></setting>
            <setting name="ImageTextAlignment" serializeAs="String"><value>1</value></setting>
            <setting name="FormMaximised" serializeAs="String"><value>False</value></setting>
            <setting name="ReportViewColumnWidthCreateDate" serializeAs="String"><value>200</value></setting>
            <setting name="ReportAdHocColumnWidthDescription" serializeAs="String"><value>300</value></setting>
            <setting name="TransactionHistoryRunningBalanceView" serializeAs="String"><value>0</value></setting>
            <setting name="TransactionHistoryAccountType" serializeAs="String"><value>2</value></setting>
            <setting name="ReportViewColumnWidthCreateBy" serializeAs="String"><value>200</value></setting>
            <setting name="RunReportsAdHocCategory" serializeAs="String"><value>2</value></setting>
            <setting name="ReportAdHocColumnWidthName" serializeAs="String"><value>300</value></setting>
            <setting name="LegalDiaryActivitiesHeight" serializeAs="String"><value>254</value></setting>
            <setting name="ReportAdHocReference" serializeAs="String"><value>80</value></setting>
            <setting name="LegalDiaryAdHocHeight" serializeAs="String"><value>187</value></setting>
            <setting name="ReportAdHocType" serializeAs="String"><value>80</value></setting>
            <setting name="TransactionHistoryPortrait" serializeAs="String"><value>Landscape</value></setting>
            <setting name="ReportViewColumnWidthHelpAdHoc" serializeAs="String"><value>40</value></setting>
            <setting name="ReportStoredColumnWidthDescription" serializeAs="String"><value>300</value></setting>
            <setting name="ReportStoredColumnWidthArchive" serializeAs="String"><value>100</value></setting>
            <setting name="ReportStoredColumnWidthReference" serializeAs="String"><value>300</value></setting>
            <setting name="ReportStoredColumnWidthDate" serializeAs="String"><value>120</value></setting>
        </Korbicom.Mustang.ApplicationUserSettings>
    </userSettings>
</configuration>
"@
}

# =============================================================================
#  UI HELPERS
# =============================================================================

# -----------------------------------------------------------------------------
# Deterministically pick a stripe color for a server string. The same server
# string always maps to the same color, so connections on the same server are
# visually grouped.
# -----------------------------------------------------------------------------
function Get-ServerColor($server) {
    if (-not $server) { return '#888AA8' }
    $palette = @('#6d1652','#cc3369','#1f6feb','#3fb950','#d29922','#bc8cff','#ff7b72','#79c0ff','#56d4dd','#f778ba')
    $hash = 0
    foreach ($c in $server.ToCharArray()) {
        $hash = (($hash * 31) + [int]$c) -band 0x7FFFFFFF
    }
    return $palette[$hash % $palette.Count]
}

# Convenience: produce a SolidColorBrush from a #RRGGBB string.
function New-Brush($hex) {
    return New-Object System.Windows.Media.SolidColorBrush(
        [System.Windows.Media.ColorConverter]::ConvertFromString($hex))
}

# -----------------------------------------------------------------------------
# Database cylinder icon (programmatic WPF drawing). Drawn programmatically so
# we don't need an external image file.
# -----------------------------------------------------------------------------
function New-DbIconDrawing {
    $dg = New-Object System.Windows.Media.DrawingGroup
    $bodyBrush = [System.Windows.Media.Brushes]::Transparent
    $strokePen = New-Object System.Windows.Media.Pen([System.Windows.Media.Brushes]::White, 1.2)
    $strokePen.LineJoin = [System.Windows.Media.PenLineJoin]::Round

    $bodyGeo = [System.Windows.Media.Geometry]::Parse("M 4,6 L 4,18 A 8,3 0 0 0 20,18 L 20,6")
    [void]$dg.Children.Add((New-Object System.Windows.Media.GeometryDrawing($bodyBrush, $strokePen, $bodyGeo)))

    $topGeo = New-Object System.Windows.Media.EllipseGeometry((New-Object System.Windows.Point(12, 6)), 8, 3)
    [void]$dg.Children.Add((New-Object System.Windows.Media.GeometryDrawing($bodyBrush, $strokePen, $topGeo)))

    $midGeo = [System.Windows.Media.Geometry]::Parse("M 4,12 A 8,3 0 0 0 20,12")
    [void]$dg.Children.Add((New-Object System.Windows.Media.GeometryDrawing($bodyBrush, $strokePen, $midGeo)))

    return $dg
}
$dbIconDrawing = New-DbIconDrawing

# -----------------------------------------------------------------------------
# Try to load gp-logo.png from disk. Returns a BitmapImage on success, $null
# otherwise. The main window then falls back to a brand-colored "GP" monogram.
# -----------------------------------------------------------------------------
function Get-LogoBitmap {
    if (-not (Test-Path $LogoPath)) { return $null }
    try {
        $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
        $bmp.BeginInit()
        $bmp.UriSource = New-Object System.Uri($LogoPath, [System.UriKind]::Absolute)
        $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bmp.EndInit()
        $bmp.Freeze()
        return $bmp
    } catch {
        Write-Log ("Logo load failed: " + $_)
        return $null
    }
}

# -----------------------------------------------------------------------------
# Hyperlink helpers for the credits / about lines.
#
# Build-CreatedByHyperlink: returns a Hyperlink containing "Created by AZ".
# Clicking it triggers the easter egg - the "AZ" letters spin three full
# rotations using a CubicEase RotateTransform animation.
#
# Build-EmailHyperlink: returns a Hyperlink wired to mailto: that opens the
# user's default mail client via Start-Process.
#
# Set-*-Append: helpers that append a hyperlink to an existing TextBlock's
# Inlines collection (used by the multi-part main window credits line).
#
# Set-CreatedByInline / Set-EmailInline: clear-then-set helpers used by the
# Settings About section where each line is its own dedicated TextBlock.
# -----------------------------------------------------------------------------
function Build-CreatedByHyperlink {
    $link = New-Object System.Windows.Documents.Hyperlink
    $link.TextDecorations = $null
    $link.Cursor = [System.Windows.Input.Cursors]::Hand
    $link.ToolTip = "Click for a small surprise"

    [void]$link.Inlines.Add((New-Object System.Windows.Documents.Run("Created by ")))

    # "AZ" lives inside an InlineUIContainer so we can attach a RotateTransform.
    $azBlock = New-Object System.Windows.Controls.TextBlock
    $azBlock.Text = "AZ"
    $azBlock.FontWeight = [System.Windows.FontWeights]::SemiBold
    $azBlock.RenderTransformOrigin = New-Object System.Windows.Point(0.5, 0.5)
    $azRotate = New-Object System.Windows.Media.RotateTransform(0)
    $azBlock.RenderTransform = $azRotate

    $azContainer = New-Object System.Windows.Documents.InlineUIContainer
    $azContainer.BaselineAlignment = [System.Windows.Documents.BaselineAlignment]::Center
    $azContainer.Child = $azBlock
    [void]$link.Inlines.Add($azContainer)

    # Click handler: spin AZ + maybe pop a snarky message + arm the reset timer.
    $link.Add_Click({
        # Spin: 1.8s, 4 full rotations, QuinticEase EaseOut so it bursts off and
        # decelerates aggressively at the end - wheel-of-fortune feel.
        $anim = New-Object System.Windows.Media.Animation.DoubleAnimation
        $anim.From = 0
        $anim.To = 3600  # 10 full rotations
        $anim.Duration = New-Object System.Windows.Duration([TimeSpan]::FromMilliseconds(4500))
        $ease = New-Object System.Windows.Media.Animation.QuinticEase
        $ease.EasingMode = [System.Windows.Media.Animation.EasingMode]::EaseOut
        $anim.EasingFunction = $ease
        # BeginAnimation auto-replaces any in-flight animation so rapid clicks
        # restart the spin cleanly (no stutter).
        $azRotate.BeginAnimation([System.Windows.Media.RotateTransform]::AngleProperty, $anim)

        # Bump the click counter and look up the message (if any) for this index.
        $script:AZClickCount++
        $msg = $null
        if ($script:AZClickCount -le $AZScriptedSnark.Count) {
            $msg = $AZScriptedSnark[$script:AZClickCount - 1]
        } else {
            $msg = $AZWildcardSnark | Get-Random
        }
        if ($msg) { Show-Snark $msg }

        # Arm/reset the inactivity timer so the joke restarts after 60s of quiet.
        if ($script:AZResetTimer) { $script:AZResetTimer.Stop() }
        $script:AZResetTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:AZResetTimer.Interval = [TimeSpan]::FromSeconds(60)
        $script:AZResetTimer.Add_Tick({
            $script:AZClickCount = 0
            $script:AZResetTimer.Stop()
        })
        $script:AZResetTimer.Start()
    }.GetNewClosure())

    return $link
}

function Build-EmailHyperlink($email) {
    $link = New-Object System.Windows.Documents.Hyperlink
    $link.TextDecorations = $null
    $link.Cursor = [System.Windows.Input.Cursors]::Hand
    $link.NavigateUri = New-Object System.Uri("mailto:$email")
    $link.ToolTip = "Send email to $email"
    [void]$link.Inlines.Add((New-Object System.Windows.Documents.Run($email)))
    $link.Add_RequestNavigate({
        param($s, $e)
        try { Start-Process $e.Uri.AbsoluteUri } catch { Write-Log ("mailto failed: " + $_) }
        $e.Handled = $true
    })
    return $link
}

function Set-CreatedByInline-Append($tb) {
    [void]$tb.Inlines.Add((Build-CreatedByHyperlink))
}
function Set-EmailInline-Append($tb, $email) {
    [void]$tb.Inlines.Add((Build-EmailHyperlink $email))
}
function Set-CreatedByInline($tb) {
    $tb.Inlines.Clear()
    [void]$tb.Inlines.Add((Build-CreatedByHyperlink))
}
function Set-EmailInline($tb, $prefix, $email) {
    $tb.Inlines.Clear()
    if ($prefix) { [void]$tb.Inlines.Add((New-Object System.Windows.Documents.Run($prefix))) }
    [void]$tb.Inlines.Add((Build-EmailHyperlink $email))
}

# =============================================================================
#  MAIN WINDOW XAML
# =============================================================================
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="GhostPractice Launcher - $AppVersion"
    Width="460" Height="640"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    WindowStyle="None"
    AllowsTransparency="True"
    Background="Transparent"
    Foreground="White"
    FontFamily="Segoe UI">

    <Window.Resources>
        <!-- Icon button: square, transparent, hover-highlights blue. -->
        <Style x:Key="IconButton" TargetType="Button">
            <Setter Property="Width" Value="28"/>
            <Setter Property="Height" Value="24"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#16213e"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <!-- Close variant: circular red hover with centered content. -->
        <Style x:Key="IconButtonClose" TargetType="Button" BasedOn="{StaticResource IconButton}">
            <Setter Property="Width" Value="26"/>
            <Setter Property="Height" Value="26"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="13" Width="26" Height="26">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="$BrandPink"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Connection launcher button. Hover slides up 1px and brightens. -->
        <Style x:Key="DbButton" TargetType="Button">
            <Setter Property="Background" Value="#16213e"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Height" Value="64"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#2a2a4a"/>
            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="RenderTransformOrigin" Value="0.5,0.5"/>
            <Setter Property="RenderTransform">
                <Setter.Value>
                    <TranslateTransform Y="0"/>
                </Setter.Value>
            </Setter>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="8"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Stretch" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#1f2950"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="$BrandPink"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="$BrandPurple"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <EventTrigger RoutedEvent="MouseEnter">
                    <BeginStoryboard>
                        <Storyboard>
                            <DoubleAnimation Storyboard.TargetProperty="(UIElement.RenderTransform).(TranslateTransform.Y)"
                                             To="-1.5" Duration="0:0:0.10"/>
                        </Storyboard>
                    </BeginStoryboard>
                </EventTrigger>
                <EventTrigger RoutedEvent="MouseLeave">
                    <BeginStoryboard>
                        <Storyboard>
                            <DoubleAnimation Storyboard.TargetProperty="(UIElement.RenderTransform).(TranslateTransform.Y)"
                                             To="0" Duration="0:0:0.10"/>
                        </Storyboard>
                    </BeginStoryboard>
                </EventTrigger>
            </Style.Triggers>
        </Style>

        <!-- Settings text+gear button at the bottom. Green gear icon. -->
        <Style x:Key="GearButton" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#3fb950"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#16213e"/>
                                <Setter Property="Foreground" Value="#4ee868"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <!-- Outer margin gives the drop shadow room to render. -->
    <Grid Margin="16">
        <Border Background="#1a1a2e" CornerRadius="12" Padding="0">
            <Border.Effect>
                <DropShadowEffect Color="Black" BlurRadius="22" Opacity="0.55" ShadowDepth="0"/>
            </Border.Effect>

            <Grid Margin="24,18,24,14">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>  <!-- header -->
                    <RowDefinition Height="Auto"/>  <!-- subtitle -->
                    <RowDefinition Height="Auto"/>  <!-- search box (collapsed when not needed) -->
                    <RowDefinition Height="*"/>     <!-- connection list -->
                    <RowDefinition Height="Auto"/>  <!-- status pill -->
                    <RowDefinition Height="Auto"/>  <!-- footer (count + settings) -->
                    <RowDefinition Height="Auto"/>  <!-- credits -->
                </Grid.RowDefinitions>

                <!-- Header: logo centered across the whole row, close X overlaid top-right. -->
                <Grid Grid.Row="0" Margin="0,0,0,4" Height="72">
                    <ContentControl x:Name="LogoSlot"
                                    HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    <Button x:Name="CloseBtn" Style="{StaticResource IconButtonClose}"
                            HorizontalAlignment="Right" VerticalAlignment="Top"
                            ToolTip="Close (Esc)" Panel.ZIndex="10">
                        <Path Stroke="White" StrokeThickness="1.5"
                              StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                              Width="12" Height="12" Stretch="Uniform"
                              HorizontalAlignment="Center" VerticalAlignment="Center"
                              Data="M 0,0 L 10,10 M 10,0 L 0,10"/>
                    </Button>
                </Grid>

                <TextBlock Grid.Row="1"
                           Text="Select a connection to launch"
                           FontSize="12"
                           Foreground="#888AA8"
                           HorizontalAlignment="Center"
                           Margin="0,2,0,16"/>

                <!-- Search box: only shown when there are many connections. -->
                <TextBox Grid.Row="2" x:Name="SearchBox" Visibility="Collapsed"
                         Background="#16213e" Foreground="White" BorderBrush="#2a2a4a"
                         FontSize="12" Padding="10,6" Margin="0,0,0,10"
                         Tag="Search connections..."/>

                <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                    <StackPanel x:Name="DbPanel" VerticalAlignment="Center"/>
                </ScrollViewer>

                <!-- Status pill: hidden by default. Shown briefly after launch/error. -->
                <Border Grid.Row="4" x:Name="StatusPill" Visibility="Collapsed"
                        CornerRadius="12" Padding="12,5"
                        HorizontalAlignment="Center" Margin="0,10,0,4"
                        Background="#16302a">
                    <StackPanel Orientation="Horizontal">
                        <TextBlock x:Name="StatusIcon" Text="OK" FontSize="11" FontWeight="Bold"
                                   Foreground="#3fb950" Margin="0,0,8,0" VerticalAlignment="Center"/>
                        <TextBlock x:Name="StatusText" Text="" FontSize="11"
                                   Foreground="White" VerticalAlignment="Center"/>
                    </StackPanel>
                </Border>

                <!-- Footer: connection count on the left, gear+Settings on the right. -->
                <Grid Grid.Row="5" Margin="0,8,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Grid.Column="0" x:Name="CountLabel"
                               FontSize="11" Foreground="#444460" VerticalAlignment="Center"/>
                    <Button Grid.Column="1" x:Name="SettingsBtn" Style="{StaticResource GearButton}"
                            ToolTip="Settings (F2)">
                        <StackPanel Orientation="Horizontal">
                            <!-- Gear icon: outer ring with 8 teeth + inner circle. -->
                            <Path Stroke="{Binding Foreground, RelativeSource={RelativeSource AncestorType=Button}}"
                                  StrokeThickness="1.6" StrokeLineJoin="Round" Fill="Transparent"
                                  Width="16" Height="16" Stretch="Uniform"
                                  VerticalAlignment="Center" Margin="0,0,6,0"
                                  Data="M 12,8.5 A 3.5,3.5 0 1 0 12,15.5 A 3.5,3.5 0 1 0 12,8.5 Z
                                        M 12,3 L 12,5.5 M 12,18.5 L 12,21
                                        M 3,12 L 5.5,12 M 18.5,12 L 21,12
                                        M 5.6,5.6 L 7.4,7.4 M 16.6,16.6 L 18.4,18.4
                                        M 18.4,5.6 L 16.6,7.4 M 7.4,16.6 L 5.6,18.4"/>
                            <TextBlock Text="Settings" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Button>
                </Grid>

                <!-- Snark popup for the AZ easter egg. Floats just above the credits. -->
                <Border Grid.Row="6" x:Name="SnarkPill" Visibility="Collapsed" Opacity="0"
                        CornerRadius="10" Padding="12,6"
                        HorizontalAlignment="Center" Margin="0,0,0,4"
                        MaxWidth="380"
                        Background="#1f2235" BorderBrush="#2a2a4a" BorderThickness="1">
                    <TextBlock x:Name="SnarkText" FontSize="11" Foreground="#cfcfdc"
                               TextWrapping="Wrap" TextAlignment="Center"/>
                </Border>

                <TextBlock Grid.Row="6" x:Name="CreditsLabel"
                           FontSize="9" Foreground="#333350"
                           HorizontalAlignment="Center" TextAlignment="Center"
                           Margin="0,8,0,0" VerticalAlignment="Bottom"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# =============================================================================
#  ADD / EDIT CONNECTION DIALOG
# =============================================================================
$script:DialogConnResult = $null
function Show-ConnectionDialog($parentWindow, $existing, $existingNames) {
    $script:DialogConnResult = $null
    $isEdit = $null -ne $existing
    $title  = if ($isEdit) { "Edit Connection" } else { "Add Connection" }

    [xml]$dlgXaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="$title"
    Width="380" Height="420"
    WindowStartupLocation="CenterOwner"
    ResizeMode="NoResize"
    WindowStyle="None"
    AllowsTransparency="True"
    Background="Transparent"
    Foreground="White"
    FontFamily="Segoe UI">

    <Window.Resources>
        <Style x:Key="SmallBtn" TargetType="Button">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="14,7">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.85"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Border Background="#1a1a2e" CornerRadius="12" Padding="20">
            <Border.Effect>
                <DropShadowEffect Color="Black" BlurRadius="22" Opacity="0.55" ShadowDepth="0"/>
            </Border.Effect>
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/><RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Grid.Row="0" Text="$title"
                           FontSize="16" FontWeight="Bold" Foreground="White"
                           Margin="0,0,0,16" HorizontalAlignment="Center"/>

                <TextBlock Grid.Row="1" Text="Display Name" FontSize="12" Foreground="#888" Margin="0,0,0,4"/>
                <TextBox   Grid.Row="2" x:Name="NameBox"
                           Background="#16213e" Foreground="White" BorderBrush="#2a2a4a"
                           FontSize="13" Padding="8,6" Margin="0,0,0,12"/>

                <TextBlock Grid.Row="3" Text="Server \ Instance" FontSize="12" Foreground="#888" Margin="0,0,0,4"/>
                <TextBox   Grid.Row="4" x:Name="ServerBox"
                           Background="#16213e" Foreground="White" BorderBrush="#2a2a4a"
                           FontSize="13" Padding="8,6" Margin="0,0,0,12"/>

                <TextBlock Grid.Row="5" Text="Database" FontSize="12" Foreground="#888" Margin="0,0,0,4"/>
                <TextBox   Grid.Row="6" x:Name="DbBox"
                           Background="#16213e" Foreground="White" BorderBrush="#2a2a4a"
                           FontSize="13" Padding="8,6" Margin="0,0,0,12"/>

                <!-- Row 7: regenerate-reports checkbox. -->
                <CheckBox Grid.Row="7" x:Name="RegenChk"
                          Content="Regenerate reports on launch"
                          Foreground="#cfcfdc" FontSize="12"
                          Margin="0,4,0,8"
                          ToolTip="Sets RegenerateReport=True in GhostPractice's user.config every time you launch this connection. Leave on if this database needs reports refreshed; turn off for production databases where you don't want reports rebuilt every launch."/>

                <TextBlock Grid.Row="8" x:Name="ErrorText" Text="" FontSize="11"
                           Foreground="$BrandPink" HorizontalAlignment="Center" Margin="0,4,0,8" TextWrapping="Wrap"/>

                <StackPanel Grid.Row="9" Orientation="Horizontal" HorizontalAlignment="Center">
                    <Button x:Name="CancelBtn" Content="Cancel"
                            Background="#2a2a4a" Style="{StaticResource SmallBtn}" Margin="0,0,8,0"/>
                    <Button x:Name="SaveBtn" Content="Save"
                            Background="$BrandPurple" Style="{StaticResource SmallBtn}"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $dlgXaml
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader)

    $nameBox   = $dlg.FindName("NameBox")
    $serverBox = $dlg.FindName("ServerBox")
    $dbBox     = $dlg.FindName("DbBox")
    $regenChk  = $dlg.FindName("RegenChk")
    $errorText = $dlg.FindName("ErrorText")
    $cancelBtn = $dlg.FindName("CancelBtn")
    $saveBtn   = $dlg.FindName("SaveBtn")

    # Window-wide drag and Esc-to-close.
    $dlg.Add_MouseLeftButtonDown({ if ($_.ButtonState -eq 'Pressed') { $dlg.DragMove() } })
    $dlg.Add_KeyDown({ if ($_.Key -eq 'Escape') { $dlg.Close() } })

    if ($isEdit) {
        $nameBox.Text   = $existing.name
        $serverBox.Text = $existing.server
        $dbBox.Text     = $existing.database
        # Honor whatever the existing record says (default to true if missing).
        $regenChk.IsChecked = if ($existing.PSObject.Properties['regenerateReport']) { [bool]$existing.regenerateReport } else { $true }
    } else {
        # Brand-new connection - default to on, matching pre-0.0.8 behavior.
        $regenChk.IsChecked = $true
    }

    $cancelBtn.Add_Click({ $dlg.Close() })

    $script:DlgIsEdit        = $isEdit
    $script:DlgExistingName  = if ($isEdit) { $existing.name } else { $null }
    $script:DlgExistingNames = @($existingNames)

    $saveBtn.Add_Click({
        $n = $nameBox.Text.Trim()
        $s = $serverBox.Text.Trim()
        $d = $dbBox.Text.Trim()

        if (-not $n -or -not $s -or -not $d) {
            $errorText.Text = "All fields are required."
            return
        }
        $clash = $false
        foreach ($nm in $script:DlgExistingNames) {
            if (-not $nm) { continue }
            if ($nm.ToLower() -eq $n.ToLower()) {
                if (-not $script:DlgIsEdit -or $nm -ne $script:DlgExistingName) {
                    $clash = $true ; break
                }
            }
        }
        if ($clash) {
            $errorText.Text = "A connection named '$n' already exists."
            return
        }
        $script:DialogConnResult = New-Connection $n $s $d ([bool]$regenChk.IsChecked)
        $dlg.Close()
    })

    if ($parentWindow) { $dlg.Owner = $parentWindow }
    [void]$dlg.ShowDialog()
    return $script:DialogConnResult
}

# -----------------------------------------------------------------------------
# Build a ListBoxItem for a connection (used by Settings).
# -----------------------------------------------------------------------------
function New-ConnectionListItem($conn) {
    $item = New-Object System.Windows.Controls.ListBoxItem
    $sp = New-Object System.Windows.Controls.StackPanel
    $sp.Orientation = [System.Windows.Controls.Orientation]::Vertical

    $t1 = New-Object System.Windows.Controls.TextBlock
    $t1.Text = [string]$conn.name
    $t1.FontWeight = [System.Windows.FontWeights]::SemiBold
    $t1.Foreground = [System.Windows.Media.Brushes]::White
    [void]$sp.Children.Add($t1)

    $t2 = New-Object System.Windows.Controls.TextBlock
    $t2.Text = "$($conn.server)  -  $($conn.database)"
    $t2.FontSize = 11
    $t2.Foreground = (New-Brush '#888AA8')
    [void]$sp.Children.Add($t2)

    $item.Content = $sp
    $item.Tag = $conn
    return $item
}

# =============================================================================
#  SETTINGS WINDOW
# =============================================================================
function Show-SettingsWindow($parentWindow, $cfg) {
    [xml]$settingsXaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Settings - $AppVersion"
    Width="460" Height="600"
    WindowStartupLocation="CenterOwner"
    ResizeMode="NoResize"
    WindowStyle="None"
    AllowsTransparency="True"
    Background="Transparent"
    Foreground="White"
    FontFamily="Segoe UI">

    <Window.Resources>
        <Style x:Key="IconButton" TargetType="Button">
            <Setter Property="Width" Value="28"/>
            <Setter Property="Height" Value="24"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="#16213e"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="IconButtonClose" TargetType="Button" BasedOn="{StaticResource IconButton}">
            <Setter Property="Width" Value="26"/>
            <Setter Property="Height" Value="26"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="13" Width="26" Height="26">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Background" Value="$BrandPink"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SmallBtn" TargetType="Button">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="12,6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.85"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="16">
        <Border Background="#1a1a2e" CornerRadius="12" Padding="20">
            <Border.Effect>
                <DropShadowEffect Color="Black" BlurRadius="22" Opacity="0.55" ShadowDepth="0"/>
            </Border.Effect>
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>   <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Header: Back | Title | Close -->
                <Grid Grid.Row="0" Margin="0,0,0,16">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="32"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="32"/>
                    </Grid.ColumnDefinitions>
                    <Button Grid.Column="0" x:Name="BackBtn" Style="{StaticResource IconButton}"
                            HorizontalAlignment="Left" VerticalAlignment="Center"
                            ToolTip="Back (Esc, discards changes)">
                        <Path Stroke="White" StrokeThickness="1.5"
                              StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
                              Data="M 11,4 L 5,10 L 11,16 M 5,10 L 17,10"/>
                    </Button>
                    <TextBlock Grid.Column="1" Text="Settings"
                               FontSize="16" FontWeight="Bold" Foreground="White"
                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    <Button Grid.Column="2" x:Name="SettingsCloseBtn" Style="{StaticResource IconButtonClose}"
                            HorizontalAlignment="Right" VerticalAlignment="Center"
                            ToolTip="Close (discards changes)">
                        <Path Stroke="White" StrokeThickness="1.5"
                              StrokeStartLineCap="Round" StrokeEndLineCap="Round"
                              Width="12" Height="12" Stretch="Uniform"
                              HorizontalAlignment="Center" VerticalAlignment="Center"
                              Data="M 0,0 L 10,10 M 10,0 L 0,10"/>
                    </Button>
                </Grid>

                <TextBlock Grid.Row="1" Text="GP Executable Path" FontSize="12" Foreground="#888" Margin="0,0,0,4"/>
                <TextBox   Grid.Row="2" x:Name="GpPathBox"
                           Background="#16213e" Foreground="White" BorderBrush="#2a2a4a"
                           FontSize="13" Padding="8,6" Margin="0,0,0,16"/>

                <TextBlock Grid.Row="3" Text="Connections" FontSize="12" Foreground="#888" Margin="0,0,0,4"/>
                <Grid Grid.Row="4">
                    <ListBox x:Name="ConnList"
                             Background="#16213e" Foreground="White" BorderBrush="#2a2a4a"
                             FontSize="13" Padding="4">
                        <ListBox.ItemContainerStyle>
                            <Style TargetType="ListBoxItem">
                                <Setter Property="Padding" Value="6,6"/>
                            </Style>
                        </ListBox.ItemContainerStyle>
                    </ListBox>
                    <!-- Empty-state placeholder layered on top of an empty ListBox. -->
                    <TextBlock x:Name="ConnEmptyMsg" Visibility="Collapsed"
                               Text="No connections yet.&#x0A;Click + Add to create one."
                               FontSize="12" Foreground="#666680" TextAlignment="Center"
                               HorizontalAlignment="Center" VerticalAlignment="Center"
                               IsHitTestVisible="False"/>
                </Grid>

                <StackPanel Grid.Row="5" Orientation="Horizontal" Margin="0,8,0,0" HorizontalAlignment="Center">
                    <Button x:Name="AddBtn"    Content="+ Add"  Background="$BrandPurple" Style="{StaticResource SmallBtn}" Margin="0,0,8,0"/>
                    <Button x:Name="EditBtn"   Content="Edit"   Background="#2a2a4a"     Style="{StaticResource SmallBtn}" Margin="0,0,8,0"/>
                    <Button x:Name="RemoveBtn" Content="Remove" Background="$BrandPink"  Style="{StaticResource SmallBtn}"/>
                </StackPanel>

                <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,16,0,0">
                    <Button x:Name="SaveBtn" Content="Save"
                            Background="$BrandPurple" Foreground="White" FontSize="13" FontWeight="SemiBold"
                            Width="140" Style="{StaticResource SmallBtn}"/>
                </StackPanel>

                <StackPanel Grid.Row="7" Margin="0,12,0,0" HorizontalAlignment="Center">
                    <TextBlock x:Name="AboutVer" FontSize="10" Foreground="#444460" HorizontalAlignment="Center"/>
                    <TextBlock x:Name="AboutCreatedBy" FontSize="10" Foreground="#444460"
                               HorizontalAlignment="Center" Margin="0,2,0,0"/>
                    <TextBlock x:Name="AboutEmail" FontSize="10" Foreground="#555580"
                               HorizontalAlignment="Center" Margin="0,2,0,0"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $settingsXaml
    $settingsWin = [System.Windows.Markup.XamlReader]::Load($reader)

    $gpPathBox        = $settingsWin.FindName("GpPathBox")
    $connList         = $settingsWin.FindName("ConnList")
    $connEmptyMsg     = $settingsWin.FindName("ConnEmptyMsg")
    $addBtn           = $settingsWin.FindName("AddBtn")
    $editBtn          = $settingsWin.FindName("EditBtn")
    $removeBtn        = $settingsWin.FindName("RemoveBtn")
    $backBtn          = $settingsWin.FindName("BackBtn")
    $saveBtn          = $settingsWin.FindName("SaveBtn")
    $settingsCloseBtn = $settingsWin.FindName("SettingsCloseBtn")
    $aboutVer         = $settingsWin.FindName("AboutVer")

    $aboutVer.Text = "GP Launcher $AppVersion"

    # Wire up the easter-egg "Created by AZ" + email hyperlink in the About section.
    $aboutCreatedBy = $settingsWin.FindName("AboutCreatedBy")
    $aboutEmail     = $settingsWin.FindName("AboutEmail")
    Set-CreatedByInline    $aboutCreatedBy
    Set-EmailInline        $aboutEmail "Report issues to " "aubrey.zemba@dyedurham.com"

    # Window-wide drag.
    $settingsWin.Add_MouseLeftButtonDown({
        if ($_.ButtonState -eq 'Pressed') { $settingsWin.DragMove() }
    })

    # Esc closes without saving (matches Back/Close behavior).
    $closeWithoutSaving = {
        Write-Log "Settings closed without saving"
        $settingsWin.DialogResult = $false
        $settingsWin.Close()
    }
    $backBtn.Add_Click($closeWithoutSaving)
    $settingsCloseBtn.Add_Click($closeWithoutSaving)
    $settingsWin.Add_KeyDown({ if ($_.Key -eq 'Escape') { & $closeWithoutSaving } })

    $gpPathBox.Text = $cfg.gpPath

    # Toggle the empty-state placeholder whenever the list count changes.
    $script:RefreshEmptyMsg = {
        if ($connList.Items.Count -eq 0) { $connEmptyMsg.Visibility = 'Visible' }
        else                              { $connEmptyMsg.Visibility = 'Collapsed' }
    }

    foreach ($c in @($cfg.connections)) {
        if ($c -and $c.name) {
            [void]$connList.Items.Add((New-ConnectionListItem (New-Connection $c.name $c.server $c.database ([bool]$c.regenerateReport))))
        }
    }
    & $script:RefreshEmptyMsg
    Write-Log ("Settings opened with {0} connections" -f $connList.Items.Count)

    $addBtn.Add_Click({
        try {
            $names = @()
            foreach ($it in $connList.Items) { $names += $it.Tag.name }
            $new = Show-ConnectionDialog $settingsWin $null $names
            if ($new) {
                [void]$connList.Items.Add((New-ConnectionListItem $new))
                & $script:RefreshEmptyMsg
                Write-Log ("Added connection '{0}' (list now {1})" -f $new.name, $connList.Items.Count)
            }
        } catch { Write-Log ("Add error: " + $_) ; [System.Windows.MessageBox]::Show("Add failed: $_") | Out-Null }
    })

    $editBtn.Add_Click({
        try {
            $sel = $connList.SelectedItem
            if (-not $sel) { return }
            $existing = $sel.Tag
            $names = @()
            foreach ($it in $connList.Items) { $names += $it.Tag.name }
            $updated = Show-ConnectionDialog $settingsWin $existing $names
            if ($updated) {
                $idx = $connList.Items.IndexOf($sel)
                $connList.Items.RemoveAt($idx)
                $connList.Items.Insert($idx, (New-ConnectionListItem $updated))
                Write-Log ("Edited connection '{0}' -> '{1}'" -f $existing.name, $updated.name)
            }
        } catch { Write-Log ("Edit error: " + $_) ; [System.Windows.MessageBox]::Show("Edit failed: $_") | Out-Null }
    })

    $removeBtn.Add_Click({
        try {
            $sel = $connList.SelectedItem
            if (-not $sel) { return }
            $name = $sel.Tag.name
            $connList.Items.Remove($sel)
            & $script:RefreshEmptyMsg
            Write-Log ("Removed connection '{0}' (list now {1})" -f $name, $connList.Items.Count)
        } catch { Write-Log ("Remove error: " + $_) ; [System.Windows.MessageBox]::Show("Remove failed: $_") | Out-Null }
    })

    $saveBtn.Add_Click({
        try {
            $list = New-Object System.Collections.ArrayList
            foreach ($it in $connList.Items) {
                $t = $it.Tag
                if ($t -and $t.name) {
                    [void]$list.Add((New-Connection $t.name $t.server $t.database ([bool]$t.regenerateReport)))
                }
            }

            $cfg.gpPath      = $gpPathBox.Text.Trim()
            $cfg.connections = [object[]]$list.ToArray()

            $newUsage = [PSCustomObject]@{}
            foreach ($c in $cfg.connections) {
                $count = 0
                if ($cfg.usage.PSObject.Properties[$c.name]) {
                    $count = [int]$cfg.usage.$($c.name)
                }
                $newUsage | Add-Member -NotePropertyName $c.name -NotePropertyValue $count
            }
            $cfg.usage = $newUsage

            # Clear lastLaunched if that connection no longer exists.
            $stillExists = $false
            foreach ($c in $cfg.connections) { if ($c.name -eq $cfg.lastLaunched) { $stillExists = $true; break } }
            if (-not $stillExists) { $cfg.lastLaunched = "" }

            Save-Config $cfg
            Write-Log ("Saved settings with {0} connections" -f $cfg.connections.Count)
            $settingsWin.DialogResult = $true
            $settingsWin.Close()
        } catch { Write-Log ("Save error: " + $_) ; [System.Windows.MessageBox]::Show("Save failed: $_") | Out-Null }
    })

    if ($parentWindow) { $settingsWin.Owner = $parentWindow }
    return $settingsWin.ShowDialog()
}

# =============================================================================
#  MAIN WINDOW BOOTSTRAP
# =============================================================================

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

if (Test-Path $IconPath) {
    $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create(
        (New-Object System.Uri($IconPath, [System.UriKind]::Absolute)))
}

# Resolve named controls.
$dbPanel       = $window.FindName("DbPanel")
$searchBox     = $window.FindName("SearchBox")
$statusPill    = $window.FindName("StatusPill")
$statusIcon    = $window.FindName("StatusIcon")
$statusText    = $window.FindName("StatusText")
$countLabel    = $window.FindName("CountLabel")
$settingsBtn   = $window.FindName("SettingsBtn")
$closeBtn      = $window.FindName("CloseBtn")
$creditsLabel  = $window.FindName("CreditsLabel")
$logoSlot      = $window.FindName("LogoSlot")

$creditsLabel.Inlines.Clear()
[void]$creditsLabel.Inlines.Add((New-Object System.Windows.Documents.Run("$AppVersion  |  ")))
Set-CreatedByInline-Append $creditsLabel
[void]$creditsLabel.Inlines.Add((New-Object System.Windows.Documents.Run("  |  ")))
Set-EmailInline-Append    $creditsLabel "aubrey.zemba@dyedurham.com"

# -----------------------------------------------------------------------------
# Snark popup for the AZ easter egg. Show-Snark fades a small pill in/out
# above the credits line. Multiple rapid calls cancel any in-flight fade and
# restart so the latest message always wins.
# -----------------------------------------------------------------------------
$snarkPill = $window.FindName("SnarkPill")
$snarkText = $window.FindName("SnarkText")
$script:SnarkHideTimer = $null

function Show-Snark($message) {
    $snarkText.Text = $message
    $snarkPill.Visibility = 'Visible'

    # Cancel any prior fade so this one starts fresh.
    $snarkPill.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $null)

    $fadeIn = New-Object System.Windows.Media.Animation.DoubleAnimation
    $fadeIn.From = $snarkPill.Opacity
    $fadeIn.To = 1
    $fadeIn.Duration = New-Object System.Windows.Duration([TimeSpan]::FromMilliseconds(150))
    $snarkPill.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fadeIn)

    if ($script:SnarkHideTimer) { $script:SnarkHideTimer.Stop() }
    $script:SnarkHideTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:SnarkHideTimer.Interval = [TimeSpan]::FromMilliseconds(2500)
    $script:SnarkHideTimer.Add_Tick({
        $script:SnarkHideTimer.Stop()
        $fadeOut = New-Object System.Windows.Media.Animation.DoubleAnimation
        $fadeOut.From = $snarkPill.Opacity
        $fadeOut.To = 0
        $fadeOut.Duration = New-Object System.Windows.Duration([TimeSpan]::FromMilliseconds(300))
        $fadeOut.Add_Completed({ $snarkPill.Visibility = 'Collapsed' })
        $snarkPill.BeginAnimation([System.Windows.UIElement]::OpacityProperty, $fadeOut)
    })
    $script:SnarkHideTimer.Start()
}

# -----------------------------------------------------------------------------
# Logo: prefer gp-logo.png if present, otherwise render a brand-colored
# "GP" monogram inline.
# -----------------------------------------------------------------------------
$logoBitmap = Get-LogoBitmap
if ($logoBitmap) {
    # Real GP logo (1742x417, ~4.18:1 aspect). Stretch=Uniform preserves the
    # ratio; Height drives the actual rendered size.
    $img = New-Object System.Windows.Controls.Image
    $img.Source = $logoBitmap
    $img.Height = 64
    $img.Stretch = [System.Windows.Media.Stretch]::Uniform
    $img.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
    $img.VerticalAlignment   = [System.Windows.VerticalAlignment]::Center
    $logoSlot.Content = $img
} else {
    # Monogram fallback (when gp-logo.png isn't present): same centered placement.
    $row = New-Object System.Windows.Controls.StackPanel
    $row.Orientation = [System.Windows.Controls.Orientation]::Horizontal
    $row.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
    $row.VerticalAlignment   = [System.Windows.VerticalAlignment]::Center

    $badge = New-Object System.Windows.Controls.Border
    $badge.Background = New-Brush $BrandPurple
    $badge.CornerRadius = New-Object System.Windows.CornerRadius(8)
    $badge.Width = 36; $badge.Height = 36
    $badge.VerticalAlignment = [System.Windows.VerticalAlignment]::Center

    $gp = New-Object System.Windows.Controls.TextBlock
    $gp.Text = "GP"
    $gp.Foreground = New-Brush $BrandPink
    $gp.FontWeight = [System.Windows.FontWeights]::Bold
    $gp.FontSize = 16
    $gp.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Center
    $gp.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $badge.Child = $gp
    [void]$row.Children.Add($badge)

    $word = New-Object System.Windows.Controls.TextBlock
    $word.Text = "GhostPractice"
    $word.Foreground = [System.Windows.Media.Brushes]::White
    $word.FontWeight = [System.Windows.FontWeights]::Bold
    $word.FontSize = 20
    $word.Margin = New-Object System.Windows.Thickness(10, 0, 0, 0)
    $word.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    [void]$row.Children.Add($word)

    $logoSlot.Content = $row
}

# Window-wide drag.
$window.Add_MouseLeftButtonDown({
    if ($_.ButtonState -eq 'Pressed') { $window.DragMove() }
})

# Header close icon exits the app.
$closeBtn.Add_Click({ $window.Close() })

# Frozen DrawingImage shared by every connection-button icon.
$dbDrawingImage = New-Object System.Windows.Media.DrawingImage($dbIconDrawing)
$dbDrawingImage.Freeze()

# -----------------------------------------------------------------------------
# Sort connections most-used-first, then alphabetically.
# -----------------------------------------------------------------------------
function Get-SortedConnections($cfg) {
    $cfg.connections | Sort-Object {
        $count = 0
        if ($cfg.usage.PSObject.Properties[$_.name]) {
            $count = [int]$cfg.usage.$($_.name)
        }
        -$count
    }, { $_.name }
}

# -----------------------------------------------------------------------------
# Show the status pill with an icon, message, and color. Auto-hide after 3.5s.
# -----------------------------------------------------------------------------
$script:StatusTimer = $null
function Show-Status($iconText, $message, $color) {
    $statusIcon.Text = $iconText
    $statusIcon.Foreground = New-Brush $color
    $statusText.Text = $message
    $statusPill.Background = New-Brush "#1f2235"
    $statusPill.Visibility = 'Visible'

    if ($script:StatusTimer) { $script:StatusTimer.Stop() }
    $script:StatusTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:StatusTimer.Interval = [TimeSpan]::FromSeconds(3.5)
    $script:StatusTimer.Add_Tick({
        $statusPill.Visibility = 'Collapsed'
        $script:StatusTimer.Stop()
    })
    $script:StatusTimer.Start()
}

# -----------------------------------------------------------------------------
# Build a launcher button for one connection. Includes:
#   - Left color stripe (one color per server, deterministic)
#   - Database icon
#   - Bold name + grey subtitle (server  -  database)
#   - "last" green dot if this was the most recently launched connection
#   - Tooltip with full server\database
#   - Right-click context menu (Edit / Remove / Reset usage)
# -----------------------------------------------------------------------------
function New-ConnectionButton($conn, $isLast) {
    $btn = New-Object System.Windows.Controls.Button
    $btn.Style = $window.FindResource("DbButton")
    $btn.Tag   = $conn
    $regenLine = if ($conn.regenerateReport) { "Reports: regenerate on launch" } else { "Reports: don't regenerate" }
    $btn.ToolTip = "$($conn.server)`n$($conn.database)`n$regenLine"

    # Outer 3-column layout: stripe | content | last-indicator
    $grid = New-Object System.Windows.Controls.Grid
    $col1 = New-Object System.Windows.Controls.ColumnDefinition; $col1.Width = '4'
    $col2 = New-Object System.Windows.Controls.ColumnDefinition; $col2.Width = '*'
    $col3 = New-Object System.Windows.Controls.ColumnDefinition; $col3.Width = 'Auto'
    [void]$grid.ColumnDefinitions.Add($col1)
    [void]$grid.ColumnDefinitions.Add($col2)
    [void]$grid.ColumnDefinitions.Add($col3)

    # Left stripe colored per server.
    $stripe = New-Object System.Windows.Controls.Border
    $stripe.Background = New-Brush (Get-ServerColor $conn.server)
    $stripe.CornerRadius = New-Object System.Windows.CornerRadius(8, 0, 0, 8)
    [System.Windows.Controls.Grid]::SetColumn($stripe, 0)
    [void]$grid.Children.Add($stripe)

    # Content column: icon + text.
    $inner = New-Object System.Windows.Controls.StackPanel
    $inner.Orientation = [System.Windows.Controls.Orientation]::Horizontal
    $inner.Margin = New-Object System.Windows.Thickness(14, 0, 12, 0)
    [System.Windows.Controls.Grid]::SetColumn($inner, 1)

    $img = New-Object System.Windows.Controls.Image
    $img.Source = $dbDrawingImage
    $img.Width  = 22; $img.Height = 22
    $img.Margin = New-Object System.Windows.Thickness(0, 0, 12, 0)
    $img.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    [void]$inner.Children.Add($img)

    $textPanel = New-Object System.Windows.Controls.StackPanel
    $textPanel.Orientation = [System.Windows.Controls.Orientation]::Vertical
    $textPanel.VerticalAlignment = [System.Windows.VerticalAlignment]::Center

    $nameText = New-Object System.Windows.Controls.TextBlock
    $nameText.Text = $conn.name
    $nameText.Foreground = [System.Windows.Media.Brushes]::White
    $nameText.FontSize = 14
    $nameText.FontWeight = [System.Windows.FontWeights]::SemiBold
    [void]$textPanel.Children.Add($nameText)

    $subText = New-Object System.Windows.Controls.TextBlock
    $subText.Text = "$($conn.server)  -  $($conn.database)"
    $subText.FontSize = 11
    $subText.Foreground = New-Brush '#888AA8'
    $subText.Margin = New-Object System.Windows.Thickness(0, 2, 0, 0)
    $subText.TextTrimming = [System.Windows.TextTrimming]::CharacterEllipsis
    [void]$textPanel.Children.Add($subText)

    [void]$inner.Children.Add($textPanel)
    [void]$grid.Children.Add($inner)

    # Right column: green "last launched" dot (only for the most recent click).
    if ($isLast) {
        $dot = New-Object System.Windows.Shapes.Ellipse
        $dot.Width = 9; $dot.Height = 9
        $dot.Fill = New-Brush '#3fb950'
        $dot.Margin = New-Object System.Windows.Thickness(0, 0, 12, 0)
        $dot.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $dot.ToolTip = "Last launched"
        [System.Windows.Controls.Grid]::SetColumn($dot, 2)
        [void]$grid.Children.Add($dot)
    }

    $btn.Content = $grid

    # Right-click context menu: quick edit / remove / reset usage.
    $menu = New-Object System.Windows.Controls.ContextMenu
    $miEdit = New-Object System.Windows.Controls.MenuItem; $miEdit.Header = "Edit..."
    $miRem  = New-Object System.Windows.Controls.MenuItem; $miRem.Header  = "Remove"
    $miSep  = New-Object System.Windows.Controls.Separator
    $miReset= New-Object System.Windows.Controls.MenuItem; $miReset.Header= "Reset usage count"
    [void]$menu.Items.Add($miEdit); [void]$menu.Items.Add($miRem)
    [void]$menu.Items.Add($miSep);  [void]$menu.Items.Add($miReset)
    $btn.ContextMenu = $menu

    $miEdit.Add_Click({
        $names = @($Config.connections | ForEach-Object { $_.name })
        $updated = Show-ConnectionDialog $window $conn $names
        if ($updated) {
            for ($i = 0; $i -lt $Config.connections.Count; $i++) {
                if ($Config.connections[$i].name -eq $conn.name) {
                    $Config.connections[$i] = $updated
                    if ($Config.usage.PSObject.Properties[$conn.name] -and $conn.name -ne $updated.name) {
                        $Config.usage | Add-Member -NotePropertyName $updated.name -NotePropertyValue ([int]$Config.usage.$($conn.name)) -Force
                        $Config.usage.PSObject.Properties.Remove($conn.name)
                    }
                    if ($Config.lastLaunched -eq $conn.name) { $Config.lastLaunched = $updated.name }
                    break
                }
            }
            Save-Config $Config
            Populate-Connections
        }
    }.GetNewClosure())

    $miRem.Add_Click({
        $Config.connections = @($Config.connections | Where-Object { $_.name -ne $conn.name })
        if ($Config.usage.PSObject.Properties[$conn.name]) { $Config.usage.PSObject.Properties.Remove($conn.name) }
        if ($Config.lastLaunched -eq $conn.name) { $Config.lastLaunched = "" }
        Save-Config $Config
        Populate-Connections
    }.GetNewClosure())

    $miReset.Add_Click({
        if ($Config.usage.PSObject.Properties[$conn.name]) { $Config.usage.$($conn.name) = 0 }
        Save-Config $Config
        Populate-Connections
    }.GetNewClosure())

    return $btn
}

# -----------------------------------------------------------------------------
# Filter helper for the search box.
# -----------------------------------------------------------------------------
function Match-ConnectionFilter($conn, $needle) {
    if (-not $needle) { return $true }
    $n = $needle.ToLower()
    return ($conn.name.ToLower()    -like "*$n*") -or
           ($conn.server.ToLower()  -like "*$n*") -or
           ($conn.database.ToLower() -like "*$n*")
}

# -----------------------------------------------------------------------------
# Re-build the list of connection buttons. Called at startup, after Settings
# Save, after right-click actions, and after the search box changes.
# -----------------------------------------------------------------------------
function Populate-Connections {
    $dbPanel.Children.Clear()
    $allSorted = @(Get-SortedConnections $Config) | Where-Object { $_ -and $_.name }

    # Empty-state
    if (-not $allSorted -or @($allSorted).Count -eq 0) {
        $empty = New-Object System.Windows.Controls.TextBlock
        $empty.Text = "No connections configured.`nClick Settings (or press F2) to add one."
        $empty.Foreground = New-Brush '#666680'
        $empty.FontSize = 12
        $empty.TextAlignment = [System.Windows.TextAlignment]::Center
        $empty.Margin = New-Object System.Windows.Thickness(0, 40, 0, 0)
        [void]$dbPanel.Children.Add($empty)
        $countLabel.Text = "0 connections"
        $searchBox.Visibility = 'Collapsed'
        return
    }

    # Show / hide the search box based on list length.
    $searchBox.Visibility = if (@($allSorted).Count -gt 8) { 'Visible' } else { 'Collapsed' }

    # Apply current search filter (if any).
    $needle = $searchBox.Text
    if ($needle -and $needle -eq $searchBox.Tag) { $needle = '' }  # ignore placeholder

    $filtered = @($allSorted | Where-Object { Match-ConnectionFilter $_ $needle })

    foreach ($conn in $filtered) {
        $isLast = ($Config.lastLaunched -and $conn.name -eq $Config.lastLaunched)
        $btn = New-ConnectionButton $conn $isLast

        $btn.Add_Click({
            param($sender, $e)
            $selected = $sender.Tag

            $targetDir = Get-GPConfigTargetPath
            if (-not $targetDir) {
                Show-Status "ERR" "Could not find GP config path. Run GhostPractice manually once first." '#ff7b72'
                return
            }
            if (-not (Test-Path $Config.gpPath)) {
                Show-Status "ERR" "GhostPractice.exe not found at configured path." '#ff7b72'
                return
            }

            $version = Split-Path $targetDir -Leaf
            $configFile = Join-Path $targetDir "user.config"
            $regen = if ($selected.PSObject.Properties['regenerateReport']) { [bool]$selected.regenerateReport } else { $true }

            # Surgical edit if user.config already exists; full-template fallback
            # only for brand-new installs that haven't run GP yet.
            try {
                if (Test-Path $configFile) {
                    Backup-UserConfigOnce $configFile
                    Update-UserConfig $configFile $selected.server $selected.database $version $regen
                    Write-Log ("Updated user.config: Server={0}  Database={1}  Version={2}  RegenerateReport={3}" -f $selected.server, $selected.database, $version, $regen)
                } else {
                    $xml = Build-UserConfig $selected.server $selected.database $version $regen
                    $xml | Set-Content $configFile -Encoding UTF8 -Force
                    Write-Log ("Created new user.config from template (Server={0}, Database={1}, RegenerateReport={2})" -f $selected.server, $selected.database, $regen)
                }
            }
            catch { Show-Status "ERR" "Error updating config: $_" '#ff7b72'; return }

            try { Start-Process $Config.gpPath }
            catch { Show-Status "ERR" "Error launching: $_" '#ff7b72'; return }

            # Bump usage and mark as last-launched.
            if (-not $Config.usage.PSObject.Properties[$selected.name]) {
                $Config.usage | Add-Member -NotePropertyName $selected.name -NotePropertyValue 0
            }
            $Config.usage.$($selected.name) = [int]$Config.usage.$($selected.name) + 1
            $Config.lastLaunched = $selected.name
            Save-Config $Config

            Show-Status "OK" ("Launched " + $selected.name) '#3fb950'
        })

        [void]$dbPanel.Children.Add($btn)
    }

    # Footer count: shown / total when filtering is active.
    $total = @($allSorted).Count
    $shown = @($filtered).Count
    if ($needle) {
        $countLabel.Text = "$shown of $total connections"
    } else {
        $countLabel.Text = if ($total -eq 1) { "1 connection" } else { "$total connections" }
    }
}

# -----------------------------------------------------------------------------
# Search box: placeholder text + live filtering.
# -----------------------------------------------------------------------------
$searchBox.Text = $searchBox.Tag  # initial placeholder
$searchBox.Foreground = New-Brush '#666680'
$searchBox.Add_GotFocus({
    if ($searchBox.Text -eq $searchBox.Tag) {
        $searchBox.Text = ''
        $searchBox.Foreground = [System.Windows.Media.Brushes]::White
    }
})
$searchBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($searchBox.Text)) {
        $searchBox.Text = $searchBox.Tag
        $searchBox.Foreground = New-Brush '#666680'
    }
})
$searchBox.Add_TextChanged({ Populate-Connections })

# -----------------------------------------------------------------------------
# Settings button: open dialog. If user clicks Save, reload + re-render.
# -----------------------------------------------------------------------------
$settingsBtn.Add_Click({
    $result = Show-SettingsWindow $window $Config
    if ($result) {
        $script:Config = Load-Config
        Write-Log ("Reloaded config: {0} connections" -f (@($script:Config.connections)).Count)
        Populate-Connections
    }
})

# -----------------------------------------------------------------------------
# Keyboard shortcuts on the main window:
#   Esc  = close       F2 = Settings
# -----------------------------------------------------------------------------
$window.Add_KeyDown({
    if ($_.Key -eq 'Escape') { $window.Close(); return }
    if ($_.Key -eq 'F2')     { $settingsBtn.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent))) }
})

# -----------------------------------------------------------------------------
# First-run wizard: if this is a brand-new install with no connections,
# show the Add Connection dialog before the main window appears.
# -----------------------------------------------------------------------------
if (@($Config.connections).Count -eq 0) {
    Write-Log "First run detected (no connections); prompting to add one"
    $first = Show-ConnectionDialog $null $null @()
    if ($first) {
        $Config.connections = @($first)
        $Config.usage | Add-Member -NotePropertyName $first.name -NotePropertyValue 0 -Force
        Save-Config $Config
    }
}

# Initial render.
Populate-Connections

# Show the window modally; script exits when the window is closed.
$window.ShowDialog() | Out-Null
