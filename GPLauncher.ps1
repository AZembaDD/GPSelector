Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# --- Resolve paths relative to script location ---
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ConfigPath = Join-Path $ScriptDir "config.json"
$IconPath = Join-Path $ScriptDir "icon.ico"

# --- Load or create config ---
function Load-Config {
    if (Test-Path $ConfigPath) {
        $cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json
        # Ensure usage property exists
        if (-not $cfg.PSObject.Properties['usage']) {
            $usageObj = @{}
            foreach ($db in $cfg.databases) { $usageObj[$db] = 0 }
            $cfg | Add-Member -NotePropertyName 'usage' -NotePropertyValue ([PSCustomObject]$usageObj)
            Save-Config $cfg
        }
        return $cfg
    }
    $default = @{
        server    = "10.10.10.21"
        gpPath    = "C:\Program Files (x86)\Korbicom\GhostPractice\GhostPractice.exe"
        databases = @("GD", "GhostpracticeCanada", "KLS", "Korbicom", "KorbicomSA")
        usage     = @{ GD = 0; GhostpracticeCanada = 0; KLS = 0; Korbicom = 0; KorbicomSA = 0 }
    }
    $default | ConvertTo-Json -Depth 3 | Set-Content $ConfigPath -Encoding UTF8
    return [PSCustomObject]$default
}

function Save-Config($cfg) {
    $cfg | ConvertTo-Json -Depth 3 | Set-Content $ConfigPath -Encoding UTF8
}

$Config = Load-Config

# --- Detect GP version path ---
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

function Get-DetectedVersion {
    $target = Get-GPConfigTargetPath
    if ($target) { return (Split-Path $target -Leaf) }
    return "0.0.0.0"
}

# --- Generate user.config XML ---
function Build-UserConfig($server, $database, $version) {
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
            <setting name="FormLocation" serializeAs="String">
                <value>0, 0</value>
            </setting>
            <setting name="HasUpgraded" serializeAs="String">
                <value>True</value>
            </setting>
            <setting name="Server" serializeAs="String">
                <value>$server</value>
            </setting>
            <setting name="Database" serializeAs="String">
                <value>$database</value>
            </setting>
            <setting name="ApplicationVersion" serializeAs="String">
                <value>$version</value>
            </setting>
            <setting name="MessagingProgram" serializeAs="String">
                <value>1</value>
            </setting>
            <setting name="ShowDBConnectionDetails" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="FormSize" serializeAs="String">
                <value>88, 28</value>
            </setting>
            <setting name="ShowConversionTools" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="ApplicationDeviceID" serializeAs="String">
                <value>00000000-0000-0000-0000-000000000000</value>
            </setting>
            <setting name="InvoiceWizardOpenMaximized" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="AllowAdvancedSupport" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="EnableGenerateStatementDocuments" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="AllowAllocationOfOldMatters" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="ShowHMRCTestOptions" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="ShowMonthEndGB" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="ShowDocumentDataMigrationVersion10" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="ShowDocumentMetaDataMigration" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="ShowTransactionID" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="ResyncDocumentNames" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="GenerateDataset" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="LogProgress" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="EnableEditingE4XML" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="RegenerateReport" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="AllowAgingRecalculate" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="ShowWarningForUncomittedTransactions" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="ImageTextAlignment" serializeAs="String">
                <value>1</value>
            </setting>
            <setting name="FormMaximised" serializeAs="String">
                <value>False</value>
            </setting>
            <setting name="ReportViewColumnWidthCreateDate" serializeAs="String">
                <value>200</value>
            </setting>
            <setting name="ReportAdHocColumnWidthDescription" serializeAs="String">
                <value>300</value>
            </setting>
            <setting name="TransactionHistoryRunningBalanceView" serializeAs="String">
                <value>0</value>
            </setting>
            <setting name="TransactionHistoryAccountType" serializeAs="String">
                <value>2</value>
            </setting>
            <setting name="ReportViewColumnWidthCreateBy" serializeAs="String">
                <value>200</value>
            </setting>
            <setting name="RunReportsAdHocCategory" serializeAs="String">
                <value>2</value>
            </setting>
            <setting name="ReportAdHocColumnWidthName" serializeAs="String">
                <value>300</value>
            </setting>
            <setting name="LegalDiaryActivitiesHeight" serializeAs="String">
                <value>254</value>
            </setting>
            <setting name="ReportAdHocReference" serializeAs="String">
                <value>80</value>
            </setting>
            <setting name="LegalDiaryAdHocHeight" serializeAs="String">
                <value>187</value>
            </setting>
            <setting name="ReportAdHocType" serializeAs="String">
                <value>80</value>
            </setting>
            <setting name="TransactionHistoryPortrait" serializeAs="String">
                <value>Landscape</value>
            </setting>
            <setting name="ReportViewColumnWidthHelpAdHoc" serializeAs="String">
                <value>40</value>
            </setting>
            <setting name="ReportStoredColumnWidthDescription" serializeAs="String">
                <value>300</value>
            </setting>
            <setting name="ReportStoredColumnWidthArchive" serializeAs="String">
                <value>100</value>
            </setting>
            <setting name="ReportStoredColumnWidthReference" serializeAs="String">
                <value>300</value>
            </setting>
            <setting name="ReportStoredColumnWidthDate" serializeAs="String">
                <value>120</value>
            </setting>
        </Korbicom.Mustang.ApplicationUserSettings>
    </userSettings>
</configuration>
"@
}

# --- SVG database icon as base64-encoded PNG (rendered via WPF drawing) ---
function New-DbIconDrawing {
    # Draw a simple database cylinder icon programmatically
    $dg = New-Object System.Windows.Media.DrawingGroup

    $bodyBrush = [System.Windows.Media.Brushes]::Transparent
    $strokePen = New-Object System.Windows.Media.Pen([System.Windows.Media.Brushes]::White, 1.2)
    $strokePen.LineJoin = [System.Windows.Media.PenLineJoin]::Round

    # Cylinder body (rectangle with side lines)
    $bodyGeo = [System.Windows.Media.Geometry]::Parse("M 4,6 L 4,18 A 8,3 0 0 0 20,18 L 20,6")
    $bodyDrawing = New-Object System.Windows.Media.GeometryDrawing($bodyBrush, $strokePen, $bodyGeo)
    $dg.Children.Add($bodyDrawing)

    # Top ellipse
    $topGeo = New-Object System.Windows.Media.EllipseGeometry(
        (New-Object System.Windows.Point(12, 6)), 8, 3)
    $topDrawing = New-Object System.Windows.Media.GeometryDrawing($bodyBrush, $strokePen, $topGeo)
    $dg.Children.Add($topDrawing)

    # Middle line
    $midGeo = [System.Windows.Media.Geometry]::Parse("M 4,12 A 8,3 0 0 0 20,12")
    $midDrawing = New-Object System.Windows.Media.GeometryDrawing($bodyBrush, $strokePen, $midGeo)
    $dg.Children.Add($midDrawing)

    return $dg
}

# --- Build the WPF window ---
$dbIconDrawing = New-DbIconDrawing

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="GhostPractice Launcher - Beta 0.0.1"
    Width="380" Height="540"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    WindowStyle="None"
    AllowsTransparency="True"
    Background="Transparent"
    Foreground="White"
    FontFamily="Segoe UI">

    <Window.Resources>
        <Style x:Key="DbButton" TargetType="Button">
            <Setter Property="Background" Value="#16213e"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Height" Value="52"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#2a2a4a"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Padding" Value="16,0,0,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="8"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}"
                                              VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#0f3460"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#e94560"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#e94560"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SettingsButton" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#666680"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border" Background="Transparent" Padding="8,4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#16213e"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Border Background="#1a1a2e" CornerRadius="12" Padding="0">
    <Grid Margin="24,20,24,16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Grid Grid.Row="0" x:Name="TitleBar" Margin="0,0,0,0">
            <TextBlock Text="GhostPractice"
                       FontSize="22" FontWeight="Bold"
                       Foreground="White"
                       HorizontalAlignment="Center"/>
            <TextBlock x:Name="CloseBtn" Text="X"
                       FontSize="14" Foreground="#666680"
                       HorizontalAlignment="Right" VerticalAlignment="Top"
                       Cursor="Hand" Margin="0,-4,0,0"/>
        </Grid>

        <TextBlock Grid.Row="1"
                   Text="Select a database to launch"
                   FontSize="12"
                   Foreground="#666680"
                   HorizontalAlignment="Center"
                   Margin="0,2,0,20"/>

        <!-- Database buttons container -->
        <StackPanel Grid.Row="2" x:Name="DbPanel" VerticalAlignment="Top"/>

        <!-- Status bar -->
        <TextBlock Grid.Row="3"
                   x:Name="StatusText"
                   Text=""
                   FontSize="11"
                   Foreground="#e94560"
                   HorizontalAlignment="Center"
                   Margin="0,8,0,8"
                   TextWrapping="Wrap"/>

        <!-- Footer: server info + settings -->
        <Grid Grid.Row="4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>

            <TextBlock Grid.Column="0"
                       x:Name="ServerLabel"
                       FontSize="11"
                       Foreground="#444460"
                       VerticalAlignment="Center"/>

            <Button Grid.Column="1"
                    x:Name="SettingsBtn"
                    Content="Settings"
                    Style="{StaticResource SettingsButton}"/>
        </Grid>

        <!-- Credits -->
        <TextBlock Grid.Row="5"
                   Text="Beta 0.0.1  |  Created by AZ  |  aubrey.zemba@dyedurham.com"
                   FontSize="9"
                   Foreground="#333350"
                   HorizontalAlignment="Center"
                   Margin="0,8,0,0"/>
    </Grid>
    </Border>
</Window>
"@

# --- Settings window XAML ---
function Show-SettingsWindow($parentWindow, $cfg) {
    [xml]$settingsXaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Settings - Beta 0.0.1"
    Width="360" Height="500"
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
                        <Border x:Name="border"
                                Background="{TemplateBinding Background}"
                                CornerRadius="6"
                                Padding="12,6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.85"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Opacity" Value="0.7"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Border Background="#1a1a2e" CornerRadius="12" Padding="20">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Draggable title bar -->
        <TextBlock Grid.Row="0" x:Name="SettingsTitleBar" Text="Settings"
                   FontSize="16" FontWeight="Bold" Foreground="White"
                   Margin="0,0,0,16" HorizontalAlignment="Center"/>

        <TextBlock Grid.Row="0" Text="Server" FontSize="12" Foreground="#888" Margin="0,30,0,4"/>
        <TextBox Grid.Row="1" x:Name="ServerBox"
                 Background="#16213e" Foreground="White" BorderBrush="#2a2a4a"
                 FontSize="13" Padding="8,6" Margin="0,0,0,16"/>

        <TextBlock Grid.Row="2" Text="GP Executable Path" FontSize="12" Foreground="#888" Margin="0,0,0,4"/>
        <TextBox Grid.Row="3" x:Name="GpPathBox"
                 Background="#16213e" Foreground="White" BorderBrush="#2a2a4a"
                 FontSize="13" Padding="8,6" Margin="0,0,0,16"/>

        <Grid Grid.Row="4">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <TextBlock Grid.Row="0" Text="Databases" FontSize="12" Foreground="#888" Margin="0,0,0,4"/>
            <ListBox Grid.Row="1" x:Name="DbList"
                     Background="#16213e" Foreground="White" BorderBrush="#2a2a4a"
                     FontSize="13" Padding="4"/>

            <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,0">
                <TextBox x:Name="NewDbBox"
                         Background="#16213e" Foreground="White" BorderBrush="#2a2a4a"
                         FontSize="13" Padding="8,4" Width="160"/>
                <Button x:Name="AddDbBtn" Content="Add" Margin="8,0,0,0"
                        Background="#0f3460" Style="{StaticResource SmallBtn}"/>
                <Button x:Name="RemoveDbBtn" Content="Remove" Margin="8,0,0,0"
                        Background="#e94560" Style="{StaticResource SmallBtn}"/>
            </StackPanel>
        </Grid>

        <Button Grid.Row="5" x:Name="SaveBtn" Content="Save"
                Background="#0f3460" Foreground="White" FontSize="13" FontWeight="SemiBold"
                Width="100" HorizontalAlignment="Center"
                Margin="0,16,0,0" Style="{StaticResource SmallBtn}"/>

        <!-- About -->
        <StackPanel Grid.Row="6" Margin="0,12,0,0" HorizontalAlignment="Center">
            <TextBlock Text="GP Launcher Beta 0.0.1" FontSize="10" Foreground="#444460" HorizontalAlignment="Center"/>
            <TextBlock Text="Created by AZ" FontSize="10" Foreground="#444460" HorizontalAlignment="Center" Margin="0,2,0,0"/>
            <TextBlock Text="Report issues to aubrey.zemba@dyedurham.com" FontSize="10" Foreground="#555580" HorizontalAlignment="Center" Margin="0,2,0,0"/>
        </StackPanel>
    </Grid>
    </Border>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $settingsXaml
    $settingsWin = [System.Windows.Markup.XamlReader]::Load($reader)

    $serverBox = $settingsWin.FindName("ServerBox")
    $gpPathBox = $settingsWin.FindName("GpPathBox")
    $dbList = $settingsWin.FindName("DbList")
    $newDbBox = $settingsWin.FindName("NewDbBox")
    $addDbBtn = $settingsWin.FindName("AddDbBtn")
    $removeDbBtn = $settingsWin.FindName("RemoveDbBtn")
    $saveBtn = $settingsWin.FindName("SaveBtn")
    $settingsTitleBar = $settingsWin.FindName("SettingsTitleBar")

    # Drag settings window
    $settingsTitleBar.Add_MouseLeftButtonDown({ $settingsWin.DragMove() })

    $serverBox.Text = $cfg.server
    $gpPathBox.Text = $cfg.gpPath
    foreach ($db in $cfg.databases) { $dbList.Items.Add($db) | Out-Null }

    $addDbBtn.Add_Click({
        $name = $newDbBox.Text.Trim()
        if ($name -and -not $dbList.Items.Contains($name)) {
            $dbList.Items.Add($name) | Out-Null
            $newDbBox.Text = ""
        }
    })

    $removeDbBtn.Add_Click({
        if ($dbList.SelectedItem) {
            $dbList.Items.Remove($dbList.SelectedItem)
        }
    })

    $saveBtn.Add_Click({
        $cfg.server = $serverBox.Text.Trim()
        $cfg.gpPath = $gpPathBox.Text.Trim()
        $cfg.databases = @($dbList.Items | ForEach-Object { $_.ToString() })
        # Ensure usage entries exist for all databases
        foreach ($db in $cfg.databases) {
            if (-not $cfg.usage.PSObject.Properties[$db]) {
                $cfg.usage | Add-Member -NotePropertyName $db -NotePropertyValue 0
            }
        }
        Save-Config $cfg
        $settingsWin.DialogResult = $true
        $settingsWin.Close()
    })

    $settingsWin.Owner = $parentWindow
    return $settingsWin.ShowDialog()
}

# --- Create the main window ---
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# Set window icon
if (Test-Path $IconPath) {
    $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create(
        (New-Object System.Uri($IconPath, [System.UriKind]::Absolute)))
}

$dbPanel = $window.FindName("DbPanel")
$statusText = $window.FindName("StatusText")
$serverLabel = $window.FindName("ServerLabel")
$settingsBtn = $window.FindName("SettingsBtn")
$titleBar = $window.FindName("TitleBar")
$closeBtn = $window.FindName("CloseBtn")

# Drag window by title area
$titleBar.Add_MouseLeftButtonDown({ $window.DragMove() })

# Close button
$closeBtn.Add_MouseLeftButtonDown({ $window.Close() })
$closeBtn.Add_MouseEnter({ $closeBtn.Foreground = [System.Windows.Media.Brushes]::White })
$closeBtn.Add_MouseLeave({ $closeBtn.Foreground = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString('#666680'))) })

$serverLabel.Text = "Server: $($Config.server)"

# --- Create a WPF DrawingImage for the DB icon ---
$dbDrawingImage = New-Object System.Windows.Media.DrawingImage($dbIconDrawing)
$dbDrawingImage.Freeze()

# --- Get databases sorted by usage (descending), then alphabetical ---
function Get-SortedDatabases($cfg) {
    $cfg.databases | Sort-Object {
        $count = 0
        if ($cfg.usage.PSObject.Properties[$_]) {
            $count = [int]$cfg.usage.$_
        }
        -$count  # negative for descending
    }, { $_ }
}

# --- Populate database buttons ---
function Populate-DbButtons {
    $dbPanel.Children.Clear()
    $sorted = Get-SortedDatabases $Config
    foreach ($db in $sorted) {
        $btn = New-Object System.Windows.Controls.Button
        $btn.Style = $window.FindResource("DbButton")
        $btn.Tag = $db

        # Build button content with icon + text
        $sp = New-Object System.Windows.Controls.StackPanel
        $sp.Orientation = [System.Windows.Controls.Orientation]::Horizontal

        $img = New-Object System.Windows.Controls.Image
        $img.Source = $dbDrawingImage
        $img.Width = 20
        $img.Height = 20
        $img.Margin = New-Object System.Windows.Thickness(0, 0, 12, 0)
        $sp.Children.Add($img)

        $txt = New-Object System.Windows.Controls.TextBlock
        $txt.Text = $db
        $txt.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $txt.Foreground = [System.Windows.Media.Brushes]::White
        $txt.FontSize = 14
        $txt.FontWeight = [System.Windows.FontWeights]::SemiBold
        $sp.Children.Add($txt)

        $btn.Content = $sp

        $btn.Add_Click({
            param($sender, $e)
            $selectedDb = $sender.Tag

            # Detect version path
            $targetDir = Get-GPConfigTargetPath
            if (-not $targetDir) {
                $statusText.Text = "Error: Could not find GP config path. Run GhostPractice manually once first."
                $statusText.Foreground = [System.Windows.Media.Brushes]::Red
                return
            }

            # Check GP exe exists
            if (-not (Test-Path $Config.gpPath)) {
                $statusText.Text = "Error: GhostPractice.exe not found at configured path."
                $statusText.Foreground = [System.Windows.Media.Brushes]::Red
                return
            }

            # Generate and write config
            $version = Split-Path $targetDir -Leaf
            $xml = Build-UserConfig $Config.server $selectedDb $version
            $configFile = Join-Path $targetDir "user.config"

            try {
                $xml | Set-Content $configFile -Encoding UTF8 -Force
            }
            catch {
                $statusText.Text = "Error writing config: $_"
                $statusText.Foreground = [System.Windows.Media.Brushes]::Red
                return
            }

            # Launch GP
            try {
                Start-Process $Config.gpPath
            }
            catch {
                $statusText.Text = "Error launching GhostPractice: $_"
                $statusText.Foreground = [System.Windows.Media.Brushes]::Red
                return
            }

            # Track usage
            if (-not $Config.usage.PSObject.Properties[$selectedDb]) {
                $Config.usage | Add-Member -NotePropertyName $selectedDb -NotePropertyValue 0
            }
            $Config.usage.$selectedDb = [int]$Config.usage.$selectedDb + 1
            Save-Config $Config

            # Show confirmation
            $statusText.Text = "Launched $selectedDb"
            $statusText.Foreground = [System.Windows.Media.Brushes]::LightGreen
        })

        $dbPanel.Children.Add($btn)
    }
}

Populate-DbButtons

# --- Settings button handler ---
$settingsBtn.Add_Click({
    $result = Show-SettingsWindow $window $Config
    if ($result) {
        $script:Config = Load-Config
        $serverLabel.Text = "Server: $($Config.server)"
        Populate-DbButtons
    }
})

# --- Show the window ---
$window.ShowDialog() | Out-Null
