# Requires -Version 5.1
param(
    [switch]$RunHidden
)

$isActualAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# --- Auto-Elevate and Self-Relaunch to Hide Console ---
if (-not $RunHidden) {
    $psArgs = @("-ExecutionPolicy", "Bypass", "-File", $PSCommandPath, "-RunHidden")
    if (-not $isActualAdmin) {
        try {
            # Ask for Admin permission
            Start-Process powershell.exe -Verb RunAs -WindowStyle Hidden -ArgumentList $psArgs
            exit
        } catch {
            # User clicked 'No' on the UAC prompt. Fall back to standard user mode.
            Start-Process powershell.exe -WindowStyle Hidden -ArgumentList $psArgs
            exit
        }
    } else {
        # Already Admin, just hide the console window
        Start-Process powershell.exe -WindowStyle Hidden -ArgumentList $psArgs
        exit
    }
}
# ------------------------------------------

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Drawing

# 1. Define the UI layout using XAML
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Mini WingetUI" Height="600" Width="850" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <!-- Dark Theme Dictionary -->
        <SolidColorBrush x:Key="WindowBackground" Color="#1E1E1E"/>
        <SolidColorBrush x:Key="TextForeground" Color="#D4D4D4"/>
        <SolidColorBrush x:Key="ControlBackground" Color="#2D2D30"/>
        <SolidColorBrush x:Key="ControlBorder" Color="#3F3F46"/>
        <SolidColorBrush x:Key="AccentColor" Color="#007ACC"/>
        
        <Style TargetType="Window">
            <Setter Property="Background" Value="{StaticResource WindowBackground}"/>
            <Setter Property="Foreground" Value="{StaticResource TextForeground}"/>
        </Style>
        
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="{StaticResource WindowBackground}"/>
            <Setter Property="BorderBrush" Value="{StaticResource ControlBorder}"/>
        </Style>
        
        <Style TargetType="TabItem">
            <Setter Property="Background" Value="{StaticResource ControlBackground}"/>
            <Setter Property="Foreground" Value="{StaticResource TextForeground}"/>
            <Setter Property="BorderBrush" Value="{StaticResource ControlBorder}"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border Name="TabBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1,1,1,0" Margin="0,0,2,0">
                            <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="{StaticResource WindowBackground}"/>
                                <Setter Property="Foreground" Value="White"/>
                                <Setter Property="FontWeight" Value="SemiBold"/>
                            </Trigger>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsMouseOver" Value="True"/>
                                    <Condition Property="IsSelected" Value="False"/>
                                </MultiTrigger.Conditions>
                                <Setter TargetName="TabBorder" Property="Background" Value="{StaticResource ControlBorder}"/>
                            </MultiTrigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="Button">
            <Setter Property="Background" Value="{StaticResource ControlBackground}"/>
            <Setter Property="Foreground" Value="{StaticResource TextForeground}"/>
            <Setter Property="BorderBrush" Value="{StaticResource ControlBorder}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="3">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="{StaticResource ControlBorder}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="{StaticResource AccentColor}"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="{StaticResource ControlBackground}"/>
            <Setter Property="Foreground" Value="{StaticResource TextForeground}"/>
            <Setter Property="BorderBrush" Value="{StaticResource ControlBorder}"/>
            <Setter Property="CaretBrush" Value="{StaticResource TextForeground}"/>
        </Style>

        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="{StaticResource WindowBackground}"/>
            <Setter Property="Foreground" Value="{StaticResource TextForeground}"/>
            <Setter Property="RowBackground" Value="{StaticResource WindowBackground}"/>
            <Setter Property="AlternatingRowBackground" Value="#252526"/>
            <Setter Property="BorderBrush" Value="{StaticResource ControlBorder}"/>
            <Setter Property="HorizontalGridLinesBrush" Value="{StaticResource ControlBorder}"/>
            <Setter Property="VerticalGridLinesBrush" Value="{StaticResource ControlBorder}"/>
        </Style>

        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="{StaticResource ControlBackground}"/>
            <Setter Property="Foreground" Value="{StaticResource TextForeground}"/>
            <Setter Property="BorderBrush" Value="{StaticResource ControlBorder}"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
            <Setter Property="Padding" Value="5"/>
        </Style>

        <Style TargetType="DataGridRow">
            <Setter Property="Background" Value="Transparent"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{StaticResource AccentColor}"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style TargetType="ContextMenu">
            <Setter Property="Background" Value="{StaticResource ControlBackground}"/>
            <Setter Property="Foreground" Value="{StaticResource TextForeground}"/>
            <Setter Property="BorderBrush" Value="{StaticResource ControlBorder}"/>
        </Style>

        <Style TargetType="MenuItem">
            <Setter Property="Foreground" Value="{StaticResource TextForeground}"/>
        </Style>

        <Style TargetType="Expander">
            <Setter Property="Foreground" Value="{StaticResource TextForeground}"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>

        <Style TargetType="StatusBar">
            <Setter Property="Background" Value="{StaticResource AccentColor}"/>
            <Setter Property="Foreground" Value="White"/>
        </Style>
    </Window.Resources>
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TabControl Name="MainTabs" Grid.Row="0" Margin="0,0,0,10">
            <!-- Discover Tab -->
            <TabItem Header="Discover / Search">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="3*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="2*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <!-- Row 0: Search Area -->
                    <StackPanel Orientation="Horizontal" Grid.Row="0" Margin="0,0,0,10">
                        <TextBox Name="SearchBox" Width="400" Margin="0,0,10,0" Padding="5" VerticalContentAlignment="Center"/>
                        <Button Name="SearchBtn" Content="Search Winget" Width="120" Padding="5"/>
                    </StackPanel>
                    
                    <!-- Row 1: Search Results Grid -->
                    <DataGrid Name="DiscoverGrid" Grid.Row="1" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Extended">
                        <DataGrid.ContextMenu>
                            <ContextMenu>
                                <MenuItem Name="DiscoverMenuDetails" Header="Show App Details..." />
                            </ContextMenu>
                        </DataGrid.ContextMenu>
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2*"/>
                            <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="1.5*"/>
                            <DataGridTextColumn Header="Version" Binding="{Binding Version}" Width="*"/>
                            <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    
                    <!-- Row 2: Queue Controls -->
                    <StackPanel Orientation="Horizontal" Grid.Row="2" Margin="0,10,0,10" HorizontalAlignment="Right">
                        <Button Name="AddToQueueBtn" Content="Add Selected to Install Queue &#x2193;" Padding="8" Width="250" FontWeight="Bold"/>
                    </StackPanel>
                    
                    <!-- Row 3: Queue Grid -->
                    <DataGrid Name="QueueGrid" Grid.Row="3" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Extended">
                        <DataGrid.ContextMenu>
                            <ContextMenu>
                                <MenuItem Name="QueueMenuDetails" Header="Show App Details..." />
                            </ContextMenu>
                        </DataGrid.ContextMenu>
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Queued App Name" Binding="{Binding Name}" Width="2*"/>
                            <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="1.5*"/>
                            <DataGridTextColumn Header="Target Version" Binding="{Binding Version}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    
                    <!-- Row 4: Final Actions -->
                    <StackPanel Orientation="Horizontal" Grid.Row="4" Margin="0,10,0,0" HorizontalAlignment="Right">
                        <CheckBox Name="AdminInstallCheck" Content="Install for All Users (Admin)" IsChecked="True" VerticalAlignment="Center" Margin="0,0,15,0" Foreground="{StaticResource TextForeground}"/>
                        <Button Name="ImportQueueBtn" Content="Import Queue" Padding="8" Width="110" Margin="0,0,10,0"/>
                        <Button Name="ExportQueueBtn" Content="Export Queue" Padding="8" Width="110" Margin="0,0,10,0"/>
                        <Button Name="RemoveFromQueueBtn" Content="Remove from Queue" Padding="8" Width="130" Margin="0,0,10,0"/>
                        <Button Name="InstallBtn" Content="Install Queued Packages" Padding="8" Width="180" FontWeight="Bold" Foreground="#73D96B"/>
                    </StackPanel>
                </Grid>
            </TabItem>
            
            <!-- Combined Installed & Updates Tab -->
            <TabItem Header="Installed &amp; Updates">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <StackPanel Orientation="Horizontal" Grid.Row="0" Margin="0,0,0,10">
                        <Button Name="RefreshInstalledBtn" Content="Load Installed &amp; Check Updates" Padding="5" Width="220" Margin="0,0,10,0"/>
                    </StackPanel>
                    
                    <TabControl Name="InstalledTabs" Grid.Row="1" Margin="0,5,0,0">
                        <TabItem Header="Desktop / External Apps">
                            <DataGrid Name="InstalledGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Extended" BorderThickness="0">
                                <DataGrid.RowStyle>
                                    <Style TargetType="DataGridRow" BasedOn="{StaticResource {x:Type DataGridRow}}">
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding HasUpdate}" Value="True">
                                                <Setter Property="Background" Value="#2A4032"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </DataGrid.RowStyle>
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2*"/>
                                    <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="1.5*"/>
                                    <DataGridTextColumn Header="Current Version" Binding="{Binding Version}" Width="*"/>
                                    <DataGridTextColumn Header="Available Update" Binding="{Binding Extra}" Width="*"/>
                                    <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="*"/>
                                </DataGrid.Columns>
                                <DataGrid.ContextMenu>
                                    <ContextMenu>
                                        <MenuItem Name="InstalledMenuDetails" Header="Show App Details..." />
                                    </ContextMenu>
                                </DataGrid.ContextMenu>
                            </DataGrid>
                        </TabItem>
                        <TabItem Header="Windows / Store Apps">
                            <DataGrid Name="WindowsAppsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Extended" BorderThickness="0">
                                <DataGrid.RowStyle>
                                    <Style TargetType="DataGridRow" BasedOn="{StaticResource {x:Type DataGridRow}}">
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding HasUpdate}" Value="True">
                                                <Setter Property="Background" Value="#2A4032"/>
                                            </DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </DataGrid.RowStyle>
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2*"/>
                                    <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="1.5*"/>
                                    <DataGridTextColumn Header="Current Version" Binding="{Binding Version}" Width="*"/>
                                    <DataGridTextColumn Header="Available Update" Binding="{Binding Extra}" Width="*"/>
                                    <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="*"/>
                                </DataGrid.Columns>
                                <DataGrid.ContextMenu>
                                    <ContextMenu>
                                        <MenuItem Name="WindowsAppsMenuDetails" Header="Show App Details..." />
                                    </ContextMenu>
                                </DataGrid.ContextMenu>
                            </DataGrid>
                        </TabItem>
                    </TabControl>

                    <StackPanel Orientation="Horizontal" Grid.Row="2" Margin="0,10,0,0" HorizontalAlignment="Right">
                        <Button Name="UninstallBtn" Content="Uninstall Selected" Padding="8" Width="140" Foreground="#FF6B6B" FontWeight="Bold" Margin="0,0,10,0"/>
                        <Button Name="UpdateBtn" Content="Update Selected" Padding="8" Width="140" Foreground="#6BA4FF" FontWeight="Bold" Margin="0,0,10,0"/>
                        <Button Name="UpdateAllBtn" Content="Update All Apps" Padding="8" Width="140" Foreground="#73D96B" FontWeight="Bold"/>
                    </StackPanel>
                </Grid>
            </TabItem>
        </TabControl>
        
        <!-- Live Console Log Expander -->
        <Expander Name="LogExpander" Grid.Row="1" Header="Live Console Log" Margin="0,0,0,5">
            <TextBox Name="LogTextBox" Height="140" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="12" Background="#0C0C0C" Foreground="#CCCCCC" BorderBrush="{StaticResource ControlBorder}" TextWrapping="NoWrap" Margin="0,5,0,0"/>
        </Expander>

        <StatusBar Grid.Row="2">
            <StatusBarItem>
                <ProgressBar Name="JobProgress" Width="150" Height="15" IsIndeterminate="False" Visibility="Hidden" Margin="0,0,10,0"/>
            </StatusBarItem>
            <StatusBarItem>
                <Button Name="StopJobBtn" Content="Stop" Padding="15,2" Background="#FF4444" Foreground="White" FontWeight="Bold" Visibility="Hidden" Margin="0,0,10,0"/>
            </StatusBarItem>
            <StatusBarItem>
                <TextBlock Name="StatusText" Text="Ready." FontWeight="SemiBold" Foreground="White"/>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
"@

# 2. Load the XAML into PowerShell
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Map XAML elements to PowerShell variables
$SearchBox = $Window.FindName("SearchBox")
$SearchBtn = $Window.FindName("SearchBtn")
$DiscoverGrid = $Window.FindName("DiscoverGrid")
$DiscoverMenuDetails = $Window.FindName("DiscoverMenuDetails")

$AddToQueueBtn = $Window.FindName("AddToQueueBtn")
$QueueGrid = $Window.FindName("QueueGrid")
$QueueMenuDetails = $Window.FindName("QueueMenuDetails")
$RemoveFromQueueBtn = $Window.FindName("RemoveFromQueueBtn")
$ImportQueueBtn = $Window.FindName("ImportQueueBtn")
$ExportQueueBtn = $Window.FindName("ExportQueueBtn")
$AdminInstallCheck = $Window.FindName("AdminInstallCheck")
$InstallBtn = $Window.FindName("InstallBtn")

# --- Apply Admin Status to UI ---
if ($isActualAdmin) {
    $Window.Title = "Mini WingetUI (Administrator)"
    $AdminInstallCheck.IsChecked = $true
    $AdminInstallCheck.IsEnabled = $true
} else {
    $Window.Title = "Mini WingetUI (Standard User)"
    $AdminInstallCheck.IsChecked = $false
    $AdminInstallCheck.IsEnabled = $false
    $AdminInstallCheck.Content = "Install for Current User (No Admin)"
    $AdminInstallCheck.ToolTip = "You declined the Administrator prompt. Installations are limited to the current user."
    $AdminInstallCheck.Foreground = "#888888"
}

$RefreshInstalledBtn = $Window.FindName("RefreshInstalledBtn")
$UninstallBtn = $Window.FindName("UninstallBtn")
$UpdateBtn = $Window.FindName("UpdateBtn")
$UpdateAllBtn = $Window.FindName("UpdateAllBtn")
$InstalledGrid = $Window.FindName("InstalledGrid")
$InstalledMenuDetails = $Window.FindName("InstalledMenuDetails")

$WindowsAppsGrid = $Window.FindName("WindowsAppsGrid")
$WindowsAppsMenuDetails = $Window.FindName("WindowsAppsMenuDetails")

$LogExpander = $Window.FindName("LogExpander")
$LogTextBox = $Window.FindName("LogTextBox")

$JobProgress = $Window.FindName("JobProgress")
$StopJobBtn = $Window.FindName("StopJobBtn")
$StatusText = $Window.FindName("StatusText")

# --- Initialize the Observable Queue ---
# This allows the UI to automatically update when we add/remove items from the queue
$script:InstallQueue = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$QueueGrid.ItemsSource = $script:InstallQueue

# 3. Setup the Background Runspace (The Backend Engine)
$syncHash = [hashtable]::Synchronized(@{})
$syncHash.LogQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$runspace = [runspacefactory]::CreateRunspace()
$runspace.Open()
$runspace.SessionStateProxy.SetVariable("syncHash", $syncHash)

$script:psInstance = $null
$script:asyncResult = $null
$script:IsJobRunning = $false
$script:AllInstalledApps = $null

# 4. Define the Universal Background Job
$bgJobBlock = {
    param($Action, $Query, $Id, $Hash, $IsAdmin)
    
    # Force PowerShell to read external Winget output using UTF-8 to prevent 'ΓÇª' encoding issues
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
    
    # Internal parser for tabular winget output
    function ConvertFrom-WingetOutput($raw) {
        $parsed = @()
        
        # 1. Find the header line and dashes line
        $headerIdx = -1
        for ($i = 0; $i -lt $raw.Count; $i++) {
            if ($raw[$i] -match "^---+") {
                $headerIdx = $i - 1
                break
            }
        }
        
        if ($headerIdx -lt 0) { return $parsed }
        
        $headerLine = $raw[$headerIdx]
        
        # 2. Find column start indices using word boundaries (\S+ matches any contiguous word)
        # This completely solves the "missing column" bug, regardless of Winget's spacing.
        $colIndices = @()
        $headerMatches = [regex]::Matches($headerLine, '\S+')
        foreach ($m in $headerMatches) {
            $colIndices += $m.Index
        }
        
        if ($colIndices.Count -eq 0) { return $parsed }
        
        # 3. Parse data lines using strict fixed-width indexing
        for ($i = $headerIdx + 2; $i -lt $raw.Count; $i++) {
            $line = $raw[$i]
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if ($line -match "^[0-9]+ upgrades available") { continue }
            
            $cols = @()
            for ($c = 0; $c -lt $colIndices.Count; $c++) {
                $start = $colIndices[$c]
                if ($start -ge $line.Length) {
                    $cols += ""
                } else {
                    $len = if ($c -eq $colIndices.Count - 1) { $line.Length - $start } else { $colIndices[$c+1] - $start }
                    if ($start + $len -gt $line.Length) { $len = $line.Length - $start }
                    
                    $val = $line.Substring($start, $len).Trim()
                    
                    # Strip the ugly ellipsis and OEM encoding artifacts safely using Regex Unicode Hex
                    $val = $val -replace '(\u0393\u00C7\u00AA)+$', '' # Removes ΓÇª
                    $val = $val -replace '(\u0393\u00C7\u00F6)+$', '' # Removes ΓÇö
                    $val = $val -replace '\u2026+$', ''             # Removes …
                    
                    $cols += $val.Trim()
                }
            }
            
            if ($cols.Count -ge 3) {
                $name = $cols[0]
                $id = $cols[1]
                $version = $cols[2]
                $extra = ""
                $source = ""
                $hasUpdate = $false
                
                # Assign Extra (Available) and Source dynamically based on the parsed headers
                if ($cols.Count -ge 5) {
                    $extra = $cols[3]
                    $source = $cols[4]
                    if (-not [string]::IsNullOrWhiteSpace($extra) -and $extra -ne "Unknown" -and $extra -notmatch "^<") {
                        $hasUpdate = $true
                    }
                } elseif ($cols.Count -eq 4) {
                    $source = $cols[3]
                }

                $parsed += [PSCustomObject]@{
                    Name      = $name
                    Id        = $id
                    Version   = $version
                    Extra     = $extra
                    Source    = $source
                    HasUpdate = $hasUpdate
                }
            }
        }
        return $parsed
    }

    try {
        switch ($Action) {
            'Search' {
                $sysLocale = (Get-Culture).Name                      
                $langOnly = (Get-Culture).TwoLetterISOLanguageName   
                
                $raw = @()
                $Hash.LogQueue.Enqueue(">>> Executing: winget search ""$Query""")
                $wingetArgs = @("search", $Query, "--count", "40", "--accept-source-agreements")
                & winget @wingetArgs 2>&1 | ForEach-Object {
                    $line = $_.ToString()
                    $Hash.LogQueue.Enqueue($line)
                    $raw += $line
                }
                
                $parsed = ConvertFrom-WingetOutput $raw
                
                $filtered = @()
                foreach ($item in $parsed) {
                    # Detect if the ID ends with a common locale suffix (like .zh-CN, .fr, .ja-JP)
                    $hasLocaleSuffix = $item.Id -match '\.([a-z]{2}-[A-Z]{2}|[a-z]{2})$'
                    
                    if (-not $hasLocaleSuffix) {
                        # 1. Keep standard/neutral packages that don't have a locale tag
                        $filtered += $item
                    } elseif ($item.Id -match "\.($sysLocale|$langOnly|en-US|en-GB|en)$") {
                        # 2. Keep packages that match your local OS language or English
                        $filtered += $item
                    }
                    # 3. Everything else (foreign languages) is quietly discarded to keep the UI clean
                }
                
                $Hash.Result = $filtered
            }
            'Installed' {
                $raw = @()
                $Hash.LogQueue.Enqueue(">>> Executing: winget list")
                & winget list --accept-source-agreements 2>&1 | ForEach-Object {
                    $line = $_.ToString()
                    $Hash.LogQueue.Enqueue($line)
                    $raw += $line
                }
                $Hash.Result = ConvertFrom-WingetOutput $raw
            }
            'Install' {
                foreach ($targetId in @($Id)) {
                    $Hash.LogQueue.Enqueue("`r`n>>> Executing: winget install $targetId")
                    $wingetArgs = @("install", "--id", $targetId, "--exact", "--accept-source-agreements", "--accept-package-agreements", "--silent", "--disable-interactivity")
                    if ($IsAdmin) { $wingetArgs += "--scope"; $wingetArgs += "machine" }
                    
                    & winget @wingetArgs 2>&1 | ForEach-Object {
                        $l = $_.ToString().Trim()
                        # Ignore Winget's 1-character ASCII animation frames
                        if ($l -in @('\', '|', '/', '-')) { return }
                        
                        # Intercept the Percentage for the UI Progress Bar
                        if ($l -match '(?<pct>\d{1,3})\s*%') {
                            $Hash.Progress = [int]$matches['pct']
                            return # Hide this line from the text log to keep it clean
                        }
                        
                        # Also filter block progress bars if they appear without percentages
                        if ($l -match '█') { return }
                        if (-not [string]::IsNullOrWhiteSpace($l)) { $Hash.LogQueue.Enqueue($l) }
                    }
                }
                $Hash.Result = "Success"
            }
            'Uninstall' {
                foreach ($targetId in @($Id)) {
                    $Hash.LogQueue.Enqueue("`r`n>>> Executing: winget uninstall $targetId")
                    $wingetArgs = @("uninstall", "--id", $targetId, "--exact", "--silent", "--disable-interactivity")
                    & winget @wingetArgs 2>&1 | ForEach-Object {
                        $l = $_.ToString().Trim()
                        if ($l -in @('\', '|', '/', '-')) { return }
                        if ($l -match '(?<pct>\d{1,3})\s*%') {
                            $Hash.Progress = [int]$matches['pct']
                            return
                        }
                        if ($l -match '█') { return }
                        if (-not [string]::IsNullOrWhiteSpace($l)) { $Hash.LogQueue.Enqueue($l) }
                    }
                }
                $Hash.Result = "Success"
            }
            'Update' {
                foreach ($targetId in @($Id)) {
                    $Hash.LogQueue.Enqueue("`r`n>>> Executing: winget upgrade $targetId")
                    $wingetArgs = @("upgrade", "--id", $targetId, "--exact", "--accept-source-agreements", "--accept-package-agreements", "--silent", "--disable-interactivity")
                    & winget @wingetArgs 2>&1 | ForEach-Object {
                        $l = $_.ToString().Trim()
                        if ($l -in @('\', '|', '/', '-')) { return }
                        if ($l -match '(?<pct>\d{1,3})\s*%') {
                            $Hash.Progress = [int]$matches['pct']
                            return
                        }
                        if ($l -match '█') { return }
                        if (-not [string]::IsNullOrWhiteSpace($l)) { $Hash.LogQueue.Enqueue($l) }
                    }
                }
                $Hash.Result = "Success"
            }
            'ShowDetails' {
                $raw = @()
                $Hash.LogQueue.Enqueue(">>> Executing: winget show $Id")
                $wingetArgs = @("show", "--id", $Id, "--exact", "--accept-source-agreements")
                & winget @wingetArgs 2>&1 | ForEach-Object {
                    $line = $_.ToString()
                    $Hash.LogQueue.Enqueue($line)
                    $raw += $line
                }
                $Hash.Result = $raw -join "`r`n"
            }
        }
    } catch {
        $Hash.Result = "Error: $($_.Exception.Message)"
    }
}

# 5. Helper Function to safely dispatch jobs to the Runspace
function Start-WingetJob($Action, $Query, $Id, $StatusMsg, $IsAdmin = $false) {
    # Check if a job is already running and warn the user
    if ($script:IsJobRunning) {
        [System.Windows.MessageBox]::Show("A Winget task is currently running in the background.`n`nPlease wait for it to finish or click 'Stop' before starting a new action.", "Task in Progress", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    $script:IsJobRunning = $true
    
    # Clear the live log UI and Queue
    $LogTextBox.Clear()
    $dummy = [string]::Empty
    while ($syncHash.LogQueue.TryDequeue([ref]$dummy)) {}

    # Auto-expand the log panel if we are making system changes
    if ($Action -in @('Install', 'Uninstall', 'Update')) {
        $LogExpander.IsExpanded = $true
    }
    
    # Activate Progress Bar and Stop Button
    $JobProgress.Visibility = 'Visible'
    $JobProgress.IsIndeterminate = $true
    $JobProgress.Value = 0
    $StopJobBtn.Visibility = 'Visible'
    $StopJobBtn.IsEnabled = $true
    
    $StatusText.Text = $StatusMsg
    $syncHash.Action = $Action
    $syncHash.Result = $null
    $syncHash.Progress = $null
    $syncHash.StatusMsg = $StatusMsg # Store the base message so we can append % to it later

    $script:psInstance = [PowerShell]::Create().AddScript($bgJobBlock).AddArgument($Action).AddArgument($Query).AddArgument($Id).AddArgument($syncHash).AddArgument($IsAdmin)
    $script:psInstance.Runspace = $runspace
    $script:asyncResult = $script:psInstance.BeginInvoke()
    $timer.Start()
}

# --- Heuristic Leftover Scanner Function ---
function Get-SafeLeftoverPaths {
    param($AppId, $AppName)
    
    $terms = @()
    $parts = $AppId -split '\.'
    
    if ($parts.Count -ge 2) {
        $terms += "$($parts[0])\$($parts[1])" # e.g. Mozilla\Firefox
        $terms += $parts[1] # e.g. Firefox
    } else {
        $terms += $AppId
    }
    
    # Extract the first distinct word of the App Name
    $nameWord = ($AppName -split '\s+') | Where-Object { $_.Length -gt 3 } | Select-Object -First 1
    if ($nameWord) { $terms += $nameWord }
    
    # CRITICAL: Strict blacklist to prevent deleting OS/Shared components
    $unsafe = @('microsoft','windows','intel','amd','nvidia','system','software','common','program','google','apple','adobe','oracle','java','video','music','documents','desktop','downloads','users','admin','local','roaming','temp')
    $validTerms = $terms | Where-Object { $_.Length -gt 3 -and $_.ToLower() -notin $unsafe } | Select-Object -Unique
    
    $leftovers = @()
    $baseDirs = @($env:LOCALAPPDATA, $env:APPDATA, $env:ProgramData, $env:ProgramFiles, ${env:ProgramFiles(x86)})
    $regBases = @("HKCU:\SOFTWARE", "HKLM:\SOFTWARE")
    
    foreach ($term in $validTerms) {
        foreach ($dir in $baseDirs) {
            $p = Join-Path $dir $term
            if (Test-Path $p -ErrorAction SilentlyContinue) {
                # Protect root directories from accidental matches
                if ($p.Length -gt ($dir.Length + 2)) {
                    $leftovers += [PSCustomObject]@{ Selected=$true; Type="Folder"; Path=$p }
                }
            }
        }
        foreach ($reg in $regBases) {
            $p = Join-Path $reg $term
            if (Test-Path $p -ErrorAction SilentlyContinue) {
                # Protect base registry nodes
                if ($p.Length -gt ($reg.Length + 2)) {
                    $leftovers += [PSCustomObject]@{ Selected=$true; Type="Registry"; Path=$p }
                }
            }
        }
    }
    # Deduplicate matches
    return $leftovers | Group-Object Path | ForEach-Object { $_.Group[0] }
}

# 6. Setup the UI Timer to check job status
$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromMilliseconds(200)

$timer.Add_Tick({
    # Drain the live log queue and append to the TextBox
    $newLogs = [System.Text.StringBuilder]::new()
    $logLine = [string]::Empty
    while ($syncHash.LogQueue.TryDequeue([ref]$logLine)) {
        [void]$newLogs.AppendLine($logLine)
    }
    if ($newLogs.Length -gt 0) {
        $LogTextBox.AppendText($newLogs.ToString())
        $LogTextBox.ScrollToEnd()
    }
    
    # Check if a live percentage was reported back from Winget
    if ($syncHash.Progress -ne $null) {
        if ($JobProgress.IsIndeterminate) {
            # Switch from spinning mode to solid fill mode
            $JobProgress.IsIndeterminate = $false
        }
        $JobProgress.Value = $syncHash.Progress
        $StatusText.Text = "$($syncHash.StatusMsg) - $($syncHash.Progress)%"
    }

    if ($script:asyncResult -ne $null -and $script:asyncResult.IsCompleted) {
        $timer.Stop()
        
        try {
            $script:psInstance.EndInvoke($script:asyncResult)
        } catch {
            # This triggers if the user clicks the "Stop" button and aborts the pipeline
            $syncHash.Result = "Error: Operation was cancelled by the user."
        }
        
        $script:psInstance.Dispose()
        
        # Mark the job as finished EARLY so chained jobs (like Refreshing) can start successfully
        $script:IsJobRunning = $false
        
        # Disable Progress Bar and Stop button
        $JobProgress.IsIndeterminate = $false
        $JobProgress.Visibility = 'Hidden'
        $StopJobBtn.Visibility = 'Hidden'
        
        $res = $syncHash.Result
        $action = $syncHash.Action
        
        if ($res -is [string] -and $res -match "^Error") {
            $StatusText.Text = "Operation failed: $res"
        } else {
            # Route the data to the correct Grid based on what action just finished
            switch ($action) {
                'Search' {
                    $DiscoverGrid.ItemsSource = $res
                    $StatusText.Text = "Search complete. Found $($res.Count) packages."
                }
                'Installed' {
                    # Sort so packages with updates are on top, then sort alphabetically
                    $script:AllInstalledApps = @($res | Sort-Object -Property @{Expression="HasUpdate"; Descending=$true}, Name)
                    
                    # Split into two lists based on standard Windows Store / Appx / MSIX heuristics
                    $desktopApps = @($script:AllInstalledApps | Where-Object { $_.Source -ne 'msstore' -and $_.Id -notmatch '_[a-zA-Z0-9]{13}$' -and $_.Id -notmatch '^MSIX\\' })
                    $windowsApps = @($script:AllInstalledApps | Where-Object { $_.Source -eq 'msstore' -or $_.Id -match '_[a-zA-Z0-9]{13}$' -or $_.Id -match '^MSIX\\' })
                    
                    $InstalledGrid.ItemsSource = $desktopApps
                    $WindowsAppsGrid.ItemsSource = $windowsApps
                    
                    # Count how many packages have updates
                    $updateCount = @($res | Where-Object { $_.HasUpdate -eq $true }).Count
                    $StatusText.Text = "Loaded $($desktopApps.Count) Desktop Apps and $($windowsApps.Count) Windows Apps. Found $updateCount available updates."
                }
                'Install' { 
                    $StatusText.Text = "Installation finished successfully." 
                    $script:InstallQueue.Clear() # Empty the queue when finished
                }
                'Uninstall' { 
                    $StatusText.Text = "Uninstallation finished. Scanning for leftovers..."
                    
                    # Run the Heuristic Leftover Scanner
                    $allLeftovers = @()
                    if ($syncHash.TargetApps) {
                        foreach ($app in $syncHash.TargetApps) {
                            $leftovers = Get-SafeLeftoverPaths -AppId $app.Id -AppName $app.Name
                            $allLeftovers += $leftovers
                        }
                    }
                    
                    if ($allLeftovers.Count -gt 0) {
                        # Display the Custom Cleanup UI
                        $leftoverXaml = @"
                        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Cleanup Leftovers" Width="650" Height="400" Background="#1E1E1E" Foreground="#D4D4D4" WindowStartupLocation="CenterScreen">
                            <Window.Resources>
                                <Style TargetType="DataGridColumnHeader">
                                    <Setter Property="Background" Value="#2D2D30"/>
                                    <Setter Property="Foreground" Value="#D4D4D4"/>
                                    <Setter Property="BorderBrush" Value="#3F3F46"/>
                                    <Setter Property="BorderThickness" Value="0,0,1,1"/>
                                    <Setter Property="Padding" Value="5,8"/>
                                    <Setter Property="FontWeight" Value="SemiBold"/>
                                </Style>
                            </Window.Resources>
                            <Grid Margin="15">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <TextBlock Text="The following leftover folders and registry keys were found. Select items to permanently remove:" FontSize="14" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,0,0,10"/>
                                <DataGrid Name="LeftoversGrid" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" Background="#2D2D30" Foreground="#D4D4D4" BorderBrush="#3F3F46" HeadersVisibility="Column" RowBackground="#2D2D30" AlternatingRowBackground="#252526" GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#3F3F46">
                                    <DataGrid.Columns>
                                        <DataGridCheckBoxColumn Header="Remove" Binding="{Binding Selected, UpdateSourceTrigger=PropertyChanged}" Width="70">
                                            <DataGridCheckBoxColumn.ElementStyle>
                                                <Style TargetType="CheckBox"><Setter Property="HorizontalAlignment" Value="Center"/><Setter Property="VerticalAlignment" Value="Center"/></Style>
                                            </DataGridCheckBoxColumn.ElementStyle>
                                        </DataGridCheckBoxColumn>
                                        <DataGridTextColumn Header="Type" Binding="{Binding Type}" IsReadOnly="True" Width="80"/>
                                        <DataGridTextColumn Header="Path" Binding="{Binding Path}" IsReadOnly="True" Width="*"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                                <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,15,0,0">
                                    <Button Name="SkipBtn" Content="Skip" Width="80" Padding="8" Margin="0,0,10,0" Background="#3F3F46" Foreground="White" BorderThickness="0" Cursor="Hand"/>
                                    <Button Name="DeleteBtn" Content="Delete Selected" Width="140" Padding="8" Background="#FF4444" Foreground="White" BorderThickness="0" FontWeight="Bold" Cursor="Hand"/>
                                </StackPanel>
                            </Grid>
                        </Window>
"@
                        $lReader = (New-Object System.Xml.XmlNodeReader ([xml]$leftoverXaml))
                        $lWindow = [Windows.Markup.XamlReader]::Load($lReader)
                        
                        $LeftoversGrid = $lWindow.FindName("LeftoversGrid")
                        $SkipBtn = $lWindow.FindName("SkipBtn")
                        $DeleteBtn = $lWindow.FindName("DeleteBtn")
                        
                        # Populate DataGrid
                        $obsLeftovers = New-Object System.Collections.ObjectModel.ObservableCollection[object]
                        foreach ($l in $allLeftovers) { $obsLeftovers.Add($l) }
                        $LeftoversGrid.ItemsSource = $obsLeftovers
                        
                        $SkipBtn.Add_Click({ $lWindow.Close() })
                        $DeleteBtn.Add_Click({
                            foreach ($item in $obsLeftovers) {
                                if ($item.Selected) {
                                    try {
                                        Remove-Item -Path $item.Path -Recurse -Force -ErrorAction SilentlyContinue
                                    } catch { }
                                }
                            }
                            $lWindow.Close()
                        })
                        
                        $lWindow.ShowDialog() | Out-Null
                    }
                    
                    $syncHash.TargetApps = $null
                    Start-WingetJob -Action "Installed" -Query "" -Id "" -StatusMsg "Refreshing installed packages..."
                    return # Skip re-enabling UI so the refresh can begin immediately
                }
                'Update' { 
                    $StatusText.Text = "Update finished."
                    Start-WingetJob -Action "Installed" -Query "" -Id "" -StatusMsg "Refreshing installed packages..."
                    return 
                }
                'ShowDetails' {
                    $StatusText.Text = "App details loaded."
                    # Display the details in a new lightweight popup window
                    $detailXaml = @"
                    <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                            Title="App Details" Height="450" Width="650" WindowStartupLocation="CenterScreen" Background="#1E1E1E">
                        <TextBox Text="{Binding Mode=OneWay}" IsReadOnly="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="13" Margin="10" Background="#2D2D30" Foreground="#D4D4D4" BorderBrush="#3F3F46"/>
                    </Window>
"@
                    $detailReader = (New-Object System.Xml.XmlNodeReader ([xml]$detailXaml))
                    $detailWindow = [Windows.Markup.XamlReader]::Load($detailReader)
                    $detailWindow.DataContext = $res
                    $detailWindow.ShowDialog() | Out-Null
                }
            }
        }
    }
})

# 7. Map Button Clicks
$StopJobBtn.Add_Click({
    if ($script:psInstance -ne $null -and $script:asyncResult.IsCompleted -eq $false) {
        $StatusText.Text = "Stopping operation... (Killing Winget)"
        $StopJobBtn.IsEnabled = $false
        
        # 1. Stop the PowerShell runspace pipeline immediately
        $script:psInstance.Stop()
        
        # 2. Force kill the underlying winget process to prevent orphaned background downloads/installers
        Get-Process -Name winget -ErrorAction SilentlyContinue | Stop-Process -Force
    }
})

$SearchBtn.Add_Click({
    $query = $SearchBox.Text
    if (![string]::IsNullOrWhiteSpace($query)) {
        $DiscoverGrid.ItemsSource = $null
        Start-WingetJob -Action "Search" -Query $query -Id "" -StatusMsg "Searching Winget for '$query'..."
    }
})

$SearchBox.Add_KeyDown({
    if ($_.Key -eq 'Enter') {
        $query = $SearchBox.Text
        if (![string]::IsNullOrWhiteSpace($query)) {
            $DiscoverGrid.ItemsSource = $null
            Start-WingetJob -Action "Search" -Query $query -Id "" -StatusMsg "Searching Winget for '$query'..."
        }
    }
})

$AddToQueueBtn.Add_Click({
    if ($DiscoverGrid.SelectedItems.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please select packages from the search results first.")
        return
    }
    foreach ($item in $DiscoverGrid.SelectedItems) {
        # Check for duplicates before adding to the queue
        $exists = $false
        foreach ($q in $script:InstallQueue) {
            if ($q.Id -eq $item.Id) { $exists = $true; break }
        }
        if (-not $exists) {
            $script:InstallQueue.Add($item)
        }
    }
})

$DiscoverGrid.Add_MouseDoubleClick({
    if ($DiscoverGrid.SelectedItem -ne $null) {
        $item = $DiscoverGrid.SelectedItem
        # Check for duplicates before adding to the queue
        $exists = $false
        foreach ($q in $script:InstallQueue) {
            if ($q.Id -eq $item.Id) { $exists = $true; break }
        }
        if (-not $exists) {
            $script:InstallQueue.Add($item)
        }
    }
})

$RemoveFromQueueBtn.Add_Click({
    if ($QueueGrid.SelectedItems.Count -eq 0) { return }
    # Copy to array to prevent modifying collection while iterating
    $toRemove = @($QueueGrid.SelectedItems)
    foreach ($item in $toRemove) {
        $script:InstallQueue.Remove($item)
    }
})

$ExportQueueBtn.Add_Click({
    if ($script:InstallQueue.Count -eq 0) {
        [System.Windows.MessageBox]::Show("The install queue is empty.")
        return
    }
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
    $dialog.FileName = "Winget-Install-Queue.json"
    if ($dialog.ShowDialog() -eq $true) {
        $script:InstallQueue | Select-Object Name, Id | ConvertTo-Json | Set-Content $dialog.FileName
        [System.Windows.MessageBox]::Show("Install Queue exported successfully to $($dialog.FileName)")
    }
})

$ImportQueueBtn.Add_Click({
    $dialog = New-Object Microsoft.Win32.OpenFileDialog
    $dialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
    if ($dialog.ShowDialog() -eq $true) {
        try {
            # Force conversion to an array in case there is only 1 item in the JSON
            $imported = @(Get-Content $dialog.FileName | ConvertFrom-Json)
            $addedCount = 0
            foreach ($item in $imported) {
                # Prevent duplicates
                $exists = $false
                foreach ($q in $script:InstallQueue) { if ($q.Id -eq $item.Id) { $exists = $true; break } }
                if (-not $exists) {
                    $script:InstallQueue.Add([PSCustomObject]@{Name=$item.Name; Id=$item.Id; Version="Latest"})
                    $addedCount++
                }
            }
            [System.Windows.MessageBox]::Show("Successfully imported $addedCount new apps into the queue.")
        } catch {
            [System.Windows.MessageBox]::Show("Error reading the JSON file. Ensure it is a valid exported queue.")
        }
    }
})

$InstallBtn.Add_Click({
    if ($script:InstallQueue.Count -gt 0) {
        [string[]]$ids = @($script:InstallQueue | ForEach-Object { $_.Id })
        $msg = ""
        if ($ids.Count -eq 1) { 
            $msg = "Installing $($script:InstallQueue[0].Name)..." 
        } else { 
            $msg = "Installing $($ids.Count) packages... Please wait." 
        }
        
        $isAdmin = $AdminInstallCheck.IsChecked -eq $true
        Start-WingetJob -Action "Install" -Query "" -Id $ids -StatusMsg $msg -IsAdmin $isAdmin
    } else { 
        [System.Windows.MessageBox]::Show("Please add packages to the install queue before clicking Install.") 
    }
})

$RefreshInstalledBtn.Add_Click({
    $InstalledGrid.ItemsSource = $null
    $WindowsAppsGrid.ItemsSource = $null
    $script:AllInstalledApps = $null
    Start-WingetJob -Action "Installed" -Query "" -Id "" -StatusMsg "Loading installed packages and checking for updates... This might take a moment."
})

$UninstallBtn.Add_Click({
    $selected = @($InstalledGrid.SelectedItems) + @($WindowsAppsGrid.SelectedItems)
    if ($selected.Count -gt 0) { 
        [string[]]$ids = @($selected | ForEach-Object { $_.Id })
        
        # Save the selected app objects to the background syncHash so the scanner knows what to look for
        $syncHash.TargetApps = $selected
        
        $msg = ""
        if ($ids.Count -eq 1) { 
            $msg = "Uninstalling $($selected[0].Name)..." 
        } else { 
            $msg = "Uninstalling $($ids.Count) packages... Please wait." 
        }
        Start-WingetJob -Action "Uninstall" -Query "" -Id $ids -StatusMsg $msg
    } else { 
        [System.Windows.MessageBox]::Show("Please select at least one installed package to uninstall.") 
    }
})

$UpdateBtn.Add_Click({
    $selected = @($InstalledGrid.SelectedItems) + @($WindowsAppsGrid.SelectedItems)
    if ($selected.Count -gt 0) { 
        [string[]]$ids = @($selected | ForEach-Object { $_.Id })
        $msg = ""
        if ($ids.Count -eq 1) { 
            $msg = "Updating $($selected[0].Name)..." 
        } else { 
            $msg = "Updating $($ids.Count) packages... Please wait." 
        }
        Start-WingetJob -Action "Update" -Query "" -Id $ids -StatusMsg $msg
    } else { 
        [System.Windows.MessageBox]::Show("Please select at least one package to update.") 
    }
})

$UpdateAllBtn.Add_Click({
    if ($script:AllInstalledApps -eq $null) {
        [System.Windows.MessageBox]::Show("Please load the installed packages first.")
        return
    }
    # Filter the overall items source for packages that have an available update
    [string[]]$ids = @($script:AllInstalledApps | Where-Object { $_.HasUpdate -eq $true } | ForEach-Object { $_.Id })
    
    if ($ids.Count -gt 0) {
        $msg = if ($ids.Count -eq 1) { "Updating 1 package..." } else { "Updating all $($ids.Count) available updates... Please wait." }
        Start-WingetJob -Action "Update" -Query "" -Id $ids -StatusMsg $msg
    } else {
        [System.Windows.MessageBox]::Show("No available updates found.")
    }
})

# --- Context Menu Event Handlers ---
$showDetailsAction = {
    param($GridControl)
    $sel = $GridControl.SelectedItem
    if ($sel -ne $null) {
        Start-WingetJob -Action "ShowDetails" -Query "" -Id $sel.Id -StatusMsg "Loading details for $($sel.Name)..."
    }
}
$DiscoverMenuDetails.Add_Click({ &$showDetailsAction $DiscoverGrid })
$QueueMenuDetails.Add_Click({ &$showDetailsAction $QueueGrid })
$InstalledMenuDetails.Add_Click({ &$showDetailsAction $InstalledGrid })
$WindowsAppsMenuDetails.Add_Click({ &$showDetailsAction $WindowsAppsGrid })

# --- Auto-Load Installed Apps on Startup ---
$Window.Add_Loaded({
    Start-WingetJob -Action "Installed" -Query "" -Id "" -StatusMsg "Loading installed packages and checking for updates... This might take a moment."
})

# 6. Show the Window and clean up when closed
$Window.ShowDialog() | Out-Null

# Cleanup memory once the window is closed
$runspace.Close()
$runspace.Dispose()