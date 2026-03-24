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

# --- Pre-flight Check: Ensure Winget is Installed ---
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    $msgResult = [System.Windows.MessageBox]::Show("Windows Package Manager (Winget) was not found on this system.`n`nWould you like to automatically download and install it now?", "Winget Missing", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    
    if ($msgResult -eq 'Yes') {
        # Create a mini Dark Mode popup to show while downloading
        $dlWindowXaml = @"
        <Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Installing Winget..." Width="400" Height="120" WindowStartupLocation="CenterScreen" WindowStyle="ToolWindow" Background="#1E1E1E" Foreground="#D4D4D4">
            <StackPanel VerticalAlignment="Center" Margin="20">
                <TextBlock Text="Downloading and installing Winget..." Margin="0,0,0,10" HorizontalAlignment="Center" FontSize="14" FontWeight="SemiBold"/>
                <ProgressBar IsIndeterminate="True" Height="15" Background="#2D2D30" Foreground="#007ACC" BorderThickness="0"/>
            </StackPanel>
        </Window>
"@
        $dlReader = (New-Object System.Xml.XmlNodeReader ([xml]$dlWindowXaml))
        $dlWindow = [Windows.Markup.XamlReader]::Load($dlReader)
        $dlWindow.Show()
        
        # Force the UI to draw on the screen before the thread freezes for the download
        try { $dlWindow.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Render) } catch {}
        
        try {
            $ProgressPreference = 'SilentlyContinue' # Hides the raw console download bar to speed up Invoke-WebRequest
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            
            $bundlePath = Join-Path $env:TEMP "winget.msixbundle"
            
            # Download the latest offline installer bundle directly from Microsoft's GitHub
            Invoke-WebRequest -Uri "https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle" -OutFile $bundlePath -UseBasicParsing
            
            # Install the Appx Package
            Add-AppxPackage -Path $bundlePath
            Remove-Item $bundlePath -ErrorAction SilentlyContinue
            
            $dlWindow.Close()
            [System.Windows.MessageBox]::Show("Winget was installed successfully! Starting WinToolsUI...", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } catch {
            $dlWindow.Close()
            [System.Windows.MessageBox]::Show("Failed to install Winget. Error: $($_.Exception.Message)`n`nPlease install 'App Installer' manually from the Microsoft Store.", "Install Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            exit
        }
    } else {
        # User clicked No, so we must exit as the app cannot run without Winget
        exit
    }
}
# ----------------------------------------------------

# 1. Define the UI layout using XAML
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="WinToolsUI" Height="600" Width="850" WindowStartupLocation="CenterScreen">
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
                    <WrapPanel Grid.Row="4" Margin="0,10,0,0" HorizontalAlignment="Right">
                        <CheckBox Name="CreateRestorePointInstallCheck" Content="Create Restore Point" IsChecked="False" VerticalAlignment="Center" Margin="0,0,15,5" Foreground="{StaticResource TextForeground}"/>
                        <CheckBox Name="AdminInstallCheck" Content="Install for All Users (Admin)" IsChecked="True" VerticalAlignment="Center" Margin="0,0,15,5" Foreground="{StaticResource TextForeground}"/>
                        <Button Name="ImportQueueBtn" Content="Import Queue" Padding="8" Width="110" Margin="0,0,10,5"/>
                        <Button Name="ExportQueueBtn" Content="Export Queue" Padding="8" Width="110" Margin="0,0,10,5"/>
                        <Button Name="RemoveFromQueueBtn" Content="Remove from Queue" Padding="8" Width="130" Margin="0,0,10,5"/>
                        <Button Name="InstallBtn" Content="Install Queued Packages" Padding="8" Width="180" FontWeight="Bold" Foreground="#73D96B" Margin="0,0,0,5"/>
                    </WrapPanel>
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

                    <WrapPanel Grid.Row="2" Margin="0,10,0,0" HorizontalAlignment="Right">
                        <CheckBox Name="CreateRestorePointUpdateCheck" Content="Create Restore Point" IsChecked="False" VerticalAlignment="Center" Margin="0,0,15,5" Foreground="{StaticResource TextForeground}"/>
                        <Button Name="UninstallBtn" Content="Uninstall Selected" Padding="8" Width="140" Foreground="#FF6B6B" FontWeight="Bold" Margin="0,0,10,5"/>
                        <Button Name="UpdateBtn" Content="Update Selected" Padding="8" Width="140" Foreground="#6BA4FF" FontWeight="Bold" Margin="0,0,10,5"/>
                        <Button Name="UpdateAllBtn" Content="Update All Apps" Padding="8" Width="140" Foreground="#73D96B" FontWeight="Bold" Margin="0,0,0,5"/>
                    </WrapPanel>
                </Grid>
            </TabItem>
            
            <!-- NEW: Privacy & Ads Tab -->
            <TabItem Header="Privacy &amp; Ads">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <ScrollViewer VerticalScrollBarVisibility="Auto" Grid.Row="0">
                        <StackPanel Margin="5,10,5,15">
                            <TextBlock Text="Privacy &amp; Ad Blocker" FontWeight="Bold" FontSize="18" Foreground="{StaticResource AccentColor}" Margin="0,0,0,5"/>
                            <TextBlock Text="Toggle the switches below to disable telemetry tracking and system-wide advertisements." Foreground="#AAAAAA" Margin="0,0,0,10"/>
                            <TextBlock Name="PrivacyAdminWarning" Text="Administrator privileges are required to apply system-level privacy settings." Foreground="#FF6B6B" FontWeight="SemiBold" Visibility="Collapsed" Margin="0,0,0,15"/>
                            
                            <!-- Telemetry Box -->
                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                <StackPanel>
                                    <TextBlock Text="Telemetry &amp; Data Collection" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                    <CheckBox Name="ChkTelemetry" Content="Disable Diagnostic Data &amp; Telemetry (DiagTrack)" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                    <CheckBox Name="ChkActivity" Content="Disable Activity History &amp; Timeline Tracking" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                    <CheckBox Name="ChkTailoredExp" Content="Disable Tailored Experiences (Diagnostic data-based ads)" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                    <CheckBox Name="ChkWER" Content="Disable Windows Error Reporting (Prevent crash dump uploads)" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                    <CheckBox Name="ChkFeedback" Content="Disable Feedback Prompts (Stop Microsoft surveys)" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                </StackPanel>
                            </Border>
                            
                            <!-- Ads Box -->
                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                <StackPanel>
                                    <TextBlock Text="System Annoyances &amp; Ads" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                    <CheckBox Name="ChkBingSearch" Content="Disable Bing Web Search in Start Menu" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                    <CheckBox Name="ChkStartAds" Content="Disable Start Menu Suggestions (Promoted apps)" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                    <CheckBox Name="ChkLockScreenAds" Content="Disable Lock Screen Tips &amp; Fun Facts (Spotlight ads)" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                    <CheckBox Name="ChkExplorerAds" Content="Disable File Explorer Notifications (OneDrive/Office 365 banners)" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                    <CheckBox Name="ChkWelcomeExp" Content="Disable Windows Welcome Experience (Post-update nagging)" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                    <CheckBox Name="ChkAdId" Content="Disable Advertising ID (Targeted app ads)" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                </StackPanel>
                            </Border>

                            <!-- AI & Features Box -->
                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                <StackPanel>
                                    <TextBlock Text="Windows Features &amp; AI" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                    <CheckBox Name="ChkCopilot" Content="Disable Windows Copilot &amp; AI Features" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                    <CheckBox Name="ChkWidgets" Content="Disable Taskbar Widgets / News &amp; Interests" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                </StackPanel>
                            </Border>

                            <!-- Bloatware Box -->
                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                <StackPanel>
                                    <TextBlock Text="Automatic Installations" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                    <CheckBox Name="ChkConsumer" Content="Disable Windows Consumer Features (Prevents Candy Crush/TikTok auto-installs)" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                </StackPanel>
                            </Border>
                            
                            <!-- Network Box -->
                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                <StackPanel>
                                    <TextBlock Text="Network Privacy" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                    <CheckBox Name="ChkWifiSense" Content="Disable Wi-Fi Sense (Stops background open-network connections)" Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                </StackPanel>
                            </Border>
                        </StackPanel>
                    </ScrollViewer>
                    
                    <!-- Action Buttons -->
                    <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                        <CheckBox Name="CreateRestorePointPrivacyCheck" Content="Create Restore Point" IsChecked="True" VerticalAlignment="Center" Margin="0,0,15,0" Foreground="{StaticResource TextForeground}"/>
                        <Button Name="RefreshPrivacyBtn" Content="Refresh Status" Padding="15,8" Margin="0,0,10,0"/>
                        <Button Name="ApplyPrivacyBtn" Content="Apply Privacy Settings" Padding="15,8" Width="200" FontWeight="Bold" Foreground="#FF6B6B"/>
                    </StackPanel>
                </Grid>
            </TabItem>

            <!-- NEW: Utilities Tab -->
            <TabItem Header="Utilities">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <StackPanel Margin="15,20,15,15">
                        <TextBlock Text="System Utilities" FontWeight="Bold" FontSize="18" Foreground="{StaticResource AccentColor}" Margin="0,0,0,5"/>
                        <TextBlock Name="UtilAdminWarning" Text="Administrator privileges are required for these tools." Foreground="#FF6B6B" FontWeight="SemiBold" Visibility="Collapsed" Margin="0,0,0,15"/>
                        
                        <UniformGrid Columns="2">
                            <!-- Box 1 -->
                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,10,15">
                                <StackPanel>
                                    <TextBlock Text="System Repair &amp; Maintenance" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                    <TextBlock Text="Run system corruption scans and fix damaged Windows components." TextWrapping="Wrap" Foreground="#AAAAAA" Margin="0,0,0,10" Height="35"/>
                                    <WrapPanel>
                                        <Button Name="UtilSysScanBtn" Content="Run System Scan (SFC &amp; DISM)" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilResetWUBtn" Content="Reset Windows Update" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilRestorePointBtn" Content="Create System Restore Point" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilOpenRestoreBtn" Content="Open System Restore" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilLongPathBtn" Content="Enable Long Paths (Remove 260 Char Limit)" Padding="10,8" Margin="0,0,10,10"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>

                            <!-- Box 2 -->
                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="10,0,0,15">
                                <StackPanel>
                                    <TextBlock Text="App &amp; Package Managers" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                    <TextBlock Text="Repair broken Store apps or reset Winget repositories." TextWrapping="Wrap" Foreground="#AAAAAA" Margin="0,0,0,10" Height="35"/>
                                    <WrapPanel>
                                        <Button Name="UtilWingetRepairBtn" Content="Repair Winget Sources" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilStoreRepairBtn" Content="Repair Microsoft Store" Padding="10,8" Margin="0,0,10,10"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>

                            <!-- Box 3 -->
                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,10,15">
                                <StackPanel>
                                    <TextBlock Text="System Cleanup" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                    <TextBlock Text="Free up disk space and remove unnecessary system logs." TextWrapping="Wrap" Foreground="#AAAAAA" Margin="0,0,0,10" Height="35"/>
                                    <WrapPanel>
                                        <Button Name="UtilDiskCleanupBtn" Content="Deep Disk Cleanup" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilClearLogsBtn" Content="Clear Event Viewer Logs" Padding="10,8" Margin="0,0,10,10"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>

                            <!-- Box 4 -->
                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="10,0,0,15">
                                <StackPanel>
                                    <TextBlock Text="Desktop &amp; UI Repair" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                    <TextBlock Text="Fix blank icons and broken image thumbnails by rebuilding the cache." TextWrapping="Wrap" Foreground="#AAAAAA" Margin="0,0,0,10" Height="35"/>
                                    <WrapPanel>
                                        <Button Name="UtilIconCacheBtn" Content="Rebuild Icon &amp; Thumbnail Cache" Padding="10,8" Margin="0,0,10,10"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>

                            <!-- Box 5 -->
                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,10,15">
                                <StackPanel>
                                    <TextBlock Text="Network Tools" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                    <TextBlock Text="Fix network connectivity issues or enable legacy sharing protocols." TextWrapping="Wrap" Foreground="#AAAAAA" Margin="0,0,0,10" Height="35"/>
                                    <WrapPanel>
                                        <Button Name="UtilResetNetBtn" Content="Reset Network Adapters" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilSMBBtn" Content="Enable SMBv1 Protocol" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilFileSharingBtn" Content="Enable File &amp; Printer Sharing" Padding="10,8" Margin="0,0,10,10"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>
                            
                            <!-- Box 6 -->
                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="10,0,0,15">
                                <StackPanel>
                                    <TextBlock Text="Hardware &amp; Drivers" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                    <TextBlock Text="Download missing or outdated hardware drivers via Microsoft Update or Snappy Driver Installer." TextWrapping="Wrap" Foreground="#AAAAAA" Margin="0,0,0,10" Height="35"/>
                                    <WrapPanel>
                                        <Button Name="UtilDriverBtn" Content="Official Microsoft Drivers" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilSDIOBtn" Content="Snappy Driver Installer (SDIO)" Padding="10,8" Margin="0,0,10,10"/>
                                    </WrapPanel>
                                </StackPanel>
                            </Border>
                        </UniformGrid>
                        
                    </StackPanel>
                </ScrollViewer>
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
$CreateRestorePointInstallCheck = $Window.FindName("CreateRestorePointInstallCheck")
$AdminInstallCheck = $Window.FindName("AdminInstallCheck")
$InstallBtn = $Window.FindName("InstallBtn")

# Utilities UI Map
$UtilAdminWarning = $Window.FindName("UtilAdminWarning")
$UtilSysScanBtn = $Window.FindName("UtilSysScanBtn")
$UtilResetWUBtn = $Window.FindName("UtilResetWUBtn")
$UtilRestorePointBtn = $Window.FindName("UtilRestorePointBtn")
$UtilOpenRestoreBtn = $Window.FindName("UtilOpenRestoreBtn")
$UtilLongPathBtn = $Window.FindName("UtilLongPathBtn")
$UtilResetNetBtn = $Window.FindName("UtilResetNetBtn")
$UtilSMBBtn = $Window.FindName("UtilSMBBtn")
$UtilFileSharingBtn = $Window.FindName("UtilFileSharingBtn")
$UtilDriverBtn = $Window.FindName("UtilDriverBtn")
$UtilSDIOBtn = $Window.FindName("UtilSDIOBtn")
$UtilWingetRepairBtn = $Window.FindName("UtilWingetRepairBtn")
$UtilStoreRepairBtn = $Window.FindName("UtilStoreRepairBtn")
$UtilDiskCleanupBtn = $Window.FindName("UtilDiskCleanupBtn")
$UtilClearLogsBtn = $Window.FindName("UtilClearLogsBtn")
$UtilIconCacheBtn = $Window.FindName("UtilIconCacheBtn")

# Privacy UI Map
$PrivacyAdminWarning = $Window.FindName("PrivacyAdminWarning")
$ChkTelemetry = $Window.FindName("ChkTelemetry")
$ChkActivity = $Window.FindName("ChkActivity")
$ChkTailoredExp = $Window.FindName("ChkTailoredExp")
$ChkWER = $Window.FindName("ChkWER")
$ChkFeedback = $Window.FindName("ChkFeedback")
$ChkBingSearch = $Window.FindName("ChkBingSearch")
$ChkStartAds = $Window.FindName("ChkStartAds")
$ChkLockScreenAds = $Window.FindName("ChkLockScreenAds")
$ChkExplorerAds = $Window.FindName("ChkExplorerAds")
$ChkWelcomeExp = $Window.FindName("ChkWelcomeExp")
$ChkAdId = $Window.FindName("ChkAdId")
$ChkCopilot = $Window.FindName("ChkCopilot")
$ChkWidgets = $Window.FindName("ChkWidgets")
$ChkConsumer = $Window.FindName("ChkConsumer")
$ChkWifiSense = $Window.FindName("ChkWifiSense")
$RefreshPrivacyBtn = $Window.FindName("RefreshPrivacyBtn")
$ApplyPrivacyBtn = $Window.FindName("ApplyPrivacyBtn")
$CreateRestorePointPrivacyCheck = $Window.FindName("CreateRestorePointPrivacyCheck")

$CreateRestorePointUpdateCheck = $Window.FindName("CreateRestorePointUpdateCheck")

# --- Apply Admin Status to UI ---
if ($isActualAdmin) {
    $Window.Title = "WinToolsUI (Administrator)"
    $AdminInstallCheck.IsChecked = $true
    $AdminInstallCheck.IsEnabled = $true
    
    $CreateRestorePointInstallCheck.IsEnabled = $true
    $CreateRestorePointUpdateCheck.IsEnabled = $true
    $CreateRestorePointPrivacyCheck.IsEnabled = $true
} else {
    $Window.Title = "WinToolsUI (Standard User)"
    $AdminInstallCheck.IsChecked = $false
    $AdminInstallCheck.IsEnabled = $false
    $AdminInstallCheck.Content = "Install for Current User (No Admin)"
    $AdminInstallCheck.ToolTip = "You declined the Administrator prompt. Installations are limited to the current user."
    $AdminInstallCheck.Foreground = "#888888"
    
    $CreateRestorePointInstallCheck.IsEnabled = $false
    $CreateRestorePointInstallCheck.ToolTip = "Administrator privileges are required to create Restore Points."
    $CreateRestorePointInstallCheck.Foreground = "#888888"
    
    $CreateRestorePointUpdateCheck.IsEnabled = $false
    $CreateRestorePointUpdateCheck.ToolTip = "Administrator privileges are required to create Restore Points."
    $CreateRestorePointUpdateCheck.Foreground = "#888888"
    
    $CreateRestorePointPrivacyCheck.IsEnabled = $false
    $CreateRestorePointPrivacyCheck.ToolTip = "Administrator privileges are required to create Restore Points."
    $CreateRestorePointPrivacyCheck.Foreground = "#888888"
    
    # Disable Utilities if not admin
    $UtilAdminWarning.Visibility = 'Visible'
    $UtilSysScanBtn.IsEnabled = $false
    $UtilResetWUBtn.IsEnabled = $false
    $UtilRestorePointBtn.IsEnabled = $false
    $UtilOpenRestoreBtn.IsEnabled = $false
    $UtilLongPathBtn.IsEnabled = $false
    $UtilResetNetBtn.IsEnabled = $false
    $UtilSMBBtn.IsEnabled = $false
    $UtilFileSharingBtn.IsEnabled = $false
    $UtilDriverBtn.IsEnabled = $false
    $UtilSDIOBtn.IsEnabled = $false
    $UtilWingetRepairBtn.IsEnabled = $false
    $UtilStoreRepairBtn.IsEnabled = $false
    $UtilDiskCleanupBtn.IsEnabled = $false
    $UtilClearLogsBtn.IsEnabled = $false
    $UtilIconCacheBtn.IsEnabled = $false
    
    # Disable Privacy if not admin
    $PrivacyAdminWarning.Visibility = 'Visible'
    $ApplyPrivacyBtn.IsEnabled = $false
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
$syncHash.AppPath = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { $PWD.Path }
$runspace = [runspacefactory]::CreateRunspace()
$runspace.Open()
$runspace.SessionStateProxy.SetVariable("syncHash", $syncHash)

$script:psInstance = $null
$script:asyncResult = $null
$script:IsJobRunning = $false
$script:AllInstalledApps = $null

# 4. Define the Universal Background Job
$bgJobBlock = {
    param($Action, $Query, $Id, $Hash, $IsAdmin, $CreateRestore)
    
    # Force PowerShell to read external Winget output using UTF-8 to prevent 'ΓÇª' encoding issues
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
    
    # Helper to create restore points and bypass the Windows 24-hour limit
    function New-BypassRestorePoint($Desc) {
        try {
            Enable-ComputerRestore -Drive "$env:SystemDrive\" -ErrorAction SilentlyContinue
            
            # Bypass the Windows 1-per-24-hours limit
            $srKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore"
            if (-not (Test-Path $srKey)) { New-Item -Path $srKey -Force -ErrorAction SilentlyContinue | Out-Null }
            Set-ItemProperty -Path $srKey -Name "SystemRestorePointCreationFrequency" -Value 0 -Type DWord -ErrorAction SilentlyContinue
            
            Checkpoint-Computer -Description $Desc -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
            $Hash.LogQueue.Enqueue("System Restore Point created successfully.")
            return $true
        } catch {
            $Hash.LogQueue.Enqueue("Warning: Failed to create restore point ($($_.Exception.Message)).")
            return $false
        }
    }

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
                $wingetArgs = @("search", $Query, "--count", "40", "--accept-source-agreements", "--disable-interactivity")
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
                $wingetArgs = @("list", "--accept-source-agreements", "--disable-interactivity")
                & winget @wingetArgs 2>&1 | ForEach-Object {
                    $line = $_.ToString()
                    $Hash.LogQueue.Enqueue($line)
                    $raw += $line
                }
                $Hash.Result = ConvertFrom-WingetOutput $raw
            }
            'Install' {
                if ($CreateRestore) {
                    $Hash.LogQueue.Enqueue(">>> Creating System Restore Point before installation...")
                    New-BypassRestorePoint -Desc "WinToolsUI Install" | Out-Null
                }
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
                if ($CreateRestore) {
                    $Hash.LogQueue.Enqueue(">>> Creating System Restore Point before uninstallation...")
                    New-BypassRestorePoint -Desc "WinToolsUI Uninstall" | Out-Null
                }
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
                if ($CreateRestore) {
                    $Hash.LogQueue.Enqueue(">>> Creating System Restore Point before update...")
                    New-BypassRestorePoint -Desc "WinToolsUI Update" | Out-Null
                }
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
                $wingetArgs = @("show", "--id", $Id, "--exact", "--accept-source-agreements", "--disable-interactivity")
                & winget @wingetArgs 2>&1 | ForEach-Object {
                    $line = $_.ToString()
                    $Hash.LogQueue.Enqueue($line)
                    $raw += $line
                }
                $Hash.Result = $raw -join "`r`n"
            }
            # --- NEW UTILITY COMMANDS ---
            'UtilSystemScan' {
                $Hash.LogQueue.Enqueue(">>> Running DISM Component Store Cleanup and Repair...")
                $Hash.LogQueue.Enqueue(">>> (This may take several minutes to complete)")
                & DISM.exe /Online /Cleanup-image /Restorehealth 2>&1 | ForEach-Object {
                    $l = $_.ToString().Trim()
                    if (-not [string]::IsNullOrWhiteSpace($l)) { $Hash.LogQueue.Enqueue($l) }
                }
                
                $Hash.LogQueue.Enqueue("`r`n>>> Running System File Checker (SFC)...")
                & sfc /scannow 2>&1 | ForEach-Object {
                    $l = $_.ToString().Trim()
                    if (-not [string]::IsNullOrWhiteSpace($l)) { $Hash.LogQueue.Enqueue($l) }
                }
                $Hash.Result = "Success"
            }
            'UtilResetWU' {
                $Hash.LogQueue.Enqueue(">>> Stopping Windows Update Services...")
                $services = @("wuauserv", "cryptSvc", "bits", "msiserver")
                foreach ($svc in $services) {
                    $Hash.LogQueue.Enqueue("Stopping $svc...")
                    Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
                }
                
                $Hash.LogQueue.Enqueue(">>> Renaming SoftwareDistribution and catroot2 folders...")
                Rename-Item -Path "$env:windir\SoftwareDistribution" -NewName "SoftwareDistribution.old" -ErrorAction SilentlyContinue
                Rename-Item -Path "$env:windir\System32\catroot2" -NewName "catroot2.old" -ErrorAction SilentlyContinue
                
                $Hash.LogQueue.Enqueue(">>> Restarting Windows Update Services...")
                foreach ($svc in $services) {
                    $Hash.LogQueue.Enqueue("Starting $svc...")
                    Start-Service -Name $svc -ErrorAction SilentlyContinue
                }
                $Hash.Result = "Success"
            }
            'UtilRestorePoint' {
                $Hash.LogQueue.Enqueue(">>> Initializing System Restore Point creation...")
                $success = New-BypassRestorePoint -Desc "WinToolsUI Checkpoint"
                if ($success) {
                    $Hash.Result = "Success"
                } else {
                    $Hash.Result = "Error"
                }
            }
            'UtilResetNet' {
                $Hash.LogQueue.Enqueue(">>> Releasing and Renewing IP...")
                & ipconfig /release 2>&1 | Out-Null
                & ipconfig /flushdns 2>&1 | ForEach-Object { if (-not [string]::IsNullOrWhiteSpace($_)) { $Hash.LogQueue.Enqueue($_) } }
                & ipconfig /renew 2>&1 | Out-Null
                
                $Hash.LogQueue.Enqueue(">>> Resetting Winsock and IP Configuration...")
                & netsh winsock reset 2>&1 | ForEach-Object { if (-not [string]::IsNullOrWhiteSpace($_)) { $Hash.LogQueue.Enqueue($_) } }
                & netsh int ip reset 2>&1 | ForEach-Object { if (-not [string]::IsNullOrWhiteSpace($_)) { $Hash.LogQueue.Enqueue($_) } }
                $Hash.Result = "Success"
            }
            'UtilSMB' {
                $Hash.LogQueue.Enqueue(">>> Enabling SMBv1 Protocol...")
                try {
                    Enable-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -All -NoRestart -ErrorAction Stop | Out-Null
                    $Hash.LogQueue.Enqueue("Successfully enabled SMBv1. A system restart may be required.")
                } catch {
                    $Hash.LogQueue.Enqueue("Error enabling SMBv1: $($_.Exception.Message)")
                }
                $Hash.Result = "Success"
            }
            'UtilFileSharing' {
                $Hash.LogQueue.Enqueue(">>> Enabling File and Printer Sharing rules...")
                try {
                    # Enable the display group (this handles both inbound and outbound rules in the group)
                    Enable-NetFirewallRule -DisplayGroup "File and Printer Sharing" -ErrorAction Stop | Out-Null
                    
                    # Update scope to Any IP for all network profiles
                    Set-NetFirewallRule -DisplayGroup "File and Printer Sharing" -Profile Any -LocalAddress Any -RemoteAddress Any -ErrorAction Stop | Out-Null
                    
                    $Hash.LogQueue.Enqueue("Successfully enabled 'File and Printer Sharing' for all profiles (Public, Private, Domain).")
                    $Hash.LogQueue.Enqueue("Scope successfully expanded to Any IP Address (Inbound & Outbound).")
                    $Hash.Result = "Success"
                } catch {
                    $Hash.LogQueue.Enqueue("PowerShell Cmdlet Error: $($_.Exception.Message)")
                    $Hash.LogQueue.Enqueue("Attempting fallback using netsh...")
                    & netsh advfirewall firewall set rule group="File and Printer Sharing" new enable=Yes
                    $Hash.Result = "Success (Fallback)"
                }
            }
            'UtilWingetRepair' {
                $Hash.LogQueue.Enqueue(">>> Resetting Winget Sources...")
                & winget source reset --force 2>&1 | ForEach-Object {
                    if (-not [string]::IsNullOrWhiteSpace($_)) { $Hash.LogQueue.Enqueue($_) }
                }
                $Hash.Result = "Success"
            }
            'UtilStoreRepair' {
                $Hash.LogQueue.Enqueue(">>> Running wsreset.exe to clear Microsoft Store cache...")
                try {
                    Start-Process wsreset.exe -Wait -WindowStyle Hidden
                    $Hash.LogQueue.Enqueue("Microsoft Store cache cleared successfully.")
                    $Hash.Result = "Success"
                } catch {
                    $Hash.Result = "Error: $($_.Exception.Message)"
                }
            }
            'UtilDiskCleanup' {
                $Hash.LogQueue.Enqueue(">>> Emptying Windows System Temp Folder...")
                Remove-Item -Path "$env:windir\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue
                
                $Hash.LogQueue.Enqueue(">>> Emptying User Local Temp Folder...")
                Remove-Item -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue
                
                $Hash.LogQueue.Enqueue(">>> Emptying Recycle Bin...")
                Clear-RecycleBin -Force -ErrorAction SilentlyContinue
                
                $Hash.LogQueue.Enqueue("Disk cleanup completed.")
                $Hash.Result = "Success"
            }
            'UtilIconCache' {
                $Hash.LogQueue.Enqueue(">>> Stopping Explorer.exe...")
                Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
                Start-Sleep -Seconds 1
                
                $Hash.LogQueue.Enqueue(">>> Deleting Icon and Thumbnail Caches...")
                $cachePath = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"
                Remove-Item -Path "$cachePath\iconcache*" -Force -ErrorAction SilentlyContinue
                Remove-Item -Path "$cachePath\thumbcache*" -Force -ErrorAction SilentlyContinue
                
                $Hash.LogQueue.Enqueue(">>> Restarting Explorer.exe...")
                Start-Process explorer.exe
                
                $Hash.Result = "Success"
            }
            'UtilClearLogs' {
                $Hash.LogQueue.Enqueue(">>> Clearing all Event Viewer logs (this may take a minute)...")
                try {
                    $logs = wevtutil el
                    $count = 0
                    foreach ($log in $logs) {
                        wevtutil cl "$log" 2>$null
                        $count++
                    }
                    $Hash.LogQueue.Enqueue("Successfully cleared $count logs.")
                    $Hash.Result = "Success"
                } catch {
                    $Hash.Result = "Error: $($_.Exception.Message)"
                }
            }
            'UtilLongPath' {
                $Hash.LogQueue.Enqueue(">>> Enabling Win32 Long Paths (Removing MAX_PATH limit)...")
                try {
                    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1 -ErrorAction Stop
                    $Hash.LogQueue.Enqueue("Successfully enabled Long Paths in the registry.")
                    $Hash.Result = "Success"
                } catch {
                    $Hash.Result = "Error: $($_.Exception.Message)"
                }
            }
            'UtilDriverUpdate' {
                try {
                    $Hash.LogQueue.Enqueue(">>> Connecting to Microsoft Update Catalog...")
                    $UpdateSvc = New-Object -ComObject Microsoft.Update.ServiceManager
                    $UpdateSvc.AddService2("7971f918-a847-4430-9279-4a52d1efe18d", 7, "") | Out-Null
                    
                    $Session = New-Object -ComObject Microsoft.Update.Session
                    $Searcher = $Session.CreateUpdateSearcher()
                    $Searcher.ServiceID = '7971f918-a847-4430-9279-4a52d1efe18d'
                    $Searcher.SearchScope = 1
                    $Searcher.ServerSelection = 3 
                    
                    $Hash.LogQueue.Enqueue(">>> Scanning hardware for missing or outdated drivers. This may take a few minutes...")
                    $Criteria = "IsInstalled=0 and Type='Driver' and IsHidden=0"
                    $SearchResult = $Searcher.Search($Criteria)
                    $Updates = $SearchResult.Updates
                    
                    if ($Updates.Count -eq 0) {
                        $Hash.LogQueue.Enqueue("[+] Your system is fully up to date! No missing drivers found.")
                        $Hash.Result = "Success"
                        return
                    }
                    
                    $Hash.LogQueue.Enqueue("`r`nFound $($Updates.Count) driver update(s):")
                    
                    # Prepare to Download
                    $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
                    for ($i = 0; $i -lt $Updates.Count; $i++) {
                        $Update = $Updates.Item($i)
                        $Hash.LogQueue.Enqueue("  -> $($Update.Title)")
                        $UpdatesToDownload.Add($Update) | Out-Null
                    }
                    
                    $Hash.LogQueue.Enqueue("`r`n>>> Downloading Drivers...")
                    $Downloader = $Session.CreateUpdateDownloader()
                    $Downloader.Updates = $UpdatesToDownload
                    $Downloader.Download() | Out-Null
                    $Hash.LogQueue.Enqueue("[+] Download Complete.")
                    
                    # Filter for successfully downloaded drivers to install
                    $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
                    for ($i = 0; $i -lt $Updates.Count; $i++) {
                        $Update = $Updates.Item($i)
                        if ($Update.IsDownloaded) {
                            $UpdatesToInstall.Add($Update) | Out-Null
                        }
                    }
                    
                    if ($UpdatesToInstall.Count -gt 0) {
                        $Hash.LogQueue.Enqueue(">>> Installing Drivers...")
                        $Installer = $Session.CreateUpdateInstaller()
                        $Installer.Updates = $UpdatesToInstall
                        
                        $InstallationResult = $Installer.Install()
                        
                        $Hash.LogQueue.Enqueue("[+] Installation Process Finished.")
                        
                        if ($InstallationResult.RebootRequired) {
                            $Hash.LogQueue.Enqueue("===================================================")
                            $Hash.LogQueue.Enqueue("[!] REBOOT REQUIRED: Please restart your computer to apply the new drivers.")
                            $Hash.LogQueue.Enqueue("===================================================")
                            $Hash.Result = "Success (Reboot Required)"
                        } else {
                            $Hash.LogQueue.Enqueue("[+] All drivers installed successfully. No reboot required.")
                            $Hash.Result = "Success"
                        }
                    } else {
                        $Hash.LogQueue.Enqueue("[-] Could not verify downloaded drivers. Installation aborted.")
                        $Hash.Result = "Success"
                    }
                } catch {
                    $Hash.LogQueue.Enqueue("[-] Driver update failed: $($_.Exception.Message)")
                    $Hash.Result = "Error: $($_.Exception.Message)"
                }
            }
            'UtilSDIO' {
                $Hash.LogQueue.Enqueue(">>> Initializing Snappy Driver Installer Origin (SDIO)...")
                
                # Create the dedicated folder
                $SDIODir = Join-Path $Hash.AppPath "SDIO"
                if (-not (Test-Path $SDIODir)) {
                    New-Item -ItemType Directory -Path $SDIODir -Force | Out-Null
                }
                
                # Check for existing executables
                $SDIO_x64 = (Get-ChildItem -Path $SDIODir -Filter "SDIO_x64*.exe" | Select-Object -First 1).FullName
                $SDIO_x86 = (Get-ChildItem -Path $SDIODir -Filter "SDIO_R*.exe" | Select-Object -First 1).FullName
                $SDIOPath = if ($SDIO_x64) { $SDIO_x64 } else { $SDIO_x86 }

                if (-not $SDIOPath) {
                    $Hash.LogQueue.Enqueue(">>> SDIO not found locally. Downloading the latest version...")
                    $ZipPath = Join-Path $SDIODir "SDIO_Latest.zip"
                    $DownloadUrl = "https://www.glenn.delahoy.com/downloads/sdio/SDIO.zip"
                    
                    try {
                        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                        Invoke-WebRequest -Uri $DownloadUrl -OutFile $ZipPath -UseBasicParsing
                        
                        $Hash.LogQueue.Enqueue(">>> Extracting SDIO to dedicated folder...")
                        Expand-Archive -Path $ZipPath -DestinationPath $SDIODir -Force
                        Remove-Item -Path $ZipPath -Force
                        
                        # Recheck after extraction
                        $SDIO_x64 = (Get-ChildItem -Path $SDIODir -Filter "SDIO_x64*.exe" | Select-Object -First 1).FullName
                        $SDIO_x86 = (Get-ChildItem -Path $SDIODir -Filter "SDIO_R*.exe" | Select-Object -First 1).FullName
                        $SDIOPath = if ($SDIO_x64) { $SDIO_x64 } else { $SDIO_x86 }
                        
                        if ($SDIOPath) {
                            $Hash.LogQueue.Enqueue("[+] SDIO successfully downloaded and extracted!")
                        }
                    } catch {
                        $Hash.LogQueue.Enqueue("[-] Error downloading SDIO: $($_.Exception.Message)")
                        $Hash.Result = "Error: $($_.Exception.Message)"
                        return
                    }
                }
                
                if ($SDIOPath) {
                    $Hash.LogQueue.Enqueue(">>> Launching Snappy Driver Installer...")
                    $Hash.LogQueue.Enqueue("    -> Updating indexes, analyzing hardware, and installing drivers.")
                    $Hash.LogQueue.Enqueue("    -> (App will run minimized in the taskbar. This may take several minutes...)")
                    
                    # Changed from Hidden to Minimized because SDIO overrides strict hidden modes
                    Start-Process -FilePath $SDIOPath -ArgumentList "-autoupdate", "-autoinstall", "-autoclose" -WindowStyle Minimized -Wait
                    
                    $Hash.LogQueue.Enqueue("[+] SDIO Process Finished.")
                    $Hash.Result = "Success (SDIO)"
                } else {
                    $Hash.LogQueue.Enqueue("[-] SDIO execution aborted. Executable not found after extraction.")
                    $Hash.Result = "Error: SDIO executable missing."
                }
            }
            'ApplyPrivacy' {
                if ($CreateRestore) {
                    $Hash.LogQueue.Enqueue(">>> Creating System Restore Point before applying privacy settings...")
                    New-BypassRestorePoint -Desc "WinToolsUI Privacy Settings" | Out-Null
                }
                
                # Helper function for setting registry keys deeply
                function Set-PrivacyRegKey($Path, $Name, $Value, $Type = "DWord") {
                    if (-not (Test-Path $Path)) { New-Item -Path $Path -Force -ErrorAction SilentlyContinue | Out-Null }
                    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -ErrorAction SilentlyContinue
                }

                $cfg = $Query # Query holds our hashtable of checkbox states
                $Hash.LogQueue.Enqueue(">>> Applying Privacy and Ad-blocking settings...")
                
                # 1. Telemetry
                if ($cfg.Telemetry) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Windows Telemetry (DiagTrack)")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" "AllowTelemetry" 0
                    Set-PrivacyRegKey "HKLM:\SYSTEM\CurrentControlSet\Services\DiagTrack" "Start" 4
                } else {
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" "AllowTelemetry" 3
                    Set-PrivacyRegKey "HKLM:\SYSTEM\CurrentControlSet\Services\DiagTrack" "Start" 2
                }

                # 2. Activity History
                if ($cfg.Activity) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Activity History")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" "EnableActivityFeed" 0
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" "PublishUserActivities" 0
                } else {
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" "EnableActivityFeed" 1
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" "PublishUserActivities" 1
                }

                # 3. Tailored Experiences
                if ($cfg.TailoredExp) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Tailored Experiences")
                    Set-PrivacyRegKey "HKCU:\Software\Policies\Microsoft\Windows\CloudContent" "DisableTailoredExperiencesWithDiagnosticData" 1
                } else {
                    Set-PrivacyRegKey "HKCU:\Software\Policies\Microsoft\Windows\CloudContent" "DisableTailoredExperiencesWithDiagnosticData" 0
                }

                # 4. Start Menu Ads
                if ($cfg.StartAds) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Start Menu Suggested Apps")
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338388Enabled" 0
                } else {
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338388Enabled" 1
                }

                # 5. Lock Screen Ads
                if ($cfg.LockScreenAds) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Lock Screen Tips")
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338387Enabled" 0
                } else {
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338387Enabled" 1
                }

                # 6. File Explorer Ads
                if ($cfg.ExplorerAds) {
                    $Hash.LogQueue.Enqueue("    -> Disabling File Explorer Notifications")
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "ShowSyncProviderNotifications" 0
                } else {
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "ShowSyncProviderNotifications" 1
                }

                # 7. Welcome Experience
                if ($cfg.WelcomeExp) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Windows Welcome Experience")
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-310093Enabled" 0
                } else {
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-310093Enabled" 1
                }

                # 8. Advertising ID
                if ($cfg.AdId) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Advertising ID")
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" "Enabled" 0
                } else {
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" "Enabled" 1
                }

                # 9. Consumer Features (Bloatware)
                if ($cfg.Consumer) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Windows Consumer Features (Auto-Installs)")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" "DisableWindowsConsumerFeatures" 1
                } else {
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" "DisableWindowsConsumerFeatures" 0
                }
                
                # 10. Windows Error Reporting
                if ($cfg.WER) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Windows Error Reporting (Crash Dumps)")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting" "Disabled" 1
                } else {
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting" "Disabled" 0
                }

                # 11. Feedback Prompts
                if ($cfg.Feedback) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Feedback Prompts")
                    Set-PrivacyRegKey "HKCU:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" "DoNotShowFeedbackNotifications" 1
                } else {
                    Set-PrivacyRegKey "HKCU:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" "DoNotShowFeedbackNotifications" 0
                }

                # 12. Bing Web Search
                if ($cfg.BingSearch) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Bing Web Search in Start Menu")
                    Set-PrivacyRegKey "HKCU:\SOFTWARE\Policies\Microsoft\Windows\Explorer" "DisableSearchBoxSuggestions" 1
                } else {
                    Set-PrivacyRegKey "HKCU:\SOFTWARE\Policies\Microsoft\Windows\Explorer" "DisableSearchBoxSuggestions" 0
                }

                # 13. Copilot
                if ($cfg.Copilot) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Windows Copilot & AI Features")
                    Set-PrivacyRegKey "HKCU:\Software\Policies\Microsoft\Windows\WindowsCopilot" "TurnOffWindowsCopilot" 1
                } else {
                    Set-PrivacyRegKey "HKCU:\Software\Policies\Microsoft\Windows\WindowsCopilot" "TurnOffWindowsCopilot" 0
                }

                # 14. Widgets / News & Interests
                if ($cfg.Widgets) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Taskbar Widgets / News & Interests")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" "AllowNewsAndInterests" 0
                } else {
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" "AllowNewsAndInterests" 1
                }

                # 15. Wi-Fi Sense
                if ($cfg.WifiSense) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Wi-Fi Sense (Shared Hotspots)")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config" "AutoConnectAllowedOEM" 0
                } else {
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config" "AutoConnectAllowedOEM" 1
                }

                $Hash.LogQueue.Enqueue("[+] Privacy settings applied successfully.")
                $Hash.Result = "Success"
            }
        }
    } catch {
        $Hash.Result = "Error: $($_.Exception.Message)"
    }
}

# 5. Helper Function to safely dispatch jobs to the Runspace
function Start-WingetJob($Action, $Query, $Id, $StatusMsg, $IsAdmin = $false, $CreateRestore = $false) {
    # Check if a job is already running and warn the user
    if ($script:IsJobRunning) {
        [System.Windows.MessageBox]::Show("A task is currently running in the background.`n`nPlease wait for it to finish or click 'Stop' before starting a new action.", "Task in Progress", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    $script:IsJobRunning = $true
    
    # Clear the live log UI and Queue
    $LogTextBox.Clear()
    $dummy = [string]::Empty
    while ($syncHash.LogQueue.TryDequeue([ref]$dummy)) {}

    # Auto-expand the log panel if we are making system changes
    if ($Action -in @('Install', 'Uninstall', 'Update', 'UtilSystemScan', 'UtilResetWU', 'UtilRestorePoint', 'UtilLongPath', 'UtilResetNet', 'UtilSMB', 'UtilFileSharing', 'UtilDriverUpdate', 'UtilSDIO', 'UtilWingetRepair', 'UtilStoreRepair', 'UtilDiskCleanup', 'UtilIconCache', 'UtilClearLogs', 'ApplyPrivacy')) {
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

    $script:psInstance = [PowerShell]::Create().AddScript($bgJobBlock).AddArgument($Action).AddArgument($Query).AddArgument($Id).AddArgument($syncHash).AddArgument($IsAdmin).AddArgument($CreateRestore)
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
                    if ($res -ne $null -and $res.Count -gt 0) {
                        $DiscoverGrid.ItemsSource = $res
                        $StatusText.Text = "Search complete. Found $($res.Count) packages."
                    } else {
                        $DiscoverGrid.ItemsSource = $null
                        $StatusText.Text = "No results found. (If you are offline, search will not work)."
                    }
                }
                'Installed' {
                    if ($res -ne $null -and $res.Count -gt 0) {
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
                    } else {
                        $InstalledGrid.ItemsSource = $null
                        $WindowsAppsGrid.ItemsSource = $null
                        $StatusText.Text = "No apps found (or Winget failed to connect to the network)."
                    }
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
                'UtilSystemScan' { $StatusText.Text = "System scan and repair completed." }
                'UtilResetWU'    { $StatusText.Text = "Windows Update services reset completed." }
                'UtilRestorePoint' { $StatusText.Text = "System restore point created successfully." }
                'UtilLongPath'   { $StatusText.Text = "Long Paths have been enabled successfully." }
                'UtilResetNet'   { $StatusText.Text = "Network adapters reset successfully." }
                'UtilSMB'        { $StatusText.Text = "SMBv1 protocol has been enabled." }
                'UtilFileSharing'{ $StatusText.Text = "File and Printer Sharing enabled globally." }
                'UtilWingetRepair' { $StatusText.Text = "Winget repositories have been reset." }
                'UtilStoreRepair'  { $StatusText.Text = "Microsoft Store cache cleared." }
                'UtilDiskCleanup'  { $StatusText.Text = "Deep disk cleanup completed." }
                'UtilIconCache'    { $StatusText.Text = "Icon and thumbnail cache rebuilt." }
                'UtilClearLogs'    { $StatusText.Text = "Event Viewer logs successfully cleared." }
                'UtilDriverUpdate' { 
                    $StatusText.Text = if ($res -match "Reboot") { "Drivers installed! System reboot required." } else { "Driver scan and update process completed." } 
                }
                'UtilSDIO'         { $StatusText.Text = "Snappy Driver Installer process completed." }
                'ApplyPrivacy'     { $StatusText.Text = "Privacy and Ad settings applied successfully." }
            }
        }
    }
})

# --- Privacy Check/Read Function ---
function Read-PrivacyStates {
    # Helper to read registry silently
    function Get-PrivacyValue($Path, $Name, $ExpectedDisabledValue) {
        try {
            $val = (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name
            return ($val -eq $ExpectedDisabledValue)
        } catch { return $false }
    }
    
    $ChkTelemetry.IsChecked   = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" "AllowTelemetry" 0
    $ChkActivity.IsChecked    = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" "EnableActivityFeed" 0
    $ChkTailoredExp.IsChecked = Get-PrivacyValue "HKCU:\Software\Policies\Microsoft\Windows\CloudContent" "DisableTailoredExperiencesWithDiagnosticData" 1
    $ChkWER.IsChecked         = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting" "Disabled" 1
    $ChkFeedback.IsChecked    = Get-PrivacyValue "HKCU:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" "DoNotShowFeedbackNotifications" 1
    $ChkBingSearch.IsChecked  = Get-PrivacyValue "HKCU:\SOFTWARE\Policies\Microsoft\Windows\Explorer" "DisableSearchBoxSuggestions" 1
    $ChkStartAds.IsChecked    = Get-PrivacyValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338388Enabled" 0
    $ChkLockScreenAds.IsChecked = Get-PrivacyValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338387Enabled" 0
    $ChkExplorerAds.IsChecked = Get-PrivacyValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "ShowSyncProviderNotifications" 0
    $ChkWelcomeExp.IsChecked  = Get-PrivacyValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-310093Enabled" 0
    $ChkAdId.IsChecked        = Get-PrivacyValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" "Enabled" 0
    $ChkCopilot.IsChecked     = Get-PrivacyValue "HKCU:\Software\Policies\Microsoft\Windows\WindowsCopilot" "TurnOffWindowsCopilot" 1
    $ChkWidgets.IsChecked     = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" "AllowNewsAndInterests" 0
    $ChkConsumer.IsChecked    = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" "DisableWindowsConsumerFeatures" 1
    $ChkWifiSense.IsChecked   = Get-PrivacyValue "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config" "AutoConnectAllowedOEM" 0
}

$RefreshPrivacyBtn.Add_Click({
    Read-PrivacyStates
    $StatusText.Text = "Privacy settings refreshed from Registry."
})

$ApplyPrivacyBtn.Add_Click({
    $cfg = @{
        Telemetry     = $ChkTelemetry.IsChecked -eq $true
        Activity      = $ChkActivity.IsChecked -eq $true
        TailoredExp   = $ChkTailoredExp.IsChecked -eq $true
        WER           = $ChkWER.IsChecked -eq $true
        Feedback      = $ChkFeedback.IsChecked -eq $true
        BingSearch    = $ChkBingSearch.IsChecked -eq $true
        StartAds      = $ChkStartAds.IsChecked -eq $true
        LockScreenAds = $ChkLockScreenAds.IsChecked -eq $true
        ExplorerAds   = $ChkExplorerAds.IsChecked -eq $true
        WelcomeExp    = $ChkWelcomeExp.IsChecked -eq $true
        AdId          = $ChkAdId.IsChecked -eq $true
        Copilot       = $ChkCopilot.IsChecked -eq $true
        Widgets       = $ChkWidgets.IsChecked -eq $true
        Consumer      = $ChkConsumer.IsChecked -eq $true
        WifiSense     = $ChkWifiSense.IsChecked -eq $true
    }
    $createRestore = $CreateRestorePointPrivacyCheck.IsChecked -eq $true
    Start-WingetJob -Action "ApplyPrivacy" -Query $cfg -Id "" -StatusMsg "Applying privacy settings... Please wait." -CreateRestore $createRestore
})

# 7. Map Button Clicks
$StopJobBtn.Add_Click({
    if ($script:psInstance -ne $null -and $script:asyncResult.IsCompleted -eq $false) {
        $StatusText.Text = "Stopping operation..."
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
        $promptMsg = ""
        if ($ids.Count -eq 1) { 
            $msg = "Installing $($script:InstallQueue[0].Name)..." 
            $promptMsg = "Are you sure you want to install $($script:InstallQueue[0].Name)?"
        } else { 
            $msg = "Installing $($ids.Count) packages... Please wait." 
            $promptMsg = "Are you sure you want to install $($ids.Count) packages?"
        }
        
        $msgResult = [System.Windows.MessageBox]::Show($promptMsg, "Confirm Install", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        
        if ($msgResult -eq 'Yes') {
            $isAdmin = $AdminInstallCheck.IsChecked -eq $true
            $createRestore = $CreateRestorePointInstallCheck.IsChecked -eq $true
            Start-WingetJob -Action "Install" -Query "" -Id $ids -StatusMsg $msg -IsAdmin $isAdmin -CreateRestore $createRestore
        }
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
        
        $msg = ""
        $promptMsg = ""
        if ($ids.Count -eq 1) { 
            $msg = "Uninstalling $($selected[0].Name)..." 
            $promptMsg = "Are you sure you want to uninstall $($selected[0].Name)?"
        } else { 
            $msg = "Uninstalling $($ids.Count) packages... Please wait." 
            $promptMsg = "Are you sure you want to uninstall $($ids.Count) packages?"
        }

        $msgResult = [System.Windows.MessageBox]::Show($promptMsg, "Confirm Uninstall", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        
        if ($msgResult -eq 'Yes') {
            # Save the selected app objects to the background syncHash so the scanner knows what to look for
            $syncHash.TargetApps = $selected
            $createRestore = $CreateRestorePointUpdateCheck.IsChecked -eq $true
            Start-WingetJob -Action "Uninstall" -Query "" -Id $ids -StatusMsg $msg -CreateRestore $createRestore
        }
    } else { 
        [System.Windows.MessageBox]::Show("Please select at least one installed package to uninstall.") 
    }
})

$UpdateBtn.Add_Click({
    $selected = @($InstalledGrid.SelectedItems) + @($WindowsAppsGrid.SelectedItems)
    if ($selected.Count -gt 0) { 
        [string[]]$ids = @($selected | ForEach-Object { $_.Id })
        $msg = ""
        $promptMsg = ""
        if ($ids.Count -eq 1) { 
            $msg = "Updating $($selected[0].Name)..." 
            $promptMsg = "Are you sure you want to update $($selected[0].Name)?"
        } else { 
            $msg = "Updating $($ids.Count) packages... Please wait." 
            $promptMsg = "Are you sure you want to update $($ids.Count) packages?"
        }
        
        $msgResult = [System.Windows.MessageBox]::Show($promptMsg, "Confirm Update", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        
        if ($msgResult -eq 'Yes') {
            $createRestore = $CreateRestorePointUpdateCheck.IsChecked -eq $true
            Start-WingetJob -Action "Update" -Query "" -Id $ids -StatusMsg $msg -CreateRestore $createRestore
        }
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
        $promptMsg = if ($ids.Count -eq 1) { "Are you sure you want to update 1 package?" } else { "Are you sure you want to update all $($ids.Count) available packages?" }
        
        $msgResult = [System.Windows.MessageBox]::Show($promptMsg, "Confirm Update All", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        
        if ($msgResult -eq 'Yes') {
            $createRestore = $CreateRestorePointUpdateCheck.IsChecked -eq $true
            Start-WingetJob -Action "Update" -Query "" -Id $ids -StatusMsg $msg -CreateRestore $createRestore
        }
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

# --- Utility Button Event Handlers ---
$UtilSysScanBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will run DISM /RestoreHealth and SFC /scannow.`nThis process checks for system corruption and repairs missing Windows components.`n`nThis process can take up to 20 minutes and might cause high CPU usage. Continue?", "System Scan", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilSystemScan" -Query "" -Id "" -StatusMsg "Running System Scan (DISM & SFC)... Please wait." }
})

$UtilResetWUBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will stop all Windows Update services, clear the software distribution cache, and restart the services.`n`nUse this if updates are failing to download or install. Continue?", "Reset Windows Update", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilResetWU" -Query "" -Id "" -StatusMsg "Resetting Windows Update components..." }
})

$UtilRestorePointBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will create a new Windows System Restore Point.`n`nContinue?", "Create Restore Point", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilRestorePoint" -Query "" -Id "" -StatusMsg "Creating System Restore Point... Please wait." }
})

$UtilOpenRestoreBtn.Add_Click({
    try {
        Start-Process "$env:windir\System32\rstrui.exe"
    } catch {
        [System.Windows.MessageBox]::Show("Failed to open System Restore. It may be disabled on this system.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$UtilLongPathBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will modify the registry to remove the 260-character path limit (MAX_PATH) in Windows.`n`nThis allows applications to access deeply nested files without errors. Continue?", "Enable Long Paths", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilLongPath" -Query "" -Id "" -StatusMsg "Enabling Long Paths... Please wait." }
})

$UtilResetNetBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will flush DNS, release/renew your IP, and completely reset Winsock and TCP/IP configurations.`n`nYou may briefly lose internet connection. Continue?", "Reset Network", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilResetNet" -Query "" -Id "" -StatusMsg "Resetting Network Adapters and configurations..." }
})

$UtilSMBBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("Are you sure you want to enable the legacy SMBv1 protocol?`n`nNote: SMBv1 is considered insecure and should only be enabled if required for connecting to legacy NAS drives or older network printers.", "Enable SMBv1", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilSMB" -Query "" -Id "" -StatusMsg "Enabling SMBv1 Protocol..." }
})

$UtilFileSharingBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will enable the 'File and Printer Sharing' firewall rules for all network profiles (Public, Private, Domain) and allow inbound/outbound traffic from ANY IP address.`n`nWarning: Expanding this scope on Public networks can be a security risk. Continue?", "Enable File & Printer Sharing", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilFileSharing" -Query "" -Id "" -StatusMsg "Enabling File and Printer Sharing..." }
})

$UtilWingetRepairBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will force-reset the Winget package repositories.`n`nUse this if searches are failing or apps refuse to download. Continue?", "Repair Winget Sources", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilWingetRepair" -Query "" -Id "" -StatusMsg "Resetting Winget Sources..." }
})

$UtilStoreRepairBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will run wsreset.exe to clear the Microsoft Store cache.`n`nUse this if Store Apps are stuck on 'Pending' or fail to update. Continue?", "Repair Microsoft Store", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilStoreRepair" -Query "" -Id "" -StatusMsg "Clearing Microsoft Store Cache..." }
})

$UtilDiskCleanupBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will permanently empty the Recycle Bin and delete temporary files in both your User Temp and Windows Temp folders.`n`nContinue?", "Deep Disk Cleanup", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilDiskCleanup" -Query "" -Id "" -StatusMsg "Cleaning up disk space..." }
})

$UtilClearLogsBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will clear all background Windows Event Viewer logs.`n`nThis is useful for freeing up space or starting with a clean slate for troubleshooting. Continue?", "Clear Event Logs", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilClearLogs" -Query "" -Id "" -StatusMsg "Clearing Event Viewer Logs... This may take a moment." }
})

$UtilIconCacheBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will restart the Windows taskbar (explorer.exe) and delete the hidden icon/thumbnail databases to force Windows to rebuild them.`n`nYour screen will blink during this process. Continue?", "Rebuild Icon Cache", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilIconCache" -Query "" -Id "" -StatusMsg "Rebuilding Icon and Thumbnail caches..." }
})

$UtilDriverBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will scan Microsoft Update for missing or outdated hardware drivers, download them, and automatically install them.`n`nThe scan process can take several minutes. Continue?", "Update Hardware Drivers", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilDriverUpdate" -Query "" -Id "" -StatusMsg "Scanning for missing drivers... Please wait." }
})

$UtilSDIOBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will launch Snappy Driver Installer Origin (SDIO).`n`nIf this is your first time, it will automatically download the tool and extract it to an 'SDIO' folder next to this app.`n`nContinue?", "Snappy Driver Installer", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilSDIO" -Query "" -Id "" -StatusMsg "Initializing Snappy Driver Installer... Please wait." }
})

# --- Auto-Load Installed Apps on Startup ---
$Window.Add_Loaded({
    Read-PrivacyStates # Read initial privacy states
    Start-WingetJob -Action "Installed" -Query "" -Id "" -StatusMsg "Loading installed packages and checking for updates... This might take a moment."
})

# 6. Show the Window and clean up when closed
$Window.ShowDialog() | Out-Null

# Cleanup memory once the window is closed
$runspace.Close()
$runspace.Dispose()