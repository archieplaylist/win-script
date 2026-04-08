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
Add-Type -AssemblyName System.Windows.Forms

# --- Pre-flight Check: Ensure Winget is Installed ---
$script:WingetMissing = $false
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
            [System.Windows.MessageBox]::Show("Failed to install Winget. Error: $($_.Exception.Message)`n`nPlease install 'App Installer' manually from the Microsoft Store.`n`nWinToolsUI will launch with Discovery features disabled.", "Install Failed", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            $script:WingetMissing = $true
        }
    } else {
        # User clicked No, flag winget as missing to disable UI features
        $script:WingetMissing = $true
    }
}
# ----------------------------------------------------

# 1. Define the UI layout using XAML
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="WinToolsUI" Height="620" Width="870" WindowStartupLocation="CenterScreen">
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
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="TabBorder" Property="Background" Value="#1A1A1C"/>
                                <Setter Property="Foreground" Value="#666666"/>
                            </Trigger>
                            <MultiTrigger>
                                <MultiTrigger.Conditions>
                                    <Condition Property="IsMouseOver" Value="True"/>
                                    <Condition Property="IsSelected" Value="False"/>
                                    <Condition Property="IsEnabled" Value="True"/>
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
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#2D2D30"/>
                                <Setter Property="Foreground" Value="#666666"/>
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
            <TabItem Name="DiscoverTab" Header="Discover / Install">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="3*"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="2*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TabControl Name="DiscoverSubTabs" Grid.Row="0" Margin="0,0,0,10" Background="Transparent" BorderThickness="0">
                        <!-- Sub Tab: Search -->
                        <TabItem Header="Search Winget">
                            <Grid Margin="0,10,0,0">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <!-- Search Area -->
                                <StackPanel Orientation="Horizontal" Grid.Row="0" Margin="0,0,0,10">
                                    <TextBox Name="SearchBox" Width="400" Margin="0,0,10,0" Padding="5" VerticalContentAlignment="Center"/>
                                    <Button Name="SearchBtn" Content="Search" Width="120" Padding="5"/>
                                </StackPanel>
                                
                                <!-- Search Results Grid -->
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
                                
                                <!-- Add To Queue Button -->
                                <StackPanel Orientation="Horizontal" Grid.Row="2" Margin="0,10,0,0" HorizontalAlignment="Right">
                                    <Button Name="AddToQueueBtn" Content="Add Selected to Install Queue &#x2193;" Padding="8" Width="250" FontWeight="Bold"/>
                                </StackPanel>
                            </Grid>
                        </TabItem>
                        
                        <!-- Sub Tab: Templates -->
                        <TabItem Header="Quick Templates">
                            <Grid Margin="0,10,0,0">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                
                                <ScrollViewer Grid.Row="0" VerticalScrollBarVisibility="Auto" Padding="0,0,10,0">
                                    <Grid>
                                        <Grid.ColumnDefinitions>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                            <ColumnDefinition Width="*"/>
                                        </Grid.ColumnDefinitions>
                                        
                                        <!-- Column 1 -->
                                        <StackPanel Grid.Column="0" VerticalAlignment="Top">
                                            <!-- Utilities -->
                                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,10,10">
                                                <StackPanel>
                                                    <TextBlock Text="Utilities" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,8" Foreground="{StaticResource AccentColor}"/>
                                                    <CheckBox Name="TplEverything" Content="Everything Search" Tag="Winget|voidtools.Everything" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplAnyDesk" Content="AnyDesk" Tag="Winget|AnyDesk.AnyDesk" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplTightVNC" Content="TightVNC" Tag="Winget|GlavSoft.TightVNC" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplWizTree" Content="WizTree" Tag="Winget|AntibodySoftware.WizTree" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplQBittorrent" Content="qBittorrent" Tag="Winget|qBittorrent.qBittorrent" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplLocalSend" Content="LocalSend" Tag="Winget|LocalSend.LocalSend" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplFileZilla" Content="FileZilla (Choco)" Tag="Chocolatey|filezilla" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplWinSCP" Content="WinSCP" Tag="Winget|WinSCP.WinSCP" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplHeidiSQL" Content="HeidiSQL" Tag="Winget|HeidiSQL.HeidiSQL" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplDBeaver" Content="DBeaver" Tag="Winget|DBeaver.DBeaver.Community" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplVSCode" Content="VS Code" Tag="Winget|Microsoft.VisualStudioCode" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplPython" Content="Python Manager" Tag="Winget|Python.PythonInstallManager" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplGit" Content="Git" Tag="Winget|Git.Git" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                </StackPanel>
                                            </Border>
                                        </StackPanel>
                                        
                                        <!-- Column 2 -->
                                        <StackPanel Grid.Column="1" VerticalAlignment="Top">
                                            <!-- Browsers -->
                                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,10,10">
                                                <StackPanel>
                                                    <TextBlock Text="Web Browsers" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,8" Foreground="{StaticResource AccentColor}"/>
                                                    <CheckBox Name="TplChrome" Content="Google Chrome" Tag="Winget|Google.Chrome" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplFirefox" Content="Firefox" Tag="Winget|Mozilla.Firefox" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplEdge" Content="Microsoft Edge" Tag="Winget|Microsoft.Edge" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplBrave" Content="Brave" Tag="Winget|Brave.Brave" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplVivaldi" Content="Vivaldi" Tag="Winget|Vivaldi.Vivaldi" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                </StackPanel>
                                            </Border>
                                            
                                            <!-- Media -->
                                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,10,10">
                                                <StackPanel>
                                                    <TextBlock Text="Media" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,8" Foreground="{StaticResource AccentColor}"/>
                                                    <CheckBox Name="TpliTunes" Content="iTunes" Tag="Winget|Apple.iTunes" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplVLC" Content="VLC Media Player" Tag="Winget|VideoLAN.VLC" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplWinamp" Content="Winamp" Tag="Winget|Winamp.Winamp" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplAudacity" Content="Audacity" Tag="Winget|Audacity.Audacity" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplKLite" Content="K-Lite Codecs" Tag="Winget|CodecGuide.K-LiteCodecPack.Standard" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplHandBrake" Content="HandBrake" Tag="Winget|HandBrake.HandBrake" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                </StackPanel>
                                            </Border>
                                        </StackPanel>
                                        
                                        <!-- Column 3 -->
                                        <StackPanel Grid.Column="2" VerticalAlignment="Top">
                                            <!-- Documents -->
                                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,10,10">
                                                <StackPanel>
                                                    <TextBlock Text="Documents" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,8" Foreground="{StaticResource AccentColor}"/>
                                                    <CheckBox Name="TplLibreOffice" Content="LibreOffice" Tag="Winget|TheDocumentFoundation.LibreOffice" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplOnlyOffice" Content="OnlyOffice" Tag="Winget|ONLYOFFICE.DesktopEditors" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplOpenOffice" Content="OpenOffice" Tag="Winget|Apache.OpenOffice" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplPDF24" Content="PDF24 Creator" Tag="Winget|geeksoftwareGmbH.PDF24Creator" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                </StackPanel>
                                            </Border>
                                            
                                            <!-- Imaging -->
                                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,10,10">
                                                <StackPanel>
                                                    <TextBlock Text="Imaging" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,8" Foreground="{StaticResource AccentColor}"/>
                                                    <CheckBox Name="TplKrita" Content="Krita" Tag="Winget|KDE.Krita" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplPaintNet" Content="Paint.Net" Tag="Winget|dotPDN.PaintDotNet" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplGimp" Content="GIMP" Tag="Winget|GIMP.GIMP.3" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplInkscape" Content="Inkscape" Tag="Winget|Inkscape.Inkscape" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplFastStone" Content="FastStone Viewer" Tag="Winget|FastStone.Viewer" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                </StackPanel>
                                            </Border>

                                            <!-- Compression -->
                                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,10,10">
                                                <StackPanel>
                                                    <TextBlock Text="Compression" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,8" Foreground="{StaticResource AccentColor}"/>
                                                    <CheckBox Name="Tpl7zip" Content="7-Zip" Tag="Winget|7zip.7zip" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplPeaZip" Content="PeaZip" Tag="Winget|Giorgiotani.Peazip" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplWinRAR" Content="WinRAR" Tag="Winget|RARLab.WinRAR" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                </StackPanel>
                                            </Border>
                                        </StackPanel>

                                        <!-- Column 4 -->
                                        <StackPanel Grid.Column="3" VerticalAlignment="Top">
                                            <!-- Messaging -->
                                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,10,10">
                                                <StackPanel>
                                                    <TextBlock Text="Messaging" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,8" Foreground="{StaticResource AccentColor}"/>
                                                    <CheckBox Name="TplZoom" Content="Zoom" Tag="Winget|Zoom.Zoom" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplDiscord" Content="Discord" Tag="Winget|Discord.Discord" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplTeams" Content="Microsoft Teams" Tag="Winget|Microsoft.Teams" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplThunderbird" Content="Thunderbird" Tag="Winget|Mozilla.Thunderbird" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                </StackPanel>
                                            </Border>

                                            <!-- Gaming -->
                                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,10,10">
                                                <StackPanel>
                                                    <TextBlock Text="Gaming" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,8" Foreground="{StaticResource AccentColor}"/>
                                                    <CheckBox Name="TplSteam" Content="Steam" Tag="Winget|Valve.Steam" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplEpic" Content="Epic Games Launcher" Tag="Winget|EpicGames.EpicGamesLauncher" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                    <CheckBox Name="TplGOG" Content="GOG Galaxy" Tag="Winget|GOG.Galaxy" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                                                </StackPanel>
                                            </Border>
                                        </StackPanel>
                                    </Grid>
                                </ScrollViewer>
                            </Grid>
                        </TabItem>
                    </TabControl>
                    
                    <!-- Global Queue UI (Shared across sub-tabs) -->
                    <TextBlock Grid.Row="1" Text="Installation Queue" FontWeight="SemiBold" Margin="0,0,0,5" Foreground="{StaticResource TextForeground}"/>
                    <DataGrid Name="QueueGrid" Grid.Row="2" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Extended">
                        <DataGrid.ContextMenu>
                            <ContextMenu>
                                <MenuItem Name="QueueMenuDetails" Header="Show App Details..." />
                            </ContextMenu>
                        </DataGrid.ContextMenu>
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Queued App Name" Binding="{Binding Name}" Width="2*"/>
                            <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="1.5*"/>
                            <DataGridTextColumn Header="Target Version" Binding="{Binding Version}" Width="*"/>
                            <DataGridTextColumn Header="Manager" Binding="{Binding Manager}" Width="80"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    
                    <!-- Final Actions -->
                    <WrapPanel Grid.Row="3" Margin="0,10,0,0" HorizontalAlignment="Right">
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
                        <TextBox Name="SearchInstalledBox" Width="300" Margin="0,0,10,0" Padding="5" VerticalContentAlignment="Center"/>
                        <Button Name="RefreshInstalledBtn" Content="Load Installed &amp; Check Updates" Padding="5" Width="220" Margin="0,0,10,0"/>
                    </StackPanel>
                    
                    <TabControl Name="InstalledTabs" Grid.Row="1" Margin="0,5,0,0">
                        <TabItem Header="Desktop / External Apps">
                            <DataGrid Name="InstalledGrid" AutoGenerateColumns="False" IsReadOnly="False" SelectionMode="Extended" BorderThickness="0">
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
                                    <DataGridCheckBoxColumn Binding="{Binding IsSelected, UpdateSourceTrigger=PropertyChanged}" Width="40" Header="&#x2713;">
                                        <DataGridCheckBoxColumn.ElementStyle>
                                            <Style TargetType="CheckBox">
                                                <Setter Property="HorizontalAlignment" Value="Center"/>
                                                <Setter Property="VerticalAlignment" Value="Center"/>
                                            </Style>
                                        </DataGridCheckBoxColumn.ElementStyle>
                                    </DataGridCheckBoxColumn>
                                    <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2*" IsReadOnly="True"/>
                                    <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="1.5*" IsReadOnly="True"/>
                                    <DataGridTextColumn Header="Current Version" Binding="{Binding Version}" Width="*" IsReadOnly="True"/>
                                    <DataGridTextColumn Header="Available Update" Binding="{Binding Extra}" Width="*" IsReadOnly="True"/>
                                    <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="*" IsReadOnly="True"/>
                                </DataGrid.Columns>
                                <DataGrid.ContextMenu>
                                    <ContextMenu>
                                        <MenuItem Name="InstalledMenuDetails" Header="Show App Details..." />
                                        <MenuItem Name="InstalledMenuCopyId" Header="Copy App ID" />
                                        <MenuItem Name="InstalledMenuUninstall" Header="Uninstall" />
                                    </ContextMenu>
                                </DataGrid.ContextMenu>
                            </DataGrid>
                        </TabItem>
                        <TabItem Header="Windows / Store Apps">
                            <DataGrid Name="WindowsAppsGrid" AutoGenerateColumns="False" IsReadOnly="False" SelectionMode="Extended" BorderThickness="0">
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
                                    <DataGridCheckBoxColumn Binding="{Binding IsSelected, UpdateSourceTrigger=PropertyChanged}" Width="40" Header="&#x2713;">
                                        <DataGridCheckBoxColumn.ElementStyle>
                                            <Style TargetType="CheckBox">
                                                <Setter Property="HorizontalAlignment" Value="Center"/>
                                                <Setter Property="VerticalAlignment" Value="Center"/>
                                            </Style>
                                        </DataGridCheckBoxColumn.ElementStyle>
                                    </DataGridCheckBoxColumn>
                                    <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="2*" IsReadOnly="True"/>
                                    <DataGridTextColumn Header="ID" Binding="{Binding Id}" Width="1.5*" IsReadOnly="True"/>
                                    <DataGridTextColumn Header="Current Version" Binding="{Binding Version}" Width="*" IsReadOnly="True"/>
                                    <DataGridTextColumn Header="Available Update" Binding="{Binding Extra}" Width="*" IsReadOnly="True"/>
                                    <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="*" IsReadOnly="True"/>
                                </DataGrid.Columns>
                                <DataGrid.ContextMenu>
                                    <ContextMenu>
                                        <MenuItem Name="WindowsAppsMenuDetails" Header="Show App Details..." />
                                        <MenuItem Name="WindowsAppsMenuCopyId" Header="Copy App ID" />
                                        <MenuItem Name="WindowsAppsMenuUninstall" Header="Uninstall" />
                                    </ContextMenu>
                                </DataGrid.ContextMenu>
                            </DataGrid>
                        </TabItem>
                    </TabControl>

                    <WrapPanel Grid.Row="2" Margin="0,10,0,0" HorizontalAlignment="Right">
                        <Button Name="SelectAllInstalledBtn" Content="Select All" Padding="8" Width="90" Margin="0,0,10,5"/>
                        <Button Name="DeselectAllInstalledBtn" Content="Deselect" Padding="8" Width="90" Margin="0,0,10,5"/>
                        <CheckBox Name="CreateRestorePointUpdateCheck" Content="Create Restore Point" IsChecked="False" VerticalAlignment="Center" Margin="0,0,15,5" Foreground="{StaticResource TextForeground}"/>
                        <Button Name="UninstallBtn" Content="Uninstall Selected" Padding="8" Width="140" Foreground="#FF6B6B" FontWeight="Bold" Margin="0,0,10,5"/>
                        <Button Name="UpdateBtn" Content="Update Selected" Padding="8" Width="140" Foreground="#6BA4FF" FontWeight="Bold" Margin="0,0,10,5"/>
                        <Button Name="UpdateAllBtn" Content="Update All Apps" Padding="8" Width="140" Foreground="#73D96B" FontWeight="Bold" Margin="0,0,0,5"/>
                    </WrapPanel>
                </Grid>
            </TabItem>
            
            <!-- Optimize Tab (Combines Privacy, Performance, Customization) -->
            <TabItem Header="Optimize">
                <Grid Margin="10">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    
                    <TabControl Name="OptimizeSubTabs" Grid.Row="0" Margin="0,0,0,10" BorderThickness="0" Background="Transparent">
                        <!-- Customization & Theme Sub-Tab -->
                        <TabItem Header="Customization">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <StackPanel Margin="5,10,5,15">
                                    <TextBlock Text="Theme &amp; Customization" FontWeight="Bold" FontSize="18" Foreground="{StaticResource AccentColor}" Margin="0,0,0,5"/>
                                    <TextBlock Text="Personalize Windows appearance and theming (System &amp; App level)." Foreground="#AAAAAA" Margin="0,0,0,10"/>
                                    
                                    <!-- Dark Mode Box -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="Dark Mode &amp; Theme" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkDarkSystem" Content="Enable Dark Mode for Windows (Taskbar, Start Menu)" ToolTip="Forces the Windows OS UI elements to use dark theme." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkDarkApps" Content="Enable Dark Mode for Default Apps" ToolTip="Forces supported applications and Explorer to use dark theme." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>

                                    <!-- Visual Effects Box -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="Visual Effects &amp; Colors" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkDisableTransparency" Content="Disable Transparency Effects" ToolTip="Disables acrylic/blur effects across Windows. This can improve UI performance." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkTaskbarAccent" Content="Show Accent Color on Start Menu, Taskbar, and Action Center" ToolTip="Applies your current accent color to the Start menu and Taskbar backgrounds." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkTitlebarAccent" Content="Show Accent Color on Title Bars and Window Borders" ToolTip="Applies your current accent color to application window title bars and borders." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>
                                    
                                    <!-- Taskbar Box -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="Taskbar Preferences" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkTaskbarLeft" Content="Align Taskbar to the Left (Windows 11)" ToolTip="Moves the Windows 11 Start Menu and taskbar icons to the left side." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkHideSearch" Content="Hide Search on Taskbar" ToolTip="Removes the Search box or icon from the taskbar." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkHideTaskView" Content="Hide Task View Button" ToolTip="Removes the Task View (virtual desktops) button from the taskbar." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkHideChat" Content="Hide Chat Button (Windows 11)" ToolTip="Removes the Microsoft Teams Chat button from the taskbar." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>
                                    
                                    <!-- Start Menu Box -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="Start Menu Preferences" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkStartMorePins" Content="Show More Pins (Windows 11)" ToolTip="Changes the Start Menu layout to show more pinned apps and fewer recommendations." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkStartHideRecentApps" Content="Hide Recently Added Apps" ToolTip="Removes the 'Recently added' list from the Start Menu." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkStartHideMostUsed" Content="Hide Most Used Apps" ToolTip="Removes the 'Most used' list from the Start Menu." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkStartHideRecentDocs" Content="Hide Recommended/Recent Files" ToolTip="Hides recently opened items in the Start Menu, Jump Lists, and File Explorer." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>
                                    
                                    <!-- File Explorer Box -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="File Explorer Preferences" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkExpHidden" Content="Show Hidden Files, Folders, and Drives" ToolTip="Displays files and folders that are marked as hidden by the system." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkExpExt" Content="Show File Extensions for Known File Types" ToolTip="Displays the file extensions (like .txt, .exe) at the end of file names." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkExpThisPC" Content="Open File Explorer to 'This PC'" ToolTip="Changes the default opening folder of File Explorer from 'Home' or 'Quick Access' to 'This PC'." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkExpCompact" Content="Enable Compact View" ToolTip="Decreases space between items in File Explorer to show more files at once." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkExpClassicMenu" Content="Use Classic Context Menu (Windows 11)" ToolTip="Restores the classic Windows 10 full right-click menu in File Explorer (Requires Explorer restart)." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>

                        <!-- Privacy & Ads Sub-Tab -->
                        <TabItem Header="Privacy &amp; Ads">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <StackPanel Margin="5,10,5,15">
                                    <TextBlock Text="Privacy &amp; Ad Blocker" FontWeight="Bold" FontSize="18" Foreground="{StaticResource AccentColor}" Margin="0,0,0,5"/>
                                    <TextBlock Text="Toggle the switches below to disable telemetry tracking and system-wide advertisements. Hover over an item for details." Foreground="#AAAAAA" Margin="0,0,0,10"/>
                                    <TextBlock Name="PrivacyAdminWarning" Text="Administrator privileges are required to apply system-level privacy settings." Foreground="#FF6B6B" FontWeight="SemiBold" Visibility="Collapsed" Margin="0,0,0,15"/>
                                    
                                    <!-- Templates Box -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="Privacy Templates" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <TextBlock Text="Quickly apply predefined combinations of privacy settings." Foreground="#AAAAAA" Margin="0,0,0,10"/>
                                            <WrapPanel>
                                                <Button Name="BtnPrivacyMinimal" Content="Minimal" ToolTip="Disables basic telemetry and OS ads without breaking functionality." Padding="15,8" Margin="0,0,10,0"/>
                                                <Button Name="BtnPrivacyStandard" Content="Standard (Recommended)" ToolTip="Balances privacy and usability. Disables tracking, AI telemetry, and suggestions." Padding="15,8" Margin="0,0,10,0" Foreground="#6BA4FF"/>
                                                <Button Name="BtnPrivacyMaximal" Content="Maximal" ToolTip="Disables all tracking, telemetry, AI, app permissions, and background features." Padding="15,8" Foreground="#FF6B6B"/>
                                            </WrapPanel>
                                        </StackPanel>
                                    </Border>

                                    <!-- Telemetry Box -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="Telemetry &amp; Data Collection" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkTelemetry" Content="Disable Diagnostic Data &amp; Telemetry (DiagTrack)" ToolTip="Stops Windows from sending diagnostic data, typing data, and usage metrics to Microsoft." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkLocation" Content="Disable System-wide Location Services" ToolTip="Prevents Windows and apps from accessing your device's geographical location." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkActivity" Content="Disable Activity History &amp; Timeline Tracking" ToolTip="Stops Windows from tracking the files you open and websites you visit in the Timeline." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkAppLaunch" Content="Disable App Launch Tracking (Start/Search history)" ToolTip="Prevents Windows from personalizing your Start menu based on the apps you launch." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkLanguage" Content="Disable Language List Website Access" ToolTip="Prevents websites from seeing your locally installed languages." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkSpeech" Content="Disable Online Speech Recognition" ToolTip="Disables cloud-based speech recognition services used for dictation and Cortana." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkTailoredExp" Content="Disable Tailored Experiences (Diagnostic data-based ads)" ToolTip="Stops Microsoft from using your diagnostic data to offer tailored tips, ads, and recommendations." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkWER" Content="Disable Windows Error Reporting (Prevent crash dump uploads)" ToolTip="Stops Windows from automatically uploading crash dumps and error logs to Microsoft." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkFeedback" Content="Disable Feedback Prompts (Stop Microsoft surveys)" ToolTip="Disables those annoying 'How likely are you to recommend Windows?' popup surveys." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>

                                    <!-- O&O App Privacy Box -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="App Permissions &amp; Hardware Access" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkAppCamera" Content="Deny Apps Access to Camera" ToolTip="Prevents Windows apps from using your camera." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkAppMic" Content="Deny Apps Access to Microphone" ToolTip="Prevents Windows apps from using your microphone." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkAppAccountInfo" Content="Deny Apps Access to Account Info" ToolTip="Prevents apps from accessing your account name, picture, and other info." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkAppContacts" Content="Deny Apps Access to Contacts" ToolTip="Prevents Windows apps from reading your contacts list." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkAppCalendar" Content="Deny Apps Access to Calendar" ToolTip="Prevents Windows apps from reading your calendar events." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkAppEmail" Content="Deny Apps Access to Email" ToolTip="Prevents Windows apps from accessing and reading your emails." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkAppCallHistory" Content="Deny Apps Access to Call History" ToolTip="Prevents Windows apps from accessing your phone call history." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkAppTasks" Content="Deny Apps Access to Tasks" ToolTip="Prevents Windows apps from accessing your task lists." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkAppMessages" Content="Deny Apps Access to Messages" ToolTip="Prevents Windows apps from reading or sending SMS/MMS messages." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>
                                    
                                    <!-- O&O Security & Updates Box -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="Security Telemetry &amp; Updates" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkDefenderTelemetry" Content="Disable Windows Defender Telemetry (MAPS/Spynet)" ToolTip="Stops Defender from sending telemetry and sample files to Microsoft." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkSmartScreen" Content="Disable SmartScreen Filter" ToolTip="Disables the SmartScreen filter which sends visited URLs and downloaded file hashes to Microsoft." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkP2PUpdate" Content="Disable Windows Update Peer-to-Peer (Delivery Optimization)" ToolTip="Stops your PC from uploading Windows Updates to other PCs on the internet." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkCortana" Content="Disable Cortana (Legacy)" ToolTip="Completely disables the Cortana assistant." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkCEIP" Content="Disable Customer Experience Improvement Program (CEIP)" ToolTip="Stops Windows and installed programs from participating in the CEIP data collection." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkHandwriting" Content="Disable Typing and Inking Telemetry" ToolTip="Stops Microsoft from collecting your handwriting, typing, and dictation samples." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkAutoDriverUpdate" Content="Disable Automatic Driver Updates via Windows Update" ToolTip="Prevents Windows Update from automatically overwriting your hardware drivers." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>
                                    
                                    <!-- O&O Edge Privacy Box -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="Microsoft Edge Privacy (O&amp;O)" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkEdgeTelemetry" Content="Disable Edge Telemetry &amp; Site Tracking" ToolTip="Stops MS Edge from sending browsing data and usage metrics to Microsoft." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkEdgeCopilot" Content="Disable Edge Copilot / Discover Sidebar" ToolTip="Removes the Bing Copilot 'B' icon and sidebar from Microsoft Edge." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkEdgeSearchAds" Content="Disable Edge Search Suggestions" ToolTip="Prevents your keystrokes in the address bar from being sent to the search engine for suggestions." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>
                                    
                                    <!-- Ads Box -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="System Annoyances &amp; Ads" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkWorkplace" Content="Block 'Allow organization to manage device' Prompts" ToolTip="Stops the 'Allow my organization to manage my device' prompt when signing into work/school accounts." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkOneDrive" Content="Disable OneDrive Automatic Folder Backups" ToolTip="Prevents OneDrive from automatically hijacking your Desktop, Documents, and Pictures folders for backup." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkBingSearch" Content="Disable Bing Web Search in Start Menu" ToolTip="Removes Bing web search results from the Start Menu, making local searches much faster and private." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkStartAds" Content="Disable Start Menu Suggestions (Promoted apps)" ToolTip="Removes 'suggested' (promoted) apps from appearing in your Start Menu." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkLockScreenAds" Content="Disable Lock Screen Tips &amp; Fun Facts (Spotlight ads)" ToolTip="Disables tips, tricks, and promotional content on the Windows Lock Screen." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkExplorerAds" Content="Disable File Explorer Notifications (OneDrive/Office 365 banners)" ToolTip="Disables promotional banners (like Office 365 offers) inside File Explorer." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkWelcomeExp" Content="Disable Windows Welcome Experience (Post-update nagging)" ToolTip="Stops the 'Let's finish setting up your device' full-screen prompt after major Windows updates." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkAdId" Content="Disable Advertising ID (Targeted app ads)" ToolTip="Prevents apps from using a unique advertising ID to show you targeted ads." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>

                                    <!-- AI & Features Box -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="Windows Features, Security &amp; AI" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkBitLocker" Content="Prevent BitLocker Auto-Encryption (Windows 11)" ToolTip="(Windows 11) Prevents Windows from automatically encrypting your drives, which can slightly reduce SSD performance." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkCopilot" Content="Disable Windows Copilot &amp; AI Features" ToolTip="Removes the Windows Copilot icon and disables its integrated AI features." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkWidgets" Content="Disable Taskbar Widgets / News &amp; Interests" ToolTip="Removes the News &amp; Interests or Widgets panel from the Taskbar." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkMaintenance" Content="Disable Automatic System Maintenance" ToolTip="Stops Windows from running defrags, scans, and updates while the PC is idle." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>

                                    <!-- Bloatware Box -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="Automatic Installations" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkConsumer" Content="Disable Windows Consumer Features (Prevents Candy Crush/TikTok auto-installs)" ToolTip="Prevents Windows from automatically downloading bloatware like Candy Crush or TikTok on new accounts." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>
                                    
                                    <!-- Network Box -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="Network &amp; Remote Access" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkWifiSense" Content="Disable Wi-Fi Sense (Stops background open-network connections)" ToolTip="Stops your PC from automatically connecting to suggested open hotspots." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkRemoteAssist" Content="Disable Remote Assistance Connections" ToolTip="Disables the legacy Remote Assistance feature to reduce background listening ports." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                        
                        <!-- Performance & Gaming Sub-Tab -->
                        <TabItem Header="Performance">
                            <ScrollViewer VerticalScrollBarVisibility="Auto">
                                <StackPanel Margin="5,10,5,15">
                                    <TextBlock Text="Gaming &amp; Performance" FontWeight="Bold" FontSize="18" Foreground="{StaticResource AccentColor}" Margin="0,0,0,5"/>
                                    <TextBlock Text="Maximize your system's hardware potential, reduce input lag, and eliminate stutters. Hover over an item for details." Foreground="#AAAAAA" Margin="0,0,0,10"/>
                                    <TextBlock Name="PerfAdminWarning" Text="Administrator privileges are required to apply system-level performance settings." Foreground="#FF6B6B" FontWeight="SemiBold" Visibility="Collapsed" Margin="0,0,0,15"/>
                                    
                                    <!-- Gaming & CPU -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="Gaming &amp; Processing" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkGameMode" Content="Enable Windows Game Mode" ToolTip="Stops background activities (like Windows Update) and allocates more system resources to your active game." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkCpuPriority" Content="Optimize CPU/GPU Priority &amp; System Responsiveness for Gaming" ToolTip="Adjusts Windows thread scheduling to prioritize foreground gaming apps over background tasks." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkHwSchMode" Content="Enable Hardware-Accelerated GPU Scheduling (HAGS) (Requires Restart)" ToolTip="Offloads memory management from the CPU to the GPU, improving framerates and lowering latency." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkFSO" Content="Disable Fullscreen Optimizations (Fixes stutters in older games)" ToolTip="Disables Fullscreen Optimizations globally. This can fix micro-stutters and input lag in older DirectX 9/11 games." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkGameDVR" Content="Disable Xbox Game DVR &amp; Game Bar (Reduces background usage)" ToolTip="Disables the Xbox background recording service to free up CPU/GPU resources and reduce stuttering." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>
                                    
                                    <!-- Input & UI Responsiveness -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="Input &amp; UI Responsiveness" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkMouseAccel" Content="Disable Enhance Pointer Precision (Mouse Acceleration) (Requires Restart)" ToolTip="Disables 'Enhance Pointer Precision', ensuring 1:1 raw mouse input for consistent aiming in FPS games." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkMenuDelay" Content="Disable Menu Show Delay (Instant menus)" ToolTip="Changes the default 400ms delay to 0ms so right-click context menus and submenus open instantly." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkStartupDelay" Content="Disable Startup Delay for Apps (Faster boot login)" ToolTip="Removes the default 10-second delay Windows adds before launching your startup apps, speeding up login." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>
                                    
                                    <!-- System & Network -->
                                    <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,0,15">
                                        <StackPanel>
                                            <TextBlock Text="System &amp; Network Background Tasks" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                            <CheckBox Name="ChkBackgroundApps" Content="Disable Apps Running in Background" ToolTip="Stops UWP (Store) apps from running, updating, and sending notifications in the background while closed." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkNetworkThrottling" Content="Disable Network Throttling for Gaming" ToolTip="Removes the default limit on network packet processing, improving latency and ping for online games." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                            <CheckBox Name="ChkStorageSense" Content="Enable Storage Sense (Auto-cleanup)" ToolTip="Automatically cleans up temporary files and the recycle bin when disk space runs low." Margin="0,0,0,8" Foreground="{StaticResource TextForeground}"/>
                                        </StackPanel>
                                    </Border>
                                </StackPanel>
                            </ScrollViewer>
                        </TabItem>
                    </TabControl>
                    
                    <!-- Action Buttons -->
                    <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                        <CheckBox Name="CreateRestorePointOptimizeCheck" Content="Create Restore Point" IsChecked="True" VerticalAlignment="Center" Margin="0,0,15,0" Foreground="{StaticResource TextForeground}"/>
                        <Button Name="RefreshOptimizeBtn" Content="Refresh Status" Padding="15,8" Margin="0,0,10,0"/>
                        <Button Name="ApplyOptimizeBtn" Content="Apply Optimizations" Padding="15,8" Width="200" FontWeight="Bold" Foreground="#FF6B6B"/>
                    </StackPanel>
                </Grid>
            </TabItem>

            <!-- Utilities Tab -->
            <TabItem Header="Utilities">
                <TabItem.Resources>
                    <!-- Make all utility buttons pop more by defining a custom ControlTemplate for them -->
                    <Style TargetType="Button">
                        <Setter Property="Background" Value="#3E3E42"/>
                        <Setter Property="Foreground" Value="{StaticResource TextForeground}"/>
                        <Setter Property="BorderBrush" Value="#555555"/>
                        <Setter Property="BorderThickness" Value="1"/>
                        <Setter Property="FontWeight" Value="SemiBold"/>
                        <Setter Property="Cursor" Value="Hand"/>
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="Button">
                                    <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="3">
                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter Property="Background" Value="#54545C"/>
                                        </Trigger>
                                        <Trigger Property="IsPressed" Value="True">
                                            <Setter Property="Background" Value="{StaticResource AccentColor}"/>
                                            <Setter Property="Foreground" Value="White"/>
                                        </Trigger>
                                        <Trigger Property="IsEnabled" Value="False">
                                            <Setter Property="Background" Value="#2D2D30"/>
                                            <Setter Property="Foreground" Value="#666666"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                </TabItem.Resources>
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
                                        <Button Name="UtilLongPathBtn" Content="Enable Long Paths" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilDisableLongPathBtn" Content="Disable Long Paths" Padding="10,8" Margin="0,0,10,10"/>
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
                                        <Button Name="UtilInstallChocoBtn" Content="Install Chocolatey" Padding="10,8" Margin="0,0,10,10"/>
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
                                    <TextBlock Text="Fix network connectivity issues or enable/disable sharing protocols." TextWrapping="Wrap" Foreground="#AAAAAA" Margin="0,0,0,10" Height="35"/>
                                    <WrapPanel>
                                        <Button Name="UtilResetNetBtn" Content="Reset Network Adapters" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilSMBBtn" Content="Enable SMBv1" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilDisableSMBBtn" Content="Disable SMBv1" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilFileSharingBtn" Content="Enable File &amp; Printer Sharing" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilDisableFileSharingBtn" Content="Disable File &amp; Printer Sharing" Padding="10,8" Margin="0,0,10,10"/>
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

                            <!-- Box 7 -->
                            <Border Background="{StaticResource ControlBackground}" BorderBrush="{StaticResource ControlBorder}" BorderThickness="1" CornerRadius="5" Padding="15" Margin="0,0,10,15">
                                <StackPanel>
                                    <TextBlock Text="ODBC &amp; Database Tools" FontWeight="SemiBold" FontSize="14" Margin="0,0,0,10"/>
                                    <TextBlock Text="Backup and restore 32-bit and 64-bit ODBC Data Sources (System &amp; User DSNs)." TextWrapping="Wrap" Foreground="#AAAAAA" Margin="0,0,0,10" Height="35"/>
                                    <WrapPanel>
                                        <Button Name="UtilExportODBCBtn" Content="Export ODBC Data Sources" Padding="10,8" Margin="0,0,10,10"/>
                                        <Button Name="UtilImportODBCBtn" Content="Import ODBC Data Sources" Padding="10,8" Margin="0,0,10,10"/>
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
            <StatusBarItem HorizontalAlignment="Right" HorizontalContentAlignment="Right">
                <TextBlock Name="NetworkStatusText" Text="" FontWeight="Bold" Margin="15,0,10,0"/>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
"@

# 2. Load the XAML into PowerShell
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$Window = [Windows.Markup.XamlReader]::Load($reader)

# Map XAML elements to PowerShell variables
$MainTabs = $Window.FindName("MainTabs")
$DiscoverTab = $Window.FindName("DiscoverTab")
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

# Fetch all template checkboxes programmatically
$TemplateNames = @(
    "TplChrome", "TplFirefox", "TplEdge", "TplBrave", "TplVivaldi",
    "TplLibreOffice", "TplOnlyOffice", "TplOpenOffice", "TplPDF24",
    "TplZoom", "TplDiscord", "TplTeams", "TplThunderbird",
    "Tpl7zip", "TplPeaZip", "TplWinRAR",
    "TpliTunes", "TplVLC", "TplWinamp", "TplAudacity", "TplKLite", "TplHandBrake",
    "TplKrita", "TplPaintNet", "TplGimp", "TplInkscape", "TplFastStone",
    "TplEverything", "TplAnyDesk", "TplTightVNC", "TplWizTree", "TplQBittorrent", "TplLocalSend", "TplFileZilla", "TplWinSCP", "TplHeidiSQL", "TplDBeaver", "TplVSCode", "TplPython", "TplGit",
    "TplSteam", "TplEpic", "TplGOG"
)

$TemplateCheckboxes = @()
foreach ($name in $TemplateNames) {
    $chk = $Window.FindName($name)
    if ($chk) { $TemplateCheckboxes += $chk }
}

# Utilities UI Map
$UtilAdminWarning = $Window.FindName("UtilAdminWarning")
$UtilSysScanBtn = $Window.FindName("UtilSysScanBtn")
$UtilResetWUBtn = $Window.FindName("UtilResetWUBtn")
$UtilRestorePointBtn = $Window.FindName("UtilRestorePointBtn")
$UtilOpenRestoreBtn = $Window.FindName("UtilOpenRestoreBtn")
$UtilLongPathBtn = $Window.FindName("UtilLongPathBtn")
$UtilDisableLongPathBtn = $Window.FindName("UtilDisableLongPathBtn")
$UtilResetNetBtn = $Window.FindName("UtilResetNetBtn")
$UtilSMBBtn = $Window.FindName("UtilSMBBtn")
$UtilDisableSMBBtn = $Window.FindName("UtilDisableSMBBtn")
$UtilFileSharingBtn = $Window.FindName("UtilFileSharingBtn")
$UtilDisableFileSharingBtn = $Window.FindName("UtilDisableFileSharingBtn")
$UtilDriverBtn = $Window.FindName("UtilDriverBtn")
$UtilSDIOBtn = $Window.FindName("UtilSDIOBtn")
$UtilWingetRepairBtn = $Window.FindName("UtilWingetRepairBtn")
$UtilStoreRepairBtn = $Window.FindName("UtilStoreRepairBtn")
$UtilInstallChocoBtn = $Window.FindName("UtilInstallChocoBtn")
$UtilDiskCleanupBtn = $Window.FindName("UtilDiskCleanupBtn")
$UtilClearLogsBtn = $Window.FindName("UtilClearLogsBtn")
$UtilIconCacheBtn = $Window.FindName("UtilIconCacheBtn")
$UtilExportODBCBtn = $Window.FindName("UtilExportODBCBtn")
$UtilImportODBCBtn = $Window.FindName("UtilImportODBCBtn")

# Privacy & Optimize UI Map
$PrivacyAdminWarning = $Window.FindName("PrivacyAdminWarning")
$PerfAdminWarning = $Window.FindName("PerfAdminWarning")

$BtnPrivacyMinimal = $Window.FindName("BtnPrivacyMinimal")
$BtnPrivacyStandard = $Window.FindName("BtnPrivacyStandard")
$BtnPrivacyMaximal = $Window.FindName("BtnPrivacyMaximal")

# Customization Checkboxes
$ChkDarkSystem = $Window.FindName("ChkDarkSystem")
$ChkDarkApps = $Window.FindName("ChkDarkApps")
$ChkDisableTransparency = $Window.FindName("ChkDisableTransparency")
$ChkTaskbarAccent = $Window.FindName("ChkTaskbarAccent")
$ChkTitlebarAccent = $Window.FindName("ChkTitlebarAccent")
$ChkTaskbarLeft = $Window.FindName("ChkTaskbarLeft")
$ChkHideSearch = $Window.FindName("ChkHideSearch")
$ChkHideTaskView = $Window.FindName("ChkHideTaskView")
$ChkHideChat = $Window.FindName("ChkHideChat")
$ChkStartMorePins = $Window.FindName("ChkStartMorePins")
$ChkStartHideRecentApps = $Window.FindName("ChkStartHideRecentApps")
$ChkStartHideMostUsed = $Window.FindName("ChkStartHideMostUsed")
$ChkStartHideRecentDocs = $Window.FindName("ChkStartHideRecentDocs")
$ChkExpHidden = $Window.FindName("ChkExpHidden")
$ChkExpExt = $Window.FindName("ChkExpExt")
$ChkExpThisPC = $Window.FindName("ChkExpThisPC")
$ChkExpCompact = $Window.FindName("ChkExpCompact")
$ChkExpClassicMenu = $Window.FindName("ChkExpClassicMenu")

# Privacy Checkboxes
$ChkTelemetry = $Window.FindName("ChkTelemetry")
$ChkLocation = $Window.FindName("ChkLocation")
$ChkActivity = $Window.FindName("ChkActivity")
$ChkAppLaunch = $Window.FindName("ChkAppLaunch")
$ChkLanguage = $Window.FindName("ChkLanguage")
$ChkSpeech = $Window.FindName("ChkSpeech")
$ChkTailoredExp = $Window.FindName("ChkTailoredExp")
$ChkWER = $Window.FindName("ChkWER")
$ChkFeedback = $Window.FindName("ChkFeedback")
$ChkWorkplace = $Window.FindName("ChkWorkplace")
$ChkOneDrive = $Window.FindName("ChkOneDrive")
$ChkBingSearch = $Window.FindName("ChkBingSearch")
$ChkStartAds = $Window.FindName("ChkStartAds")
$ChkLockScreenAds = $Window.FindName("ChkLockScreenAds")
$ChkExplorerAds = $Window.FindName("ChkExplorerAds")
$ChkWelcomeExp = $Window.FindName("ChkWelcomeExp")
$ChkAdId = $Window.FindName("ChkAdId")
$ChkBitLocker = $Window.FindName("ChkBitLocker")
$ChkCopilot = $Window.FindName("ChkCopilot")
$ChkWidgets = $Window.FindName("ChkWidgets")
$ChkMaintenance = $Window.FindName("ChkMaintenance")
$ChkConsumer = $Window.FindName("ChkConsumer")
$ChkWifiSense = $Window.FindName("ChkWifiSense")
$ChkRemoteAssist = $Window.FindName("ChkRemoteAssist")

# O&O Privacy Checkboxes
$ChkAppCamera = $Window.FindName("ChkAppCamera")
$ChkAppMic = $Window.FindName("ChkAppMic")
$ChkAppAccountInfo = $Window.FindName("ChkAppAccountInfo")
$ChkAppContacts = $Window.FindName("ChkAppContacts")
$ChkAppCalendar = $Window.FindName("ChkAppCalendar")
$ChkAppEmail = $Window.FindName("ChkAppEmail")
$ChkAppCallHistory = $Window.FindName("ChkAppCallHistory")
$ChkAppTasks = $Window.FindName("ChkAppTasks")
$ChkAppMessages = $Window.FindName("ChkAppMessages")
$ChkDefenderTelemetry = $Window.FindName("ChkDefenderTelemetry")
$ChkSmartScreen = $Window.FindName("ChkSmartScreen")
$ChkP2PUpdate = $Window.FindName("ChkP2PUpdate")
$ChkCortana = $Window.FindName("ChkCortana")
$ChkCEIP = $Window.FindName("ChkCEIP")
$ChkHandwriting = $Window.FindName("ChkHandwriting")
$ChkAutoDriverUpdate = $Window.FindName("ChkAutoDriverUpdate")
$ChkEdgeTelemetry = $Window.FindName("ChkEdgeTelemetry")
$ChkEdgeCopilot = $Window.FindName("ChkEdgeCopilot")
$ChkEdgeSearchAds = $Window.FindName("ChkEdgeSearchAds")

# Performance Checkboxes
$ChkGameMode = $Window.FindName("ChkGameMode")
$ChkCpuPriority = $Window.FindName("ChkCpuPriority")
$ChkHwSchMode = $Window.FindName("ChkHwSchMode")
$ChkFSO = $Window.FindName("ChkFSO")
$ChkGameDVR = $Window.FindName("ChkGameDVR")
$ChkMouseAccel = $Window.FindName("ChkMouseAccel")
$ChkMenuDelay = $Window.FindName("ChkMenuDelay")
$ChkStartupDelay = $Window.FindName("ChkStartupDelay")
$ChkBackgroundApps = $Window.FindName("ChkBackgroundApps")
$ChkNetworkThrottling = $Window.FindName("ChkNetworkThrottling")
$ChkStorageSense = $Window.FindName("ChkStorageSense")

$RefreshOptimizeBtn = $Window.FindName("RefreshOptimizeBtn")
$ApplyOptimizeBtn = $Window.FindName("ApplyOptimizeBtn")
$CreateRestorePointOptimizeCheck = $Window.FindName("CreateRestorePointOptimizeCheck")

$CreateRestorePointUpdateCheck = $Window.FindName("CreateRestorePointUpdateCheck")

# --- Privacy Templates Logic ---
$AllPrivacyChecks = @(
    $ChkTelemetry, $ChkLocation, $ChkActivity, $ChkAppLaunch, $ChkLanguage, $ChkSpeech, $ChkTailoredExp, 
    $ChkWER, $ChkFeedback, $ChkWorkplace, $ChkOneDrive, $ChkBingSearch, $ChkStartAds, $ChkLockScreenAds, 
    $ChkExplorerAds, $ChkWelcomeExp, $ChkAdId, $ChkBitLocker, $ChkCopilot, $ChkWidgets, $ChkMaintenance, 
    $ChkConsumer, $ChkWifiSense, $ChkRemoteAssist,
    $ChkAppCamera, $ChkAppMic, $ChkAppAccountInfo, $ChkAppContacts, $ChkAppCalendar, $ChkAppEmail, 
    $ChkAppCallHistory, $ChkAppTasks, $ChkAppMessages,
    $ChkDefenderTelemetry, $ChkSmartScreen, $ChkP2PUpdate, $ChkCortana, $ChkCEIP, $ChkHandwriting, 
    $ChkAutoDriverUpdate, $ChkEdgeTelemetry, $ChkEdgeCopilot, $ChkEdgeSearchAds
)

$BtnPrivacyMinimal.Add_Click({
    $MinimalChecks = @(
        $ChkTelemetry, $ChkTailoredExp, $ChkStartAds, $ChkLockScreenAds, $ChkExplorerAds, $ChkWelcomeExp, 
        $ChkAdId, $ChkConsumer, $ChkCEIP, $ChkEdgeTelemetry, $ChkEdgeSearchAds
    )
    foreach ($chk in $AllPrivacyChecks) { $chk.IsChecked = $false }
    foreach ($chk in $MinimalChecks) { $chk.IsChecked = $true }
    $StatusText.Text = "Minimal privacy template selected. Click 'Apply Optimizations' to save."
})

$BtnPrivacyStandard.Add_Click({
    $StandardChecks = @(
        $ChkTelemetry, $ChkTailoredExp, $ChkStartAds, $ChkLockScreenAds, $ChkExplorerAds, $ChkWelcomeExp, 
        $ChkAdId, $ChkConsumer, $ChkCEIP, $ChkEdgeTelemetry, $ChkEdgeSearchAds,
        $ChkLocation, $ChkActivity, $ChkAppLaunch, $ChkLanguage, $ChkSpeech, $ChkWER, $ChkFeedback, 
        $ChkBingSearch, $ChkWidgets, $ChkWifiSense, $ChkP2PUpdate, $ChkCortana, $ChkHandwriting
    )
    foreach ($chk in $AllPrivacyChecks) { $chk.IsChecked = $false }
    foreach ($chk in $StandardChecks) { $chk.IsChecked = $true }
    $StatusText.Text = "Standard privacy template selected. Click 'Apply Optimizations' to save."
})

$BtnPrivacyMaximal.Add_Click({
    foreach ($chk in $AllPrivacyChecks) { $chk.IsChecked = $true }
    $StatusText.Text = "Maximal privacy template selected. Click 'Apply Optimizations' to save."
})

# --- Apply Admin Status to UI ---
if ($isActualAdmin) {
    $Window.Title = "WinToolsUI (Administrator)"
    $AdminInstallCheck.IsChecked = $true
    $AdminInstallCheck.IsEnabled = $true
    
    $CreateRestorePointInstallCheck.IsEnabled = $true
    $CreateRestorePointUpdateCheck.IsEnabled = $true
    $CreateRestorePointOptimizeCheck.IsEnabled = $true
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
    
    $CreateRestorePointOptimizeCheck.IsEnabled = $false
    $CreateRestorePointOptimizeCheck.ToolTip = "Administrator privileges are required to create Restore Points."
    $CreateRestorePointOptimizeCheck.Foreground = "#888888"
    
    # Disable Utilities if not admin
    $UtilAdminWarning.Visibility = 'Visible'
    $UtilSysScanBtn.IsEnabled = $false
    $UtilResetWUBtn.IsEnabled = $false
    $UtilRestorePointBtn.IsEnabled = $false
    $UtilOpenRestoreBtn.IsEnabled = $false
    $UtilLongPathBtn.IsEnabled = $false
    $UtilDisableLongPathBtn.IsEnabled = $false
    $UtilResetNetBtn.IsEnabled = $false
    $UtilSMBBtn.IsEnabled = $false
    $UtilDisableSMBBtn.IsEnabled = $false
    $UtilFileSharingBtn.IsEnabled = $false
    $UtilDisableFileSharingBtn.IsEnabled = $false
    $UtilDriverBtn.IsEnabled = $false
    $UtilSDIOBtn.IsEnabled = $false
    $UtilWingetRepairBtn.IsEnabled = $false
    $UtilStoreRepairBtn.IsEnabled = $false
    $UtilInstallChocoBtn.IsEnabled = $false
    $UtilDiskCleanupBtn.IsEnabled = $false
    $UtilClearLogsBtn.IsEnabled = $false
    $UtilIconCacheBtn.IsEnabled = $false
    $UtilExportODBCBtn.IsEnabled = $false
    $UtilImportODBCBtn.IsEnabled = $false
    
    # Disable Optimize (Privacy/Performance) if not admin
    $PrivacyAdminWarning.Visibility = 'Visible'
    $PerfAdminWarning.Visibility = 'Visible'
    $ApplyOptimizeBtn.IsEnabled = $false
}

# --- Apply Winget Missing Status to UI ---
if ($script:WingetMissing) {
    # Disable the Discover tab completely
    $DiscoverTab.IsEnabled = $false
    $DiscoverTab.Header = "Discover (Winget Required)"
    $DiscoverTab.ToolTip = "Winget is not installed on this system."
    
    # Switch the starting active tab to the 'Installed & Updates' tab
    $MainTabs.SelectedIndex = 1 
    
    # Disable the Winget repair utility button
    $UtilWingetRepairBtn.IsEnabled = $false
}

$RefreshInstalledBtn = $Window.FindName("RefreshInstalledBtn")
$SelectAllInstalledBtn = $Window.FindName("SelectAllInstalledBtn")
$DeselectAllInstalledBtn = $Window.FindName("DeselectAllInstalledBtn")
$UninstallBtn = $Window.FindName("UninstallBtn")
$UpdateBtn = $Window.FindName("UpdateBtn")
$UpdateAllBtn = $Window.FindName("UpdateAllBtn")
$SearchInstalledBox = $Window.FindName("SearchInstalledBox")
$InstalledGrid = $Window.FindName("InstalledGrid")
$InstalledMenuDetails = $Window.FindName("InstalledMenuDetails")
$InstalledMenuCopyId = $Window.FindName("InstalledMenuCopyId")
$InstalledMenuUninstall = $Window.FindName("InstalledMenuUninstall")

$WindowsAppsGrid = $Window.FindName("WindowsAppsGrid")
$WindowsAppsMenuDetails = $Window.FindName("WindowsAppsMenuDetails")
$WindowsAppsMenuCopyId = $Window.FindName("WindowsAppsMenuCopyId")
$WindowsAppsMenuUninstall = $Window.FindName("WindowsAppsMenuUninstall")

$LogExpander = $Window.FindName("LogExpander")
$LogTextBox = $Window.FindName("LogTextBox")

$JobProgress = $Window.FindName("JobProgress")
$StopJobBtn = $Window.FindName("StopJobBtn")
$StatusText = $Window.FindName("StatusText")
$NetworkStatusText = $Window.FindName("NetworkStatusText")

# --- Initialize the Observable Queue ---
$script:InstallQueue = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$QueueGrid.ItemsSource = $script:InstallQueue

# --- Bind Quick Templates Checkboxes to Queue ---
foreach ($chk in $TemplateCheckboxes) {
    # When checked, add to queue
    $chk.Add_Checked({
        $source = $this
        $tagData = $source.Tag -split '\|'
        $mgr = $tagData[0]
        $id = $tagData[1]
        $name = $source.Content

        $exists = $false
        foreach ($q in $script:InstallQueue) {
            if ($q.Id -eq $id) { $exists = $true; break }
        }
        if (-not $exists) {
            $script:InstallQueue.Add([PSCustomObject]@{Name=$name; Id=$id; Version="Latest"; Manager=$mgr})
        }
    })

    # When unchecked, remove from queue
    $chk.Add_Unchecked({
        $source = $this
        $tagData = $source.Tag -split '\|'
        $id = $tagData[1]

        $toRemove = $null
        foreach ($q in $script:InstallQueue) {
            if ($q.Id -eq $id) {
                $toRemove = $q
                break
            }
        }
        if ($toRemove -ne $null) {
            $script:InstallQueue.Remove($toRemove) | Out-Null
        }
    })
}

# 3. Setup the Background Runspace (The Backend Engine)
$syncHash = [hashtable]::Synchronized(@{})
$syncHash.LogQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()
$syncHash.AppPath = if ($PSCommandPath) { Split-Path -Parent $PSCommandPath } else { $PWD.Path }
$syncHash.IsOnline = $null

$runspace = [runspacefactory]::CreateRunspace()
$runspace.Open()
$runspace.SessionStateProxy.SetVariable("syncHash", $syncHash)

$script:psInstance = $null
$script:asyncResult = $null
$script:IsJobRunning = $false
$script:AllInstalledApps = $null

# 4. Define the Universal Background Job
$bgJobBlock = {
    param($Action, $Query, $Id, $Hash, $IsAdmin, $CreateRestore, $Manager)
    
    $ProgressPreference = 'SilentlyContinue' # Hide default PS progress bar
    # Force PowerShell to read external Winget output using UTF-8 to prevent 'ΓÇª' encoding issues
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
    
    # --- Safe logger to filter out progress percentage spam ---
    function Write-FilteredLog {
        param($raw)
        $str = $raw.ToString()
        
        # 1. Strip ANSI escape sequences
        $str = $str -replace "\x1B\[[0-9;]*[a-zA-Z]", ""
        
        # 2. Extract Percentage and convert to a clean, single-line UI element
        if ($str -match '(?<pct>\d{1,3})(?:\.\d+)?\s*%') {
            $pct = [int]$matches['pct']
            $Hash.Progress = $pct
            
            # Format a clean progress bar string
            $barCount = [math]::Floor($pct / 5)
            if ($barCount -gt 20) { $barCount = 20 }
            if ($barCount -lt 0) { $barCount = 0 }
            
            # Safely generate blocks using Unicode to prevent encoding parser errors
            $bars = ([char]0x2588).ToString() * $barCount
            $spaces = ([char]0x2592).ToString() * (20 - $barCount)
            
            # Send with a special [PROGRESS] tag so the UI knows to overwrite the line
            $Hash.LogQueue.Enqueue("[PROGRESS] Progress: [$bars$spaces] $pct%")
            return 
        }

        # 2.5 Catch Winget's raw MB/MB download output (Calculates percentage dynamically)
        if ($str -match '(?<curr>\d+(?:\.\d+)?)\s*[A-Z]B\s*/\s*(?<tot>\d+(?:\.\d+)?)\s*[A-Z]B') {
            try {
                $c = [double]$matches['curr']
                $t = [double]$matches['tot']
                if ($t -gt 0) {
                    $pct = [int][math]::Round(($c / $t) * 100)
                    $Hash.Progress = $pct
                    
                    $barCount = [math]::Floor($pct / 5)
                    if ($barCount -gt 20) { $barCount = 20 }
                    if ($barCount -lt 0) { $barCount = 0 }
                    
                    # Safely generate blocks using Unicode to prevent encoding parser errors
                    $bars = ([char]0x2588).ToString() * $barCount
                    $spaces = ([char]0x2592).ToString() * (20 - $barCount)
                    
                    $Hash.LogQueue.Enqueue("[PROGRESS] Downloading: [$bars$spaces] $pct% ($c / $t)")
                }
            } catch {}
            return
        }

        # 3. Strip Control Characters (Backspaces, Carriage Returns)
        $str = $str -replace "[\b\r]", ""
        $str = $str.Trim()
        
        if ([string]::IsNullOrWhiteSpace($str)) { return }
        
        # 4. Ignore isolated spinner characters and stray raw blocks
        if ($str -in @('\', '|', '/', '-')) { return }
        
        # Safely match block characters using regex Unicode escapes
        if ($str -match '[\u2588\u2591\u2592\u2593]') { return }
        
        $Hash.LogQueue.Enqueue($str)
    }
    
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
                    $val = $val -replace '(\u0393\u00C7\u00AA)+$', '' # Removes OEM ellipsis artifact
                    $val = $val -replace '(\u0393\u00C7\u00F6)+$', '' # Removes OEM dash artifact
                    $val = $val -replace '\u2026+$', ''             # Removes standard ellipsis
                    
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
                    Manager   = "Winget"
                    IsSelected = $false
                }
            }
        }
        return $parsed
    }
    
    # Parser for Chocolatey output
    function ConvertFrom-ChocoOutput($raw) {
        $parsed = @()
        foreach ($line in $raw) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            # choco list -l -r output: name|version
            # choco search -r output: name|version|...
            $parts = $line -split '\|'
            if ($parts.Count -ge 2) {
                $parsed += [PSCustomObject]@{
                    Name      = $parts[0]
                    Id        = $parts[0]
                    Version   = $parts[1]
                    Extra     = ""
                    Source    = "chocolatey"
                    HasUpdate = $false
                    Manager   = "Chocolatey"
                    IsSelected = $false
                }
            }
        }
        return $parsed
    }
    
    # Parser for Chocolatey Outdated
    function ConvertFrom-ChocoOutdated($raw) {
        # choco outdated -r output: Name|Current Version|Available Version|Pinned
        $updates = @{}
        foreach ($line in $raw) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            if ($line -notmatch '\|') { continue } # Ignore non-tabular lines (banners/warnings)
            $parts = $line -split '\|'
            if ($parts.Count -ge 3) {
                # Trim whitespace to handle non-semantic or padded version strings safely
                $updates[$parts[0].Trim()] = $parts[2].Trim()
            }
        }
        return $updates
    }

    try {
        switch ($Action) {
            'Search' {
                $combinedResults = @()
                
                # 1. Winget Search
                $sysLocale = (Get-Culture).Name                     
                $langOnly = (Get-Culture).TwoLetterISOLanguageName   
                
                $rawWinget = @()
                $Hash.LogQueue.Enqueue(">>> Executing: winget search ""$Query""")
                $wingetArgs = @("search", $Query, "--count", "40", "--accept-source-agreements", "--disable-interactivity")
                & winget @wingetArgs 2>&1 | ForEach-Object {
                    $line = $_.ToString()
                    Write-FilteredLog $line
                    $rawWinget += $line
                }
                
                $parsedWinget = ConvertFrom-WingetOutput $rawWinget
                
                foreach ($item in $parsedWinget) {
                    $hasLocaleSuffix = $item.Id -match '\.([a-z]{2}-[A-Z]{2}|[a-z]{2})$'
                    if (-not $hasLocaleSuffix) {
                        $combinedResults += $item
                    } elseif ($item.Id -match "\.($sysLocale|$langOnly|en-US|en-GB|en)$") {
                        $combinedResults += $item
                    }
                }
                
                # 2. Chocolatey Search (if available)
                if (Get-Command choco -ErrorAction SilentlyContinue) {
                    $Hash.LogQueue.Enqueue(">>> Executing: choco search $Query")
                    $rawChoco = @()
                    & choco search $Query -r 2>&1 | ForEach-Object {
                        $line = $_.ToString()
                        $rawChoco += $line
                        Write-FilteredLog $line
                    }
                    
                    # Parse the raw output and add it to the combined results
                    $parsedChoco = ConvertFrom-ChocoOutput $rawChoco
                    $combinedResults += $parsedChoco
                }
                
                $Hash.Result = $combinedResults
            }
            'Installed' {
                $combinedResults = @()
                
                # 1. Fetch Chocolatey Packages (if available)
                if (Get-Command choco -ErrorAction SilentlyContinue) {
                    $Hash.LogQueue.Enqueue(">>> Executing: choco list -l")
                    $rawChoco = @()
                    & choco list -l -r 2>&1 | ForEach-Object {
                        $line = $_.ToString()
                        $rawChoco += $line
                    }
                    $parsedChoco = ConvertFrom-ChocoOutput $rawChoco
                    
                    $Hash.LogQueue.Enqueue(">>> Checking for outdated packages...")
                    $rawOutdated = @()
                    & choco outdated -r 2>&1 | ForEach-Object {
                        $line = $_.ToString()
                        $rawOutdated += $line
                    }
                    $updates = ConvertFrom-ChocoOutdated $rawOutdated
                    
                    # Merge updates
                    foreach ($p in $parsedChoco) {
                        if ($updates.ContainsKey($p.Id)) {
                            $p.HasUpdate = $true
                            $p.Extra = $updates[$p.Id]
                        }
                    }
                    $combinedResults += $parsedChoco
                }
                
                # 2. Fetch Winget Packages (or Fallback to Native)
                if (Get-Command winget -ErrorAction SilentlyContinue) {
                    $rawWinget = @()
                    $Hash.LogQueue.Enqueue(">>> Executing: winget list")
                    $wingetArgs = @("list", "--accept-source-agreements", "--disable-interactivity")
                    & winget @wingetArgs 2>&1 | ForEach-Object {
                        $line = $_.ToString()
                        Write-FilteredLog $line
                        $rawWinget += $line
                    }
                    $parsedWinget = ConvertFrom-WingetOutput $rawWinget
                    
                    # Add Winget packages (filtering out duplicates that claim to be from chocolatey source to avoid double listing)
                    foreach ($p in $parsedWinget) {
                        if ($p.Source -ne 'chocolatey') {
                            $combinedResults += $p
                        }
                    }
                } else {
                    $Hash.LogQueue.Enqueue(">>> Winget not found. Loading installed apps natively via Registry & Appx...")
                    
                    # Native Fallback 1: Desktop Apps via Registry
                    $regPaths = @(
                        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
                        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
                    )
                    
                    $nativeApps = @()
                    foreach ($path in $regPaths) {
                        Get-ItemProperty $path -ErrorAction SilentlyContinue | Where-Object { -not [string]::IsNullOrWhiteSpace($_.DisplayName) } | ForEach-Object {
                            $uString = if ($_.QuietUninstallString) { $_.QuietUninstallString } else { $_.UninstallString }
                            if (-not [string]::IsNullOrWhiteSpace($uString)) {
                                $nativeApps += [PSCustomObject]@{
                                    Name      = $_.DisplayName
                                    Id        = $uString # We store the uninstall command in the ID field to use it later
                                    Version   = if ($_.DisplayVersion) { $_.DisplayVersion } else { "Unknown" }
                                    Extra     = ""
                                    Source    = "Registry (Native)"
                                    HasUpdate = $false
                                    Manager   = "Native"
                                    IsSelected = $false
                                }
                            }
                        }
                    }
                    # Deduplicate native registry apps by Name to prevent visual clutter
                    $combinedResults += @($nativeApps | Group-Object Name | ForEach-Object { $_.Group[0] })
                    
                    # Native Fallback 2: Store Apps via AppxPackage
                    if ($IsAdmin) {
                        $appxPackages = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue
                    } else {
                        $appxPackages = Get-AppxPackage -ErrorAction SilentlyContinue
                    }
                    
                    $appxPackages | Where-Object { $_.IsFramework -eq $false -and $_.NonRemovable -eq $false } | ForEach-Object {
                        $combinedResults += [PSCustomObject]@{
                            Name      = $_.Name
                            Id        = $_.PackageFullName
                            Version   = $_.Version
                            Extra     = ""
                            Source    = "Appx (Native)"
                            HasUpdate = $false
                            Manager   = "NativeAppx"
                            IsSelected = $false
                        }
                    }
                }
                
                $Hash.Result = $combinedResults
            }
            'Install' {
                if ($CreateRestore) {
                    $Hash.LogQueue.Enqueue(">>> Creating System Restore Point before installation...")
                    New-BypassRestorePoint -Desc "WinToolsUI Install" | Out-Null
                }

                # --- Auto-Install Chocolatey if needed by the Queue ---
                $needsChoco = $false
                foreach ($item in @($Id)) {
                    $mgr = if ($item -is [string]) { $Manager } else { $item.Manager }
                    if ($mgr -eq 'Chocolatey') { $needsChoco = $true; break }
                }

                if ($needsChoco -and -not (Get-Command choco -ErrorAction SilentlyContinue)) {
                    $Hash.LogQueue.Enqueue("`r`n>>> Chocolatey is required for queued packages but is not installed.")
                    $Hash.LogQueue.Enqueue(">>> Auto-installing Chocolatey Package Manager first...")
                    try {
                        Set-ExecutionPolicy Bypass -Scope Process -Force
                        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
                        $installCmd = "iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))"
                        
                        Invoke-Expression $installCmd | Out-String -Stream | ForEach-Object { Write-FilteredLog $_ }
                        
                        # Refresh environment path so the 'choco' command is immediately available to this background runspace
                        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
                        
                        $Hash.LogQueue.Enqueue("[+] Chocolatey installed successfully. Continuing with package queue...")
                    } catch {
                        $Hash.LogQueue.Enqueue("[-] Failed to install Chocolatey: $($_.Exception.Message)")
                    }
                }
                # ------------------------------------------------------

                foreach ($item in @($Id)) {
                    if ($item -is [string]) { $targetId = $item; $mgr = $Manager } else { $targetId = $item.Id; $mgr = $item.Manager }
                    
                    if ($mgr -eq 'Chocolatey') {
                        $Hash.LogQueue.Enqueue("`r`n>>> Executing: choco install $targetId")
                        & choco install $targetId -y 2>&1 | ForEach-Object { Write-FilteredLog $_ }
                    } else {
                        $Hash.LogQueue.Enqueue("`r`n>>> Executing: winget install $targetId")
                        $wingetArgs = @("install", "--id", $targetId, "--exact", "--accept-source-agreements", "--accept-package-agreements", "--silent", "--disable-interactivity")
                        if ($IsAdmin) { $wingetArgs += "--scope"; $wingetArgs += "machine" }
                        
                        & winget @wingetArgs 2>&1 | ForEach-Object { Write-FilteredLog $_ }
                    }
                }
                $Hash.Result = "Success"
            }
            'Uninstall' {
                if ($CreateRestore) {
                    $Hash.LogQueue.Enqueue(">>> Creating System Restore Point before uninstallation...")
                    New-BypassRestorePoint -Desc "WinToolsUI Uninstall" | Out-Null
                }
                foreach ($item in @($Id)) {
                    if ($item -is [string]) { $targetId = $item; $mgr = $Manager } else { $targetId = $item.Id; $mgr = $item.Manager }
                    
                    if ($mgr -eq 'Chocolatey') {
                        $Hash.LogQueue.Enqueue("`r`n>>> Executing: choco uninstall $targetId")
                        & choco uninstall $targetId -y 2>&1 | ForEach-Object { Write-FilteredLog $_ }
                    } elseif ($mgr -eq 'Native') {
                        $Hash.LogQueue.Enqueue("`r`n>>> Executing Native Uninstaller...")
                        $Hash.LogQueue.Enqueue("Command: $targetId")
                        try {
                            Start-Process cmd.exe -ArgumentList "/c", "$targetId" -Wait -WindowStyle Hidden -ErrorAction Stop
                            $Hash.LogQueue.Enqueue("Uninstall triggered successfully. (Note: Some installers may pop up in the background).")
                        } catch {
                            $Hash.LogQueue.Enqueue("Error triggering uninstall: $($_.Exception.Message)")
                        }
                    } elseif ($mgr -eq 'NativeAppx') {
                        $Hash.LogQueue.Enqueue("`r`n>>> Removing Appx Package: $targetId")
                        try {
                            if ($IsAdmin) { Remove-AppxPackage -Package $targetId -AllUsers -ErrorAction Stop } 
                            else { Remove-AppxPackage -Package $targetId -ErrorAction Stop }
                            $Hash.LogQueue.Enqueue("Appx Package removed successfully.")
                        } catch {
                            $Hash.LogQueue.Enqueue("Error: $($_.Exception.Message)")
                        }
                    } else {
                        $Hash.LogQueue.Enqueue("`r`n>>> Executing: winget uninstall $targetId")
                        $wingetArgs = @("uninstall", "--id", $targetId, "--exact", "--silent", "--disable-interactivity")
                        & winget @wingetArgs 2>&1 | ForEach-Object { Write-FilteredLog $_ }
                    }
                }
                $Hash.Result = "Success"
            }
            'Update' {
                if ($CreateRestore) {
                    $Hash.LogQueue.Enqueue(">>> Creating System Restore Point before update...")
                    New-BypassRestorePoint -Desc "WinToolsUI Update" | Out-Null
                }
                foreach ($item in @($Id)) {
                    if ($item -is [string]) { $targetId = $item; $mgr = $Manager } else { $targetId = $item.Id; $mgr = $item.Manager }
                    
                    if ($mgr -eq 'Chocolatey') {
                        $Hash.LogQueue.Enqueue("`r`n>>> Executing: choco upgrade $targetId")
                        & choco upgrade $targetId -y 2>&1 | ForEach-Object { Write-FilteredLog $_ }
                    } else {
                        $Hash.LogQueue.Enqueue("`r`n>>> Executing: winget upgrade $targetId")
                        $wingetArgs = @("upgrade", "--id", $targetId, "--exact", "--accept-source-agreements", "--accept-package-agreements", "--silent", "--disable-interactivity")
                        & winget @wingetArgs 2>&1 | ForEach-Object { Write-FilteredLog $_ }
                    }
                }
                $Hash.Result = "Success"
            }
            'ShowDetails' {
                if ($Manager -eq 'Chocolatey') {
                    $Hash.LogQueue.Enqueue(">>> Executing: choco info $Id")
                    $raw = @()
                    & choco info $Id 2>&1 | ForEach-Object {
                        $line = $_.ToString()
                        $raw += $line
                        Write-FilteredLog $line
                    }
                    $Hash.Result = $raw -join "`r`n"
                } else {
                    $raw = @()
                    $Hash.LogQueue.Enqueue(">>> Executing: winget show $Id")
                    $wingetArgs = @("show", "--id", $Id, "--exact", "--accept-source-agreements", "--disable-interactivity")
                    & winget @wingetArgs 2>&1 | ForEach-Object {
                        $line = $_.ToString()
                        Write-FilteredLog $line
                        $raw += $line
                    }
                    $Hash.Result = $raw -join "`r`n"
                }
            }
            'UtilSystemScan' {
                $Hash.LogQueue.Enqueue(">>> Running DISM Component Store Cleanup and Repair...")
                $Hash.LogQueue.Enqueue(">>> (This may take several minutes to complete)")
                & DISM.exe /Online /Cleanup-image /Restorehealth 2>&1 | ForEach-Object { Write-FilteredLog $_ }
                
                $Hash.LogQueue.Enqueue("`r`n>>> Running System File Checker (SFC)...")
                & sfc /scannow 2>&1 | ForEach-Object { Write-FilteredLog $_ }
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
                & ipconfig /flushdns 2>&1 | ForEach-Object { Write-FilteredLog $_ }
                & ipconfig /renew 2>&1 | Out-Null
                
                $Hash.LogQueue.Enqueue(">>> Resetting Winsock and IP Configuration...")
                & netsh winsock reset 2>&1 | ForEach-Object { Write-FilteredLog $_ }
                & netsh int ip reset 2>&1 | ForEach-Object { Write-FilteredLog $_ }
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
            'UtilDisableSMB' {
                $Hash.LogQueue.Enqueue(">>> Disabling SMBv1 Protocol...")
                try {
                    Disable-WindowsOptionalFeature -Online -FeatureName SMB1Protocol -NoRestart -ErrorAction Stop | Out-Null
                    $Hash.LogQueue.Enqueue("Successfully disabled SMBv1. A system restart may be required.")
                    $Hash.Result = "Success"
                } catch {
                    $Hash.LogQueue.Enqueue("Error disabling SMBv1: $($_.Exception.Message)")
                    $Hash.Result = "Error: $($_.Exception.Message)"
                }
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
            'UtilDisableFileSharing' {
                $Hash.LogQueue.Enqueue(">>> Disabling File and Printer Sharing rules...")
                try {
                    Disable-NetFirewallRule -DisplayGroup "File and Printer Sharing" -ErrorAction Stop | Out-Null
                    $Hash.LogQueue.Enqueue("Successfully disabled 'File and Printer Sharing' for all profiles.")
                    $Hash.Result = "Success"
                } catch {
                    $Hash.LogQueue.Enqueue("PowerShell Cmdlet Error: $($_.Exception.Message)")
                    $Hash.LogQueue.Enqueue("Attempting fallback using netsh...")
                    & netsh advfirewall firewall set rule group="File and Printer Sharing" new enable=No
                    $Hash.Result = "Success (Fallback)"
                }
            }
            'UtilWingetRepair' {
                $Hash.LogQueue.Enqueue(">>> Resetting Winget Sources...")
                & winget source reset --force 2>&1 | ForEach-Object { Write-FilteredLog $_ }
                $Hash.Result = "Success"
            }
            'UtilInstallChoco' {
                $Hash.LogQueue.Enqueue(">>> Installing Chocolatey Package Manager...")
                try {
                    # Standard Chocolatey Install Script (Official Method)
                    Set-ExecutionPolicy Bypass -Scope Process -Force
                    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
                    $installCmd = "iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))"
                    
                    Invoke-Expression $installCmd | Out-String -Stream | ForEach-Object { Write-FilteredLog $_ }
                    
                    $Hash.LogQueue.Enqueue(">>> Installation command executed.")
                    $Hash.LogQueue.Enqueue("Note: You may need to restart the application or computer for 'choco' to appear in path.")
                    $Hash.Result = "Success"
                } catch {
                    $Hash.LogQueue.Enqueue("Error installing Chocolatey: $($_.Exception.Message)")
                    $Hash.Result = "Error: $($_.Exception.Message)"
                }
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
            'UtilDisableLongPath' {
                $Hash.LogQueue.Enqueue(">>> Disabling Win32 Long Paths...")
                try {
                    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 0 -ErrorAction Stop
                    $Hash.LogQueue.Enqueue("Successfully disabled Long Paths in the registry.")
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
                    $Hash.LogQueue.Enqueue("    -> Opening SDIO interface.")
                    
                    Start-Process -FilePath $SDIOPath
                    
                    $Hash.LogQueue.Enqueue("[+] SDIO launched successfully.")
                    $Hash.Result = "Success (SDIO)"
                } else {
                    $Hash.LogQueue.Enqueue("[-] SDIO execution aborted. Executable not found after extraction.")
                    $Hash.Result = "Error: SDIO executable missing."
                }
            }
            'UtilExportODBC' {
                $folder = $Query
                $Hash.LogQueue.Enqueue(">>> Exporting ODBC Data Sources to: $folder")
                
                $exports = @(
                    @{ Name = "System_64bit"; Path = "HKLM\SOFTWARE\ODBC\ODBC.INI" },
                    @{ Name = "System_32bit"; Path = "HKLM\SOFTWARE\WOW6432Node\ODBC\ODBC.INI" },
                    @{ Name = "User"; Path = "HKCU\Software\ODBC\ODBC.INI" }
                )
                
                $successCount = 0
                foreach ($exp in $exports) {
                    $file = Join-Path $folder "ODBC_Backup_$($exp.Name).reg"
                    $Hash.LogQueue.Enqueue("-> Exporting $($exp.Name) to $file...")
                    
                    # Run reg.exe to export safely
                    $procInfo = New-Object System.Diagnostics.ProcessStartInfo
                    $procInfo.FileName = "reg.exe"
                    $procInfo.Arguments = "export `"$($exp.Path)`" `"$file`" /y"
                    $procInfo.RedirectStandardOutput = $true
                    $procInfo.RedirectStandardError = $true
                    $procInfo.UseShellExecute = $false
                    $procInfo.CreateNoWindow = $true
                    
                    $process = [System.Diagnostics.Process]::Start($procInfo)
                    $process.WaitForExit()
                    
                    if ($process.ExitCode -eq 0) {
                        $Hash.LogQueue.Enqueue("   [OK] Exported successfully.")
                        $successCount++
                    } else {
                        $err = $process.StandardError.ReadToEnd().Trim()
                        $Hash.LogQueue.Enqueue("   [WARNING] Export failed or key does not exist (Code: $($process.ExitCode)). Details: $err")
                    }
                }
                $Hash.Result = "Success"
            }
            'UtilImportODBC' {
                $files = $Query
                $Hash.LogQueue.Enqueue(">>> Importing ODBC Data Sources from $($files.Count) file(s)...")
                
                foreach ($file in $files) {
                    $Hash.LogQueue.Enqueue("-> Importing: $file")
                    
                    # Run reg.exe to import safely
                    $procInfo = New-Object System.Diagnostics.ProcessStartInfo
                    $procInfo.FileName = "reg.exe"
                    $procInfo.Arguments = "import `"$file`""
                    $procInfo.RedirectStandardOutput = $true
                    $procInfo.RedirectStandardError = $true
                    $procInfo.UseShellExecute = $false
                    $procInfo.CreateNoWindow = $true
                    
                    $process = [System.Diagnostics.Process]::Start($procInfo)
                    $process.WaitForExit()
                    
                    if ($process.ExitCode -eq 0) {
                        $Hash.LogQueue.Enqueue("   [OK] Imported successfully.")
                    } else {
                        $Hash.LogQueue.Enqueue("   [ERROR] Import failed. Ensure it is a valid .reg file.")
                    }
                }
                $Hash.Result = "Success"
            }
            'ApplyOptimizations' {
                if ($CreateRestore) {
                    $Hash.LogQueue.Enqueue(">>> Creating System Restore Point before applying optimizations...")
                    New-BypassRestorePoint -Desc "WinToolsUI Optimizations" | Out-Null
                }
                
                # Helper function for setting registry keys deeply
                function Set-PrivacyRegKey($Path, $Name, $Value, $Type = "DWord") {
                    if (-not (Test-Path $Path)) { New-Item -Path $Path -Force -ErrorAction SilentlyContinue | Out-Null }
                    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -ErrorAction SilentlyContinue
                }

                $cfg = $Query # Query holds our hashtable of checkbox states
                $Hash.LogQueue.Enqueue(">>> Applying Customizations, Privacy, and Performance settings...")
                
                # --- CUSTOMIZATION ---
                $Hash.LogQueue.Enqueue(">>> Applying Theme & Customization settings...")
                
                if ($cfg.DarkSystem) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "SystemUsesLightTheme" 0 }
                else { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "SystemUsesLightTheme" 1 }
                
                if ($cfg.DarkApps) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "AppsUseLightTheme" 0 }
                else { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "AppsUseLightTheme" 1 }
                
                if ($cfg.DisableTransparency) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "EnableTransparency" 0 }
                else { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "EnableTransparency" 1 }
                
                if ($cfg.TaskbarAccent) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "ColorPrevalence" 1 }
                else { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "ColorPrevalence" 0 }
                
                if ($cfg.TitlebarAccent) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\DWM" "ColorPrevalence" 1 }
                else { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\DWM" "ColorPrevalence" 0 }

                $Hash.LogQueue.Enqueue(">>> Applying Taskbar settings...")
                if ($cfg.TaskbarLeft) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "TaskbarAl" 0 }
                else { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "TaskbarAl" 1 }

                if ($cfg.HideSearch) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" "SearchboxTaskbarMode" 0 }
                else { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" "SearchboxTaskbarMode" 1 }

                if ($cfg.HideTaskView) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "ShowTaskViewButton" 0 }
                else { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "ShowTaskViewButton" 1 }

                if ($cfg.HideChat) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "TaskbarMn" 0 }
                else { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "TaskbarMn" 1 }
                
                $Hash.LogQueue.Enqueue(">>> Applying Start Menu settings...")
                if ($cfg.StartMorePins) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "Start_Layout" 1 }
                else { try { Remove-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "Start_Layout" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.StartHideRecentApps) { Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer" "HideRecentlyAddedApps" 1 }
                else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer" -Name "HideRecentlyAddedApps" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.StartHideMostUsed) { Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer" "ShowOrHideMostUsedApps" 1 }
                else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer" -Name "ShowOrHideMostUsedApps" -ErrorAction SilentlyContinue } catch {} }
                
                if ($cfg.StartHideRecentDocs) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "Start_TrackDocs" 0 }
                else { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "Start_TrackDocs" 1 }
                
                $Hash.LogQueue.Enqueue(">>> Applying File Explorer settings...")
                if ($cfg.ExpHidden) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "Hidden" 1 }
                else { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "Hidden" 2 }

                if ($cfg.ExpExt) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "HideFileExt" 0 }
                else { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "HideFileExt" 1 }

                if ($cfg.ExpThisPC) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "LaunchTo" 1 }
                else { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "LaunchTo" 2 }

                if ($cfg.ExpCompact) { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "UseCompactMode" 1 }
                else { Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "UseCompactMode" 0 }

                if ($cfg.ExpClassicMenu) {
                    $keyPath = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
                    if (-not (Test-Path $keyPath)) { New-Item -Path $keyPath -Force -ErrorAction SilentlyContinue | Out-Null }
                    Set-ItemProperty -Path $keyPath -Name "(default)" -Value "" -ErrorAction SilentlyContinue
                } else {
                    $keyPath = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}"
                    if (Test-Path $keyPath) { Remove-Item -Path $keyPath -Recurse -Force -ErrorAction SilentlyContinue }
                }

                $Hash.LogQueue.Enqueue("    -> Note: Taskbar and Explorer changes may require restarting Explorer or logging out to take full effect.")


                # --- PRIVACY & ADS ---
                $Hash.LogQueue.Enqueue(">>> Applying Privacy & Ads settings...")
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
                
                # 16. Location Services
                if ($cfg.Location) {
                    $Hash.LogQueue.Enqueue("    -> Disabling System-wide Location Services")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors" "DisableLocation" 1
                } else {
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors" "DisableLocation" 0
                }
                
                # 17. App Launch Tracking
                if ($cfg.AppLaunch) {
                    $Hash.LogQueue.Enqueue("    -> Disabling App Launch Tracking (Start/Search)")
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "Start_TrackProgs" 0
                } else {
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "Start_TrackProgs" 1
                }
                
                # 18. Language List
                if ($cfg.Language) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Language List Website Access")
                    Set-PrivacyRegKey "HKCU:\Control Panel\International\User Profile" "HttpAcceptLanguageOptOut" 1
                } else {
                    Set-PrivacyRegKey "HKCU:\Control Panel\International\User Profile" "HttpAcceptLanguageOptOut" 0
                }
                
                # 19. Online Speech Recognition
                if ($cfg.Speech) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Online Speech Recognition")
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Speech_OneCore\Settings\OnlineSpeechPrivacy" "HasAccepted" 0
                } else {
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Speech_OneCore\Settings\OnlineSpeechPrivacy" "HasAccepted" 1
                }
                
                # 20. Workplace Join Messages
                if ($cfg.Workplace) {
                    $Hash.LogQueue.Enqueue("    -> Blocking Workplace Join (Manage Device) Prompts")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin" "BlockAADWorkplaceJoin" 1
                } else {
                    try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin" -Name "BlockAADWorkplaceJoin" -ErrorAction SilentlyContinue } catch {}
                }
                
                # 21. OneDrive Auto Backup
                if ($cfg.OneDrive) {
                    $Hash.LogQueue.Enqueue("    -> Disabling OneDrive Automatic Folder Backups")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive" "KFMBlockOptIn" 1
                } else {
                    try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive" -Name "KFMBlockOptIn" -ErrorAction SilentlyContinue } catch {}
                }
                
                # 22. BitLocker Auto Encryption
                if ($cfg.BitLocker) {
                    $Hash.LogQueue.Enqueue("    -> Preventing BitLocker Auto-Encryption")
                    Set-PrivacyRegKey "HKLM:\SYSTEM\CurrentControlSet\Control\BitLocker" "PreventDeviceEncryption" 1
                } else {
                    Set-PrivacyRegKey "HKLM:\SYSTEM\CurrentControlSet\Control\BitLocker" "PreventDeviceEncryption" 0
                }
                
                # 23. Automatic Maintenance
                if ($cfg.Maintenance) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Automatic System Maintenance")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\Maintenance" "MaintenanceDisabled" 1
                } else {
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\Maintenance" "MaintenanceDisabled" 0
                }
                
                # 24. Remote Assistance
                if ($cfg.RemoteAssist) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Remote Assistance Connections")
                    Set-PrivacyRegKey "HKLM:\SYSTEM\CurrentControlSet\Control\Remote Assistance" "fAllowToGetHelp" 0
                } else {
                    Set-PrivacyRegKey "HKLM:\SYSTEM\CurrentControlSet\Control\Remote Assistance" "fAllowToGetHelp" 1
                }

                # --- ADVANCED PRIVACY (O&O) ---
                if ($cfg.AppCamera) {
                    $Hash.LogQueue.Enqueue("    -> Denying Apps Access to Camera")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessCamera" 2
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsAccessCamera" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.AppMic) {
                    $Hash.LogQueue.Enqueue("    -> Denying Apps Access to Microphone")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessMicrophone" 2
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsAccessMicrophone" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.AppAccountInfo) {
                    $Hash.LogQueue.Enqueue("    -> Denying Apps Access to Account Info")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessAccountInfo" 2
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsAccessAccountInfo" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.AppContacts) {
                    $Hash.LogQueue.Enqueue("    -> Denying Apps Access to Contacts")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessContacts" 2
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsAccessContacts" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.AppCalendar) {
                    $Hash.LogQueue.Enqueue("    -> Denying Apps Access to Calendar")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessCalendar" 2
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsAccessCalendar" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.AppEmail) {
                    $Hash.LogQueue.Enqueue("    -> Denying Apps Access to Email")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessEmail" 2
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsAccessEmail" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.AppCallHistory) {
                    $Hash.LogQueue.Enqueue("    -> Denying Apps Access to Call History")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessCallHistory" 2
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsAccessCallHistory" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.AppTasks) {
                    $Hash.LogQueue.Enqueue("    -> Denying Apps Access to Tasks")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessTasks" 2
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsAccessTasks" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.AppMessages) {
                    $Hash.LogQueue.Enqueue("    -> Denying Apps Access to Messages")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessMessaging" 2
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsAccessMessaging" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.DefenderTelemetry) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Windows Defender Telemetry (Spynet)")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" "SpynetReporting" 0
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" "SubmitSamplesConsent" 2
                } else {
                    try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" -Name "SpynetReporting" -ErrorAction SilentlyContinue } catch {}
                    try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" -Name "SubmitSamplesConsent" -ErrorAction SilentlyContinue } catch {}
                }

                if ($cfg.SmartScreen) {
                    $Hash.LogQueue.Enqueue("    -> Disabling SmartScreen Filter")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" "EnableSmartScreen" 0
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -Name "EnableSmartScreen" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.P2PUpdate) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Windows Update Peer-to-Peer (Delivery Optimization)")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization" "DODownloadMode" 0
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization" -Name "DODownloadMode" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.Cortana) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Cortana")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" "AllowCortana" 0
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name "AllowCortana" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.CEIP) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Customer Experience Improvement Program (CEIP)")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\SQMClient\Windows" "CEIPEnable" 0
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\SQMClient\Windows" -Name "CEIPEnable" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.Handwriting) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Typing and Inking Telemetry")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\TabletPC" "PreventHandwritingDataSharing" 1
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\InputPersonalization" "AllowInputPersonalization" 0
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\InputPersonalization" "RestrictImplicitTextCollection" 1
                } else { 
                    try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\TabletPC" -Name "PreventHandwritingDataSharing" -ErrorAction SilentlyContinue } catch {}
                    try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\InputPersonalization" -Name "AllowInputPersonalization" -ErrorAction SilentlyContinue } catch {}
                    try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\InputPersonalization" -Name "RestrictImplicitTextCollection" -ErrorAction SilentlyContinue } catch {}
                }

                if ($cfg.AutoDriverUpdate) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Automatic Driver Updates")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" "ExcludeWUDriversInQualityUpdate" 1
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "ExcludeWUDriversInQualityUpdate" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.EdgeTelemetry) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Microsoft Edge Telemetry")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Edge" "DiagnosticData" 0
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Edge" "SendSiteInfoToImproveServices" 0
                } else { 
                    try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "DiagnosticData" -ErrorAction SilentlyContinue } catch {}
                    try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "SendSiteInfoToImproveServices" -ErrorAction SilentlyContinue } catch {}
                }

                if ($cfg.EdgeCopilot) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Microsoft Edge Copilot & Sidebar")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Edge" "HubsSidebarEnabled" 0
                } else { try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "HubsSidebarEnabled" -ErrorAction SilentlyContinue } catch {} }

                if ($cfg.EdgeSearchAds) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Microsoft Edge Search Suggestions")
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Edge" "SearchSuggestEnabled" 0
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Edge" "ResolveNavigationErrorsUseWebService" 0
                } else { 
                    try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "SearchSuggestEnabled" -ErrorAction SilentlyContinue } catch {}
                    try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "ResolveNavigationErrorsUseWebService" -ErrorAction SilentlyContinue } catch {}
                }

                # --- PERFORMANCE & GAMING ---
                $Hash.LogQueue.Enqueue(">>> Applying Performance & Gaming settings...")
                # 25. Game Mode
                if ($cfg.GameMode) {
                    $Hash.LogQueue.Enqueue("    -> Enabling Windows Game Mode")
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\GameBar" "AutoGameModeEnabled" 1
                } else {
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\GameBar" "AutoGameModeEnabled" 0
                }

                # 26. CPU Priority & Responsiveness
                if ($cfg.CpuPriority) {
                    $Hash.LogQueue.Enqueue("    -> Optimizing CPU/GPU Priority & System Responsiveness for Gaming")
                    Set-PrivacyRegKey "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games" "Priority" 6
                    Set-PrivacyRegKey "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games" "Scheduling Category" "High" "String"
                    Set-PrivacyRegKey "HKLM:\Software\Software\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games" "GPU Priority" 8
                    Set-PrivacyRegKey "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" "SystemResponsiveness" 10
                } else {
                    Set-PrivacyRegKey "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games" "Priority" 2
                    Set-PrivacyRegKey "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games" "Scheduling Category" "Medium" "String"
                    Set-PrivacyRegKey "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games" "GPU Priority" 2
                    Set-PrivacyRegKey "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" "SystemResponsiveness" 20
                }

                # 27. Hardware-Accelerated GPU Scheduling (HAGS)
                if ($cfg.HwSchMode) {
                    $Hash.LogQueue.Enqueue("    -> Enabling Hardware-Accelerated GPU Scheduling (HAGS)")
                    Set-PrivacyRegKey "HKLM:\System\CurrentControlSet\Control\GraphicsDrivers" "HwSchMode" 2
                } else {
                    Set-PrivacyRegKey "HKLM:\System\CurrentControlSet\Control\GraphicsDrivers" "HwSchMode" 1
                }

                # 28. Fullscreen Optimizations (Disable)
                if ($cfg.FSO) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Fullscreen Optimizations")
                    Set-PrivacyRegKey "HKCU:\System\GameConfigStore" "GameDVR_FSEBehaviorMode" 2
                } else {
                    Set-PrivacyRegKey "HKCU:\System\GameConfigStore" "GameDVR_FSEBehaviorMode" 0
                }

                # 29. Xbox Game DVR & Bar (Disable)
                if ($cfg.GameDVR) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Xbox Game DVR & Game Bar")
                    Set-PrivacyRegKey "HKCU:\System\GameConfigStore" "GameDVR_Enabled" 0
                    Set-PrivacyRegKey "HKLM:\Software\Microsoft\Windows\CurrentVersion\GameConfigStore" "GameDVR_Enabled" 0
                    Set-PrivacyRegKey "HKLM:\Software\Policies\Microsoft\Windows\GameDVR" "AllowGameDVR" 0
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\GameBar" "UseNexusForGameBarEnabled" 0
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\GameBar" "ShowStartupPanel" 0
                } else {
                    Set-PrivacyRegKey "HKCU:\System\GameConfigStore" "GameDVR_Enabled" 1
                    Set-PrivacyRegKey "HKLM:\Software\Microsoft\Windows\CurrentVersion\GameConfigStore" "GameDVR_Enabled" 1
                    try { Remove-ItemProperty -Path "HKLM:\Software\Policies\Microsoft\Windows\GameDVR" -Name "AllowGameDVR" -ErrorAction SilentlyContinue } catch {}
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\GameBar" "UseNexusForGameBarEnabled" 1
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\GameBar" "ShowStartupPanel" 1
                }

                # 30. Mouse Acceleration (Disable)
                if ($cfg.MouseAccel) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Enhance Pointer Precision (Mouse Acceleration)")
                    Set-PrivacyRegKey "HKCU:\Control Panel\Mouse" "MouseSpeed" "0" "String"
                    Set-PrivacyRegKey "HKCU:\Control Panel\Mouse" "MouseThreshold1" "0" "String"
                    Set-PrivacyRegKey "HKCU:\Control Panel\Mouse" "MouseThreshold2" "0" "String"
                } else {
                    Set-PrivacyRegKey "HKCU:\Control Panel\Mouse" "MouseSpeed" "1" "String"
                    Set-PrivacyRegKey "HKCU:\Control Panel\Mouse" "MouseThreshold1" "6" "String"
                    Set-PrivacyRegKey "HKCU:\Control Panel\Mouse" "MouseThreshold2" "10" "String"
                }

                # 31. Menu Show Delay
                if ($cfg.MenuDelay) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Menu Show Delay")
                    Set-PrivacyRegKey "HKCU:\Control Panel\Desktop" "MenuShowDelay" "0" "String"
                } else {
                    Set-PrivacyRegKey "HKCU:\Control Panel\Desktop" "MenuShowDelay" "400" "String"
                }

                # 32. Startup Delay
                if ($cfg.StartupDelay) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Startup Delay for Apps")
                    Set-PrivacyRegKey "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize" "StartupDelayInMSec" 0
                } else {
                    try { Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize" -Name "StartupDelayInMSec" -ErrorAction SilentlyContinue } catch {}
                }

                # 33. Background Apps (Disable)
                if ($cfg.BackgroundApps) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Apps Running in Background")
                    Set-PrivacyRegKey "HKCU:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsRunInBackground" 2
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsRunInBackground" 2
                } else {
                    try { Remove-ItemProperty -Path "HKCU:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsRunInBackground" -ErrorAction SilentlyContinue } catch {}
                    try { Remove-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" -Name "LetAppsRunInBackground" -ErrorAction SilentlyContinue } catch {}
                }

                # 34. Network Throttling (Disable)
                if ($cfg.NetworkThrottling) {
                    $Hash.LogQueue.Enqueue("    -> Disabling Network Throttling for Gaming")
                    Set-PrivacyRegKey "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" "NetworkThrottlingIndex" 10
                } else {
                    Set-PrivacyRegKey "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" "NetworkThrottlingIndex" 5
                }

                # 35. Storage Sense (Enable)
                if ($cfg.StorageSense) {
                    $Hash.LogQueue.Enqueue("    -> Enabling Storage Sense")
                    Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy" "01" 1
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\StorageSense" "AllowStorageSenseGlobal" 1
                } else {
                    Set-PrivacyRegKey "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy" "01" 0
                    Set-PrivacyRegKey "HKLM:\SOFTWARE\Policies\Microsoft\Windows\StorageSense" "AllowStorageSenseGlobal" 0
                }

                $Hash.LogQueue.Enqueue("[+] Optimization & Customization settings applied successfully.")
                $Hash.Result = "Success"
            }
        }
    } catch {
        $Hash.Result = "Error: $($_.Exception.Message)"
    }
}

# 5. Helper Function to safely dispatch jobs to the Runspace
function Start-WingetJob($Action, $Query, $Id, $StatusMsg, $IsAdmin = $false, $CreateRestore = $false, $Manager = "Winget") {
    # Check if a job is already running and warn the user
    if ($script:IsJobRunning) {
        [System.Windows.MessageBox]::Show("A task is currently running in the background.`n`nPlease wait for it to finish or click 'Stop' before starting a new action.", "Task in Progress", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    $script:IsJobRunning = $true
    $script:LastLogWasProgress = $false
    
    # Clear the live log UI and Queue
    $LogTextBox.Clear()
    $dummy = [string]::Empty
    while ($syncHash.LogQueue.TryDequeue([ref]$dummy)) {}

    # Auto-expand the log panel if we are making system changes
    if ($Action -in @('Install', 'Uninstall', 'Update', 'UtilSystemScan', 'UtilResetWU', 'UtilRestorePoint', 'UtilLongPath', 'UtilDisableLongPath', 'UtilResetNet', 'UtilSMB', 'UtilDisableSMB', 'UtilFileSharing', 'UtilDisableFileSharing', 'UtilDriverUpdate', 'UtilSDIO', 'UtilWingetRepair', 'UtilStoreRepair', 'UtilDiskCleanup', 'UtilIconCache', 'UtilClearLogs', 'ApplyOptimizations', 'UtilExportODBC', 'UtilImportODBC')) {
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

    $script:psInstance = [PowerShell]::Create().AddScript($bgJobBlock).AddArgument($Action).AddArgument($Query).AddArgument($Id).AddArgument($syncHash).AddArgument($IsAdmin).AddArgument($CreateRestore).AddArgument($Manager)
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
    # --- Monitor Background Network State ---
    if ($syncHash.IsOnline -ne $null -and $syncHash.IsOnline -ne $script:IsOnline) {
        $script:IsOnline = $syncHash.IsOnline
        Set-OfflineMode
        
        if ($script:IsOnline) {
            $syncHash.LogQueue.Enqueue("`r`n[NETWORK] Internet connection restored. Online mode activated.")
        } else {
            $syncHash.LogQueue.Enqueue("`r`n[NETWORK] Internet connection lost. Switched to Offline mode.")
        }
    }

    # Drain the live log queue
    $logLine = [string]::Empty
    $queueItems = @()
    while ($syncHash.LogQueue.TryDequeue([ref]$logLine)) {
        $queueItems += $logLine
    }
    
    # Process updates into the LogTextBox
    if ($queueItems.Count -gt 0) {
        $txt = $LogTextBox.Text
        
        foreach ($item in $queueItems) {
            # Check if this line is a streaming progress update
            if ($item -match "^\[PROGRESS\](.*)") {
                $progText = $matches[1]
                
                if ($script:LastLogWasProgress) {
                    # Overwrite the last line for a seamless streaming effect
                    $idx = $txt.LastIndexOf("`r`n")
                    if ($idx -ge 0) { $txt = $txt.Substring(0, $idx) } else { $txt = "" }
                }
                
                if ($txt.Length -gt 0) { $txt += "`r`n$progText" } else { $txt = "$progText" }
                $script:LastLogWasProgress = $true
            } else {
                # Normal log line
                if ($txt.Length -gt 0) { $txt += "`r`n$item" } else { $txt = "$item" }
                $script:LastLogWasProgress = $false
            }
        }
        
        $LogTextBox.Text = $txt
        $LogTextBox.ScrollToEnd()
    }
    
    # Check if a live percentage was reported back from Winget (for the GUI bar)
    if ($syncHash.Progress -ne $null) {
        if ($JobProgress.IsIndeterminate) {
            # Switch from spinning mode to solid fill mode
            $JobProgress.IsIndeterminate = $false
        }
        $JobProgress.Value = $syncHash.Progress
        $StatusText.Text = "$($syncHash.StatusMsg) - $($syncHash.Progress)%"
    }

    if ($script:asyncResult -ne $null -and $script:asyncResult.IsCompleted) {
        # REMOVED $timer.Stop() here so the UI continues checking network state even when idle!
        
        try {
            $script:psInstance.EndInvoke($script:asyncResult)
        } catch {
            # This triggers if the user clicks the "Stop" button and aborts the pipeline
            $syncHash.Result = "Error: Operation was cancelled by the user."
        }
        
        $script:psInstance.Dispose()
        
        # Reset states so we don't process completion twice and don't freeze the progress bar
        $script:asyncResult = $null 
        $syncHash.Progress = $null
        
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
                    
                    # Uncheck all template checkboxes to match the empty queue
                    foreach ($chk in $TemplateCheckboxes) {
                        $chk.IsChecked = $false
                    }
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
                'UtilDisableLongPath' { $StatusText.Text = "Long Paths have been disabled successfully." }
                'UtilResetNet'   { $StatusText.Text = "Network adapters reset successfully." }
                'UtilSMB'        { $StatusText.Text = "SMBv1 protocol has been enabled." }
                'UtilDisableSMB' { $StatusText.Text = "SMBv1 protocol has been disabled." }
                'UtilFileSharing'{ $StatusText.Text = "File and Printer Sharing enabled globally." }
                'UtilDisableFileSharing'{ $StatusText.Text = "File and Printer Sharing disabled globally." }
                'UtilWingetRepair' { $StatusText.Text = "Winget repositories have been reset." }
                'UtilStoreRepair'  { $StatusText.Text = "Microsoft Store cache cleared." }
                'UtilInstallChoco' { $StatusText.Text = "Chocolatey installation completed." }
                'UtilDiskCleanup'  { $StatusText.Text = "Deep disk cleanup completed." }
                'UtilIconCache'    { $StatusText.Text = "Icon and thumbnail cache rebuilt." }
                'UtilClearLogs'    { $StatusText.Text = "Event Viewer logs successfully cleared." }
                'UtilDriverUpdate' { 
                    $StatusText.Text = if ($res -match "Reboot") { "Drivers installed! System reboot required." } else { "Driver scan and update process completed." } 
                }
                'UtilSDIO'         { $StatusText.Text = "Snappy Driver Installer opened." }
                'UtilExportODBC'   { $StatusText.Text = "ODBC Data Sources exported successfully." }
                'UtilImportODBC'   { $StatusText.Text = "ODBC Data Sources imported successfully." }
                'ApplyOptimizations' { $StatusText.Text = "Optimization & Customization settings applied successfully." }
            }
        }
    }
})

# --- Privacy & Customization Check/Read Function ---
function Read-OptimizeStates {
    # Helper to read registry silently
    function Get-PrivacyValue($Path, $Name, $ExpectedValue) {
        try {
            $val = (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name
            return ("$val" -eq "$ExpectedValue")
        } catch { return $false }
    }
    
    # Read Customization States
    $ChkDarkSystem.IsChecked = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "SystemUsesLightTheme" 0
    $ChkDarkApps.IsChecked = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "AppsUseLightTheme" 0
    $ChkDisableTransparency.IsChecked = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "EnableTransparency" 0
    $ChkTaskbarAccent.IsChecked = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize" "ColorPrevalence" 1
    $ChkTitlebarAccent.IsChecked = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\DWM" "ColorPrevalence" 1
    
    $ChkTaskbarLeft.IsChecked = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "TaskbarAl" 0
    $ChkHideSearch.IsChecked = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search" "SearchboxTaskbarMode" 0
    $ChkHideTaskView.IsChecked = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "ShowTaskViewButton" 0
    $ChkHideChat.IsChecked = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "TaskbarMn" 0

    $ChkStartMorePins.IsChecked = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "Start_Layout" 1
    $ChkStartHideRecentApps.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer" "HideRecentlyAddedApps" 1
    $ChkStartHideMostUsed.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer" "ShowOrHideMostUsedApps" 1
    $ChkStartHideRecentDocs.IsChecked = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "Start_TrackDocs" 0

    $ChkExpHidden.IsChecked = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "Hidden" 1
    $ChkExpExt.IsChecked = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "HideFileExt" 0
    $ChkExpThisPC.IsChecked = (Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "LaunchTo" 1)
    $ChkExpCompact.IsChecked = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "UseCompactMode" 1
    try {
        $classicMenuExists = Test-Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
        $ChkExpClassicMenu.IsChecked = $classicMenuExists
    } catch { $ChkExpClassicMenu.IsChecked = $false }

    # Read Privacy States
    $ChkTelemetry.IsChecked   = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" "AllowTelemetry" 0
    $ChkLocation.IsChecked    = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors" "DisableLocation" 1
    $ChkActivity.IsChecked    = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" "EnableActivityFeed" 0
    $ChkAppLaunch.IsChecked   = Get-PrivacyValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "Start_TrackProgs" 0
    $ChkLanguage.IsChecked    = Get-PrivacyValue "HKCU:\Control Panel\International\User Profile" "HttpAcceptLanguageOptOut" 1
    $ChkSpeech.IsChecked      = Get-PrivacyValue "HKCU:\Software\Microsoft\Speech_OneCore\Settings\OnlineSpeechPrivacy" "HasAccepted" 0
    $ChkTailoredExp.IsChecked = Get-PrivacyValue "HKCU:\Software\Policies\Microsoft\Windows\CloudContent" "DisableTailoredExperiencesWithDiagnosticData" 1
    $ChkWER.IsChecked         = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting" "Disabled" 1
    $ChkFeedback.IsChecked    = Get-PrivacyValue "HKCU:\SOFTWARE\Policies\Microsoft\Windows\DataCollection" "DoNotShowFeedbackNotifications" 1
    $ChkWorkplace.IsChecked   = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin" "BlockAADWorkplaceJoin" 1
    $ChkOneDrive.IsChecked    = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive" "KFMBlockOptIn" 1
    $ChkBingSearch.IsChecked  = Get-PrivacyValue "HKCU:\SOFTWARE\Policies\Microsoft\Windows\Explorer" "DisableSearchBoxSuggestions" 1
    $ChkStartAds.IsChecked    = Get-PrivacyValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338388Enabled" 0
    $ChkLockScreenAds.IsChecked = Get-PrivacyValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338387Enabled" 0
    $ChkExplorerAds.IsChecked = Get-PrivacyValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" "ShowSyncProviderNotifications" 0
    $ChkWelcomeExp.IsChecked  = Get-PrivacyValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-310093Enabled" 0
    $ChkAdId.IsChecked        = Get-PrivacyValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" "Enabled" 0
    $ChkBitLocker.IsChecked   = Get-PrivacyValue "HKLM:\SYSTEM\CurrentControlSet\Control\BitLocker" "PreventDeviceEncryption" 1
    $ChkCopilot.IsChecked     = Get-PrivacyValue "HKCU:\Software\Policies\Microsoft\Windows\WindowsCopilot" "TurnOffWindowsCopilot" 1
    $ChkWidgets.IsChecked     = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Dsh" "AllowNewsAndInterests" 0
    $ChkMaintenance.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\Maintenance" "MaintenanceDisabled" 1
    $ChkConsumer.IsChecked    = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" "DisableWindowsConsumerFeatures" 1
    $ChkWifiSense.IsChecked   = Get-PrivacyValue "HKLM:\SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config" "AutoConnectAllowedOEM" 0
    $ChkRemoteAssist.IsChecked= Get-PrivacyValue "HKLM:\SYSTEM\CurrentControlSet\Control\Remote Assistance" "fAllowToGetHelp" 0
    
    # Read Advanced Privacy (O&O) States
    $ChkAppCamera.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessCamera" 2
    $ChkAppMic.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessMicrophone" 2
    $ChkAppAccountInfo.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessAccountInfo" 2
    $ChkAppContacts.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessContacts" 2
    $ChkAppCalendar.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessCalendar" 2
    $ChkAppEmail.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessEmail" 2
    $ChkAppCallHistory.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessCallHistory" 2
    $ChkAppTasks.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessTasks" 2
    $ChkAppMessages.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsAccessMessaging" 2
    
    $ChkDefenderTelemetry.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender\Spynet" "SpynetReporting" 0
    $ChkSmartScreen.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" "EnableSmartScreen" 0
    $ChkP2PUpdate.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization" "DODownloadMode" 0
    $ChkCortana.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" "AllowCortana" 0
    $ChkCEIP.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\SQMClient\Windows" "CEIPEnable" 0
    $ChkHandwriting.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\TabletPC" "PreventHandwritingDataSharing" 1
    $ChkAutoDriverUpdate.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" "ExcludeWUDriversInQualityUpdate" 1
    
    $ChkEdgeTelemetry.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Edge" "DiagnosticData" 0
    $ChkEdgeCopilot.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Edge" "HubsSidebarEnabled" 0
    $ChkEdgeSearchAds.IsChecked = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Edge" "SearchSuggestEnabled" 0

    # Read Performance States
    $ChkGameMode.IsChecked          = Get-PrivacyValue "HKCU:\Software\Microsoft\GameBar" "AutoGameModeEnabled" 1
    $ChkCpuPriority.IsChecked       = Get-PrivacyValue "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile\Tasks\Games" "Priority" 6
    $ChkHwSchMode.IsChecked         = Get-PrivacyValue "HKLM:\System\CurrentControlSet\Control\GraphicsDrivers" "HwSchMode" 2
    $ChkFSO.IsChecked               = Get-PrivacyValue "HKCU:\System\GameConfigStore" "GameDVR_FSEBehaviorMode" 2
    $ChkGameDVR.IsChecked           = Get-PrivacyValue "HKCU:\System\GameConfigStore" "GameDVR_Enabled" 0
    $ChkMouseAccel.IsChecked        = Get-PrivacyValue "HKCU:\Control Panel\Mouse" "MouseSpeed" "0"
    $ChkMenuDelay.IsChecked         = Get-PrivacyValue "HKCU:\Control Panel\Desktop" "MenuShowDelay" "0"
    $ChkStartupDelay.IsChecked      = Get-PrivacyValue "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Serialize" "StartupDelayInMSec" 0
    $ChkBackgroundApps.IsChecked    = Get-PrivacyValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy" "LetAppsRunInBackground" 2
    $ChkNetworkThrottling.IsChecked = Get-PrivacyValue "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile" "NetworkThrottlingIndex" 10
    $ChkStorageSense.IsChecked      = Get-PrivacyValue "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy" "01" 1
}

$RefreshOptimizeBtn.Add_Click({
    Read-OptimizeStates
    $StatusText.Text = "Optimization & Customization settings refreshed from Registry."
})

$ApplyOptimizeBtn.Add_Click({
    $cfg = @{
        # Customization Elements
        DarkSystem          = $ChkDarkSystem.IsChecked -eq $true
        DarkApps            = $ChkDarkApps.IsChecked -eq $true
        DisableTransparency = $ChkDisableTransparency.IsChecked -eq $true
        TaskbarAccent       = $ChkTaskbarAccent.IsChecked -eq $true
        TitlebarAccent      = $ChkTitlebarAccent.IsChecked -eq $true
        TaskbarLeft         = $ChkTaskbarLeft.IsChecked -eq $true
        HideSearch          = $ChkHideSearch.IsChecked -eq $true
        HideTaskView        = $ChkHideTaskView.IsChecked -eq $true
        HideChat            = $ChkHideChat.IsChecked -eq $true
        StartMorePins       = $ChkStartMorePins.IsChecked -eq $true
        StartHideRecentApps = $ChkStartHideRecentApps.IsChecked -eq $true
        StartHideMostUsed   = $ChkStartHideMostUsed.IsChecked -eq $true
        StartHideRecentDocs = $ChkStartHideRecentDocs.IsChecked -eq $true
        ExpHidden           = $ChkExpHidden.IsChecked -eq $true
        ExpExt              = $ChkExpExt.IsChecked -eq $true
        ExpThisPC           = $ChkExpThisPC.IsChecked -eq $true
        ExpCompact          = $ChkExpCompact.IsChecked -eq $true
        ExpClassicMenu      = $ChkExpClassicMenu.IsChecked -eq $true
        
        # Privacy Elements
        Telemetry     = $ChkTelemetry.IsChecked -eq $true
        Location      = $ChkLocation.IsChecked -eq $true
        Activity      = $ChkActivity.IsChecked -eq $true
        AppLaunch     = $ChkAppLaunch.IsChecked -eq $true
        Language      = $ChkLanguage.IsChecked -eq $true
        Speech        = $ChkSpeech.IsChecked -eq $true
        TailoredExp   = $ChkTailoredExp.IsChecked -eq $true
        WER           = $ChkWER.IsChecked -eq $true
        Feedback      = $ChkFeedback.IsChecked -eq $true
        Workplace     = $ChkWorkplace.IsChecked -eq $true
        OneDrive      = $ChkOneDrive.IsChecked -eq $true
        BingSearch    = $ChkBingSearch.IsChecked -eq $true
        StartAds      = $ChkStartAds.IsChecked -eq $true
        LockScreenAds = $ChkLockScreenAds.IsChecked -eq $true
        ExplorerAds   = $ChkExplorerAds.IsChecked -eq $true
        WelcomeExp    = $ChkWelcomeExp.IsChecked -eq $true
        AdId          = $ChkAdId.IsChecked -eq $true
        BitLocker     = $ChkBitLocker.IsChecked -eq $true
        Copilot       = $ChkCopilot.IsChecked -eq $true
        Widgets       = $ChkWidgets.IsChecked -eq $true
        Maintenance   = $ChkMaintenance.IsChecked -eq $true
        Consumer      = $ChkConsumer.IsChecked -eq $true
        WifiSense     = $ChkWifiSense.IsChecked -eq $true
        RemoteAssist  = $ChkRemoteAssist.IsChecked -eq $true
        
        # O&O Privacy Elements
        AppCamera         = $ChkAppCamera.IsChecked -eq $true
        AppMic            = $ChkAppMic.IsChecked -eq $true
        AppAccountInfo    = $ChkAppAccountInfo.IsChecked -eq $true
        AppContacts       = $ChkAppContacts.IsChecked -eq $true
        AppCalendar       = $ChkAppCalendar.IsChecked -eq $true
        AppEmail          = $ChkAppEmail.IsChecked -eq $true
        AppCallHistory    = $ChkAppCallHistory.IsChecked -eq $true
        AppTasks          = $ChkAppTasks.IsChecked -eq $true
        AppMessages       = $ChkAppMessages.IsChecked -eq $true
        
        DefenderTelemetry = $ChkDefenderTelemetry.IsChecked -eq $true
        SmartScreen       = $ChkSmartScreen.IsChecked -eq $true
        P2PUpdate         = $ChkP2PUpdate.IsChecked -eq $true
        Cortana           = $ChkCortana.IsChecked -eq $true
        CEIP              = $ChkCEIP.IsChecked -eq $true
        Handwriting       = $ChkHandwriting.IsChecked -eq $true
        AutoDriverUpdate  = $ChkAutoDriverUpdate.IsChecked -eq $true
        
        EdgeTelemetry     = $ChkEdgeTelemetry.IsChecked -eq $true
        EdgeCopilot       = $ChkEdgeCopilot.IsChecked -eq $true
        EdgeSearchAds     = $ChkEdgeSearchAds.IsChecked -eq $true

        # Performance Elements
        GameMode          = $ChkGameMode.IsChecked -eq $true
        CpuPriority       = $ChkCpuPriority.IsChecked -eq $true
        HwSchMode         = $ChkHwSchMode.IsChecked -eq $true
        FSO               = $ChkFSO.IsChecked -eq $true
        GameDVR           = $ChkGameDVR.IsChecked -eq $true
        MouseAccel        = $ChkMouseAccel.IsChecked -eq $true
        MenuDelay         = $ChkMenuDelay.IsChecked -eq $true
        StartupDelay      = $ChkStartupDelay.IsChecked -eq $true
        BackgroundApps    = $ChkBackgroundApps.IsChecked -eq $true
        NetworkThrottling = $ChkNetworkThrottling.IsChecked -eq $true
        StorageSense      = $ChkStorageSense.IsChecked -eq $true
    }
    $createRestore = $CreateRestorePointOptimizeCheck.IsChecked -eq $true
    Start-WingetJob -Action "ApplyOptimizations" -Query $cfg -Id "" -StatusMsg "Applying optimization & customization settings... Please wait." -CreateRestore $createRestore
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
        Start-WingetJob -Action "Search" -Query $query -Id "" -StatusMsg "Searching Winget & Chocolatey for '$query'..."
    }
})

$SearchBox.Add_KeyDown({
    if ($_.Key -eq 'Enter') {
        # Prevent searching if offline
        if (-not $script:IsOnline) {
            [System.Windows.MessageBox]::Show("Search requires an active internet connection.", "Offline Mode", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            return
        }
        
        $query = $SearchBox.Text
        if (![string]::IsNullOrWhiteSpace($query)) {
            $DiscoverGrid.ItemsSource = $null
            Start-WingetJob -Action "Search" -Query $query -Id "" -StatusMsg "Searching Winget & Chocolatey for '$query'..."
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

$AddTemplatesToQueueBtn.Add_Click({
    $addedCount = 0
    foreach ($chk in $TemplateCheckboxes) {
        if ($chk.IsChecked) {
            $tagData = $chk.Tag -split '\|'
            $mgr = $tagData[0]
            $id = $tagData[1]
            $name = $chk.Content

            $exists = $false
            foreach ($q in $script:InstallQueue) {
                if ($q.Id -eq $id) { $exists = $true; break }
            }
            if (-not $exists) {
                $script:InstallQueue.Add([PSCustomObject]@{Name=$name; Id=$id; Version="Latest"; Manager=$mgr})
                $addedCount++
            }
            # Uncheck it so the user knows it has been processed
            $chk.IsChecked = $false
        }
    }
    if ($addedCount -eq 0) {
        [System.Windows.MessageBox]::Show("No new templates were added. Ensure you have checked items that aren't already in the queue below.", "No Items Added", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
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
        
        # Keep checkboxes visually in sync if an item is removed manually
        foreach ($chk in $TemplateCheckboxes) {
            $tagData = $chk.Tag -split '\|'
            if ($tagData[1] -eq $item.Id) {
                $chk.IsChecked = $false
                break
            }
        }
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
                    $mgr = if ($item.Manager) { $item.Manager } else { "Winget" } # Fallback to Winget if standard exported queue
                    $script:InstallQueue.Add([PSCustomObject]@{Name=$item.Name; Id=$item.Id; Version="Latest"; Manager=$mgr})
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
        $targets = @($script:InstallQueue | Select-Object Id, Manager, Name)
        $msg = ""
        $promptMsg = ""
        if ($targets.Count -eq 1) { 
            $msg = "Installing $($script:InstallQueue[0].Name)..." 
            $promptMsg = "Are you sure you want to install $($script:InstallQueue[0].Name)?"
        } else { 
            $msg = "Installing $($targets.Count) packages... Please wait." 
            $promptMsg = "Are you sure you want to install $($targets.Count) packages?"
        }
        
        $msgResult = [System.Windows.MessageBox]::Show($promptMsg, "Confirm Install", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        
        if ($msgResult -eq 'Yes') {
            $isAdmin = $AdminInstallCheck.IsChecked -eq $true
            $createRestore = $CreateRestorePointInstallCheck.IsChecked -eq $true
            Start-WingetJob -Action "Install" -Query "" -Id $targets -StatusMsg $msg -IsAdmin $isAdmin -CreateRestore $createRestore
        }
    } else { 
        [System.Windows.MessageBox]::Show("Please add packages to the install queue before clicking Install.") 
    }
})

$RefreshInstalledBtn.Add_Click({
    $SearchInstalledBox.Clear()
    $InstalledGrid.ItemsSource = $null
    $WindowsAppsGrid.ItemsSource = $null
    $script:AllInstalledApps = $null
    
    Set-OfflineMode
    $loadMsg = if ($script:IsOnline) { "Loading installed packages (Winget & Chocolatey) and checking for updates..." } else { "Loading installed packages locally (Offline Mode)..." }
    Start-WingetJob -Action "Installed" -Query "" -Id "" -StatusMsg $loadMsg
})

$SelectAllInstalledBtn.Add_Click({
    # Determine active grid based on selected tab
    $activeGrid = if ($InstalledTabs.SelectedIndex -eq 1) { $WindowsAppsGrid } else { $InstalledGrid }
    if ($activeGrid.ItemsSource) {
        foreach ($item in $activeGrid.ItemsSource) { $item.IsSelected = $true }
        $activeGrid.Items.Refresh()
    }
})

$DeselectAllInstalledBtn.Add_Click({
    # Determine active grid
    $activeGrid = if ($InstalledTabs.SelectedIndex -eq 1) { $WindowsAppsGrid } else { $InstalledGrid }
    if ($activeGrid.ItemsSource) {
        foreach ($item in $activeGrid.ItemsSource) { $item.IsSelected = $false }
        $activeGrid.Items.Refresh()
    }
})

$SearchInstalledBox.Add_TextChanged({
    $filter = $SearchInstalledBox.Text
    if ($script:AllInstalledApps -ne $null) {
        $filtered = if ([string]::IsNullOrWhiteSpace($filter)) {
            $script:AllInstalledApps
        } else {
            $script:AllInstalledApps | Where-Object { $_.Name -like "*$filter*" -or $_.Id -like "*$filter*" }
        }
        
        # Split into two lists based on standard Windows Store / Appx / MSIX heuristics
        $desktopApps = @($filtered | Where-Object { $_.Source -ne 'msstore' -and $_.Id -notmatch '_[a-zA-Z0-9]{13}$' -and $_.Id -notmatch '^MSIX\\' })
        $windowsApps = @($filtered | Where-Object { $_.Source -eq 'msstore' -or $_.Id -match '_[a-zA-Z0-9]{13}$' -or $_.Id -match '^MSIX\\' })
        
        $InstalledGrid.ItemsSource = $desktopApps
        $WindowsAppsGrid.ItemsSource = $windowsApps
    }
})

$UninstallBtn.Add_Click({
    # Combine explicitly checked items + standard selection (if checked is empty)
    $checked = @()
    if ($InstalledGrid.ItemsSource) { $checked += @($InstalledGrid.ItemsSource | Where-Object { $_.IsSelected }) }
    if ($WindowsAppsGrid.ItemsSource) { $checked += @($WindowsAppsGrid.ItemsSource | Where-Object { $_.IsSelected }) }
    
    # Use ArrayList to safely collect SelectedItems without += array rebuilding issues
    $hlList = New-Object System.Collections.ArrayList
    if ($InstalledGrid.SelectedItems -and $InstalledGrid.SelectedItems.Count -gt 0) {
        foreach($i in $InstalledGrid.SelectedItems) { [void]$hlList.Add($i) }
    }
    if ($WindowsAppsGrid.SelectedItems -and $WindowsAppsGrid.SelectedItems.Count -gt 0) {
        foreach($i in $WindowsAppsGrid.SelectedItems) { [void]$hlList.Add($i) }
    }
    $highlighted = @($hlList)
    
    # Prioritize Checkboxes if any are checked, otherwise use Highlighted rows
    $finalSelection = @(if ($checked.Count -gt 0) { $checked } else { $highlighted })
    
    if ($finalSelection.Count -gt 0) { 
        $targets = @($finalSelection | Select-Object -Unique Id, Manager, Name)
        $msg = ""
        $promptMsg = ""
        if ($targets.Count -eq 1) { 
            $msg = "Uninstalling $($targets[0].Name)..." 
            $promptMsg = "Are you sure you want to uninstall $($targets[0].Name)?"
        } else { 
            $msg = "Uninstalling $($targets.Count) packages... Please wait." 
            $promptMsg = "Are you sure you want to uninstall $($targets.Count) packages?"
        }

        $msgResult = [System.Windows.MessageBox]::Show($promptMsg, "Confirm Uninstall", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        
        if ($msgResult -eq 'Yes') {
            # Save the selected app objects to the background syncHash so the scanner knows what to look for
            $syncHash.TargetApps = $finalSelection
            $createRestore = $CreateRestorePointUpdateCheck.IsChecked -eq $true
            Start-WingetJob -Action "Uninstall" -Query "" -Id $targets -StatusMsg $msg -CreateRestore $createRestore
        }
    } else { 
        [System.Windows.MessageBox]::Show("Please select at least one installed package to uninstall.") 
    }
})

$UpdateBtn.Add_Click({
    $checked = @()
    if ($InstalledGrid.ItemsSource) { $checked += @($InstalledGrid.ItemsSource | Where-Object { $_.IsSelected }) }
    if ($WindowsAppsGrid.ItemsSource) { $checked += @($WindowsAppsGrid.ItemsSource | Where-Object { $_.IsSelected }) }
    
    # Use ArrayList to safely collect SelectedItems without += array rebuilding issues
    $hlList = New-Object System.Collections.ArrayList
    if ($InstalledGrid.SelectedItems -and $InstalledGrid.SelectedItems.Count -gt 0) {
        foreach($i in $InstalledGrid.SelectedItems) { [void]$hlList.Add($i) }
    }
    if ($WindowsAppsGrid.SelectedItems -and $WindowsAppsGrid.SelectedItems.Count -gt 0) {
        foreach($i in $WindowsAppsGrid.SelectedItems) { [void]$hlList.Add($i) }
    }
    $highlighted = @($hlList)
    
    $finalSelection = @(if ($checked.Count -gt 0) { $checked } else { $highlighted })
    
    if ($finalSelection.Count -gt 0) { 
        $targets = @($finalSelection | Select-Object -Unique Id, Manager, Name)
        $msg = ""
        $promptMsg = ""
        if ($targets.Count -eq 1) { 
            $msg = "Updating $($targets[0].Name)..." 
            $promptMsg = "Are you sure you want to update $($targets[0].Name)?"
        } else { 
            $msg = "Updating $($targets.Count) packages... Please wait." 
            $promptMsg = "Are you sure you want to update $($targets.Count) packages?"
        }
        
        $msgResult = [System.Windows.MessageBox]::Show($promptMsg, "Confirm Update", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        
        if ($msgResult -eq 'Yes') {
            $createRestore = $CreateRestorePointUpdateCheck.IsChecked -eq $true
            Start-WingetJob -Action "Update" -Query "" -Id $targets -StatusMsg $msg -CreateRestore $createRestore
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
    $targets = @($script:AllInstalledApps | Where-Object { $_.HasUpdate -eq $true } | Select-Object Id, Manager, Name)
    
    if ($targets.Count -gt 0) {
        $msg = if ($targets.Count -eq 1) { "Updating 1 package..." } else { "Updating all $($targets.Count) available updates... Please wait." }
        $promptMsg = if ($targets.Count -eq 1) { "Are you sure you want to update 1 package?" } else { "Are you sure you want to update all $($targets.Count) available packages?" }
        
        $msgResult = [System.Windows.MessageBox]::Show($promptMsg, "Confirm Update All", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
        
        if ($msgResult -eq 'Yes') {
            $createRestore = $CreateRestorePointUpdateCheck.IsChecked -eq $true
            Start-WingetJob -Action "Update" -Query "" -Id $targets -StatusMsg $msg -CreateRestore $createRestore
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
        $mgr = if ($sel.Manager) { $sel.Manager } else { "Winget" }
        Start-WingetJob -Action "ShowDetails" -Query "" -Id $sel.Id -StatusMsg "Loading details for $($sel.Name)..." -Manager $mgr
    }
}
$DiscoverMenuDetails.Add_Click({ &$showDetailsAction $DiscoverGrid })
$QueueMenuDetails.Add_Click({ &$showDetailsAction $QueueGrid })
$InstalledMenuDetails.Add_Click({ &$showDetailsAction $InstalledGrid })
$WindowsAppsMenuDetails.Add_Click({ &$showDetailsAction $WindowsAppsGrid })

$copyIdAction = {
    param($GridControl)
    $sel = $GridControl.SelectedItem
    if ($sel -ne $null) {
        Set-Clipboard -Value $sel.Id
        $StatusText.Text = "Copied ID: $($sel.Id)"
    }
}
$InstalledMenuCopyId.Add_Click({ &$copyIdAction $InstalledGrid })
$WindowsAppsMenuCopyId.Add_Click({ &$copyIdAction $WindowsAppsGrid })

$uninstallContextAction = {
    param($GridControl)
    $sel = $GridControl.SelectedItem
    if ($sel -ne $null) {
        $targets = @($sel | Select-Object Id, Manager, Name)
        $msgResult = [System.Windows.MessageBox]::Show("Are you sure you want to uninstall $($sel.Name)?", "Confirm Uninstall", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
        if ($msgResult -eq 'Yes') {
            $syncHash.TargetApps = @($sel)
            $createRestore = $CreateRestorePointUpdateCheck.IsChecked -eq $true
            Start-WingetJob -Action "Uninstall" -Query "" -Id $targets -StatusMsg "Uninstalling $($sel.Name)..." -CreateRestore $createRestore
        }
    }
}
$InstalledMenuUninstall.Add_Click({ &$uninstallContextAction $InstalledGrid })
$WindowsAppsMenuUninstall.Add_Click({ &$uninstallContextAction $WindowsAppsGrid })

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

$UtilDisableLongPathBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will restore the default 260-character path limit (MAX_PATH) in Windows.`n`nContinue?", "Disable Long Paths", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilDisableLongPath" -Query "" -Id "" -StatusMsg "Disabling Long Paths... Please wait." }
})

$UtilResetNetBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will flush DNS, release/renew your IP, and completely reset Winsock and TCP/IP configurations.`n`nYou may briefly lose internet connection. Continue?", "Reset Network", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilResetNet" -Query "" -Id "" -StatusMsg "Resetting Network Adapters and configurations..." }
})

$UtilSMBBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("Are you sure you want to enable the legacy SMBv1 protocol?`n`nNote: SMBv1 is considered insecure and should only be enabled if required for connecting to legacy NAS drives or older network printers.", "Enable SMBv1", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilSMB" -Query "" -Id "" -StatusMsg "Enabling SMBv1 Protocol..." }
})

$UtilDisableSMBBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will disable the legacy SMBv1 protocol for better security.`n`nContinue?", "Disable SMBv1", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilDisableSMB" -Query "" -Id "" -StatusMsg "Disabling SMBv1 Protocol..." }
})

$UtilFileSharingBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will enable the 'File and Printer Sharing' firewall rules for all network profiles (Public, Private, Domain) and allow inbound/outbound traffic from ANY IP address.`n`nWarning: Expanding this scope on Public networks can be a security risk. Continue?", "Enable File & Printer Sharing", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilFileSharing" -Query "" -Id "" -StatusMsg "Enabling File and Printer Sharing..." }
})

$UtilDisableFileSharingBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will disable the 'File and Printer Sharing' firewall rules, preventing network discovery and shared folder access.`n`nContinue?", "Disable File & Printer Sharing", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilDisableFileSharing" -Query "" -Id "" -StatusMsg "Disabling File and Printer Sharing..." }
})

$UtilWingetRepairBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will force-reset the Winget package repositories.`n`nUse this if searches are failing or apps refuse to download. Continue?", "Repair Winget Sources", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilWingetRepair" -Query "" -Id "" -StatusMsg "Resetting Winget Sources..." }
})

$UtilStoreRepairBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will run wsreset.exe to clear the Microsoft Store cache.`n`nUse this if Store Apps are stuck on 'Pending' or fail to update. Continue?", "Repair Microsoft Store", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilStoreRepair" -Query "" -Id "" -StatusMsg "Clearing Microsoft Store Cache..." }
})

$UtilInstallChocoBtn.Add_Click({
    $msgResult = [System.Windows.MessageBox]::Show("This will download and install the Chocolatey Package Manager.`n`nContinue?", "Install Chocolatey", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($msgResult -eq 'Yes') { Start-WingetJob -Action "UtilInstallChoco" -Query "" -Id "" -StatusMsg "Installing Chocolatey..." -IsAdmin $true }
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

$UtilExportODBCBtn.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = "Select a folder to save your ODBC Registry backups."
    
    if ($fbd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $folder = $fbd.SelectedPath
        Start-WingetJob -Action "UtilExportODBC" -Query $folder -Id "" -StatusMsg "Exporting ODBC Data Sources..."
    }
})

$UtilImportODBCBtn.Add_Click({
    $ofd = New-Object Microsoft.Win32.OpenFileDialog
    $ofd.Filter = "Registry Files (*.reg)|*.reg|All Files (*.*)|*.*"
    $ofd.Multiselect = $true
    $ofd.Title = "Select the ODBC .reg backup files to import"
    
    if ($ofd.ShowDialog() -eq $true) {
        $files = $ofd.FileNames
        Start-WingetJob -Action "UtilImportODBC" -Query $files -Id "" -StatusMsg "Importing ODBC Data Sources..."
    }
})

# --- Network Connection Check & Offline Restrictions ---
function Test-InternetConnection {
    try {
        if (-not [System.Net.NetworkInformation.NetworkInterface]::GetIsNetworkAvailable()) { return $false }
        # Check internet quickly against a highly available Microsoft ping endpoint
        $req = [System.Net.WebRequest]::Create("http://www.msftconnecttest.com/connecttest.txt")
        $req.Timeout = 2500
        $res = $req.GetResponse()
        $res.Close()
        return $true
    } catch {
        return $false
    }
}

# Start Background Network Monitor (Continuously polls every 3 seconds safely in the background)
$netPollerBlock = {
    param($sync)
    while ($true) {
        try {
            if (-not [System.Net.NetworkInformation.NetworkInterface]::GetIsNetworkAvailable()) {
                $sync.IsOnline = $false
            } else {
                $req = [System.Net.WebRequest]::Create("http://www.msftconnecttest.com/connecttest.txt")
                $req.Timeout = 3000
                $res = $req.GetResponse()
                $res.Close()
                $sync.IsOnline = $true
            }
        } catch {
            $sync.IsOnline = $false
        }
        Start-Sleep -Seconds 3
    }
}
$script:netPollerPS = [PowerShell]::Create().AddScript($netPollerBlock).AddArgument($syncHash)
$script:netPollerPS.BeginInvoke() | Out-Null

function Set-OfflineMode {
    if ($script:IsOnline) {
        $NetworkStatusText.Text = "Online"
        $NetworkStatusText.Foreground = "#73D96B"
        $NetworkStatusText.ToolTip = "Connected to the internet."
        
        # Re-enable UI elements if it was previously offline
        if (-not $script:WingetMissing) {
            $DiscoverTab.Header = "Discover / Install"
            $DiscoverTab.IsEnabled = $true
            $UtilWingetRepairBtn.IsEnabled = $true -and $isActualAdmin
        }
        $SearchBtn.Content = "Search"
        $SearchBtn.IsEnabled = $true
        
        $InstallBtn.Content = "Install Queued Packages"
        $InstallBtn.IsEnabled = $true
        $InstallBtn.Foreground = "#73D96B"
        
        $UpdateBtn.Content = "Update Selected"
        $UpdateBtn.IsEnabled = $true
        $UpdateBtn.Foreground = "#6BA4FF"
        
        $UpdateAllBtn.Content = "Update All Apps"
        $UpdateAllBtn.IsEnabled = $true
        $UpdateAllBtn.Foreground = "#73D96B"
        
        $RefreshInstalledBtn.Content = "Load Installed & Check Updates"
        
        $UtilInstallChocoBtn.Content = "Install Chocolatey"
        $UtilInstallChocoBtn.IsEnabled = $true -and $isActualAdmin
        
        $UtilDriverBtn.Content = "Official Microsoft Drivers"
        $UtilDriverBtn.IsEnabled = $true -and $isActualAdmin
        
        $UtilSDIOBtn.Content = "Snappy Driver Installer (SDIO)"
        $UtilSDIOBtn.IsEnabled = $true -and $isActualAdmin
        
        $UtilWingetRepairBtn.Content = "Repair Winget Sources"
    } else {
        $NetworkStatusText.Text = "Offline Mode"
        $NetworkStatusText.Foreground = "#FF6B6B"
        $NetworkStatusText.ToolTip = "No internet connection detected. Online features are disabled."
        
        # Disable Online Features and Append "(Offline)"
        $DiscoverTab.Header = "Discover (Offline)"
        $DiscoverTab.IsEnabled = $false
        
        # If the user is currently on the Discover tab when offline hits, switch to Installed tab
        if ($MainTabs.SelectedItem -eq $DiscoverTab -or $MainTabs.SelectedIndex -eq 0) {
            $MainTabs.SelectedIndex = 1
        }
        
        $SearchBtn.Content = "Search (Offline)"
        $SearchBtn.IsEnabled = $false
        
        $InstallBtn.Content = "Install Queued (Offline)"
        $InstallBtn.IsEnabled = $false
        $InstallBtn.Foreground = "#888888"
        
        $UpdateBtn.Content = "Update Selected (Offline)"
        $UpdateBtn.IsEnabled = $false
        $UpdateBtn.Foreground = "#888888"
        
        $UpdateAllBtn.Content = "Update All (Offline)"
        $UpdateAllBtn.IsEnabled = $false
        $UpdateAllBtn.Foreground = "#888888"
        
        $RefreshInstalledBtn.Content = "Load Installed (Offline)"
        
        $UtilInstallChocoBtn.Content = "Install Chocolatey (Offline)"
        $UtilInstallChocoBtn.IsEnabled = $false
        
        $UtilDriverBtn.Content = "Official MS Drivers (Offline)"
        $UtilDriverBtn.IsEnabled = $false
        
        $UtilSDIOBtn.Content = "Snappy Driver (Offline)"
        $UtilSDIOBtn.IsEnabled = $false
        
        $UtilWingetRepairBtn.Content = "Repair Winget Sources (Offline)"
        $UtilWingetRepairBtn.IsEnabled = $false
    }
}

# --- Auto-Load Installed Apps on Startup ---
$Window.Add_Loaded({
    $timer.Start() # Ensure the global UI timer starts immediately and runs continuously
    Read-OptimizeStates # Read initial privacy, custom, and performance states
    
    # Initial synchronous check so UI loads correctly before poller kicks in
    $script:IsOnline = Test-InternetConnection
    $syncHash.IsOnline = $script:IsOnline
    Set-OfflineMode  # Check network and apply UI changes
    
    # Switch to 'Installed & Updates' tab automatically if starting in Offline Mode
    if (-not $script:IsOnline) {
        $MainTabs.SelectedIndex = 1
    }
    
    $loadMsg = if ($script:IsOnline) { "Loading installed packages and checking for updates... This might take a moment." } else { "Loading installed packages locally (Offline Mode)..." }
    Start-WingetJob -Action "Installed" -Query "" -Id "" -StatusMsg $loadMsg
})

# 6. Show the Window and clean up when closed
$Window.ShowDialog() | Out-Null

# Cleanup memory once the window is closed
if ($script:netPollerPS) {
    $script:netPollerPS.Stop()
    $script:netPollerPS.Dispose()
}
$runspace.Close()
$runspace.Dispose()