# Console Hider - Needed because Connect-ExchangeOnline hides parameters if it doesn't see a ConsoleHost
try {
    $consoleCode = @"
    using System;
    using System.Runtime.InteropServices;
    public class WinConsole {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();
        public static void Hide() {
            var handle = GetConsoleWindow();
            if (handle != IntPtr.Zero) ShowWindow(handle, 0); // SW_HIDE = 0
        }
    }
"@
    Add-Type -TypeDefinition $consoleCode -ErrorAction SilentlyContinue
    [WinConsole]::Hide()
}
catch {}

# Required Assemblies (v3.0.0+ required for Standalone EXE GUI)
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
}
"@

# Required Modules Check
$module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
if (-not $module -or $module.Version -lt [version]"3.0.0") {
    $msg = if (-not $module) { "ExchangeOnlineManagement module is missing." } else { "ExchangeOnlineManagement version $($module.Version) is too old for standalone EXE mode. v3.0.0+ is required." }
    $response = [System.Windows.MessageBox]::Show("$msg `n`nInstall/Update now?", "Module Update Required", "YesNo", "Warning")
    if ($response -eq "Yes") {
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser
            }
            Install-Module -Name ExchangeOnlineManagement -Force -Scope CurrentUser -AllowClobber -Confirm:$false
            $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement | Sort-Object Version -Descending | Select-Object -First 1
        }
        catch {
            [System.Windows.MessageBox]::Show("Installation failed: $($_.Exception.Message)", "Error", "OK", "Error")
            return
        }
    }
    else { return }
}

# Force import the latest version
Import-Module ExchangeOnlineManagement -MinimumVersion 3.0.0



# Load XAML (Native OS Look)
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="O365 AutoReply Manager by NetronIC (v 1.0)" Height="700" Width="1000"
        WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Height" Value="24"/>
            <Setter Property="Padding" Value="10,0"/>
            <Setter Property="Margin" Value="2"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="4,0"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Margin" Value="2"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Height" Value="24"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Margin" Value="2"/>
        </Style>
        <Style TargetType="DatePicker">
            <Setter Property="Height" Value="24"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Margin" Value="2"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Padding" Value="4,0"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/> <!-- Connection bar -->
            <RowDefinition Height="*"/>    <!-- Main View -->
            <RowDefinition Height="Auto"/> <!-- Footer -->
        </Grid.RowDefinitions>

        <!-- 1. Connection Header -->
        <GroupBox Header="Server Connection" Grid.Row="0" Margin="0,0,0,10">
            <DockPanel LastChildFill="True">
                <Button Name="btnConnect" Content="Connect to Exchange" DockPanel.Dock="Left" Width="160"/>
                <Button Name="btnHelp" Content="Help / About" DockPanel.Dock="Right" Width="100"/>
                <Label Name="lblStatus" Content="Not Connected" Margin="10,0,0,0" Foreground="Red"/>
            </DockPanel>
        </GroupBox>

        <!-- 2. Split Main View -->
        <Grid Grid.Row="1" Name="grpMain" IsEnabled="False">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="400" MinWidth="300"/> <!-- Compact List -->
                <ColumnDefinition Width="5"/>   <!-- Splitter -->
                <ColumnDefinition Width="*"/>   <!-- Detail -->
            </Grid.ColumnDefinitions>

            <!-- Left: Search & List -->
            <GroupBox Header="Users" Grid.Column="0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/> <!-- Search Box -->
                        <RowDefinition Height="Auto"/> <!-- Action Buttons -->
                        <RowDefinition Height="*"/>    <!-- DataGrid -->
                    </Grid.RowDefinitions>
                    
                    <TextBox Name="txtSearch" Grid.Row="0" Margin="0,0,0,5" Height="24"/>
                    
                    <UniformGrid Grid.Row="1" Columns="4" Margin="0,0,0,5">
                        <Button Name="btnSearch" Content="Search"/> 
                        <Button Name="btnCancelSearch" Content="Cancel" IsEnabled="False"/>
                        <Button Name="btnExport" Content="Export CSV"/>
                        <Button Name="btnImport" Content="Import CSV"/>
                    </UniformGrid>
                    
                    <DataGrid Name="dgUsers" Grid.Row="2" AutoGenerateColumns="False" IsReadOnly="True" 
                               SelectionMode="Extended" HeadersVisibility="Column" GridLinesVisibility="Horizontal">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Name" Binding="{Binding Name}" Width="*"/>
                            <DataGridTextColumn Header="Email" Binding="{Binding Email}" Width="140"/>
                            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="75"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </GroupBox>

            <GridSplitter Grid.Column="1" HorizontalAlignment="Stretch" VerticalAlignment="Stretch" Background="Transparent"/>

            <!-- Right: Settings -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/> <!-- Config Group -->
                    <RowDefinition Height="*"/>    <!-- Message Group -->
                    <RowDefinition Height="Auto"/> <!-- Bottom Actions -->
                </Grid.RowDefinitions>

                <GroupBox Header="Configuration" Grid.Row="0" Name="grpConfig" IsEnabled="False" Margin="5,0,0,0">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/> <!-- Radio buttons -->
                            <ColumnDefinition Width="*"/>    <!-- Schedule -->
                        </Grid.ColumnDefinitions>
                        
                        <StackPanel Grid.Column="0" Margin="0,0,10,0">
                            <RadioButton Name="rbDisabled" Content="Disabled" GroupName="Status" Margin="0,2"/>
                            <RadioButton Name="rbEnabled" Content="Enabled (Always)" GroupName="Status" Margin="0,2"/>
                            <RadioButton Name="rbScheduled" Content="Scheduled" GroupName="Status" Margin="0,2"/>
                        </StackPanel>

                        <Border Grid.Column="1" BorderThickness="1,0,0,0" BorderBrush="LightGray" Padding="10,0,0,0" Name="pnlSchedule" IsEnabled="False">
                            <Grid VerticalAlignment="Center">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/> <!-- Label -->
                                    <ColumnDefinition Width="Auto"/> <!-- Date -->
                                    <ColumnDefinition Width="Auto"/> <!-- HH -->
                                    <ColumnDefinition Width="Auto"/> <!-- MM -->
                                    <ColumnDefinition Width="*"/>    <!-- Spacer -->
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="30"/>
                                    <RowDefinition Height="30"/>
                                </Grid.RowDefinitions>
                                
                                <Label Content="Start:" Grid.Row="0" Grid.Column="0"/>
                                <DatePicker Name="dpStartDate" Grid.Row="0" Grid.Column="1" Width="140"/>
                                <ComboBox Name="cbStartHour" Grid.Row="0" Grid.Column="2" Width="60" Margin="4,0,0,0"/>
                                <ComboBox Name="cbStartMin" Grid.Row="0" Grid.Column="3" Width="60" Margin="4,0,0,0"/>
                                
                                <Label Content="End:" Grid.Row="1" Grid.Column="0"/>
                                <DatePicker Name="dpEndDate" Grid.Row="1" Grid.Column="1" Width="140"/>
                                <ComboBox Name="cbEndHour" Grid.Row="1" Grid.Column="2" Width="60" Margin="4,0,0,0"/>
                                <ComboBox Name="cbEndMin" Grid.Row="1" Grid.Column="3" Width="60" Margin="4,0,0,0"/>
                            </Grid>
                        </Border>
                    </Grid>
                </GroupBox>

                <GroupBox Header="Messages" Grid.Row="1" Name="grpMessages" IsEnabled="False" Margin="5,5,0,0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/> <!-- Checkbox sync -->
                            <RowDefinition Height="*"/>    <!-- TextBox/Tabs -->
                        </Grid.RowDefinitions>
                        
                        <CheckBox Name="chkSyncMsg" Content="Use same message for both (Sync)" Margin="0,0,0,5"/>
                        
                        <TextBox Name="txtCommonMsg" Grid.Row="1" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Visibility="Collapsed" Height="Auto" VerticalContentAlignment="Top" Padding="6"/>
                        
                        <TabControl Name="tabMessages" Grid.Row="1">
                            <TabItem Header="Internal">
                                <TextBox Name="txtInternalMsg" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" BorderThickness="0" Height="Auto" VerticalContentAlignment="Top" Padding="6"/>
                            </TabItem>
                            <TabItem Header="External">
                                <TextBox Name="txtExternalMsg" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" BorderThickness="0" Height="Auto" VerticalContentAlignment="Top" Padding="6"/>
                            </TabItem>
                        </TabControl>
                    </Grid>
                </GroupBox>
                
                <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,5,0,0">
                     <Button Name="btnClear" Content="Disable &amp; Clear" Width="120" Margin="5,0"/>
                     <Button Name="btnSave" Content="Save Changes" Width="120"/>
                </StackPanel>
            </Grid>
        </Grid>

        <!-- 3. Footer -->
        <StatusBar Grid.Row="2" Margin="0,5,0,0">
            <StatusBarItem>
                <TextBlock Name="lblActionStatus" Text="Ready"/>
            </StatusBarItem>
            <Separator HorizontalAlignment="Left" Width="1" Visibility="Hidden"/>
            <StatusBarItem HorizontalAlignment="Right">
                <ProgressBar Name="pbStatus" Width="100" Height="10" Visibility="Hidden"/>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
"@
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get Controls
$btnConnect = $window.FindName("btnConnect")
$btnHelp = $window.FindName("btnHelp")
$lblStatus = $window.FindName("lblStatus")
$grpMain = $window.FindName("grpMain")

# Search Controls
$txtSearch = $window.FindName("txtSearch")
$btnSearch = $window.FindName("btnSearch")
$btnCancelSearch = $window.FindName("btnCancelSearch")
$btnExport = $window.FindName("btnExport")
$btnImport = $window.FindName("btnImport")
$dgUsers = $window.FindName("dgUsers")

# Details Controls
$grpConfig = $window.FindName("grpConfig")
$rbDisabled = $window.FindName("rbDisabled")
$rbEnabled = $window.FindName("rbEnabled")
$rbScheduled = $window.FindName("rbScheduled")
$pnlSchedule = $window.FindName("pnlSchedule")
$dpStartDate = $window.FindName("dpStartDate")
$dpEndDate = $window.FindName("dpEndDate")
$cbStartHour = $window.FindName("cbStartHour")
$cbStartMin = $window.FindName("cbStartMin")
$cbEndHour = $window.FindName("cbEndHour")
$cbEndMin = $window.FindName("cbEndMin")

$grpMessages = $window.FindName("grpMessages")
$chkSyncMsg = $window.FindName("chkSyncMsg")
$tabMessages = $window.FindName("tabMessages")
# RichTextBoxes
$txtCommonMsg = $window.FindName("txtCommonMsg")
$txtInternalMsg = $window.FindName("txtInternalMsg")
$txtExternalMsg = $window.FindName("txtExternalMsg")

$btnSave = $window.FindName("btnSave")
$btnClear = $window.FindName("btnClear")
$lblActionStatus = $window.FindName("lblActionStatus")
$pbStatus = $window.FindName("pbStatus")

# Populate Time ComboBoxes
0..23 | ForEach-Object { 
    $val = "{0:D2}" -f $_
    $cbStartHour.Items.Add($val) | Out-Null
    $cbEndHour.Items.Add($val) | Out-Null
}
0..59 | ForEach-Object { 
    $val = "{0:D2}" -f $_
    $cbStartMin.Items.Add($val) | Out-Null
    $cbEndMin.Items.Add($val) | Out-Null
}
# Set defaults
$cbStartHour.SelectedIndex = 8
$cbStartMin.SelectedIndex = 0
$cbEndHour.SelectedIndex = 17
$cbEndMin.SelectedIndex = 0

# Variables
$SelectedUser = $null
$hasUnsavedChanges = $false
$previousSelection = $null
$suppressSelectionChanged = $false
$suppressListRefresh = $false
$cancelSearch = $false

# Helper Functions
function Set-Status {
    param($Text, $Color = "Black", $IsWorking = $false, $Percent = -1)
    $lblActionStatus.Text = $Text
    if ($IsWorking) {
        $pbStatus.Visibility = "Visible"
        if ($Percent -ge 0) {
            $pbStatus.IsIndeterminate = $false
            $pbStatus.Value = $Percent
            $pbStatus.Maximum = 100
        }
        else {
            $pbStatus.IsIndeterminate = $true
        }
    }
    else {
        $pbStatus.Visibility = "Hidden"
        $pbStatus.IsIndeterminate = $true # Reset for next time
    }
}

function Clear-Details {
    $rbDisabled.IsChecked = $false
    $rbEnabled.IsChecked = $false
    $rbScheduled.IsChecked = $false
    $dpStartDate.SelectedDate = $null
    $dpEndDate.SelectedDate = $null
    $txtInternalMsg.Text = ""
    $txtExternalMsg.Text = ""
    $txtCommonMsg.Text = ""
    $chkSyncMsg.IsChecked = $false
    
    # Reset view to tabs
    $tabMessages.Visibility = "Visible"
    $txtCommonMsg.Visibility = "Collapsed"
    
    $grpConfig.IsEnabled = $false
    $grpMessages.IsEnabled = $false
    $btnSave.IsEnabled = $false
    $btnClear.IsEnabled = $false
}

# Event Handlers: Message Sync
$chkSyncMsg.Add_Checked({
        # Switch to Common View
        $tabMessages.Visibility = "Collapsed"
        $txtCommonMsg.Visibility = "Visible"
    
        # Copy Internal to Common (assuming Internal is primary source)
        $txtCommonMsg.Text = $txtInternalMsg.Text
    })

$chkSyncMsg.Add_Unchecked({
        # Switch back to Tabs
        $tabMessages.Visibility = "Visible"
        $txtCommonMsg.Visibility = "Collapsed"
    
        # Copy Common back to both (or just keep what was there? usually better to copy back)
        $txtInternalMsg.Text = $txtCommonMsg.Text
        $txtExternalMsg.Text = $txtCommonMsg.Text
    })

$txtInternalMsg.Add_TextChanged({
        if ($chkSyncMsg.IsChecked) {
            $txtExternalMsg.Text = $txtInternalMsg.Text
        }
        if ($grpConfig.IsEnabled) { $script:hasUnsavedChanges = $true }
    })

$txtExternalMsg.Add_TextChanged({
        if ($grpConfig.IsEnabled) { $script:hasUnsavedChanges = $true }
    })

$txtCommonMsg.Add_TextChanged({
        if ($grpConfig.IsEnabled) { $script:hasUnsavedChanges = $true }
    })

$rbDisabled.Add_Checked({ if ($grpConfig.IsEnabled) { $script:hasUnsavedChanges = $true } })
$rbEnabled.Add_Checked({ if ($grpConfig.IsEnabled) { $script:hasUnsavedChanges = $true } })
$rbScheduled.Add_Checked({ if ($grpConfig.IsEnabled) { $script:hasUnsavedChanges = $true } })

$dpStartDate.Add_SelectedDateChanged({ if ($grpConfig.IsEnabled) { $script:hasUnsavedChanges = $true } })
$dpEndDate.Add_SelectedDateChanged({ if ($grpConfig.IsEnabled) { $script:hasUnsavedChanges = $true } })
$cbStartHour.Add_SelectionChanged({ if ($grpConfig.IsEnabled) { $script:hasUnsavedChanges = $true } })
$cbStartMin.Add_SelectionChanged({ if ($grpConfig.IsEnabled) { $script:hasUnsavedChanges = $true } })
$cbEndHour.Add_SelectionChanged({ if ($grpConfig.IsEnabled) { $script:hasUnsavedChanges = $true } })
$cbEndMin.Add_SelectionChanged({ if ($grpConfig.IsEnabled) { $script:hasUnsavedChanges = $true } })
$chkSyncMsg.Add_Checked({ if ($grpConfig.IsEnabled) { $script:hasUnsavedChanges = $true } })
$chkSyncMsg.Add_Unchecked({ if ($grpConfig.IsEnabled) { $script:hasUnsavedChanges = $true } })

function Search-Users {
    param($searchText)

    $script:cancelSearch = $false
    $btnSearch.IsEnabled = $false
    $btnCancelSearch.IsEnabled = $true

    try {
        Set-Status "Searching users..." "Black" $true
        [System.Windows.Forms.Application]::DoEvents()
    
        $users = @()

        if ([string]::IsNullOrWhiteSpace($searchText)) {
            # Limit explicitly set to 500
            $users = @(Get-ExoMailbox -ResultSize 500 -Properties UserPrincipalName, DisplayName, RecipientTypeDetails -RecipientTypeDetails UserMailbox, SharedMailbox)
        }
        else {
            # Limit explicitly set to 500
            $users = @(Get-ExoMailbox -Filter "DisplayName -like '*$searchText*' -or UserPrincipalName -like '*$searchText*'" -ResultSize 500 -Properties UserPrincipalName, DisplayName, RecipientTypeDetails -RecipientTypeDetails UserMailbox, SharedMailbox)
        }

        $userList = @()
        $totalUsers = $users.Count
        $i = 0

        foreach ($u in $users) {
            if ($script:cancelSearch) {
                Set-Status "Search cancelled by user."
                break
            }

            $i++
            $pct = 0
            if ($totalUsers -gt 0) {
                $pct = [int](($i / $totalUsers) * 100)
            }
            Set-Status "Processing user $i of $totalUsers (Limit: 500)..." "Black" $true $pct
            [System.Windows.Forms.Application]::DoEvents()

            try {
                $oof = Get-MailboxAutoReplyConfiguration -Identity $u.UserPrincipalName -ErrorAction Stop
                $status = $oof.AutoReplyState
                $internal = $oof.InternalMessage
                $external = $oof.ExternalMessage
                $start = $oof.StartTime
                $end = $oof.EndTime
            }
            catch {
                $status = "Error"
                $internal = ""
                $external = ""
                $start = $null
                $end = $null
            }

            $userList += [PSCustomObject]@{
                Name     = $u.DisplayName
                Email    = $u.UserPrincipalName
                Type     = $u.RecipientTypeDetails
                Status   = $status
                Internal = $internal
                External = $external
                Start    = $start
                End      = $end
            }
        }

        $dgUsers.ItemsSource = $userList
        if ($script:cancelSearch) {
            Set-Status "Search cancelled. Showing $i results."
        }
        else {
            Set-Status "Found $($userList.Count) users."
        }
    }
    catch {
        [System.Windows.MessageBox]::Show("Search failed: $_", "Error", "OK", "Error")
        Set-Status "Search failed."
    }
    finally {
        $btnSearch.IsEnabled = $true
        $btnCancelSearch.IsEnabled = $false
    }
}

$btnConnect.Add_Click({
        try {
            Set-Status "Connecting..." "Black" $false
            [System.Windows.Forms.Application]::DoEvents()
        
            # 1. Force reload module
            $modName = "ExchangeOnlineManagement"
            $m = Import-Module $modName -Force -PassThru -ErrorAction SilentlyContinue
            $modPath = if ($m) { $m.Path } else { "Not found" }

            # 2. Get Cmdlet Info
            $exoCmd = Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue | Select-Object -First 1
            $modVer = if ($exoCmd.Module) { $exoCmd.Module.Version } else { "Unknown" }
            $params = if ($exoCmd) { $exoCmd.Parameters } else { @{} }
            $pNames = ($params.Keys | Sort-Object) -join ", "

            # 3. Get Window Handle
            $ptrHandle = [System.IntPtr]::Zero
            try {
                $helper = New-Object System.Windows.Interop.WindowInteropHelper($window)
                $ptrHandle = $helper.EnsureHandle()
            }
            catch {}
            if ($ptrHandle -eq [System.IntPtr]::Zero) { $ptrHandle = [Win32]::GetForegroundWindow() }

            $connected = $false
            $lastError = ""

            # Strategy Tiers (v13 - Console-Aware)
            try {
                # Tier 1: Legacy Browser (No WAM) with Handle - MOST STABLE
                Set-Status "Trying Tier 1: Legacy Browser..."
                $splat = @{ ErrorAction = 'Stop'; ShowBanner = $false }
                if ($params.ContainsKey('ParentWindowHandle')) { $splat['ParentWindowHandle'] = $ptrHandle }
                if ($params.ContainsKey('UseWebAccountManager')) { $splat['UseWebAccountManager'] = $false }
            
                Connect-ExchangeOnline @splat
                $connected = $true
            }
            catch {
                $lastError = $_.Exception.Message
                try {
                    # Tier 2: Basic Interactive (Fallback)
                    Set-Status "Trying Tier 2: Basic Interactive..."
                    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
                    $connected = $true
                }
                catch {
                    $lastError = $_.Exception.Message
                    # Tier 3: Device Code (Fallback)
                    if ($params.ContainsKey('DeviceCode')) {
                        $msg = "Interactive login failed: $lastError`n`nWould you like to try Device Code (browser code) login?"
                        if ([System.Windows.MessageBox]::Show($msg, "Connection Strategy Fallback", "YesNo", "Question") -eq [System.Windows.MessageBoxResult]::Yes) {
                            Connect-ExchangeOnline -DeviceCode -ShowBanner:$false -ErrorAction Stop
                            $connected = $true
                        }
                    }
                    else {
                        $helpMsg = "Interactive login failed and modern parameters are missing.`n`nIMPORTANT: If you compiled with '-noConsole', try re-compiling WITHOUT it. v13 will auto-hide the console for you.`n`nError: $lastError"
                        throw $helpMsg
                    }
                }
            }
    
            if (-not $connected) { return }
    
            # Get Tenant Name
            $tenantName = "Office 365 Tenant"
            try {
                $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
                if ($orgConfig) {
                    $tenantName = if ($orgConfig.DisplayName) { $orgConfig.DisplayName } else { $orgConfig.Name }
                }
            }
            catch { }

            $lblStatus.Content = "Connected to: $tenantName"
            $lblStatus.Foreground = [System.Windows.Media.Brushes]::Green
            $btnConnect.IsEnabled = $false
            $grpMain.IsEnabled = $true
            Set-Status "Connected successfully to $tenantName."
        }
        catch {
            $errorMsg = if ($_.Exception.Message) { $_.Exception.Message } else { $_.ToString() }
            $psVer = $PSVersionTable.PSVersion.ToString()
            $arch = if ([IntPtr]::Size -eq 8) { "64-bit" } else { "32-bit" }
            $diag = "`n`n--- Diagnostics (v13) ---`nHost: $($Host.Name)`nPS: $psVer ($arch)`nModule: $modVer`nPath: $modPath`nHandle: $ptrHandle`nParams: $pNames"
            [System.Windows.MessageBox]::Show("Connection failed:`n`n$errorMsg $diag", "Error", "OK", "Error")
            Set-Status "Connection failed."
        }
    })

$btnHelp.Add_Click({
        try {
            # Bilingual User Guide content integrated directly into the script
            $helpContentHU = @'
# O365 AutoReply Manager - Felhasználói Leírás (v1.0)

Ez a dokumentum az **O365 AutoReply Manager** alkalmazás használatát mutatja be, amely lehetővé teszi Microsoft 365 (Exchange Online) postaládák automatikus válaszainak (OOF) egyszerű, grafikus felületen történő kezelését.

---

## 1. Rendszerkövetelmények és Előkészületek

Az alkalmazás futtatásához az alábbiak szükségesek:
- **Windows operációs rendszer**: PowerShell 5.1 vagy PowerShell 7+ támogatással.
- **ExchangeOnlineManagement modul**: Legalább **v3.0.0** vagy újabb verzió.
- **Internetkapcsolat**: Az Office 365 felhőhöz való csatlakozáshoz.
- **Megfelelő jogosultság**: Olyan fiók, amely rendelkezik jogosultsággal a szervezet postaládáinak módosításához (pl. Exchange Administrator).

---

## 2. Indítás és Csatlakozás

1. Indítsa el az `O365Manager.exe` fájlt.
2. Kattintson a bal felső sarokban található **"Connect to O365"** gombra.
3. Bejelentkezés:
   - Megnyílik egy Microsoft bejelentkező ablak. Adja meg hitelesítő adatait.
   - Ha a "Tier 1" bejelentkezés sikertelen, az alkalmazás alternatív módokat (pl. Device Code) is felajánlhat.
4. Sikeres kapcsolat esetén az állapotjelző zöldre vált, és megjelenik a bérlő (Tenant) neve.

---

## 3. Felhasználók Keresése

- A bal oldali **"Users"** panelen kereshet a szervezet felhasználói között.
- Írja be a nevet vagy az e-mail címet a keresőmezőbe, majd nyomja meg a **"Search"** gombot.
- Az alkalmazás egyszerre maximum 500 találatot jelenít meg a teljesítmény optimalizálása érdekében.
- A táblázatban látható a felhasználók neve, e-mail címe és jelenlegi automatikus válasz állapota.

---

## 4. Automatikus Válaszok Beállítása

Válasszon ki egy felhasználót a listából. A jobb oldali panelen betöltődnek a jelenlegi beállításai.

### Válasz Állapota (Status)
- **Disabled**: Alapértelmezett állapot, az automatikus válasz ki van kapcsolva.
- **Enabled (Always)**: Az automatikus válasz azonnal és határozatlan ideig bekapcsol.
- **Scheduled**: Időzített válasz. Ebben a módban meg kell adni a kezdő és záró dátumot, valamint az időpontot (óra:perc).

### Üzenetek Szerkesztése
- **Internal**: A szervezeten belüli munkatársaknak küldött üzenet.
- **External**: Külső feladóknak (pl. ügyfelek) küldött üzenet.
- **Sync funkció**: Ha a **"Use same message for both (Sync)"** opciót bepipálja, a két üzenet mező összevonásra kerül, így ugyanazt az üzenetet kapja mindenki.

### Műveletek
- **Save Changes**: Mentés. A módosítások csak akkor lépnek érvénybe, ha erre a gombra kattint.
- **Disable & Clear**: Egy kattintással kikapcsolja az automatikus választ és törli az üzenetek tartalmát.

---

## 5. Csoportos Műveletek (Batch Mode)

Az alkalmazás támogatja több felhasználó egyszerre történő kezelését is:
1. Jelöljön ki több felhasználót a listában (Shift vagy Ctrl billentyűk segítségével).
2. A jobb oldali panelen a jelenlegi beállítások nem fognak látszódni, de az új beállításokat megadhatja.
3. A **"Save Changes"** gombra kattintva a megadott módosítások minden kijelölt felhasználóra érvényesülnek.

---

## 6. Adatok Exportálása és Importálása (CSV)

### Exportálás
A **"Export CSV"** gombbal kimentheti a jelenleg listázott összes felhasználó beállításait egy .csv fájlba. Ez hasznos lehet archiváláshoz vagy jelentéskészítéshez.

### Importálás
A **"Import CSV"** gombbal tömegesen állíthat be automatikus válaszokat egy fájl alapú lista alapján.
- A CSV fájlnak tartalmaznia kell egy `Email` oszlopot.
- További támogatott oszlopok: `Status`, `Internal`, `External`, `Start`, `End`.
- Az importálás felülírja a célfelhasználók jelenlegi beállításait.

---

## 7. Hibaelhárítás

- **"Module mismatch"**: Ha az alkalmazás hibát jelez a modul verziója miatt, futtassa a következő parancsot rendszergazdai PowerShellben:
  `Install-Module -Name ExchangeOnlineManagement -Force`
- **Kapcsolódási hiba**: Ellenőrizze, hogy nincs-e tűzfal vagy proxy, amely blokkolja az Office 365 elérését.
- **Nem látszódik minden felhasználó**: A keresőmező használatával szűkítse le a találatokat, ha túl sok felhasználó van a szervezetben.

---
*Készítette: NetronIC*
'@

            $helpContentEN = @'
# O365 AutoReply Manager - User Guide (v1.0)

This document describes the use of the **O365 AutoReply Manager** application, which allows for easy management of Microsoft 365 (Exchange Online) mailbox Automatic Replies (OOF) through a graphical interface.

---

## 1. System Requirements and Prerequisites

To run the application, the following are required:
- **Windows OS**: With support for PowerShell 5.1 or PowerShell 7+.
- **ExchangeOnlineManagement module**: Version **v3.0.0** or newer.
- **Internet Connection**: For connecting to the Office 365 cloud.
- **Proper Permissions**: An account with permission to modify organization mailboxes (e.g., Exchange Administrator).

---

## 2. Start and Connection

1. Launch the `O365Manager.exe` file.
2. Click the **"Connect to O365"** button in the top left corner.
3. Login:
   - A Microsoft login window will open. Enter your credentials.
   - If "Tier 1" login fails, the application may offer alternative modes (e.g., Device Code).
4. Upon successful connection, the status indicator turns green, and the tenant name appears.

---

## 3. Searching for Users

- Search for organization users in the left **"Users"** panel.
- Enter a name or email address in the search box and press the **"Search"** button.
- The application displays a maximum of 500 results at a time to optimize performance.
- The table shows user names, email addresses, and current automatic reply status.

---

## 4. Setting Up Automatic Replies

Select a user from the list. Their current settings will load in the right panel.

### Reply Status
- **Disabled**: Default state, automatic reply is turned off.
- **Enabled (Always)**: Automatic reply turns on immediately and indefinitely.
- **Scheduled**: Timed reply. In this mode, you must specify start and end dates and times (HH:MM).

### Editing Messages
- **Internal**: Message sent to colleagues within the organization.
- **External**: Message sent to external senders (e.g., clients).
- **Sync function**: If you check the **"Use same message for both (Sync)"** option, the two message fields are merged, sending the same message to everyone.

### Actions
- **Save Changes**: Save. Changes only take effect when you click this button.
- **Disable & Clear**: Disables automatic replies and clears message content with one click.

---

## 5. Batch Operations (Batch Mode)

The application supports managing multiple users simultaneously:
1. Select multiple users in the list (using Shift or Ctrl keys).
2. Existing settings will not be shown in the right panel, but you can specify new settings.
3. Clicking **"Save Changes"** applies the changes to all selected users.

---

## 6. Exporting and Importing Data (CSV)

### Exporting
Use the **"Export CSV"** button to save settings for all currently listed users into a .csv file. This is useful for archiving or reporting.

### Importing
The **"Import CSV"** button allows for bulk setting of automatic replies based on a file-based list.
- The CSV file must contain an `Email` column.
- Other supported columns: `Status`, `Internal`, `External`, `Start`, `End`.
- Importing overwrites the current settings for target users.

---

## 7. Troubleshooting

- **"Module mismatch"**: If the application reports an error due to the module version, run the following command in an administrator PowerShell:
  `Install-Module -Name ExchangeOnlineManagement -Force`
- **Connection Error**: Check if there's a firewall or proxy blocking access to Office 365.
- **Not all users visible**: Use the search box to narrow down results if there are many users in the organization.

---
*Created by: NetronIC*
'@

            # Bilingual scrollable window for help (Tabs)
            $helpXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Help / About - O365 AutoReply Manager" Height="650" Width="850" WindowStartupLocation="CenterOwner">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TabControl Grid.Row="0">
            <TabItem Header="Magyar (HU)">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <TextBlock Name="txtHelpContentHU" TextWrapping="Wrap" Padding="15" FontFamily="Consolas" FontSize="12" />
                </ScrollViewer>
            </TabItem>
            <TabItem Header="English (EN)">
                <ScrollViewer VerticalScrollBarVisibility="Auto">
                    <TextBlock Name="txtHelpContentEN" TextWrapping="Wrap" Padding="15" FontFamily="Consolas" FontSize="12" />
                </ScrollViewer>
            </TabItem>
        </TabControl>
        
        <Button Name="btnCloseHelp" Content="Close / Bezárás" Grid.Row="1" Width="120" Height="28" HorizontalAlignment="Right" Margin="0,10,0,0"/>
    </Grid>
</Window>
"@
            $helpReader = (New-Object System.Xml.XmlNodeReader ([xml]$helpXaml))
            $helpWin = [Windows.Markup.XamlReader]::Load($helpReader)
            $helpWin.Owner = $window
            
            # Set content programmatically to avoid XML parsing issues
            ($helpWin.FindName("txtHelpContentHU")).Text = $helpContentHU
            ($helpWin.FindName("txtHelpContentEN")).Text = $helpContentEN
            
            $btnCloseHelp = $helpWin.FindName("btnCloseHelp")
            $btnCloseHelp.Add_Click({ $helpWin.Close() })
            
            $helpWin.ShowDialog() | Out-Null
        }
        catch {
            [System.Windows.MessageBox]::Show("Failed to load help: $_", "Error", "OK", "Error")
        }
    })
$btnSearch.Add_Click({
        Search-Users $txtSearch.Text
    })

$btnCancelSearch.Add_Click({
        $script:cancelSearch = $true
        Set-Status "Cancelling search..."
    })

$btnExport.Add_Click({
        $data = $dgUsers.ItemsSource
        if ($null -eq $data -or $data.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No data to export.", "Warning", "OK", "Warning")
            return
        }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $sfd.FileName = "O365_AutoReply_Export_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"

        if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                Set-Status "Exporting data..." "Black" $true
                [System.Windows.Forms.Application]::DoEvents()
                
                $data | Export-Csv -Path $sfd.FileName -NoTypeInformation -Encoding UTF8
                
                Set-Status "Export successful."
                [System.Windows.Forms.MessageBox]::Show("Data exported successfully to:`n$($sfd.FileName)", "Success", "OK", "Information")
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Export failed: $_", "Error", "OK", "Error")
                Set-Status "Export failed."
            }
        }
    })

$btnImport.Add_Click({
        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                Set-Status "Reading CSV file..." "Black" $true
                [System.Windows.Forms.Application]::DoEvents()

                $csvData = @(Import-Csv -Path $ofd.FileName -Encoding UTF8)
                
                if ($csvData.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show("CSV file is empty or invalid.", "Warning", "OK", "Warning")
                    Set-Status "Import cancelled (empty file)."
                    return
                }

                # Validate Headers (Basic check)
                $firstRow = $csvData[0]
                if (-not $firstRow.PSObject.Properties['Email']) {
                    [System.Windows.Forms.MessageBox]::Show("CSV must contain an 'Email' column.", "Error", "OK", "Error")
                    Set-Status "Import failed (missing Email column)."
                    return
                }

                # Confirmation
                $confirmMsg = "This will import settings for $($csvData.Count) users and OVERWRITE their current AutoReply settings.`n`nDo you want to continue?"
                $result = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "Confirm Import", "YesNo", "Warning")
                if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
                    Set-Status "Import cancelled by user."
                    return
                }

                $total = $csvData.Count
                $i = 0
                $successCount = 0
                $failCount = 0
                
                Set-Status "Importing settings for $total users..." "Black" $true
                [System.Windows.Forms.Application]::DoEvents()

                foreach ($row in $csvData) {
                    $i++
                    $pct = 0
                    if ($total -gt 0) {
                        $pct = [int](($i / $total) * 100)
                    }
                    $msg = "Importing user {0} of {1}: {2}" -f $i, $total, $row.Email
                    Set-Status $msg "Black" $true $pct
                    [System.Windows.Forms.Application]::DoEvents()

                    try {
                        $params = @{
                            Identity        = $row.Email
                            AutoReplyState  = if ($row.Status) { $row.Status } else { "Disabled" }
                            InternalMessage = if ($row.Internal) { $row.Internal } else { "" }
                            ExternalMessage = if ($row.External) { $row.External } else { "" }
                            ErrorAction     = "Stop"
                        }
                        
                        # Handle Schedule if present
                        if ($row.Status -eq "Scheduled" -and $row.Start -and $row.End) {
                            $s = $row.Start -as [DateTime]
                            $e = $row.End -as [DateTime]
                            if ($s -and $e) {
                                $params.StartTime = $s
                                $params.EndTime = $e
                            }
                        }

                        Set-MailboxAutoReplyConfiguration @params
                        $successCount++
                    }
                    catch {
                        Write-Host "Failed to import for $($row.Email): $($_)"
                        $failCount++
                    }
                }
                
                Set-Status "Import completed."
                $summaryMsg = "Import completed.`n`nSuccessful: $successCount`nFailed: $failCount`nTotal: $total"
                [System.Windows.Forms.MessageBox]::Show($summaryMsg, "Import Complete", "OK", "Information")
                
                # Refresh List
                Search-Users $txtSearch.Text
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Import failed: $_", "Error", "OK", "Error")
                Set-Status "Import failed."
            }
        }
    })

$dgUsers.Add_SelectionChanged({
        # Prevent recursive calls
        if ($script:suppressSelectionChanged) { return }

        # Check for unsaved changes before switching
        if ($script:hasUnsavedChanges) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "You have unsaved changes. Do you want to save them before switching?",
                "Unsaved Changes",
                "YesNo",
                "Warning"
            )
            
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                # Restore previous selection temporarily to save correct user
                $newSelection = $dgUsers.SelectedItem
                $script:suppressSelectionChanged = $true
                $dgUsers.SelectedItem = $script:previousSelection
                $script:suppressSelectionChanged = $false
                
                # Suppress list refresh after save
                $script:suppressListRefresh = $true
                
                # Trigger Save button click (saves the previous user)
                $btnSave.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
                
                # Restore flag
                $script:suppressListRefresh = $false
                
                # Now switch to new selection and let the event handler continue
                $script:suppressSelectionChanged = $true
                $dgUsers.SelectedItem = $newSelection
                $script:suppressSelectionChanged = $false
                # Don't return - let the handler continue to load the new user
            }
            # If No, continue without saving (flag will be reset below)
        }

        $count = $dgUsers.SelectedItems.Count

        if ($count -eq 0) {
            Clear-Details
            $SelectedUser = $null
            $script:hasUnsavedChanges = $false
            $script:previousSelection = $null
        }
        elseif ($count -gt 1) {
            # Batch Mode
            $SelectedUser = "BATCH" # Just to enable Save
            Clear-Details
        
            $grpConfig.IsEnabled = $true
            $grpMessages.IsEnabled = $true
            $btnSave.IsEnabled = $true
            $btnClear.IsEnabled = $true
        
            $script:hasUnsavedChanges = $false
            $script:previousSelection = $null
            Set-Status "Batch Mode: $count users selected. (Existing settings not shown)"
        }
        else {
            # Single User Mode
            $selectedItem = $dgUsers.SelectedItem
            $SelectedUser = $selectedItem.Email
            Set-Status "Loading settings for $SelectedUser..." "Black" $true
            [System.Windows.Forms.Application]::DoEvents()
        
            try {
                $config = Get-MailboxAutoReplyConfiguration -Identity $SelectedUser -ErrorAction Stop
        
                # Set controls
                if ($config.AutoReplyState -eq "Disabled") { $rbDisabled.IsChecked = $true }
                elseif ($config.AutoReplyState -eq "Enabled") { $rbEnabled.IsChecked = $true }
                elseif ($config.AutoReplyState -eq "Scheduled") { $rbScheduled.IsChecked = $true }

                if ($config.StartTime) {
                    $dpStartDate.SelectedDate = $config.StartTime
                    $cbStartHour.SelectedIndex = $config.StartTime.Hour
                    $cbStartMin.SelectedIndex = $config.StartTime.Minute
                }
                else { $dpStartDate.SelectedDate = $null }
            
                if ($config.EndTime) {
                    $dpEndDate.SelectedDate = $config.EndTime
                    $cbEndHour.SelectedIndex = $config.EndTime.Hour
                    $cbEndMin.SelectedIndex = $config.EndTime.Minute
                }
                else { $dpEndDate.SelectedDate = $null }

                $txtInternalMsg.Text = $config.InternalMessage
                $txtExternalMsg.Text = $config.ExternalMessage
            
                # Auto-detect sync state
                if ($config.InternalMessage -eq $config.ExternalMessage -and -not [string]::IsNullOrEmpty($config.InternalMessage)) {
                    $chkSyncMsg.IsChecked = $true
                    $txtCommonMsg.Text = $config.InternalMessage
                    $tabMessages.Visibility = "Collapsed"
                    $txtCommonMsg.Visibility = "Visible"
                }
                else {
                    $chkSyncMsg.IsChecked = $false
                    $tabMessages.Visibility = "Visible"
                    $txtCommonMsg.Visibility = "Collapsed"
                }

                $grpConfig.IsEnabled = $true
                $grpMessages.IsEnabled = $true
                $btnSave.IsEnabled = $true
                $btnClear.IsEnabled = $true
                $script:hasUnsavedChanges = $false
                $script:previousSelection = $selectedItem
                Set-Status "Settings loaded."
            }
            catch {
                Set-Status "Error loading settings: $($_.Exception.Message)"
            }
        }
    })

$rbScheduled.Add_Checked({ 
        $pnlSchedule.IsEnabled = $true 
    })

$rbScheduled.Add_Unchecked({ 
        $pnlSchedule.IsEnabled = $false 
    })

$btnSave.Add_Click({
        $selectedItems = $dgUsers.SelectedItems
        if ($selectedItems.Count -eq 0) { return }

        $state = "Disabled"
        if ($rbEnabled.IsChecked) { $state = "Enabled" }
        elseif ($rbScheduled.IsChecked) { $state = "Scheduled" }

        # Determine messages based on Sync
        $msgInternal = $txtInternalMsg.Text
        $msgExternal = $txtExternalMsg.Text
        
        if ($chkSyncMsg.IsChecked) {
            $msgInternal = $txtCommonMsg.Text
            $msgExternal = $txtCommonMsg.Text
        }

        $baseParams = @{
            AutoReplyState  = $state
            InternalMessage = $msgInternal
            ExternalMessage = $msgExternal
        }

        if ($state -eq "Scheduled") {
            if (-not $dpStartDate.SelectedDate -or -not $dpEndDate.SelectedDate -or 
                $cbStartHour.SelectedIndex -lt 0 -or $cbStartMin.SelectedIndex -lt 0 -or 
                $cbEndHour.SelectedIndex -lt 0 -or $cbEndMin.SelectedIndex -lt 0) {
                [System.Windows.MessageBox]::Show("All date and time fields must be selected for Scheduled mode.", "Validation", "OK", "Warning")
                return
            }
        
            # Combine Date + Time
            try {
                $sDate = $dpStartDate.SelectedDate.Date 
                $sHour = $cbStartHour.SelectedIndex
                $sMin = $cbStartMin.SelectedIndex
                $fullStart = $sDate.AddHours($sHour).AddMinutes($sMin)
        
                # End
                $eDate = $dpEndDate.SelectedDate.Date
                $eHour = $cbEndHour.SelectedIndex
                $eMin = $cbEndMin.SelectedIndex
                $fullEnd = $eDate.AddHours($eHour).AddMinutes($eMin)
        
                if ($fullEnd -le $fullStart) {
                    [System.Windows.MessageBox]::Show("End time must be after Start time.", "Validation", "OK", "Warning")
                    return
                }

                $baseParams.StartTime = $fullStart
                $baseParams.EndTime = $fullEnd
            }
            catch {
                [System.Windows.MessageBox]::Show("Invalid Time Selection. Error: $_", "Error", "OK", "Error")
                return
            }
        }

        try {
            Set-Status "Saving settings for $($selectedItems.Count) users..." "Black" $true
            [System.Windows.Forms.Application]::DoEvents()
        
            foreach ($item in $selectedItems) {
                $params = $baseParams.Clone()
                $params.Identity = $item.Email
                Set-MailboxAutoReplyConfiguration @params -ErrorAction Stop
            }
    
            Set-Status "Settings saved successfully for all selected users."
            [System.Windows.MessageBox]::Show("AutoReply settings saved successfully!", "Success", "OK", "Information")
            
            $script:hasUnsavedChanges = $false
            # Refresh List (unless suppressed)
            if (-not $script:suppressListRefresh) {
                Search-Users $txtSearch.Text
            }
        }
        catch {
            [System.Windows.MessageBox]::Show("Failed to save settings: $_", "Error", "OK", "Error")
            Set-Status "Save failed."
        }
    })

$btnClear.Add_Click({
        $selectedItems = $dgUsers.SelectedItems
        if ($selectedItems.Count -eq 0) { return }

        try {
            Set-Status "Disabling auto-reply for $($selectedItems.Count) users..." "Black" $true
            [System.Windows.Forms.Application]::DoEvents()
        
            foreach ($item in $selectedItems) {
                Set-MailboxAutoReplyConfiguration -Identity $item.Email -AutoReplyState Disabled -InternalMessage "" -ExternalMessage "" -ErrorAction Stop
            }
    
            # Update UI if single user, else just clear
            if ($selectedItems.Count -eq 1) {
                $rbDisabled.IsChecked = $true
                $txtInternalMsg.Text = ""
                $txtExternalMsg.Text = ""
                $txtCommonMsg.Text = ""
                $dpStartDate.SelectedDate = $null
                $dpEndDate.SelectedDate = $null
            }    

            Set-Status "AutoReply disabled and cleared."
            [System.Windows.Forms.MessageBox]::Show("AutoReply disabled and cleared.", "Success", "OK", "Information")
            
            $script:hasUnsavedChanges = $false
            # Refresh List
            Search-Users $txtSearch.Text
        }
        catch {
            Set-Status "Error clearing settings."
        }
    })

# Show Window
$window.ShowDialog() | Out-Null
