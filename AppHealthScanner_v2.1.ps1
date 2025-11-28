<#
App Health & Crash Diagnostic – Windows 10/11
#>

Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase
Add-Type -AssemblyName System.Xaml
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem

# ---------------------------
# XAML (light theme, dynamic)
# ---------------------------
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="App Health &amp; Crash Diagnostic"
        Width="1100"
        Height="700"
        MinWidth="900"
        MinHeight="500"
        WindowStartupLocation="CenterScreen"
        Background="White"
        Foreground="Black"
        FontFamily="Segoe UI">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto" />
      <RowDefinition Height="*" />
      <RowDefinition Height="Auto" />
    </Grid.RowDefinitions>

    <!-- TOP PANEL: Filters & Profile -->
    <Border Grid.Row="0" Padding="8" Margin="0,0,0,8" Background="#f0f0f0" CornerRadius="6">
      <ScrollViewer HorizontalScrollBarVisibility="Auto"
                    VerticalScrollBarVisibility="Disabled">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />
          </Grid.ColumnDefinitions>

          <!-- App filter + picker (fixed widths) -->
          <StackPanel Orientation="Vertical" Grid.Column="0" Margin="0,0,8,0">
            <TextBlock Text="App filter (name or exe)" FontWeight="Bold" Margin="0,0,0,2" />
            <DockPanel LastChildFill="False">
              <TextBox x:Name="txtAppFilter"
                       Width="260"
                       Height="26"
                       Padding="4"
                       ToolTip="Example: outlook, outlook.exe, chrome, Plex Media Server"
                       DockPanel.Dock="Left" />
              <Button x:Name="btnPickApp"
                      Content="Pick..."
                      Width="70"
                      Margin="4,0,0,0"
                      ToolTip="Open process list and pick a running app to use as the filter."
                      DockPanel.Dock="Right" />
            </DockPanel>
          </StackPanel>

          <!-- Lookback -->
          <StackPanel Orientation="Vertical" Grid.Column="1" Margin="0,0,8,0">
            <TextBlock Text="Lookback" FontWeight="Bold" Margin="0,0,0,2" />
            <ComboBox x:Name="cmbLookback" Height="26" Width="120" SelectedIndex="0">
              <ComboBoxItem Content="Last 24 hours" Tag="1" />
              <ComboBoxItem Content="Last 3 days"   Tag="3" />
              <ComboBoxItem Content="Last 7 days"   Tag="7" />
              <ComboBoxItem Content="Last 30 days"  Tag="30" />
            </ComboBox>
          </StackPanel>

          <!-- Profile -->
          <StackPanel Orientation="Vertical" Grid.Column="2" Margin="0,0,8,0">
            <TextBlock Text="Profile" FontWeight="Bold" Margin="0,0,0,2" />
            <ComboBox x:Name="cmbProfile" Height="26" Width="120" SelectedIndex="1">
              <ComboBoxItem Content="Basic User" Tag="Basic" />
              <ComboBoxItem Content="Power User" Tag="Power" />
              <ComboBoxItem Content="Sysadmin"   Tag="Admin" />
            </ComboBox>
          </StackPanel>

          <!-- Scan buttons -->
          <StackPanel Orientation="Vertical" Grid.Column="3" Margin="0,0,8,0">
            <TextBlock Text="Scan type" FontWeight="Bold" Margin="0,0,0,2" />
            <StackPanel Orientation="Horizontal">
              <Button x:Name="btnQuickScan"
                      Content="Quick Scan"
                      Margin="0,0,6,0"
                      Padding="12,4"
                      ToolTip="Quick Scan: system-wide Critical and Error events plus crashes for the selected time window. No app filter required." />
              <Button x:Name="btnAppScan"
                      Content="App-focused Scan"
                      Margin="0,0,6,0"
                      Padding="12,4"
                      ToolTip="App-focused Scan: filters events and crashes to the app name / exe in the 'App filter' box. Requires an App filter value." />
              <Button x:Name="btnFullScan"
                      Content="Full System Snapshot"
                      Padding="12,4"
                      ToolTip="Full System Snapshot: extended event set plus full OS / hardware snapshot. No special input required, but may take longer." />
            </StackPanel>
          </StackPanel>

          <!-- Export -->
          <StackPanel Orientation="Vertical" Grid.Column="4">
            <TextBlock Text="Export" FontWeight="Bold" Margin="0,0,0,2" />
            <Button x:Name="btnExport" Content="Export Report (ZIP)" Padding="12,4" />
          </StackPanel>

        </Grid>
      </ScrollViewer>
    </Border>

    <!-- MAIN TABS -->
    <TabControl Grid.Row="1" x:Name="tabMain"
                HorizontalAlignment="Stretch"
                VerticalAlignment="Stretch">
      <!-- Events -->
      <TabItem Header="Events">
        <Grid Background="White">
          <DataGrid x:Name="dgEvents"
                    Margin="6"
                    AutoGenerateColumns="True"
                    IsReadOnly="True"
                    CanUserSortColumns="True"
                    CanUserReorderColumns="True"
                    CanUserResizeColumns="True"
                    GridLinesVisibility="Horizontal"
                    HeadersVisibility="Column"
                    SelectionMode="Extended"
                    SelectionUnit="FullRow" />
        </Grid>
      </TabItem>

      <!-- Crashes -->
      <TabItem Header="Crashes">
        <Grid Background="White">
          <DataGrid x:Name="dgCrashes"
                    Margin="6"
                    AutoGenerateColumns="True"
                    IsReadOnly="True"
                    CanUserSortColumns="True"
                    CanUserReorderColumns="True"
                    CanUserResizeColumns="True"
                    GridLinesVisibility="Horizontal"
                    HeadersVisibility="Column"
                    SelectionMode="Extended"
                    SelectionUnit="FullRow" />
        </Grid>
      </TabItem>

      <!-- System Info -->
      <TabItem Header="System Info">
        <Grid Background="White">
          <ScrollViewer Margin="6">
            <TextBox x:Name="txtSystemInfo"
                     IsReadOnly="True"
                     TextWrapping="Wrap"
                     VerticalScrollBarVisibility="Auto"
                     HorizontalScrollBarVisibility="Auto"
                     Background="White"
                     BorderBrush="#cccccc"
                     AcceptsReturn="True" />
          </ScrollViewer>
        </Grid>
      </TabItem>

      <!-- Summary -->
      <TabItem Header="Summary">
        <Grid Background="White">
          <ScrollViewer Margin="6">
            <TextBox x:Name="txtSummary"
                     IsReadOnly="True"
                     TextWrapping="Wrap"
                     VerticalScrollBarVisibility="Auto"
                     HorizontalScrollBarVisibility="Auto"
                     Background="White"
                     BorderBrush="#cccccc"
                     AcceptsReturn="True" />
          </ScrollViewer>
        </Grid>
      </TabItem>

      <!-- Suggestions -->
      <TabItem Header="Suggestions">
        <Grid Background="White">
          <ScrollViewer Margin="6">
            <TextBox x:Name="txtSuggestions"
                     IsReadOnly="True"
                     TextWrapping="Wrap"
                     VerticalScrollBarVisibility="Auto"
                     HorizontalScrollBarVisibility="Auto"
                     Background="White"
                     BorderBrush="#cccccc"
                     AcceptsReturn="True" />
          </ScrollViewer>
        </Grid>
      </TabItem>
    </TabControl>

    <!-- STATUS BAR -->
    <Border Grid.Row="2" Margin="0,8,0,0" Padding="6" Background="#f0f0f0" CornerRadius="4">
      <DockPanel>
        <TextBlock x:Name="lblStatus" Text="Ready." DockPanel.Dock="Left" VerticalAlignment="Center" />
        <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock x:Name="lblScanInfo" Margin="0,0,6,0" />
          <ProgressBar x:Name="pbScan"
                       Width="180"
                       Height="16"
                       Minimum="0"
                       Maximum="100"
                       Value="0"
                       Visibility="Collapsed" />
        </StackPanel>
      </DockPanel>
    </Border>

    <!-- BUSY OVERLAY -->
    <Grid x:Name="overlayBusy"
          Grid.Row="0"
          Grid.RowSpan="3"
          Background="#80000000"
          Visibility="Collapsed">
      <Border Background="White"
              CornerRadius="8"
              Padding="20"
              HorizontalAlignment="Center"
              VerticalAlignment="Center">
        <StackPanel Orientation="Vertical" HorizontalAlignment="Center">
          <TextBlock x:Name="lblOverlayText"
                     Text="Working..."
                     FontSize="16"
                     FontWeight="Bold"
                     HorizontalAlignment="Center"
                     Margin="0,0,0,10" />
          <ProgressBar IsIndeterminate="True"
                       Width="220"
                       Height="16"
                       Margin="0,0,0,10" />
          <TextBlock Text="Collecting logs and system data. The window may be briefly unresponsive during heavy scans."
                     FontSize="11"
                     Foreground="Gray"
                     TextWrapping="Wrap"
                     TextAlignment="Center"
                     Width="260" />
        </StackPanel>
      </Border>
    </Grid>

  </Grid>
</Window>
"@

# Process picker window XAML
$procPickerXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select process"
        Width="800"
        Height="450"
        WindowStartupLocation="CenterOwner"
        Background="White"
        Foreground="Black"
        FontFamily="Segoe UI"
        ResizeMode="CanResizeWithGrip">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto" />
      <RowDefinition Height="*" />
      <RowDefinition Height="Auto" />
    </Grid.RowDefinitions>

    <!-- Search / scope -->
    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
      <TextBlock Text="Filter:" VerticalAlignment="Center" Margin="0,0,4,0" />
      <TextBox x:Name="txtProcSearch" Width="220" Margin="0,0,8,0" />
      <TextBlock Text="Scope:" VerticalAlignment="Center" Margin="0,0,4,0" />
      <ComboBox x:Name="cmbProcScope" Width="120" SelectedIndex="0" Margin="0,0,8,0">
        <ComboBoxItem Content="All processes"    Tag="All" />
        <ComboBoxItem Content="User processes"   Tag="User" />
        <ComboBoxItem Content="System processes" Tag="System" />
      </ComboBox>
      <Button x:Name="btnProcRefresh" Content="Refresh" Width="80" />
    </StackPanel>

    <!-- Process list -->
    <DataGrid x:Name="dgProcList"
              Grid.Row="1"
              Margin="0,0,0,8"
              AutoGenerateColumns="False"
              IsReadOnly="True"
              CanUserSortColumns="True"
              CanUserReorderColumns="True"
              CanUserResizeColumns="True"
              GridLinesVisibility="Horizontal"
              HeadersVisibility="Column"
              SelectionMode="Extended"
              SelectionUnit="FullRow">
      <DataGrid.Columns>
        <DataGridTextColumn Header="Name"        Binding="{Binding Name}"        Width="150" />
        <DataGridTextColumn Header="PID"         Binding="{Binding PID}"         Width="70" />
        <DataGridTextColumn Header="User"        Binding="{Binding User}"        Width="150" />
        <DataGridTextColumn Header="Company"     Binding="{Binding Company}"     Width="150" />
        <DataGridTextColumn Header="Description" Binding="{Binding Description}" Width="200" />
        <DataGridTextColumn Header="Path"        Binding="{Binding Path}"        Width="*" />
      </DataGrid.Columns>
    </DataGrid>

    <!-- Buttons -->
    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="btnProcUse"    Content="Use Selected" Margin="0,0,8,0" Width="110" />
      <Button x:Name="btnProcCancel" Content="Cancel" Width="80" />
    </StackPanel>
  </Grid>
</Window>
"@

# ---------------------------
# XAML loader
# ---------------------------
$xmlReaderSettings = New-Object System.Xml.XmlReaderSettings
$xmlReaderSettings.IgnoreComments   = $true
$xmlReaderSettings.IgnoreWhitespace = $true
$stringReader = New-Object System.IO.StringReader($xaml)
$xmlReader   = [System.Xml.XmlReader]::Create($stringReader, $xmlReaderSettings)

try {
    $window = [Windows.Markup.XamlReader]::Load([System.Xml.XmlReader]$xmlReader)
} catch {
    [System.Windows.MessageBox]::Show(
        "Failed to load UI: $($_.Exception.Message)",
        "Error",
        "OK",
        "Error"
    ) | Out-Null
    return
}

# ---------------------------
# Grab controls
# ---------------------------
$txtAppFilter   = $window.FindName("txtAppFilter")
$cmbLookback    = $window.FindName("cmbLookback")
$cmbProfile     = $window.FindName("cmbProfile")
$btnQuickScan   = $window.FindName("btnQuickScan")
$btnAppScan     = $window.FindName("btnAppScan")
$btnFullScan    = $window.FindName("btnFullScan")
$btnExport      = $window.FindName("btnExport")
$dgEvents       = $window.FindName("dgEvents")
$dgCrashes      = $window.FindName("dgCrashes")
$txtSystemInfo  = $window.FindName("txtSystemInfo")
$txtSummary     = $window.FindName("txtSummary")
$txtSuggestions = $window.FindName("txtSuggestions")
$lblStatus      = $window.FindName("lblStatus")
$lblScanInfo    = $window.FindName("lblScanInfo")
$pbScan         = $window.FindName("pbScan")
$tabMain        = $window.FindName("tabMain")
$btnPickApp     = $window.FindName("btnPickApp")
$overlayBusy    = $window.FindName("overlayBusy")
$lblOverlayText = $window.FindName("lblOverlayText")

# ---- Context menu for Events grid ----
$ctxEv = New-Object System.Windows.Controls.ContextMenu
$miCopyEv = New-Object System.Windows.Controls.MenuItem
$miCopyEv.Header = "Copy selected events to clipboard"
$miCopyEv.Add_Click({
    $sel = $dgEvents.SelectedItems
    if (-not $sel -or $sel.Count -eq 0) { return }

    $lines = @()
    foreach ($item in $sel) {
        $chunk = @(
            "TimeCreated : " + $item.TimeCreated,
            "Level       : " + $item.LevelDisplayName,
            "Id          : " + $item.Id,
            "Provider    : " + $item.ProviderName,
            "Message     : " + $item.Message
        ) -join "`r`n"
        $lines += $chunk
    }

    [System.Windows.Clipboard]::SetText(($lines -join "`r`n`r`n"))
})
[void]$ctxEv.Items.Add($miCopyEv)
$dgEvents.ContextMenu = $ctxEv

# ---- Context menu for Crashes grid ----
$ctxCr = New-Object System.Windows.Controls.ContextMenu
$miCopyCr = New-Object System.Windows.Controls.MenuItem
$miCopyCr.Header = "Copy selected crashes to clipboard"
$miCopyCr.Add_Click({
    $sel = $dgCrashes.SelectedItems
    if (-not $sel -or $sel.Count -eq 0) { return }

    $lines = @()
    foreach ($item in $sel) {
        $chunk = @(
            "Time        : " + $item.Time,
            "Product     : " + $item.ProductName,
            "Source      : " + $item.SourceName,
            "EventId     : " + $item.EventIdentifier,
            "Description : " + $item.Description
        ) -join "`r`n"
        $lines += $chunk
    }

    [System.Windows.Clipboard]::SetText(($lines -join "`r`n`r`n"))
})
[void]$ctxCr.Items.Add($miCopyCr)
$dgCrashes.ContextMenu = $ctxCr

# ---------------------------
# State
# ---------------------------
$script:LastEvents   = @()
$script:LastCrashes  = @()
$script:LastSysInfo  = ""
$script:LastSummary  = ""
$script:LastSuggest  = ""

# ---------------------------
# Helper: safe event message
# ---------------------------
function Get-SafeEventMessage {
    param($Event)
    try {
        return $Event.Message
    } catch {
        return "[Message formatting failed]"
    }
}

# ---------------------------
# Helper functions
# ---------------------------
function Get-LookbackStart {
    param([System.Windows.Controls.ComboBox]$Combo)
    $days = 1
    $selected = $Combo.SelectedItem
    if ($selected -and $selected.Tag) {
        [int]::TryParse($selected.Tag.ToString(), [ref]$days) | Out-Null
    }
    (Get-Date).AddDays(-$days)
}

function Get-AppFilterPattern {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return $null }
    $trim = $Raw.Trim() -replace '\.exe$',''
    [regex]::Escape($trim)
}

function Get-ProfileTag {
    param([System.Windows.Controls.ComboBox]$Combo)
    $selected = $Combo.SelectedItem
    if ($selected -and $selected.Tag) { return $selected.Tag.ToString() }
    "Power"
}

function Get-AppEvents {
    param(
        [datetime]$Since,
        [string]$AppPattern,
        [string]$ProfileTag,
        [string]$Mode
    )

    $maxEvents = switch ($ProfileTag) {
        "Basic" { 500 }
        "Power" { 2000 }
        "Admin" { 5000 }
        default { 2000 }
    }

    $interestingIds = 1000,1001,1002,1005,1008,1026,6008
    $logs = @("Application","System")

    $allEvents = New-Object System.Collections.Generic.List[object]

    $regex = $null
    if ($AppPattern) {
        $regex = [regex]$AppPattern
    }

    foreach ($log in $logs) {
        try {
            $filter = @{
                LogName   = $log
                StartTime = $Since
            }

            Get-WinEvent -FilterHashtable $filter -MaxEvents $maxEvents -ErrorAction SilentlyContinue |
            ForEach-Object {
                $ev = $_
                try {
                    # Decide if this event is interesting at all
                    $keep = $false
                    if ($Mode -eq "Quick") {
                        if ($ev.LevelDisplayName -in @("Critical","Error")) { $keep = $true }
                    } else {
                        if ($interestingIds -contains $ev.Id -or
                            $ev.LevelDisplayName -in @("Critical","Error","Warning")) {
                            $keep = $true
                        }
                    }
                    if (-not $keep) { return }

                    # App filter (optional)
                    $msg  = ""
                    $prov = $ev.ProviderName
                    $ids  = ($ev.Properties | ForEach-Object { $_.Value }) -join " "

                    if ($regex) {
                        try { $msg = $ev.Message } catch { $msg = "" }
                        if (-not ($regex.IsMatch($msg) -or
                                  $regex.IsMatch($prov) -or
                                  $regex.IsMatch($ids))) {
                            return
                        }
                    }

                    if (-not $msg) {
                        $msg = Get-SafeEventMessage -Event $ev
                    }

                    $allEvents.Add([pscustomobject]@{
                        TimeCreated      = $ev.TimeCreated
                        LogName          = $ev.LogName
                        LevelDisplayName = $ev.LevelDisplayName
                        Id               = $ev.Id
                        ProviderName     = $ev.ProviderName
                        Task             = $ev.TaskDisplayName
                        Machine          = $ev.MachineName
                        Message          = $msg
                    }) | Out-Null
                } catch [System.FormatException] {
                    # One bad event message – ignore it completely
                } catch {
                    # Any other per-event issue – also ignore, keep going
                }
            }
        } catch {
            # Ignore per-log failures and move to next log
        }
    }

    $allEvents | Sort-Object TimeCreated -Descending
}

function Get-ProcessInfoList {
    # Returns a list of processes with Name, PID, User, Kind (User/System), Company, Description, Path
    $result = New-Object System.Collections.Generic.List[object]

    $systemAccounts = @(
        "NT AUTHORITY\SYSTEM",
        "NT AUTHORITY\LOCAL SERVICE",
        "NT AUTHORITY\NETWORK SERVICE",
        "LOCAL SERVICE",
        "NETWORK SERVICE",
        "SYSTEM"
    )

    try {
        $procs = Get-CimInstance -ClassName Win32_Process -ErrorAction Stop
    } catch {
        return @()
    }

    foreach ($p in $procs) {
        $owner = ""
        try {
            $o = Invoke-CimMethod -InputObject $p -MethodName GetOwner -ErrorAction Stop
            if ($o.ReturnValue -eq 0) {
                $owner = $o.Domain + "\" + $o.User
            }
        } catch { }

        $kind = "User"
        if ([string]::IsNullOrWhiteSpace($owner)) {
            $kind = "System"
        } else {
            $uUpper = $owner.ToUpperInvariant()
            if ($systemAccounts -contains $uUpper) {
                $kind = "System"
            }
        }

        $path    = $p.ExecutablePath
        $company = ""
        $desc    = ""

        if ($path -and (Test-Path $path -ErrorAction SilentlyContinue)) {
            try {
                $fv = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($path)
                $company = $fv.CompanyName
                $desc    = $fv.FileDescription
            } catch { }
        }

        $obj = [pscustomobject]@{
            Name        = $p.Name
            PID         = $p.ProcessId
            User        = $owner
            Kind        = $kind     # "User" or "System"
            Company     = $company
            Description = $desc
            Path        = $path
        }

        $null = $result.Add($obj)
    }

    return $result
}

function Show-ProcessPicker {
    param(
        [string]$InitialText
    )

    # Load picker window from XAML
    $sr = New-Object System.IO.StringReader($procPickerXaml)
    $xr = [System.Xml.XmlReader]::Create($sr)
    $picker = [Windows.Markup.XamlReader]::Load($xr)

    # Grab controls
    $txtSearch  = $picker.FindName("txtProcSearch")
    $cmbScope   = $picker.FindName("cmbProcScope")
    $btnRefresh = $picker.FindName("btnProcRefresh")
    $btnUse     = $picker.FindName("btnProcUse")
    $btnCancel  = $picker.FindName("btnProcCancel")
    $dgProc     = $picker.FindName("dgProcList")

    # Backing data (shared inside this function via script-scoped variable)
    $script:ProcPicker_AllProcs = Get-ProcessInfoList

    # Helper: update filtered view based on scope and search text
    $updateView = {
        param($sender, $args)

        $scopeItem = $cmbScope.SelectedItem
        $scope = "All"
        if ($scopeItem -and $scopeItem.Tag) {
            $scope = $scopeItem.Tag.ToString()
        }

        $textFilter = $txtSearch.Text
        if ($textFilter) {
            $textFilter = $textFilter.Trim()
        } else {
            $textFilter = ""
        }

        $filtered = $script:ProcPicker_AllProcs

        if ($scope -eq "User") {
            $filtered = $filtered | Where-Object { $_.Kind -eq "User" }
        } elseif ($scope -eq "System") {
            $filtered = $filtered | Where-Object { $_.Kind -eq "System" }
        }

        if ($textFilter -ne "") {
            $pattern = [regex]::Escape($textFilter)
            $filtered = $filtered | Where-Object {
                $_.Name        -match $pattern -or
                $_.Description -match $pattern -or
                $_.Path        -match $pattern
            }
        }

        $dgProc.ItemsSource = $filtered | Sort-Object Name, PID
    }

    # Seed search with current App filter (pre-filters to associated processes)
    if (-not [string]::IsNullOrWhiteSpace($InitialText)) {
        $txtSearch.Text = $InitialText.Trim()
    }

    # Default scope = All
    $cmbScope.SelectedIndex = 0

    # Wire events
    $txtSearch.Add_TextChanged($updateView)
    $cmbScope.Add_SelectionChanged($updateView)

    $btnRefresh.Add_Click({
        $script:ProcPicker_AllProcs = Get-ProcessInfoList
        & $updateView $null $null
    })

    $onUse = {
        if ($dgProc.SelectedItem -ne $null) {
            $name = $dgProc.SelectedItem.Name
            if ($name) {
                # Trim .exe for cleaner filter
                $lower = $name.ToLowerInvariant()
                if ($lower.EndsWith(".exe")) {
                    $name = $name.Substring(0, $name.Length - 4)
                }
                $picker.Tag = $name
                $picker.DialogResult = $true
                $picker.Close()
            }
        }
    }

    $btnUse.Add_Click($onUse)
    $btnCancel.Add_Click({
        $picker.DialogResult = $false
        $picker.Close()
    })

    $dgProc.Add_MouseDoubleClick($onUse)

    # Initial view
    & $updateView $null $null

    # Show as modal dialog
    $null = $picker.ShowDialog()

    if ($picker.DialogResult -eq $true -and $picker.Tag) {
        return $picker.Tag
    }

    return $null
}

function Get-ReliabilityCrashes {
    param(
        [datetime]$Since,
        [string]$AppPattern
    )

    try {
        $records = Get-CimInstance -ClassName Win32_ReliabilityRecords -ErrorAction Stop
    } catch { return @() }

    $filtered = $records | Where-Object {
        ($_.TimeGenerated -as [datetime]) -ge $Since -and
        $_.SourceName -in @(
            "Application Failure",
            "Application Crashes",
            "Application Hang",
            "Windows Error Reporting",
            "Application Error"
        )
    }

    if ($AppPattern) {
        $regex = [regex]$AppPattern
        $filtered = $filtered | Where-Object {
            $app   = $_.ProductName
            $desc  = $_.Description
            $src   = $_.SourceName
            $regex.IsMatch($app) -or $regex.IsMatch($desc) -or $regex.IsMatch($src)
        }
    }

    $result =
        $filtered |
        Select-Object `
            @{Name="Time";Expression={ $_.TimeGenerated -as [datetime]}},
            ProductName,
            SourceName,
            EventIdentifier,
            InsertionStrings,
            Description,
            @{Name="RawRecordId";Expression={$_.RecordId}} |
        Sort-Object Time -Descending

    # Always return an array, even if there is only one crash
    return @($result)
}


function Get-SystemInfoText {
    try {
        $os  = Get-CimInstance -ClassName Win32_OperatingSystem
        $cs  = Get-CimInstance -ClassName Win32_ComputerSystem
        $cpu = Get-CimInstance -ClassName Win32_Processor

        $memTotalGB = [Math]::Round($cs.TotalPhysicalMemory / 1GB, 2)
        $uptime     = (Get-Date) - ($os.LastBootUpTime)

        $osLine  = ($os.Caption.Trim() + " " + $os.Version + " (Build " + $os.BuildNumber + ")")
        $upLine  = ($uptime.Days.ToString() + "d " +
                    $uptime.Hours.ToString() + "h " +
                    $uptime.Minutes.ToString() + "m")

        $sb = New-Object System.Text.StringBuilder

        [void]$sb.AppendLine("=== SYSTEM INFO ===")
        [void]$sb.AppendLine("Computer Name : " + $env:COMPUTERNAME)
        [void]$sb.AppendLine("User          : " + [System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
        [void]$sb.AppendLine("OS            : " + $osLine)
        [void]$sb.AppendLine("Uptime        : " + $upLine)
        [void]$sb.AppendLine("CPU           : " + $cpu.Name.Trim() + " x " + $cpu.NumberOfLogicalProcessors)
        [void]$sb.AppendLine("RAM           : " + $memTotalGB + " GB")

        $disks = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" | Sort-Object DeviceID
        [void]$sb.AppendLine()
        [void]$sb.AppendLine("=== DISKS ===")
        foreach ($d in $disks) {
            $sizeGB = if ($d.Size)      { [Math]::Round($d.Size/1GB,2) }      else { 0 }
            $freeGB = if ($d.FreeSpace) { [Math]::Round($d.FreeSpace/1GB,2) } else { 0 }
            $usedPct = if ($d.Size) {
                [Math]::Round((($d.Size - $d.FreeSpace)*100.0)/$d.Size,1)
            } else {
                0
            }

            $line = $d.DeviceID + "  " +
                    $sizeGB + " GB total  " +
                    $freeGB + " GB free  (" +
                    $usedPct + "% used)"
            [void]$sb.AppendLine($line)
        }

        $nics = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled=TRUE"
        [void]$sb.AppendLine()
        [void]$sb.AppendLine("=== NETWORK ===")
        foreach ($nic in $nics) {
            $ip = if ($nic.IPAddress) { $nic.IPAddress -join ", " } else { "" }
            $line = $nic.Description.Trim() + " [" + $ip + "]"
            [void]$sb.AppendLine($line)
        }

        return $sb.ToString()
    }
    catch {
        return "System info collection failed internally: " + $_.Exception.Message
    }
}

function Build-ScanSummary {
    param(
        $Events,
        $Crashes,
        [string]$AppFilter,
        [datetime]$Since
    )

    # Force to arrays so code below is predictable
    $eventsArr  = @($Events)
    $crashArr   = @($Crashes)
    $totalEv    = $eventsArr.Count
    $totalCrash = $crashArr.Count

    $critical   = $eventsArr | Where-Object LevelDisplayName -eq "Critical"
    $errors     = $eventsArr | Where-Object LevelDisplayName -eq "Error"
    $warnings   = $eventsArr | Where-Object LevelDisplayName -eq "Warning"

    $appLabel = if ([string]::IsNullOrWhiteSpace($AppFilter)) {
        "<no app filter>"
    } else {
        $AppFilter.Trim()
    }

    $sb = New-Object System.Text.StringBuilder

    [void]$sb.AppendLine("=== SCAN SUMMARY ===")
    [void]$sb.AppendLine("Machine      : " + $env:COMPUTERNAME)
    [void]$sb.AppendLine("User         : " + [System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
    [void]$sb.AppendLine("Time window  : " + $Since + " -> " + (Get-Date))
    [void]$sb.AppendLine("App filter   : " + $appLabel)
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("Events total : " + $totalEv +
                         " (Critical: " + $critical.Count +
                         ", Error: "   + $errors.Count   +
                         ", Warning: " + $warnings.Count + ")")
    [void]$sb.AppendLine("Crashes      : " + $totalCrash + " (Reliability Monitor)")
    [void]$sb.AppendLine()

    if ($totalCrash -gt 0) {
        $recent = $crashArr | Select-Object -First 5
        [void]$sb.AppendLine("Recent crashes:")
        foreach ($c in $recent) {
            $line = " - " +
                    $c.Time.ToString("u") + "  " +
                    $c.ProductName + "  [" +
                    $c.SourceName + "]  ID=" +
                    $c.EventIdentifier
            [void]$sb.AppendLine($line)
        }
        [void]$sb.AppendLine()
    }

    if ($totalEv -gt 0) {
        $topByProvider = $eventsArr |
            Group-Object ProviderName |
            Sort-Object Count -Descending |
            Select-Object -First 5

        [void]$sb.AppendLine("Top error sources:")
        foreach ($g in $topByProvider) {
            $line = " - " + $g.Name + " : " + $g.Count + " events"
            [void]$sb.AppendLine($line)
        }
    } else {
        [void]$sb.AppendLine("No relevant events found in the selected window.")
    }

    return $sb.ToString()
}


function Build-Suggestions {
    param(
        $Events,
        $Crashes,
        [string]$AppFilter,
        [string]$ProfileTag
    )

    $eventsArr = @($Events)
    $crashArr  = @($Crashes)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("=== GENERAL SUGGESTIONS ===")

    if ($eventsArr.Count -eq 0 -and $crashArr.Count -eq 0) {
        [void]$sb.AppendLine("No crash-related events found in the selected time window.")
        [void]$sb.AppendLine("- If the issue just happened, increase the lookback window.")
        [void]$sb.AppendLine("- Reproduce the issue again and run a new scan.")
        return $sb.ToString()
    }

    $appLabel = if ([string]::IsNullOrWhiteSpace($AppFilter)) { "the affected app" } else { $AppFilter.Trim() }

    $hasAppError = $eventsArr | Where-Object { $_.Id -eq 1000 -and $_.ProviderName -eq "Application Error" }
    $hasHang     = $eventsArr | Where-Object { $_.Id -eq 1002 -and $_.ProviderName -eq "Application Hang" }
    $hasWER      = $eventsArr | Where-Object {
        $_.ProviderName -eq "Windows Error Reporting" -or
        ($_.LogName -eq "Application" -and $_.Id -eq 1001)
    }

    if ($hasAppError) {
        [void]$sb.AppendLine("- The app is raising crash events (ID 1000, Application Error).")
        [void]$sb.AppendLine("  - Common causes: bad plugins, corrupted binaries, faulty updates, or AV interference.")
        [void]$sb.AppendLine("  - Try: repairing or reinstalling " + $appLabel + ", temporarily disabling third-party AV, checking for recent updates.")
        [void]$sb.AppendLine()
    }

    if ($hasHang) {
        [void]$sb.AppendLine("- Application Hang events detected (ID 1002).")
        [void]$sb.AppendLine("  - The process is not responding - usually resource or I/O contention.")
        [void]$sb.AppendLine("  - Try: closing background apps, checking disk and memory usage, and verifying network latency if it is a networked app.")
        [void]$sb.AppendLine()
    }

    if ($hasWER) {
        [void]$sb.AppendLine("- Windows Error Reporting entries exist for recent crashes.")
        [void]$sb.AppendLine("  - Advanced users can inspect WER crash dumps in:")
        [void]$sb.AppendLine("    %ProgramData%\Microsoft\Windows\WER\ReportArchive")
        [void]$sb.AppendLine()
    }

    $systemDrive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='C:'"
    if ($systemDrive -and $systemDrive.FreeSpace) {
        $freeGB = [Math]::Round($systemDrive.FreeSpace/1GB, 2)
        $usedPct = [Math]::Round((($systemDrive.Size-$systemDrive.FreeSpace)*100.0)/$systemDrive.Size, 1)
        if ($freeGB -lt 5 -or $usedPct -gt 90) {
            [void]$sb.AppendLine("- System drive has low free space.")
            [void]$sb.AppendLine("  - Free at least 10-15 GB on C: to avoid instability and crashes.")
            [void]$sb.AppendLine()
        }
    }

    if ($ProfileTag -in @("Power","Admin")) {
        [void]$sb.AppendLine("=== POWER USER / SYSADMIN NOTES ===")
        [void]$sb.AppendLine("- Use Get-WinEvent with the event IDs reported in the Events tab for deeper analysis.")
        [void]$sb.AppendLine("- For .NET apps, watch for System.* exceptions in the event messages.")
        [void]$sb.AppendLine("- Check Windows Defender / AV logs for blocked DLLs or quarantined files.")
        [void]$sb.AppendLine("- Use Reliability Monitor (perfmon /rel) for a timeline view.")
    }

    if ($ProfileTag -eq "Admin") {
        [void]$sb.AppendLine()
        [void]$sb.AppendLine("=== SYSADMIN EXTRAS ===")
        [void]$sb.AppendLine("- Consider centralizing these logs (Event Forwarding / SIEM) for correlation.")
        [void]$sb.AppendLine("- If crashes correlate with specific driver updates, check Device Manager and roll back where needed.")
        [void]$sb.AppendLine("- For recurring WER bucket IDs, search your internal KB or vendor support with the bucket value.")
    }

    $sb.ToString()
}


function Set-UiBusy {
    param([string]$StatusText = "Scanning...")

    # Update status bar
    $lblStatus.Text = $StatusText
    $pbScan.Visibility = "Visible"
    $pbScan.IsIndeterminate = $true
    $pbScan.Value = 0

    # Update overlay text + show overlay
    if ($lblOverlayText) { $lblOverlayText.Text = $StatusText }
    if ($overlayBusy)    { $overlayBusy.Visibility = "Visible" }

    # Change mouse cursor
    $window.Cursor = [System.Windows.Input.Cursors]::Wait

    # *** Force UI to repaint before we start heavy work ***
    try {
        # Let WPF process layout/render
        $window.Dispatcher.Invoke(
            [Action]{},
            [System.Windows.Threading.DispatcherPriority]::Background
        )

        # Also pump WinForms messages for good measure
        [System.Windows.Forms.Application]::DoEvents()
    } catch {
        # If this fails for any reason, just ignore – overlay still won't break anything
    }
}


function Set-UiReady {
    param([string]$StatusText = "Ready.")
    $lblStatus.Text = $StatusText
    $pbScan.Visibility = "Collapsed"
    $pbScan.IsIndeterminate = $false
    $pbScan.Value = 0
    if ($overlayBusy) { $overlayBusy.Visibility = "Collapsed" }
    $window.Cursor = [System.Windows.Input.Cursors]::Arrow
}

function Run-Scan {
    param([string]$Mode)

    Set-UiBusy "Running $Mode scan..."
    $lblScanInfo.Text = ""

    $since      = Get-LookbackStart -Combo $cmbLookback
    $appPattern = Get-AppFilterPattern -Raw $txtAppFilter.Text
    $profileTag = Get-ProfileTag     -Combo $cmbProfile

    # Requirement: App-focused Scan needs an app filter
    if ($Mode -eq "App" -and -not $appPattern) {
        [System.Windows.MessageBox]::Show(
            "App-focused Scan requires an app name or exe in the 'App filter' box." +
            "`r`nExample: outlook, outlook.exe, chrome, Plex Media Server.",
            "App-focused Scan",
            "OK",
            "Information"
        ) | Out-Null
        Set-UiReady
        return
    }

    $events      = @()
    $crashes     = @()
    $sysInfoText = ""
    $summaryText = ""
    $suggestText = ""

    # Events
    try {
        $lblScanInfo.Text = "Collecting events..."
        $events = Get-AppEvents -Since $since -AppPattern $appPattern -ProfileTag $profileTag -Mode $Mode
    } catch [System.FormatException] {
        # Swallow weird message-template errors
        $events = @()
    } catch {
        $events = @()
        $lblStatus.Text = "Event collection failed: $($_.Exception.Message)"
    }

    # Reliability crashes
    try {
        $lblScanInfo.Text = "Collecting reliability data..."
        $crashes = Get-ReliabilityCrashes -Since $since -AppPattern $appPattern
    } catch {
        $crashes = @()
        $lblStatus.Text = "Reliability data failed: $($_.Exception.Message)"
    }

    # System info
    try {
        $lblScanInfo.Text = "Collecting system info..."
        $sysInfoText = Get-SystemInfoText
    } catch {
        $sysInfoText = "Failed to collect system info: $($_.Exception.Message)"
    }

    # Summary
    try {
        $lblScanInfo.Text = "Building summary..."
        $summaryText = Build-ScanSummary -Events $events -Crashes $crashes -AppFilter $txtAppFilter.Text -Since $since
    } catch {
        $summaryText = "Failed to build summary: $($_.Exception.Message)"
    }

    # Suggestions
    try {
        $lblScanInfo.Text = "Building suggestions..."
        $suggestText = Build-Suggestions -Events $events -Crashes $crashes -AppFilter $txtAppFilter.Text -ProfileTag $profileTag
    } catch {
        $suggestText = "Failed to build suggestions: $($_.Exception.Message)"
    }

    # Update UI – this should never throw, but wrap just in case
    try {
        $dgEvents.ItemsSource   = $events
        $dgCrashes.ItemsSource  = @($crashes)
        $txtSystemInfo.Text     = $sysInfoText
        $txtSummary.Text        = $summaryText
        $txtSuggestions.Text    = $suggestText

        $script:LastEvents   = $events
        $script:LastCrashes  = $crashes
        $script:LastSysInfo  = $sysInfoText
        $script:LastSummary  = $summaryText
        $script:LastSuggest  = $suggestText

        $lblStatus.Text = "Scan complete. Events: $($events.Count), Crashes: $($crashes.Count)"
        $lblScanInfo.Text = "Window: $since | Profile: $profileTag"
    } catch {
        $msg = "Scan failed while updating UI: $($_.Exception.Message)"
        $lblStatus.Text = $msg
        [System.Windows.MessageBox]::Show($msg, "Scan error", "OK", "Error") | Out-Null
    } finally {
        Set-UiReady
    }
}

$btnPickApp.Add_Click({
    $current = $txtAppFilter.Text
    $chosen  = Show-ProcessPicker -InitialText $current
    if ($chosen) {
        $txtAppFilter.Text = $chosen
    }
})

function Export-Report {
    if (-not $script:LastSummary) {
        [System.Windows.MessageBox]::Show("Run a scan before exporting.", "Export", "OK", "Information") | Out-Null
        return
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "ZIP files (*.zip)|*.zip"
    $dialog.Title  = "Save diagnostic bundle"
    $dialog.FileName = "AppHealthReport_{0:yyyyMMdd_HHmmss}.zip" -f (Get-Date)

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $targetZip = $dialog.FileName
    $tmpRoot   = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("AppHealth_{0}" -f ([Guid]::NewGuid().ToString("N")))
    New-Item -Path $tmpRoot -ItemType Directory -Force | Out-Null

    try {
        $eventsPath = Join-Path $tmpRoot "Events.csv"
        $script:LastEvents  | Export-Csv -Path $eventsPath  -NoTypeInformation -Encoding UTF8
        $crashPath  = Join-Path $tmpRoot "Crashes.csv"
        $script:LastCrashes | Export-Csv -Path $crashPath   -NoTypeInformation -Encoding UTF8
        $sysInfoPath = Join-Path $tmpRoot "SystemInfo.txt"
        $script:LastSysInfo | Out-File -FilePath $sysInfoPath -Encoding UTF8
        $summaryPath = Join-Path $tmpRoot "Summary.txt"
        $script:LastSummary | Out-File -FilePath $summaryPath -Encoding UTF8
        $suggestPath = Join-Path $tmpRoot "Suggestions.txt"
        $script:LastSuggest | Out-File -FilePath $suggestPath -Encoding UTF8

        if (Test-Path $targetZip) { Remove-Item $targetZip -Force }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($tmpRoot, $targetZip)

        [System.Windows.MessageBox]::Show("Exported to:`n$targetZip","Export complete","OK","Information") | Out-Null
    } catch {
        [System.Windows.MessageBox]::Show("Export failed: $($_.Exception.Message)","Export error","OK","Error") | Out-Null
    } finally {
        if (Test-Path $tmpRoot) { Remove-Item $tmpRoot -Recurse -Force }
    }
}

# ---------------------------
# Wire up buttons
# ---------------------------
$btnQuickScan.Add_Click({ Run-Scan -Mode "Quick" })
$btnAppScan.Add_Click({   Run-Scan -Mode "App"   })
$btnFullScan.Add_Click({  Run-Scan -Mode "Full"  })
$btnExport.Add_Click({    Export-Report          })

# ---------------------------
# Initial UI text
# ---------------------------
$lblStatus.Text      = "Ready. Choose scan type to begin."
$lblScanInfo.Text    = ""
$txtSystemInfo.Text  = ""
$txtSummary.Text     = "Run a scan to see a summary of errors, crashes, and affected apps."
$txtSuggestions.Text = "Run a scan to get suggestions based on detected events and crashes."

# Show window
$window.Topmost = $false
$null = $window.ShowDialog()
