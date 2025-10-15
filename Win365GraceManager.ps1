#requires -Version 5.1
#requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.DeviceManagement.Actions

<#
.SYNOPSIS
    Windows 365 Business Cloud PC Grace Period Manager
.DESCRIPTION
    GUI application to manage Windows 365 Business Cloud PCs with focus on ending grace periods.
    Displays Cloud PCs with visual indicators for grace status and provides context menu for deprovisioning.
.NOTES
    Author: Windows 365 Admin Tool
    Date: October 15, 2025
    Requires: Microsoft Graph PowerShell SDK
#>

# Import required assemblies for WPF
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Global variables
$Global:CloudPCs = @()

# Function to connect to Microsoft Graph
function Connect-MgGraphIfNeeded {
    try {
        $context = Get-MgContext
        if (-not $context) {
            Write-Host "Connecting to Microsoft Graph..."
            Connect-MgGraph -Scopes "CloudPC.ReadWrite.All", "CloudPC.Read.All" -NoWelcome
            Write-Host "Connected successfully!"
        }
        return $true
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to connect to Microsoft Graph: $($_.Exception.Message)",
            "Connection Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return $false
    }
}

# Function to get Cloud PCs
function Get-CloudPCs {
    try {
        Write-Host "Retrieving Cloud PCs..."
        $cloudPCs = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs" -OutputType PSObject
        
        # Use ArrayList for better performance when adding items
        $pcList = [System.Collections.ArrayList]::new()
        foreach ($pc in $cloudPCs.value) {
            # Format the grace period end date for better readability
            $formattedDate = if ($pc.gracePeriodEndDateTime) {
                try {
                    ([DateTime]$pc.gracePeriodEndDateTime).ToString("dd/MM/yyyy HH:mm")
                }
                catch {
                    $pc.gracePeriodEndDateTime
                }
            }
            else {
                ""
            }
            
            [void]$pcList.Add([PSCustomObject]@{
                Id = $pc.id
                ManagedDeviceName = $pc.managedDeviceName
                UserPrincipalName = $pc.userPrincipalName
                Status = $pc.status
                ServicePlanName = $pc.servicePlanName
                GracePeriodEndDateTime = $formattedDate
                IsInGracePeriod = ($pc.status -eq "inGracePeriod")
            })
        }
        
        Write-Host "Retrieved $($pcList.Count) Cloud PCs"
        return $pcList
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to retrieve Cloud PCs: $($_.Exception.Message)",
            "Retrieval Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return @()
    }
}

# Function to deprovision Cloud PC (end grace period)
function Invoke-DeprovisionCloudPC {
    param(
        [string]$CloudPCId,
        [string]$ManagedDeviceName
    )
    
    try {
        $result = [System.Windows.MessageBox]::Show(
            "Are you sure you want to deprovision '$ManagedDeviceName'?`n`nThis action will end the grace period and remove the Cloud PC permanently. This cannot be undone.",
            "Confirm Deprovision",
            [System.Windows.MessageBoxButton]::OKCancel,
            [System.Windows.MessageBoxImage]::Warning
        )
        
        if ($result -eq [System.Windows.MessageBoxResult]::OK) {
            Write-Host "Deprovisioning Cloud PC: $ManagedDeviceName ($CloudPCId)"
            
            # Call the deprovision API
            $uri = "https://graph.microsoft.com/beta/deviceManagement/virtualEndpoint/cloudPCs/$CloudPCId/endGracePeriod"
            Invoke-MgGraphRequest -Method POST -Uri $uri
            
            [System.Windows.MessageBox]::Show(
                "Cloud PC '$ManagedDeviceName' has been successfully deprovisioned.",
                "Success",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            
            return $true
        }
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to deprovision Cloud PC: $($_.Exception.Message)",
            "Deprovision Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return $false
    }
    
    return $false
}

# Function to export Cloud PCs to CSV
function Export-CloudPCsToCSV {
    try {
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $saveDialog.FileName = "CloudPCs_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $saveDialog.Title = "Export Cloud PCs to CSV"
        
        if ($saveDialog.ShowDialog()) {
            $Global:CloudPCs | Select-Object ManagedDeviceName, UserPrincipalName, ServicePlanName, Status, GracePeriodEndDateTime, IsInGracePeriod | 
                Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
            
            [System.Windows.MessageBox]::Show(
                "Cloud PCs exported successfully to:`n$($saveDialog.FileName)",
                "Export Successful",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            return $true
        }
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to export Cloud PCs: $($_.Exception.Message)",
            "Export Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return $false
    }
}

# XAML definition for the GUI
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Windows 365 Grace Period Manager" Height="600" Width="950"
        WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style x:Key="GracePeriodStyle" TargetType="ListViewItem">
            <Style.Triggers>
                <DataTrigger Binding="{Binding IsInGracePeriod}" Value="True">
                    <Setter Property="Foreground" Value="Blue"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                </DataTrigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <StackPanel Grid.Row="0" Orientation="Vertical" Margin="0,0,0,10">
            <DockPanel>
                <StackPanel DockPanel.Dock="Left">
                    <TextBlock Text="Windows 365 Business Cloud PC Manager" 
                               FontSize="20" FontWeight="Bold" Margin="0,0,0,5"/>
                    <TextBlock Text="Cloud PCs in grace period are highlighted in blue. Right-click to deprovision." 
                               FontSize="12" Foreground="Gray"/>
                </StackPanel>
                <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Top">
                    <TextBlock Text="Search: " FontSize="12" VerticalAlignment="Center" Margin="0,0,5,0"/>
                    <TextBox Name="FilterTextBox" Width="200" Height="25" Margin="0,0,0,0" 
                             VerticalContentAlignment="Center" Padding="5,0"
                             ToolTip="Filter by name, user, or status"/>
                </StackPanel>
            </DockPanel>
        </StackPanel>
        
        <!-- DataGrid for Cloud PCs -->
        <ListView Grid.Row="1" Name="CloudPCListView" 
                  ItemContainerStyle="{StaticResource GracePeriodStyle}"
                  SelectionMode="Single">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Cloud PC Name" Width="200" DisplayMemberBinding="{Binding ManagedDeviceName}"/>
                    <GridViewColumn Header="Assigned User" Width="220" DisplayMemberBinding="{Binding UserPrincipalName}"/>
                    <GridViewColumn Header="Service Plan" Width="150" DisplayMemberBinding="{Binding ServicePlanName}"/>
                    <GridViewColumn Header="Status" Width="130" DisplayMemberBinding="{Binding Status}"/>
                    <GridViewColumn Header="Grace End Date" Width="150" DisplayMemberBinding="{Binding GracePeriodEndDateTime}"/>
                </GridView>
            </ListView.View>
            <ListView.ContextMenu>
                <ContextMenu Name="CloudPCContextMenu">
                    <MenuItem Name="DeprovisionMenuItem" Header="Deprovision Now (End Grace Period)" InputGestureText="Delete"/>
                    <Separator/>
                    <MenuItem Name="CopyNameMenuItem" Header="Copy Device Name" InputGestureText="Ctrl+C"/>
                </ContextMenu>
            </ListView.ContextMenu>
        </ListView>
        
        <!-- Status Bar -->
        <StatusBar Grid.Row="2" Margin="0,10,0,0">
            <StatusBarItem>
                <TextBlock Name="StatusText" Text="Ready"/>
            </StatusBarItem>
            <Separator/>
            <StatusBarItem>
                <TextBlock Name="CountText" Text="Cloud PCs: 0"/>
            </StatusBarItem>
            <StatusBarItem>
                <TextBlock Name="FilterStatusText" Text="" Foreground="DarkBlue"/>
            </StatusBarItem>
            <StatusBarItem HorizontalAlignment="Right">
                <StackPanel Orientation="Horizontal">
                    <Button Name="ExportButton" Content="Export to CSV" Width="100" Padding="5,2" Margin="0,0,5,0"/>
                    <Button Name="RefreshButton" Content="Refresh (F5)" Width="90" Padding="5,2"/>
                </StackPanel>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
"@

# Function to refresh the Cloud PC list
function Update-CloudPCList {
    param(
        $Window,
        [string]$FilterText = ""
    )
    
    $statusText = $Window.FindName("StatusText")
    $countText = $Window.FindName("CountText")
    $filterStatusText = $Window.FindName("FilterStatusText")
    $listView = $Window.FindName("CloudPCListView")
    
    $statusText.Text = "Loading Cloud PCs..."
    $listView.Items.Clear()
    
    # Only fetch from API if CloudPCs is empty or FilterText is empty (full refresh)
    if ($Global:CloudPCs.Count -eq 0 -or $FilterText -eq "") {
        $Global:CloudPCs = Get-CloudPCs
    }
    
    # Apply filter if provided
    $displayPCs = $Global:CloudPCs
    if ($FilterText) {
        $displayPCs = $Global:CloudPCs | Where-Object {
            $_.ManagedDeviceName -like "*$FilterText*" -or
            $_.UserPrincipalName -like "*$FilterText*" -or
            $_.Status -like "*$FilterText*" -or
            $_.ServicePlanName -like "*$FilterText*"
        }
        $filterStatusText.Text = "Filtered: $($displayPCs.Count) of $($Global:CloudPCs.Count)"
    }
    else {
        $filterStatusText.Text = ""
    }
    
    foreach ($pc in $displayPCs) {
        $listView.Items.Add($pc) | Out-Null
    }
    
    $graceCount = @($Global:CloudPCs | Where-Object { $_.IsInGracePeriod }).Count
    $countText.Text = "Cloud PCs: $($Global:CloudPCs.Count) (In Grace: $graceCount)"
    $statusText.Text = "Ready"
}

# Main function to show the GUI
function Show-Win365GraceManager {
    # Connect to Microsoft Graph first
    if (-not (Connect-MgGraphIfNeeded)) {
        return
    }
    
    # Load XAML
    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    # Get controls
    $listView = $window.FindName("CloudPCListView")
    $contextMenu = $window.FindName("CloudPCContextMenu")
    $deprovisionMenuItem = $window.FindName("DeprovisionMenuItem")
    $copyNameMenuItem = $window.FindName("CopyNameMenuItem")
    $refreshButton = $window.FindName("RefreshButton")
    $exportButton = $window.FindName("ExportButton")
    $filterTextBox = $window.FindName("FilterTextBox")
    $statusText = $window.FindName("StatusText")
    
    # Refresh button click event
    $refreshButton.Add_Click({
        $filterTextBox.Text = ""
        Update-CloudPCList -Window $window
    })
    
    # Export button click event
    $exportButton.Add_Click({
        Export-CloudPCsToCSV
    })
    
    # Filter textbox text changed event
    $filterTextBox.Add_TextChanged({
        Update-CloudPCList -Window $window -FilterText $filterTextBox.Text
    })
    
    # Keyboard shortcuts
    $window.Add_KeyDown({
        param($sender, $e)
        
        # F5 - Refresh
        if ($e.Key -eq [System.Windows.Input.Key]::F5) {
            $filterTextBox.Text = ""
            Update-CloudPCList -Window $window
            $e.Handled = $true
        }
        
        # Delete - Deprovision (if item selected and in grace)
        if ($e.Key -eq [System.Windows.Input.Key]::Delete) {
            $selectedItem = $listView.SelectedItem
            if ($selectedItem -and $selectedItem.IsInGracePeriod) {
                $result = Invoke-DeprovisionCloudPC -CloudPCId $selectedItem.Id -ManagedDeviceName $selectedItem.ManagedDeviceName
                if ($result) {
                    Update-CloudPCList -Window $window
                }
            }
            $e.Handled = $true
        }
        
        # Ctrl+C - Copy device name
        if ($e.Key -eq [System.Windows.Input.Key]::C -and 
            ($e.KeyboardDevice.Modifiers -band [System.Windows.Input.ModifierKeys]::Control)) {
            $selectedItem = $listView.SelectedItem
            if ($selectedItem) {
                [System.Windows.Clipboard]::SetText($selectedItem.ManagedDeviceName)
            }
            $e.Handled = $true
        }
    })
    
    # Context menu opening event - enable/disable based on grace status
    $contextMenu.Add_Opened({
        $selectedItem = $listView.SelectedItem
        if ($selectedItem -and $selectedItem.IsInGracePeriod) {
            $deprovisionMenuItem.IsEnabled = $true
        }
        else {
            $deprovisionMenuItem.IsEnabled = $false
        }
    })
    
    # Deprovision menu item click event
    $deprovisionMenuItem.Add_Click({
        $selectedItem = $listView.SelectedItem
        if ($selectedItem -and $selectedItem.IsInGracePeriod) {
            $result = Invoke-DeprovisionCloudPC -CloudPCId $selectedItem.Id -ManagedDeviceName $selectedItem.ManagedDeviceName
            if ($result) {
                # Refresh the list after successful deprovision
                $filterTextBox.Text = ""
                Update-CloudPCList -Window $window
            }
        }
    })
    
    # Copy name menu item click event
    $copyNameMenuItem.Add_Click({
        $selectedItem = $listView.SelectedItem
        if ($selectedItem) {
            [System.Windows.Clipboard]::SetText($selectedItem.ManagedDeviceName)
            $statusText.Text = "Copied: $($selectedItem.ManagedDeviceName)"
            
            # Reset status text after 2 seconds
            $timer = New-Object System.Windows.Threading.DispatcherTimer
            $timer.Interval = [TimeSpan]::FromSeconds(2)
            $timerScript = {
                param($sender, $e)
                $statusText.Text = "Ready"
                $sender.Stop()
            }
            $timer.Add_Tick($timerScript)
            $timer.Start()
        }
    })
    
    # Load initial data
    Update-CloudPCList -Window $window
    
    # Show the window
    $window.ShowDialog() | Out-Null
}

# Start the application
Show-Win365GraceManager
