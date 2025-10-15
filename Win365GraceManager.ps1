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
        
        $pcList = @()
        foreach ($pc in $cloudPCs.value) {
            $pcList += [PSCustomObject]@{
                Id = $pc.id
                ManagedDeviceName = $pc.managedDeviceName
                UserPrincipalName = $pc.userPrincipalName
                Status = $pc.status
                ServicePlanName = $pc.servicePlanName
                GracePeriodEndDateTime = $pc.gracePeriodEndDateTime
                IsInGracePeriod = ($pc.status -eq "inGracePeriod")
            }
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

# XAML definition for the GUI
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Windows 365 Grace Period Manager" Height="600" Width="900"
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
            <TextBlock Text="Windows 365 Business Cloud PC Manager" 
                       FontSize="20" FontWeight="Bold" Margin="0,0,0,5"/>
            <TextBlock Text="Cloud PCs in grace period are highlighted in blue. Right-click to deprovision." 
                       FontSize="12" Foreground="Gray"/>
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
                    <MenuItem Name="DeprovisionMenuItem" Header="Deprovision Now (End Grace Period)"/>
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
            <StatusBarItem HorizontalAlignment="Right">
                <Button Name="RefreshButton" Content="Refresh" Width="80" Padding="5,2"/>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
"@

# Function to refresh the Cloud PC list
function Update-CloudPCList {
    param($Window)
    
    $statusText = $Window.FindName("StatusText")
    $countText = $Window.FindName("CountText")
    $listView = $Window.FindName("CloudPCListView")
    
    $statusText.Text = "Loading Cloud PCs..."
    $listView.Items.Clear()
    
    $Global:CloudPCs = Get-CloudPCs
    
    foreach ($pc in $Global:CloudPCs) {
        $listView.Items.Add($pc) | Out-Null
    }
    
    $graceCount = @(Global:CloudPCs | Where-Object { $_.IsInGracePeriod }).Count
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
    $refreshButton = $window.FindName("RefreshButton")
    $statusText = $window.FindName("StatusText")
    
    # Refresh button click event
    $refreshButton.Add_Click({
        Update-CloudPCList -Window $window
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
                Update-CloudPCList -Window $window
            }
        }
    })
    
    # Load initial data
    Update-CloudPCList -Window $window
    
    # Show the window
    $window.ShowDialog() | Out-Null
}

# Start the application
Show-Win365GraceManager
