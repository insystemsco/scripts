<#
.SYNOPSIS
Displays a GUI of Active Directory users that are locked out and can be unlocked.

.DESCRIPTION
ADLockoutViewer displays a GUI of Active Directory users in specific OUs. Any user that is locked will have a check by their name. Unchecking the box will unlock the user. OUs can be chosen from the drop down menu.

The ComputerName parameter is the computer that has the ActiveDirectory PowerShell module installed if the local computer does not.

This command can be used with the CredSSP authentication method. This is useful when you want to use it on a computer that doesn't have the ActiveDirectory PowerShell module installed, but the computer being connected to is not a domain controller. When that happens the CredSSP authentication method has to be used, and the Fully Qualified Domain Name (FQDN) of the remote computer must be used for the ComputerName parameter. Also, the user will be required to enter Administrator level credentials in a separate dialog.
.PARAMETER Authentication
Specifies the mechanism that is used to authenticate the user's credentials.   Valid values are "Default", "Basic", "Credssp", "Digest", 
"Kerberos", "Negotiate", and "NegotiateWithImplicitCredential".  The default value is "Default".
        
For more information about the values of this parameter, see the description of the 
System.Management.Automation.Runspaces.AuthenticationMechanism enumeration in the MSDN (Microsoft Developer Network) library at 
http://go.microsoft.com/fwlink/?LinkID=144382.
        
Caution: Credential Security Support Provider (CredSSP) authentication, in which the user's credentials are passed to a remote computer 
to be authenticated, is designed for commands that require authentication on more than one resource, such as accessing a remote network 
share. This mechanism increases the security risk of the remote operation. If the remote computer is compromised, the credentials that 
are passed to it can be used to control the network session.
.PARAMETER ComputerName
The name of the computer that has the ActiveDirectory PowerShell module installed.

Note: If the CredSSP authentication method is used, the Fully Qualified Domain Name (FQDN) of the computer must be specified.
.PARAMETER IconPath
Specifies the path to an icon file to be used in the GUI.
.INPUTS
ADLockoutViewer accepts no inputs.
.OUTPUTS
ADLockoutViewer produces no output.
.EXAMPLE
PS C:> ADLockoutViewer

This command will display the GUI and is assuming the ActiveDirectory PowerShell module is installed on the local machine.
.EXAMPLE
PS C:> ADLockoutViewer -ComputerName DC01

This command will display the GUI by sending remote commands to the computer named DC01.
.EXAMPLE
PS C:> ADLockoutViewer -Authentication Credssp -ComputerName WK01.domain.local

This command will display the GUI by sending remote commands to the computer named WK01 using CredSSP. Note the use of the Fully Qualified Domain Name (FQDN) used which is required when using CredSSP.
.NOTES
The remote computer specified in the ComputerName parameter must have PowerShell Remoting enabled if it is not the local computer.
#>

#Requires -Version 3.0
[CmdletBinding(HelpURI='https://gallery.technet.microsoft.com/View-and-Unlock-Active-ef0d5757')]

Param(
    [System.Management.Automation.Runspaces.AuthenticationMechanism]
        $Authentication="Default",
    
    [Alias("CN")]
        [String]$ComputerName="$env:COMPUTERNAME",

    [Alias("Icon")]
        [String]$IconPath
)

#Add assemblies needed for GUI presentation
Add-Type -AssemblyName PresentationCore,PresentationFrameWork,System.Windows.Forms,System.Drawing

#------------------------------ Begin Region Helper Functions -------------------

#Helper function to create message boxes for alerts, errors, etc.
Function Create-MessageBox {
    Param (
        [Parameter (Mandatory=$true)][string]$Message,
        [Parameter (Mandatory=$true)][string]$Title,
        [Parameter (Mandatory=$false)][System.Windows.Forms.MessageBoxButtons]$Buttons="OK",
        [Parameter (Mandatory=$false)][System.Windows.Forms.MessageBoxIcon]$Icon="Information"
    )
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, $Buttons, $Icon)
}

Function Add-CommandAndKeyBinding ($CanExec,$Exec,$Key,$KeyModifiers,$KeyDesc,$CommandName,$CommandDesc,$CommandBindTarget,$KeyBindTarget) {
    #Build KeyGesture for Command
    $KeyGesture = New-Object System.Windows.Input.KeyGesture -ArgumentList `
        $Key, $KeyModifiers, $KeyDesc

    $GestColl = New-Object System.Windows.Input.InputGestureCollection
    $GestColl.Add($KeyGesture) | Out-Null

    #Create Routed Command
    $Command = New-Object System.Windows.Input.RoutedUICommand -ArgumentList $CommandDesc,$CommandName,([Type]"System.Object"),$GestColl

    #Create Command Binding
    $CommandBind = New-Object System.Windows.Input.CommandBinding -ArgumentList $Command,$Exec,$CanExec

    #Add to target
    $CommandBindTarget.CommandBindings.Add($CommandBind) | Out-Null

    #Add Key Binding to target
    if ($KeyBindTarget) {
        $KeyBindTarget.Command = $Command
    }
}

#------------------------------- End Region Helper Functions --------------------

#------------------------------- Begin Region Admin Check -----------------------

$noAdminMsg = "This script is not being run with Administrator privileges. `
This may cause it to fail when making changes to the database. `
If you choose 'Yes', you will be prompted for Administrator credentials. `
Do you want to continue?"
$noAdminTitle = "Administrator Privileges Required"

#Check for admin
if (!([System.Security.Principal.WindowsPrincipal][System.Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [System.Security.Principal.WindowsBuiltInRole]::Administrator))
{
    $dialog = Create-MessageBox -Message $noAdminMsg -Title $noAdminTitle -Buttons YesNo
    switch ($dialog) {
        "Yes" {
            #$adminCred = Get-Credential
            Start-Process -FilePath PowerShell.exe -Verb RunAs -ArgumentList "-WindowStyle Hidden", "-File $PSCommandPath"
            exit
        }
        "No" {
            exit
        }
    }
}

#------------------------------- End Region Admin Check -------------------------

#------------------------------- Begin Region Setup CredSSP and Remote Sesssion -

#Test for Remote Computer to use for Active Directory commands
if ($ComputerName -ne $env:COMPUTERNAME) {
    #Set flag to use for testing for remote session requirement
    $useRemote = $true

    #Test for CredSSP to turn it on
    if ($Authentication -eq "CredSSP") {
        #Turn on CredSSP server on management computer
        Invoke-Command -ScriptBlock {Enable-WSManCredSSP -Role Server -Force} -ComputerName $ComputerName

        #Turn on CredSSP client to allow managment computer to delegate credentials
        Enable-WSManCredSSP -Role Client -DelegateComputer $ComputerName -Force

        $credSplat = @{Credential=(Get-Credential)}
    }

    #Create a new session using CredSSP to be used when running remote commands
    $session = New-PSSession -ComputerName $ComputerName -Authentication $Authentication @credSplat
}

#-------------------------------- End Region Setup CredSSP and Remote Session ---

#------------------------------- Begin Region Load Collections -------------------

$obCollOUs = New-Object 'System.Collections.ObjectModel.ObservableCollection[System.Object]'
$obCollUsers = New-Object 'System.Collections.ObjectModel.ObservableCollection[System.Object]'

#To mitigate installing RSAT on every computer, connect to the management computer to use
#its copy of the Active Directory PowerShell Module if the local computer is not the management computer

$sb = {Get-ADOrganizationalUnit -Filter * | Where-Object {$_.Name -ne 'Domain Controllers'} |
        Select-Object -Property DistinguishedName, Name | Sort-Object -Property Name}

#Check if a remote session is being used
if ($useRemote) {
    #Use session created earlier to get the users from the management computer
    $result = Invoke-Command -Session $session -ScriptBlock $sb
}
else {
    #No remote session, invoke scriptblock
    $result = & $sb
}

foreach ($r in $result) {$obCollOUs.Add($r)}

<#
#>

#-------------------------------- End Region Load Collections --------------------

#-------------------------- Begin Region XAML -----------------------------------------

[xml]$xaml = @'
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        Title="Active Directory Lockout Viewer" Height="350" Width="525">
    
    <Window.Resources>
        <Style x:Key="ColumnsElementStyle">
            <Setter Property="FrameworkElement.VerticalAlignment" Value="Center"/>
            <Setter Property="FrameworkElement.Margin" Value="5,2"/>
        </Style>

        <Style x:Key="TextColumnsStyle" BasedOn="{StaticResource ColumnsElementStyle}">
            <Setter Property="FrameworkElement.HorizontalAlignment" Value="Left"/>
        </Style>

        <Style x:Key="CheckBoxColumnsStyle" BasedOn="{StaticResource ColumnsElementStyle}">
            <Setter Property="FrameworkElement.HorizontalAlignment" Value="Center"/>
        </Style>
    </Window.Resources>

    <DockPanel>
        <Menu DockPanel.Dock="Top">
            <MenuItem x:Name="MiHelp" Header="_Help">
                <MenuItem x:Name="MiSubHelp" Header="_Help"/>
                <MenuItem x:Name="MiAbout" Header="_About"/>
            </MenuItem>
        </Menu>
        <ComboBox x:Name="CmbOU" DockPanel.Dock="Top" Margin="10,10,0,10" Width="200" HorizontalAlignment="Left" DisplayMemberPath="Name"/>
        <DataGrid x:Name="DtgUsers" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" HorizontalGridLinesBrush="LightGray" VerticalGridLinesBrush="LightGray">
            <DataGrid.Columns>
                <DataGridCheckBoxColumn Header="Locked" Binding="{Binding LockedOut}" ElementStyle="{StaticResource CheckBoxColumnsStyle}"/>
                <DataGridTextColumn Header="Account Name" Binding="{Binding SamAccountName}" ElementStyle="{StaticResource TextColumnsStyle}" IsReadOnly="True"/>
                <DataGridTextColumn Header="Name" Binding="{Binding DisplayName}" ElementStyle="{StaticResource TextColumnsStyle}" IsReadOnly="True"/>
                <DataGridTextColumn Header="Description" Binding="{Binding Description}" ElementStyle="{StaticResource TextColumnsStyle}" IsReadOnly="True"/>
            </DataGrid.Columns>
        </DataGrid>
    </DockPanel>
</Window>
'@

#About Window XAML
[xml]$aboutXaml = @'
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="About Active Directory Lockout Viewer" ResizeMode="NoResize" SizeToContent="WidthAndHeight">

    <Window.Resources>
        <Style x:Key="{x:Type Hyperlink}" TargetType="{x:Type Hyperlink}">
            <Setter Property="Foreground">
                <Setter.Value>
                    <SolidColorBrush Color="White"/>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Foreground">
                        <Setter.Value>
                            <SolidColorBrush Color="WhiteSmoke"/>
                        </Setter.Value>
                    </Setter>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <StackPanel Width="292">
        <DockPanel>
            <Image Source="" Width="48" Height="48" Margin="25,20,15,20" Name="AboutIcon" />
            <TextBlock TextWrapping="Wrap" Text="Active Directory Lockout Viewer" FontSize="20" Margin="0" VerticalAlignment="Center" Name="AboutHeaderTitle">
                <TextBlock.Foreground>
                    <SolidColorBrush Color="#34495e"/>
                </TextBlock.Foreground>
            </TextBlock>
        </DockPanel>
        <StackPanel Background="LightBlue">
            <TextBlock TextWrapping="Wrap" Margin="25,20,0,20" FontSize="14">
                <TextBlock.Foreground>
                    <SolidColorBrush Color="White"/>
                </TextBlock.Foreground>

                <Run Text="ADLockoutViewer.ps1"/><LineBreak/>
                <Run FontSize="12" Text="Version: 1.1"/><LineBreak/>
                <Run FontSize="12" Text="License: TechNet"/><LineBreak/>
                <Run FontSize="12" Text="© 2016 "/>
                <Hyperlink x:Name="ProfileLink" NavigateUri="https://social.technet.microsoft.com/profile/chris%20carter%2079/">Chris Carter</Hyperlink>
            </TextBlock>
        </StackPanel>
    </StackPanel>

</Window>
'@

#Help Window XAML
[xml]$helpXaml = @'
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Help" Height="350" Width="366" ResizeMode="CanResize" MinWidth="366" MaxWidth="500" MinHeight="353" MaxHeight="600">
    <DockPanel HorizontalAlignment="Left" VerticalAlignment="Top">
        <DockPanel DockPanel.Dock="Top">
            <Image Source="" Width="32" Height="32" Margin="25,25,10,15" Name="HelpHeaderIcon" />
            <TextBlock TextWrapping="Wrap" Text="Active Directory Lockout Viewer Help" VerticalAlignment="Center" FontSize="16" Foreground="#34495e" Name="HelpHeaderTitle"/>
        </DockPanel>
        <ScrollViewer x:Name="SVHelp" VerticalScrollBarVisibility="Auto" DockPanel.Dock="Top">
            <DockPanel Margin="0,0,0,25">
                <TextBlock TextWrapping="Wrap" Margin="25,0,25,5" Text="Description" FontSize="14" FontWeight="Bold" DockPanel.Dock="Top"/>
                <TextBlock TextWrapping="Wrap" Margin="25,0,25,10" Text="The Active Directory Lockout Viewer is a PowerShell script that is used to view Active Directory users' lockout status, and unlock them if desired." DockPanel.Dock="Top"/>
                <TextBlock TextWrapping="Wrap" Margin="25,0,25,5" Text="Unlocking Users" FontSize="14" FontWeight="Bold" DockPanel.Dock="Top"/>
                <TextBlock TextWrapping="Wrap" Margin="25,0,25,10" DockPanel.Dock="Top" Text="Unchecking the box beside a user will unlock the user. There is no command to lock a user, so checking the box beside a user will not do anything, but unchecking it again would send the unlock command again."></TextBlock>
            </DockPanel>
        </ScrollViewer>
    </DockPanel>
</Window>
'@

#------------------------------ End Region XAML ---------------------------------

#------------------------------- Begin Region Event Handlers ---------------------

Function On-LockedCheckBoxUnchecked {
    $source = $this.DataContext

    #Check if remote session is being used
    if ($useRemote) {
        #Use the previous session to connect to the server through the management computer
        Invoke-Command -Session $session -ScriptBlock {Unlock-ADAccount -Identity $args[0]} -Args $source.SamAccountName
    }
    else {
        #No remote session, run locally
        Unlock-ADAccount -Identity $source.SamAccountName
    }
}

Function On-CmbSelectionChanged {
    #Clear the collection of previous entries
    $obCollUsers.Clear()

    $sb = {Get-ADUser -Filter * -SearchBase $args[0].SelectedItem.DistinguishedName -Properties * | 
            Where-Object {$_.Enabled} |
            Select-Object -Property LockedOut, SamAccountName, DisplayName, Description, Enabled |
            Sort-Object -Property SamAccountName 
    }

    #Check if remote session is being used
    if ($useRemote) {
        #Use session created earlier to get the users from the server through the management computer
        $result = Invoke-Command -Session $session -ScriptBlock $sb -Args $this
    }
    else {
        #No remote session, invoke scriptblock
        $result = & $sb $this
    }

    #Load the users into the observable collection
    foreach ($r in $result) {$obCollUsers.Add($r)}
}

#About Window Event Handlers
Function On-ProfileRequestNavigate {
    [System.Diagnostics.Process]::Start($_.Uri.ToString())
}

#Modal Dialog Creation Functions

Function New-AboutDialogBox {
    #Load About Window
    $aboutXmlReader = New-Object System.Xml.XmlNodeReader -ArgumentList $aboutXaml
    $AboutWindow = [System.Windows.Markup.XamlReader]::Load($aboutXmlReader)
    $AboutWindow.Owner = $this
    $AboutWindow.WindowStyle = "ToolWindow"

    #Set Big Icon for Presentation
    #Remove Title Bar Icon
    $BigIcon = $AboutWindow.FindName("AboutIcon")
    $HeaderTitle = $AboutWindow.FindName("AboutHeaderTitle")

    if ($icon) {
        $BigIcon.Source = $icon
    }
    else {
        $BigIcon.Visibility = "Collapsed"
        $HeaderTitle.Margin = New-Object System.Windows.Thickness -ArgumentList 25,20,0,20
    }

    #About Window Controls
    $ProfileLink = $AboutWindow.FindName("ProfileLink")

    #About Window Evnts
    $ProfileLink.add_RequestNavigate({On-ProfileRequestNavigate})

    $AboutWindow.ShowDialog()
}

Function New-HelpDialogBox {
    #Load Help Window
    $helpXmlReader = New-Object System.Xml.XmlNodeReader -ArgumentList $helpXaml
    $HelpWindow = [System.Windows.Markup.XamlReader]::Load($helpXmlReader)
    $HelpWindow.Owner = $this

    $BigIcon = $HelpWindow.FindName("HelpHeaderIcon")
    $HeaderTitle = $HelpWindow.FindName("HelpHeaderTitle")

    if ($icon) {
        $BigIcon.Source = $icon
        $HelpWindow.Icon = $icon
    }
    else {
        $BigIcon.Visibility = "Collapsed"
        $HeaderTitle.Margin = New-Object System.Windows.Thickness -ArgumentList 25,25,0,15
    }

    $HelpWindow.ShowDialog()
}

#------------------------------- End Region Event Handlers -----------------------

#-------------------------------- Begin Region Load Window and UI Components -----

#Read and Load the XAML definition to get the Window object
$xmlReader = New-Object System.Xml.XmlNodeReader -ArgumentList $xaml
$Window = [System.Windows.Markup.XamlReader]::Load($xmlReader)

$MenuAbout = $Window.FindName("MiAbout")
$MenuHelp = $Window.FindName("MiSubHelp")

$DtgUsers = $Window.FindName("DtgUsers")
$DtgUsers.ItemsSource = $obCollUsers

$CmbOu = $Window.FindName("CmbOU")
$CmbOu.ItemsSource = $obCollOUs
$CmbOu.add_SelectionChanged({On-CmbSelectionChanged})

#-------------------------------- End Region Load Window and UI Components -------

#-------------------------------- Begin Region Attach Event Handlers -------------

#Construct an EventSetter for the Report DataGrid's CheckBox column and add Unchecked event handler
#Unchecked Event Setter creation
$evSetterUnchkd = New-Object System.Windows.EventSetter
$evSetterUnchkd.Event = [System.Windows.Controls.CheckBox]::UncheckedEvent
$evSetterUnchkd.Handler = [System.Windows.RoutedEventHandler]{On-LockedCheckBoxUnchecked}
#Add EventSetters to Style and add Style to CheckBox column
$chkBoxCellStyle= New-Object System.Windows.Style
$chkBoxCellStyle.Setters.Add($evSetterUnchkd)
$DtgUsers.Columns[0].CellStyle = $chkBoxCellStyle

#-------------------------------- End Region Attach Event Handlers --------------

#-------------------------------- Begin Region Add Menu Commands and Key Bindings

$AutoTrueCanExec = {$_.CanExecute = $true}

#Command for About
$AboutExec = {New-AboutDialogBox}

Add-CommandAndKeyBinding $AutoTrueCanExec $AboutExec "A" ("Control","Shift") "" "About" "_About" $Window $MenuAbout

#Command for Help
$HelpExec = {New-HelpDialogBox}

Add-CommandAndKeyBinding $AutoTrueCanExec $HelpExec "F1" "" "" "Help" "_View Help" $Window $MenuHelp

#-------------------------------- End Region Add Menu Commands and Key Bindings -

#Icon File
if (Test-Path $IconPath) {
    $icon = New-Object System.Windows.Media.Imaging.BitmapImage -ArgumentList $IconPath
}
if ($icon) {$Window.Icon = $icon}

#Start UI
$Window.ShowDialog() | Out-Null

#-----------------------Begin Region Remove Session and Disable CredSSP----------

#Remove remote session and Disable CredSSP if used
if ($useRemote) {
    Remove-PSSession $session

    if ($Authentication -eq "CredSSP") {
        Disable-WSManCredSSP -Role Client
        Invoke-Command -ScriptBlock {Disable-WSManCredSSP -Role Server} -ComputerName $ComputerName
    }
}
