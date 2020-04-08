<# 
    .NAME
        Add Windows Disk

    .DESCRIPTION
       Online, Initialize, Partition and Format New Disks.

    .NOTES
        Written for use by TechOps Server Operations at Jack Henry and Associates   

        Author:             James May
        Email:              jamay@jackhenry.com
        Last Modified:      04/8/20

        Changelog:
            a1.0             Initial Development.  Script works when run locally under specific circumstances. No checks coded yet.
            a2.0             GUI integrated. GetOfflineDisks, ErrorsDetected, and ProvisionDisk functions written and functional.
            a3.0             Added checks for empty values, entries that didn't match, and existing assignments.  Corrected some error status messages.
#>


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Hide PowerShell Console
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0)

<#
# Declare Global Variables
$global:disks
$global:OfflineDisks = @()
$global:errormsg
$global:vmFQDN
$global:DomainCreds
/#>

Function GetOfflineDisks{
$global:offlinedisks = @()
$global:vmFQDN = $FQDN.text.ToString()

#Split FQDN name to FQDN and Short, as well as create a domain variable
$Split = $global:vmFQDN.IndexOf(".")

    try{$VMShort = $global:vmFQDN.Substring(0, $Split)}

    catch{
        $errormsg = "$vmFQDN does not appear to be a fully qualified domain name."
        ErrorsDetected
        return
        }

$Domain = $global:vmFQDN.Substring($Split+1)

#Prompt for Credentials to $Domain
$global:DomainCreds = $host.ui.PromptForCredential("GuestVM Credentials Needed", "Please enter your credentials to access systems joined to the $Domain domain.", "", "NetBiosUserName")

#Gather Remote Disk Information
$disks = Invoke-Command -ComputerName $global:vmFQDN -Credential $global:DomainCreds {Get-Disk|select Number,Friendlyname,OperationalStatus,Size}

#Populate Array with Data
foreach ($disk in $disks) {
    If ($disk.operationalstatus -like 'Offline') {
        $OfflineProperties = @{Number=$Disk.Number;Details=$disk.FriendlyName;Status=$disk.operationalstatus;SizeGB=($disk.size/1gb)}
        $OfflineDisk = New-Object PSobject -Property $OfflineProperties
        $global:offlinedisks += $OfflineDisk
        }
    
    }

#Check for empty values
if ($global:offlinedisks.Number -eq $null) {
    $global:errormsg = "No Offline Disks Found."
    ErrorsDetected
    return
    }   

#Show Offline Disk Selection
Else {    
    $OutBox.text = $global:offlinedisks | Format-table -AutoSize | Out-String 
    $ExecuteAllSteps.visible = $true
    $DriveletterLabel.Visible = $true
    $Driveletter.visible = $true
    $PartitionStyleLabel.Visible = $true
    $PartitionStyle.Visible = $true
    $BlockSizeLabel.Visible = $true
    $AllocationUnitSize.Visible = $true
    $DiskNumberLabel.Visible = $true
    $DiskNumber.Visible = $true
    $AddDiskForm.Refresh()
    }
}

Function ErrorsDetected{
    
    $OutBox.text = $global:errormsg | Format-table -AutoSize | Out-String 
    $ExecuteAllSteps.visible = $false
    $DriveletterLabel.Visible = $false
    $Driveletter.visible = $false
    $PartitionStyleLabel.Visible = $false
    $PartitionStyle.Visible = $false
    $BlockSizeLabel.Visible = $false
    $AllocationUnitSize.Visible = $false
    $DiskNumberLabel.Visible = $false
    $DiskNumber.Visible = $false
    $AddDiskForm.Refresh()
 
    }

Function ProvisionDisk{

        #check for empty entries
    If ($DiskNumber.TextLength -eq 0){
       $global:errormsg = "Error Detected:  DiskNumber is empty.  Please enter a disk number to provision and try again."
       $Outbox.text = $global:errormsg
       $AddDiskForm.Refresh()
       return
       }
    if ($Driveletter.TextLength -eq 0){
        $global:errormsg = "Error Detected:  Driveletter is empty. Please enter a drive letter to assign during formatting and try again."
        $Outbox.text = $global:errormsg
        $AddDiskForm.Refresh()
        return
        }
    if ($PartitionStyle.SelectedIndex -eq -1) {
        $global:errormsg = "Error Detected:  Partition Style not selected.  Please choose Partition Style and try again."
        $Outbox.text = $global:errormsg
        $AddDiskForm.Refresh()
        return
        }
    if ($AllocationUnitSize.SelectedIndex -eq -1){
        $global:errormsg = "Error Detected:  Allocation Unit not selected.  Please choose Allocation Unit and try again."
        $Outbox.text = $global:errormsg
        $AddDiskForm.Refresh()
        return
        }
    #check to see if disk number is valid
    
   if ($global:offlinedisks.number -notcontains $DiskNumber.Text){
        $OfflineList =  $global:offlinedisks | Format-table -AutoSize | Out-String 
        $global:errormsg = "Error Detected:  Disk number entered does not match any offline disks.  Please choose the number of an offline disk and try again.`r`n" + $OfflineList
        $Outbox.text = $global:errormsg 
        $AddDiskForm.Refresh()
        return
        }

    #collect volume letters in use
    try {$GetVolume = Invoke-Command -ComputerName $global:vmFQDN -Credential $global:DomainCreds {Get-Volume}}
       catch {
            $global:errormsg = "Error Detected: Could not run get-volume on remote machine."  + $_ + $GetVolume.Exception
            ErrorsDetected
            return
            }
   
    #check to see if driveletter already in use
    if ($GetVolume.Driveletter -contains $Driveletter.text){
        $global:errormsg = "Error Detected:  Drive letter is already in use.  Please choose a different drive letter and try again."
        $Outbox.text = $global:errormsg
        $AddDiskForm.Refresh()
        return
        }

    #update GUI
    $statusupdates = "Attempting to Online Disk " + $DiskNumber.text
    $Outbox.text = $statusupdates
    $AddDiskForm.Refresh()
      
    #online Disk
    try { $OnlineDisk = Invoke-Command -ComputerName $global:vmFQDN -Credential $global:DomainCreds {set-disk -Number $using:DiskNumber.Text -IsOffline $False}}
    catch { $statusupdate += "...............................[Failed]`r`n"
            $global:errormsg = $statusupdates + $_ + $OnlineDisk.Exception
            ErrorsDetected
            return
            }

    #update GUI
    $statusupdates += "...............................[Success]`r`n"
    $Outbox.text = $statusupdates
    $AddDiskForm.Refresh()
    
    #update GUI
    $statusupdates += "Initializing Disk as " + $PartitionStyle.Text
    $Outbox.text = $statusupdates
    $AddDiskForm.Refresh()

    #Initialize disk with Partition Style
    try {$InitializeDisk = Invoke-Command -ComputerName $global:vmFQDN -Credential $global:DomainCreds {Initialize-disk -Number $using:DiskNumber.Text -PartitionStyle $using:PartitionStyle.Text}}
    catch { $statusupdates += "..................................[Failed]`r`n"
            $global:errormsg = $statusupdates + $_ + $InitializeDisk.Exception
            ErrorsDetected
            return
            }

    #update GUI
    $statusupdates += "..................................[Success]`r`n"
    $Outbox.text = $statusupdates
    $AddDiskForm.Refresh()

    #update GUI
    $statusupdates += "Stopping Shell Hardware Detection Service"
    $Outbox.text = $statusupdates
    $AddDiskForm.Refresh()

    #Stop the Shell Hardware Detection Service to Suppress the "You need to Format the disk" popup dialog message.
    try {$StopService = Invoke-Command -ComputerName $global:vmFQDN -Credential $global:DomainCreds {Stop-Service -Name ShellHWDetection}}
    catch { $statusupdates += ".................[Failed]`r`n"
            $global:errormsg = $statusupdates + $_ + $StopService.Exception
            ErrorsDetected
            return
            }

    #update GUI
    $statusupdates += ".................[Success]`r`n"
    $Outbox.text = $statusupdates
    $AddDiskForm.Refresh()

    #update GUI
    $statusupdates += "Creating New Partition"
    $Outbox.text = $statusupdates
    $AddDiskForm.Refresh()

    #Create a Partition
    $DLetter = $Driveletter.Text.ToString()
    try {$createpartition = Invoke-Command -ComputerName $global:vmFQDN -Credential $global:DomainCreds {New-Partition -DiskNumber $using:DiskNumber.Text -Driveletter $using:Driveletter.Text -UseMaximumSize}}
    catch { $statusupdates += "....................................[Failed]`r`n"
            $global:errormsg = $statusupdates + $_ + $createpartition.Exception
            ErrorsDetected
            return
            }

    #update GUI
    $statusupdates += "....................................[Success]`r`n"
    $Outbox.text = $statusupdates
    $AddDiskForm.Refresh()

    $AllocationUnit = $AllocationUnitSize.Text.ToString()
    switch ($AllocationUnit){
     "1K"  {$AllocationUnit = "1024"}
     "2K"  {$AllocationUnit = "2048"}
     "4K"  {$AllocationUnit = "4096"}
     "8K"  {$AllocationUnit = "8192"}
     "16K" {$AllocationUnit = "16384"}
     "32K" {$AllocationUnit = "32768"}
     "64K" {$AllocationUnit = "65536"}
   }

    #update GUI
    $statusupdates += "Formatting Disk " + $DiskNumber.Text + " as " + $Driveletter.Text + ":\ with NTFS with " + $AllocationUnitSize.Text + " block size"
    $Outbox.text = $statusupdates
    $AddDiskForm.Refresh()

    #Format Volume
    try {$FormatDisk = Invoke-Command -ComputerName $global:vmFQDN -Credential $global:DomainCreds {Format-Volume -DriveLetter $using:Driveletter.Text -FileSystem NTFS -AllocationUnitSize $using:AllocationUnit -confirm:$false}}
    catch { $statusupdates += ".....[Failed]`r`n"
            $global:errormsg = $statusupdates + $_ + $FormatDisk.Exception
            ErrorsDetected
            return
            }
    
    #update GUI
    $statusupdates += ".....[Success]`r`n"
    $Outbox.text = $statusupdates
    $AddDiskForm.Refresh()
    
    #update GUI
    $statusupdates += "Starting Shell Hardware Detection Service"
    $Outbox.text = $statusupdates
    $AddDiskForm.Refresh()

    #Start the Shell Hardware Detection Service back up.
    try {$StartService = Invoke-Command -ComputerName $global:vmFQDN -Credential $global:DomainCreds {Start-Service -Name ShellHWDetection}}
    catch { $statusupdates += ".................[Failed]`r`n"
            $global:errormsg = $statusupdates + $_ + $StartService.Exception
            ErrorsDetected
            return
            }
 
    #update GUI
    $statusupdates += ".................[Success]`r`n"
    $Outbox.text = $statusupdates
    $AddDiskForm.Refresh()

    $statusupdates += "COMPLETE!"
    $Outbox.text = $statusupdates
    $ExecuteAllSteps.visible = $false
    $DriveletterLabel.Visible = $false
    $Driveletter.visible = $false
    $PartitionStyleLabel.Visible = $false
    $PartitionStyle.Visible = $false
    $BlockSizeLabel.Visible = $false
    $AllocationUnitSize.Visible = $false
    $DiskNumberLabel.Visible = $false
    $DiskNumber.Visible = $false
    $AddDiskForm.Refresh()
    
}


$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")

$AddDiskForm                     = New-Object system.Windows.Forms.Form
$AddDiskForm.ClientSize          = '736,291'
$AddDiskForm.text                = "Add Windows Disk"
$AddDiskForm.TopMost             = $false
$AddDiskForm.Icon                = $Icon

$FQDNLabel                       = New-Object system.Windows.Forms.Label
$FQDNLabel.text                  = "Server FQDN:"
$FQDNLabel.AutoSize              = $true
$FQDNLabel.width                 = 25
$FQDNLabel.height                = 10
$FQDNLabel.location              = New-Object System.Drawing.Point(30,20)
$FQDNLabel.Font                  = 'Microsoft Sans Serif,10'

$DriveletterLabel                = New-Object system.Windows.Forms.Label
$DriveletterLabel.text           = "Drive Letter:"
$DriveletterLabel.AutoSize       = $true
$DriveletterLabel.width          = 25
$DriveletterLabel.height         = 10
$DriveletterLabel.location       = New-Object System.Drawing.Point(207,258)
$DriveletterLabel.Font           = 'Microsoft Sans Serif,10'
$DriveletterLabel.Visible        = $false

$PartitionStyleLabel             = New-Object system.Windows.Forms.Label
$PartitionStyleLabel.text        = "Partition Style:"
$PartitionStyleLabel.AutoSize    = $true
$PartitionStyleLabel.width       = 25
$PartitionStyleLabel.height      = 10
$PartitionStyleLabel.location    = New-Object System.Drawing.Point(348,258)
$PartitionStyleLabel.Font        = 'Microsoft Sans Serif,10'
$PartitionStyleLabel.Visible     = $false

$PartitionStyle                  = New-Object system.Windows.Forms.ListBox
$PartitionStyle.width            = 57
$PartitionStyle.height           = 37
@('MBR','GPT') | ForEach-Object {[void] $PartitionStyle.Items.Add($_)}
$PartitionStyle.location         = New-Object System.Drawing.Point(447,245)
$PartitionStyle.Visible        = $false

$BlockSizeLabel                  = New-Object system.Windows.Forms.Label
$BlockSizeLabel.text             = "Allocation Unit:"
$BlockSizeLabel.AutoSize         = $true
$BlockSizeLabel.width            = 25
$BlockSizeLabel.height           = 10
$BlockSizeLabel.location         = New-Object System.Drawing.Point(524,258)
$BlockSizeLabel.Font             = 'Microsoft Sans Serif,10'
$BlockSizeLabel.Visible          = $false

$AllocationUnitSize              = New-Object system.Windows.Forms.ListBox
$AllocationUnitSize.text         = "listBox"
$AllocationUnitSize.width        = 51
$AllocationUnitSize.height       = 99
@('1K','2K','4K','8K','16K','32K','64K') | ForEach-Object {[void] $AllocationUnitSize.Items.Add($_)}
$AllocationUnitSize.location     = New-Object System.Drawing.Point(621,188)
$AllocationUnitSize.Visible      = $flase

$DiskNumberLabel                 = New-Object system.Windows.Forms.Label
$DiskNumberLabel.text            = "Disk Number: "
$DiskNumberLabel.AutoSize        = $true
$DiskNumberLabel.width           = 25
$DiskNumberLabel.height          = 10
$DiskNumberLabel.location        = New-Object System.Drawing.Point(30,258)
$DiskNumberLabel.Font            = 'Microsoft Sans Serif,10'
$DiskNumberLabel.Visible         = $false

$ExecuteGetDisk                  = New-Object system.Windows.Forms.Button
$ExecuteGetDisk.text             = "Get Offline Disks"
$ExecuteGetDisk.width            = 127
$ExecuteGetDisk.height           = 59
$ExecuteGetDisk.location         = New-Object System.Drawing.Point(600,16)
$ExecuteGetDisk.Font             = 'Microsoft Sans Serif,10'
$ExecuteGetDisk.add_click({GetOfflineDisks})

$ExecuteAllSteps                 = New-Object system.Windows.Forms.Button
$ExecuteAllSteps.text            = "Provision Disk"
$ExecuteAllSteps.width           = 127
$ExecuteAllSteps.height          = 59
$ExecuteAllSteps.location        = New-Object System.Drawing.Point(600,91)
$ExecuteAllSteps.Font            = 'Microsoft Sans Serif,10'
$ExecuteAllSteps.Visible         = $false
$ExecuteAllSteps.add_click({ProvisionDisk})

$FQDN                            = New-Object system.Windows.Forms.TextBox
$FQDN.multiline                  = $false
$FQDN.width                      = 458
$FQDN.height                     = 20
$FQDN.location                   = New-Object System.Drawing.Point(129,16)
$FQDN.Font                       = 'Microsoft Sans Serif,10'

$DiskNumber                      = New-Object system.Windows.Forms.TextBox
$DiskNumber.multiline            = $false
$DiskNumber.width                = 55
$DiskNumber.height               = 20
$DiskNumber.location             = New-Object System.Drawing.Point(124,253)
$DiskNumber.Font                 = 'Microsoft Sans Serif,10'
$DiskNumber.Visible              = $false

$Driveletter                     = New-Object system.Windows.Forms.TextBox
$Driveletter.multiline           = $false
$Driveletter.width               = 42
$Driveletter.height              = 20
$Driveletter.location            = New-Object System.Drawing.Point(288,253)
$Driveletter.Font                = 'Microsoft Sans Serif,10'
$Driveletter.Visible             = $false

$OutBox                          = New-Object system.Windows.Forms.TextBox
$OutBox.multiline                = $true
$OutBox.text                     = "Please enter a fully qualified domain name (FQDN) and click Get Offline Disks to begin."
$OutBox.width                    = 566
$OutBox.height                   = 185
$OutBox.location                 = New-Object System.Drawing.Point(22,48)
$OutBox.Font                     = 'Courier New,10'

$AddDiskForm.controls.AddRange(@($FQDNLabel,$DriveletterLabel,$PartitionStyleLabel,$PartitionStyle,$BlockSizeLabel,$AllocationUnitSize,$DiskNumberLabel,$ExecuteGetDisk,$ExecuteAllSteps,$FQDN,$DiskNumber,$Driveletter,$OutBox))
   

$AddDiskForm.topmost = $true
$result = $AddDiskForm.ShowDialog()