<# This form still needs a GUI added
    .NAME
        Add Windows Disk

    .DESCRIPTION
       Online, Initialize, Partition and Format New Disks.

    .NOTES
        Written for use by TechOps Server Operations at Jack Henry and Associates   

        Author:             James May
        Email:              jamay@jackhenry.com
        Last Modified:      02/21/20

        Changelog:
            a1.0             Initial Development.  Script works when run locally under specific circumstances. No checks coded yet.

#>

#Get Disk Variable
$Disks = get-disk

#Show what Get-Disk returned
Write-Host "`r`n`r`nHere's what Get-Disk looks like before" -ForegroundColor Green
$Disks

#Create Array
$OfflineDisks= @()

#Populate Array with Data
foreach ($disk in $disks) {
    If ($Disk.operationalstatus -like 'Offline') {
        $OfflineProperties = @{Number=$Disk.Number;Details=$disk.FriendlyName;Status=$disk.operationalstatus;SizeGB=($disk.size/1gb)}
        $OfflineDisk = New-Object PSobject -Property $OfflineProperties
        $OfflineDisks += $OfflineDisk
        }
    }

#Show Offline Disk Selection
Write-Host "`r`n`r`nHere are the Offline Disks" -ForegroundColor Green
$OfflineDisks | Format-Table -AutoSize -Wrap

#Prompt for Disk Number
$Number = Read-Host -Prompt 'Please enter the number of the disk you wish to online'

#online Disk
set-disk -Number $Number -IsOffline $False

#Prompt for Partition Style
$PartitionStyle = Read-Host -Prompt 'Please enter MBR or GPT'

#Initialize disk with Partition Style
Initialize-disk -Number $Number -PartitionStyle $PartitionStyle

#Prompt for Driveletter
$Driveletter = Read-Host -Prompt 'Please enter drive letter'

#Convert to string
$Driveletter = $Driveletter.ToString()

#Stop the Shell Hardware Detection Service to Suppress the "You need to Format the disk" popup dialog message.
Stop-Service -Name ShellHWDetection

#Create a Partition
New-Partition -DiskNumber $Number -Driveletter $Driveletter -UseMaximumSize

#Prompt for Allocation Unit Size
$AllocationUnitSize = Read-Host -Prompt 'Please enter Allocation Unit Size (4K, 64K, etc)'

#Convert text KB to bytes
switch ($AllocationUnitSize){
     "512b"{$AllocationUnitSize = "512"}
     "1K"  {$AllocationUnitSize = "1024"}
     "2K"  {$AllocationUnitSize = "2048"}
     "4K"  {$AllocationUnitSize = "4096"}
     "8K"  {$AllocationUnitSize = "8192"}
     "16K" {$AllocationUnitSize = "16384"}
     "32K" {$AllocationUnitSize = "32768"}
     "64K" {$AllocationUnitSize = "65536"}
   }

#Format Volume
Format-Volume -DriveLetter $Driveletter -FileSystem NTFS -AllocationUnitSize $AllocationUnitSize -confirm:$false 

#Start the Shell Hardware Detection Service back up.
Start-Service -Name ShellHWDetection

#show Disk has been Online'd
Write-Host "`r`n`r`nHere's your online info" -ForegroundColor Green
Get-Disk -Number $Number | Format-Table -Autosize

#Show Disk Blocksize Info
Write-Host "`r`n`r`nHere's your blocksize info" -ForegroundColor Green
Get-WmiObject -Class Win32_Volume | Select-Object Driveletter, BlockSize, DriveType | Where-Object {$_.DriveType -eq "3"} | Format-Table -AutoSize

#Show Partition Format
Write-Host "`r`n`r`nHere's your partition info" -ForegroundColor Green
get-partition -DriveLetter $Driveletter | Format-Table -AutoSize

