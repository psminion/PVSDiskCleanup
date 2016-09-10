[cmdletBinding()]param()
Set-StrictMode -Version Latest

# load Functions
$functions = Get-ChildItem -Path $PSScriptroot\Functions
foreach ($function in $functions) {
    . $function.FullName
}
Start-Transcript $TranscriptFile

# verify script executed with admin creds
Check-CLNAdmin

# validate fields in CSV file
Test-CLNCsvData -CleanupCsv $CleanupCsv

# prompt for VHD(x) file to be cleaned
$VHDFile = Get-CLNVHDFileName
$vDiskName = ($VHDFile.Split("\")[-1]).Split(".")[0]
if (-Not($VHDFile)) {
    Write-Warning "No file selected ... script aborted"
    Stop-Transcript; Throw 'No VHD file selected'
}

# verify existence of cleanupcsv & vhd files
$TestFiles = $CleanupCsv,$VHDFile
foreach ($TestFile in $TestFiles) {
    if (-Not(Test-Path $TestFile)) {
        Write-Warning "Unable to locate $TestFile ... script execution halted"
        Stop-Transcript; Throw "Unable to locate $TestFile"
    }
}

# check if VHD is local
if ($VHDFile -match "^\\\\") {
    Write-Warning "Cleanup of remote VHD files is not recommended"
    $Verify = Read-Host -Prompt "Are you certain you want to proceed using this remote VHD? (y/n)"
    if (-Not($Verify -ieq "y")) {
        Stop-Transcript; Throw "Script execution halted by user"
    }
}

# query for desired UTC setting
$SetUTC = Confirm-CLNUTC

# attempt to load PVS snapin & SQLPS module
$CheckPVS = Import-CLNMcliSnapin
if ($CheckPVS) {
    Write-Verbose "Loaded McliPSSnapin"
    Push-Location (Get-Location)
    $CheckPVS = Import-CLNSQLPS -ModuleName SQLPS
    if ($CheckPVS) {Write-Verbose "Loaded SQLPS module"}
    Pop-Location
}
Write-Verbose "Current location is $((Get-Location).Path)"

# get PVS/vDisk stats (if snapin/module loads were successful)
if ($CheckPVS) {
    # Get DB info from MCLI
    Write-Host "`nGathering PVS info for VHD file ..." -BackgroundColor White -ForegroundColor Black
    mcli-run setupconnection | Out-Null #connect to localhost
    $PVSFarm = Convert-McliToObjects -mclidata (Mcli-Get Farm)
    $PVSDBServer = $PVSFarm.databaseServerName
    $PVSDB = $PVSFarm.databaseName
    $PVSFarmName = $PVSFarm.farmName
    Write-Host "Farm=$PVSFarmName, SQLServer=$PVSDBServer, DB=$PVSDB" -BackgroundColor White -ForegroundColor Black

    # Run SQL queries to get vDisk / device data
    $AllDevices = Get-CLNDeviceData -PVSDB $PVSDB -PVSDBServer $PVSDBServer
    $InUseDevices = Get-CLNInUseDevices -PVSDB $PVSDB -PVSDBServer $PVSDBServer
    $DiskAssignments = Get-CLNDiskAssignments -PVSDB $PVSDB -PVSDBServer $PVSDBServer

    # Check cache type and number of assignments
    $PVSDisks = $DiskAssignments | Select-Object -Property diskLocatorName, writeCacheType -Unique
    $DiskList = $PVSDisks | Select-Object -ExpandProperty diskLocatorName
    $DiskInPVS = $False
    if ($vDiskName -in $DiskList) {
        $DiskInPVS = $True
        $AssignedCount = Get-CLNAssignedCount -DiskAssignments $DiskAssignments -vDiskName $vDiskName
        $InUseCount = Get-CLNInUseCount -InUseDevices $InUseDevices -vDiskName $vDiskName
        $CacheType = Get-CLNCacheType -PVSDisks $PVSDisks -vDiskName $vDiskName
    }
}

# output PVS data to user
    if ($CheckPVS) {
    Write-Host "`nVHD Info:" -BackgroundColor White -ForegroundColor Black    #prompt to continue
    if ($DiskInPVS) {
        # Write out vdisk data
        Write-Host "`tVHD: '$vDiskName'" -BackgroundColor White -ForegroundColor Black
        Write-Host "`tInUse Count:$InUseCount" -BackgroundColor White -ForegroundColor Black
        Write-Host "`tAssignment Count: $AssignedCount" -BackgroundColor White -ForegroundColor Black
        Write-Host "`tCache Type: $CacheType" -BackgroundColor White -ForegroundColor Black
        if ($SetUTC) {
            Write-Host "`tTimeZone: WILL be set to UTC-0`n" -BackgroundColor Yellow -ForegroundColor Black
        } else {
            Write-Host "`tTimeZone: WILL NOT be set to UTC-0`n" -BackgroundColor Yellow -ForegroundColor Black
        }
    } else {
        Write-Host "`n`tvDisk '$vDiskName' does not exist in`n`tPVS Farm '$PVSFarmName' or is not assigned to any devices" -BackgroundColor Yellow -ForegroundColor Black
        if ($SetUTC) {
            Write-Host "`tTimeZone: WILL be set to UTC-0`n" -BackgroundColor Yellow -ForegroundColor Black
        } else {
            Write-Host "`tTimeZone: WILL NOT be set to UTC-0`n" -BackgroundColor Yellow -ForegroundColor Black
        }
    }
    $Answer = Read-Host "Continue with cleanup process? (y/n)"
    Switch ($Answer) {
        'Y' {"Continuing ..."}
        'N' {Stop-Transcript; Throw "Script halted by user"}
        Default {Stop-Transcript; Throw "Script halted by user"}
    }
} else {
    Write-Host "PVS data not available for '$vDiskName'" -BackgroundColor Yellow -ForegroundColor Black
    $Answer = Read-Host "Continue with cleanup process? (y/n)"
    Switch ($Answer) {
        'Y' {"Continuing ..."}
        'N' {Stop-Transcript; Throw "Script halted by user"}
        Default {Stop-Transcript; Throw "Script halted by user"}
    }
}

# mount VHD
Mount-CLNVerify -VHDFile $VHDFile -ExpectedState $False
$DriveLetter = Mount-CLNVHD -VHDFile $VHDFile #Mount VHD and return drive letter
Write-Host "'$VHDFile' mounted as drive '$DriveLetter'" -BackgroundColor White -ForegroundColor Black

# registry clean Loop
$DataLocs = 'Reg-SOFTWARE','Reg-SYSTEM'
$RegImport = Import-Csv $CleanupCsv | Where-Object {$PSItem.DataLoc -in $DataLocs}
$FileSystemImport = Import-Csv $CleanupCsv | Where-Object {$PSItem.DataLoc -eq 'FileSystem'}
foreach ($RegHive in $DataLocs) {
    $Settings = $RegImport | Where-Object {$PSItem.DataLoc -match $RegHive}
    $HiveType = $RegHive.split("-")[1]
    $HiveName = "HKLM\vDisk$HiveType"
    $PSDriveName = "vDisk$HiveType"
    $RegFilePath = "{0}:\Windows\System32\config\$HiveType" -f $DriveLetter
    Load-CLNRegHive -HiveName $HiveName -RegFilePath $RegFilePath
    Create-CLNPSDrive -PSDriveName $PSDriveName -HiveName $HiveName
    if ($RegHive -eq 'Reg-SYSTEM') {
        $CurrControl = "ControlSet00$((Get-ItemProperty vDiskSYSTEM:\Select -Name Current).Current)"
        Write-Host "CurrentControlSet = $CurrControl"
    }
    Write-Host "Processing '$RegHive' settings" -BackgroundColor Blue
    foreach ($Setting in $Settings) {
        $Action = $Path = $RegValueName = $RegValueData = $RegType = $null
        $Action = $Setting.Action
        $Path = $Setting.Path
        if ($RegHive -eq 'Reg-SYSTEM') {$Path = $Path.Replace('CurrentControlStub',"$CurrControl")}
        $RegValueName = $Setting.RegValueName
        $RegValueData = $Setting.RegValueData
        $RegType = $Setting.RegType
        $Category = $Setting.Category
        #Skip UTC (if user selected skip option)
        if ($Category -eq 'UTC_Offset' -AND (-Not($SetUTC))) {
            Write-Warning "Skipping $Path\$RegValueName based on user choice"
            Continue
        }

        # SOFTWARE Hive Cleanup
        if ($HiveType -eq 'SOFTWARE') {
            if ($Action -eq 'ADD') {
                $RegType = Convert-CLNRegValueType -RegType $RegType #Validate/fix allowed reg value type
                Add-CLNRegData -PSDriveName $PSDriveName -Path $Path -RegValueName $RegValueName -RegValueData $RegValueData -RegType $RegType #Check/Add key/value
            }
            if ($Action -eq 'DEL') {
                Remove-CLNRegData -PSDriveName $PSDriveName -Path $Path -RegValueName $RegValueName #Check/Remove key/value
            }
        }

        # SYSTEM Hive Cleanup
        if ($HiveType -eq 'SYSTEM') {
            if ($Action -eq 'ADD') {
                $RegType = Convert-CLNRegValueType $RegType #Validate/fix allowed reg value type
                Add-CLNRegData -PSDriveName $PSDriveName -Path $Path -RegValueName $RegValueName -RegValueData $RegValueData -RegType $RegType #Check/Add key/value
            }
            if ($Action -eq 'DEL') {
                Remove-CLNRegData -PSDriveName $PSDriveName -Path $Path -RegValueName $RegValueName #Check/Remove key/value
            }
        }
    }
    [gc]::Collect()
    Unload-CLNRegHive -HiveName $HiveName -RegFilePath $RegFilePath -PSDriveName $PSDriveName
}

# file system cleanup loop
Write-Host "Processing FileSystem settings" -BackgroundColor Blue
foreach ($Entry in $FileSystemImport) {
    $Action = $Entry.Action
    $Path = $Entry.Path
    $TargetDir = Join-Path "$DriveLetter`:" $Path
    if (-Not($Action -eq 'DEL')) {
        Write-Host ""
        Write-Warning "The requested action ('$Action') is invalid. Only 'DEL' operations are supported."
        Write-Warning "No action taken against item $Path"
        Write-Host ""
        Return
    }
    Write-Host "Processing $Action on $TargetDir"
    if ($Path -match "\\\*$") { # Path ends with "\*" - files in dir should be deleted, folder not deleted
        Write-Host "`tChecking/deleting files and folders in directory"
        if (Test-Path ($TargetDir.TrimEnd("*"))) { # dir exists, delete files and subdirs
            $DelItems = Get-ChildItem ($TargetDir.TrimEnd("*"))
            if ($DelItems) { # files/dirs found - proceed with deletions
                $RemoveResult = Remove-Item -Path $TargetDir -Force -Recurse -ErrorAction SilentlyContinue
                if ($RemoveResult) {$RemoveResult.Handle.Close()}
                Write-Host "`t$($TargetDir.TrimEnd("*")), files and subdirs deleted" -BackgroundColor Yellow -ForegroundColor Black
            } else {
                Write-Host "`tDirectory already empty, no deletions performed"
            }
        } else {
            Write-Host "$($TargetDir.TrimEnd("*")) does not exist, no action taken"
        }
    } else { # Path does not end with "\*" - directory should be deleted
        Write-Host "`tChecking/deleting directory"
        if (Test-Path $TargetDir) {
            $RemoveResult = Remove-Item -Path $TargetDir -Force -Recurse -ErrorAction SilentlyContinue
            if ($RemoveResult) {$RemoveResult.Handle.Close()}
            Write-Host "`t$TargetDir directory deleted" -BackgroundColor Yellow -ForegroundColor Black
        } else { #dir does not exits
            Write-Host "`tDirectory does not exist, no action taken"
        }
    }
}

# script completion / cleanup
[gc]::Collect()
if (Get-Module -Name SQLPS -ErrorAction SilentlyContinue) {Remove-Module -Name SQLPS}
Dismount-CLNVHD -VHDFile $VHDFile
Write-Host "'$VHDFile' successfully dismounted" -BackgroundColor White -ForegroundColor Black
Write-Host "Script execution COMPLETE" -BackgroundColor White -ForegroundColor Black
$ScriptUser = ([Security.Principal.WindowsIdentity]::GetCurrent().Name).ToUpper()
$DateTime = Get-Date -Format g
Write-Host "`nScript Stats`n------------:"
Write-Host "vDisk = $VHDFile"
Write-Host "Executed by: $ScriptUser"
Write-Host "Execution completed at: $DateTime"
Write-Host "Cleanup data file: $CleanupCsv"
Get-ChildItem $DelFiles | Where-Object {$_.LastWriteTime -lt $DelTime} | Remove-Item -Force
Write-Host 'Deleted all Transcripts >30 days old'
Stop-Transcript
#endregion