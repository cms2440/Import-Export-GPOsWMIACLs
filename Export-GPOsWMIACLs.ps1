#List of GPOs to backup
$STIGs = @"

"@.Split("`n") | foreach {$_.trim()}

#To be decided how it will be handled
#Im thinking look for GPOS that match the STIG category, parse for the latest version and revision, and then grab that GPO name
$GPOsToExport = @()

#Since I did the work manually to export the GPOs, ill just work off that
$GPOsToExport = Get-ChildItem "C:\Users\1456084571E\Desktop\GPOs" | select -ExpandProperty fullname | foreach {
    ((Get-Content (Get-ChildItem $_ -Filter "gpreport.xml" | select -ExpandProperty fullname)) -match "  <Name>" | select -First 1).split(">")[1].split("<")[0].trim()
    }

#Set up our variables and backup folder to save the GPOs to
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$BackupFolderPath = "$env:USERPROFILE\desktop\GPOBackups_$timestamp"
New-Item -ItemType Directory -Path $BackupFolderPath -EA SilentlyContinue | Out-Null
$WMIBackupPath = "$BackupFolderPath\WMIBackups.csv"

#Start our csv file to export info to
Out-File $WMIBackupPath -Encoding default -InputObject ("GPOName`tWMIName`tDescription`tFilter")

foreach ($GPOname in $GPOsToExport) {
    $GPO = Get-GPO -DisplayName $GPOname
    $WMIFilterName = $GPO.WmiFilter.Name

    #Backup-GPO -Name $GPOname -Path $BackupFolderPath -Domain "area52.afnoapps.usaf.mil" | out-null
    
    $WMIFilter = Get-ADObject -Filter 'objectClass -eq "msWMI-Som" -and msWMI-Name -eq $WMIFilterName' -Properties "msWMI-Name","msWMI-Parm1","msWMI-Parm2"

    $NewContent = $GPOname + "`t" + $WMIFilter."msWMI-Name" + "`t" + $WMIFilter."msWMI-Parm1" + "`t" + $WMIFilter."msWMI-Parm2"
    Add-Content $NewContent -Path $WMIBackupPath
    }

Add-Type -Assembly "System.IO.Compression.FileSystem"
[System.IO.Compression.ZipFile]::CreateFromDirectory($BackupFolderPath, "$env:USERPROFILE\desktop\GPOBackups_$timestamp.zip")

<#
#How to import WMI Filters from our above export
param([String]$BackupFile)

if ([String]::IsNullOrEmpty($BackupFile))
{
    write-host -ForeGroundColor Red "BackupFile is a required parameter. Exiting Script.`n"
    return
}

import-module ActiveDirectory

#$Header = "Name","Description","Filter"
$WMIFilters = import-csv $BackupFile -Delimiter "`t"# -Header $Header

#The code below this line was modified to work with the file import code above this line. The original code can be found here.

$defaultNamingContext = (get-adrootdse).defaultnamingcontext
$configurationNamingContext = (get-adrootdse).configurationNamingContext
$msWMIAuthor = "Administrator@" + [System.DirectoryServices.ActiveDirectory.Domain]::getcurrentdomain().name

foreach ($WMIFilter in $WMIFilters) {
    $WMIGUID = [string]"{"+([System.Guid]::NewGuid())+"}"
    $WMIDN = "CN="+$WMIGUID+",CN=SOM,CN=WMIPolicy,CN=System,"+$defaultNamingContext
    $WMICN = $WMIGUID
    $WMIdistinguishedname = $WMIDN
    $WMIID = $WMIGUID
 
    $now = (Get-Date).ToUniversalTime()
    $msWMICreationDate = ($now.Year).ToString("0000") + ($now.Month).ToString("00") + ($now.Day).ToString("00") + ($now.Hour).ToString("00") + ($now.Minute).ToString("00") + ($now.Second).ToString("00") + "." + ($now.Millisecond * 1000).ToString("000000") + "-000"
    $msWMIName = $WMIFilter.WMIName
    $msWMIParm1 = $WMIFilter.Description + " "
    $msWMIParm2 = $WMIFilter.Filter

    $Attr = @{"msWMI-Name" = $msWMIName;"msWMI-Parm1" = $msWMIParm1;"msWMI-Parm2" = $msWMIParm2;"msWMI-Author" = $msWMIAuthor;"msWMI-ID"=$WMIID;"instanceType" = 4;"showInAdvancedViewOnly" = "TRUE";"distinguishedname" = $WMIdistinguishedname;"msWMI-ChangeDate" = $msWMICreationDate; "msWMI-CreationDate" = $msWMICreationDate}
    $WMIPath = ("CN=SOM,CN=WMIPolicy,CN=System,"+$defaultNamingContext)
 
    New-ADObject -name $WMICN -type "msWMI-Som" -Path $WMIPath -OtherAttributes $Attr
    }
#>