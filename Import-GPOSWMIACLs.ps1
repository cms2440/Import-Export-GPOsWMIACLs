#This script is meant to import, link, and set wmi filter on a collection of GPOs imported from NIPR
#For now, the path will be manaually set
#I plan to make a file/folder selector

$myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)
$adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator
if (-not $myWindowsPrincipal.IsInRole($adminRole)) {
    $scriptpath = $MyInvocation.MyCommand.Definition
    $scriptpaths = "'$scriptPath'"
    Start-Process -FilePath PowerShell.exe -Verb runAs -ArgumentList "& $scriptPaths"
    exit
    }

#ToDo
#Check for backup name GPO.  if exists, rename
#For this run, delete renamed GPO first before trying to rename backupGPO that exists
#Add label field to GUI to list currently linked OUs

$GPOBackupPath = "C:\Users\chris.steele.adm\Desktop\GPOBackups_2018-10-10_15-39-48"
$GPOsToImport = @()
$GPOsToImport += Get-ChildItem $GPOBackupPath -Directory | select -ExpandProperty fullname 
$WMICSV = Import-Csv "$GPOBackupPath\wmibackups.csv" -Delimiter "`t"
$timestamp = get-date -Format "yyyy/MM/dd - HH:mm:ss"

$toImport = $false #If we want to import our GPOs
$toLink = $true #If we want to link our GPOs

#Stuff needed for WMI
$defaultNamingContext = (get-adrootdse).defaultnamingcontext
$configurationNamingContext = (get-adrootdse).configurationNamingContext
$DestDomain = [System.DirectoryServices.ActiveDirectory.Domain]::getcurrentdomain().name
$msWMIAuthor = "Administrator@" + $DestDomain

#Display a List of the GPOnames we will be making/importing
<#
($GPOsToImport | foreach {
    $GPO = $_
    $GPReportXML = (Get-ChildItem $GPO -Filter "gpreport.xml" | select -ExpandProperty fullname)
    $BackupGPOName = ((Get-Content $GPReportXML) -match "<Name>" | select -First 1).split(">")[1].split("<")[0]
    $GPOName = $BackupGPOName -replace "AFNETOPS","ACC"
    $GPOName
    }) | sort
    #>

#Section to import the gpo settings, and configure WMI filter
if ($toImport -eq $true) {
    #Complicated for testing
    :main for ($i = 0; $i -lt $GPOsToImport.count; $i++) {
        $GPO = $GPOsToImport[$i]
        #Get GPO Name
        $GPReportXML = (Get-ChildItem $GPO -Filter "gpreport.xml" | select -ExpandProperty fullname)
        $BackupGPOName = ((Get-Content $GPReportXML) -match "<Name>" | select -First 1).split(">")[1].split("<")[0]
        $GPOName = $BackupGPOName -replace "AFNETOPS","ACC"

        #Decide which MigTable to use (only matters because of the local account name, but I'm not even sure migtables can change that
        if ($GPOName -match "Windows Server") {$MigrationTable = "C:\Users\chris.steele.adm\Desktop\DSCC.migtable"}
        else {$MigrationTable = "C:\Users\chris.steele.adm\Desktop\SDC.migtable"}

        #Error handling loop to see if the GPO already exists
        do {
            try {
                #If the GPO already exists, skip it
                if ((get-gpo -DisplayName $GPOname -EA STOP) -ne $null) {
                    Write-Host "Skipping `"$GPOName`""
                    #continue main
                    }
                $success = $true
                }
            catch {
                if ($_.FullyQualifiedErrorId -eq "GpoWithNameNotFound,Microsoft.GroupPolicy.Commands.GetGpoCommand") {$success = $true}
                else {$success = $false}
                }
            } until ($success -eq $true)


        #Error handling loop for the creation and importing of GPO settings
        Write-Host "Importing `"$GPOname`""
        $importOut = $false #Exit condition that i dont remember wtf i was doing with.  I wanted to see the error, but only the first time it failed.
        do {
            try {
                Import-GPO -BackupGpoName $BackupGPOName -Path $GPOBackupPath -TargetName $GPOName -MigrationTable $MigrationTable -CreateIfNeeded -EA stop | out-null
                $success = $true
                }
            catch {
                if ($importOut -eq $false) {
                    $_
                    $importOut = $true
                    }
                $success = $false
                }
            } until ($success -eq $true)

        Write-Host "Import Successful"

        #Forced wait just in case replication
        Start-Sleep -Seconds (60*5)

        #Get the WMI info for the current GPO
        $WMIFilter = $WMICSV | Where {$_.GPOname -eq $BackupGPOName}

        #See if we already have a filter in place with this exact query
        $msWMIParm2 = $WMIFilter.Filter
        $WMIFltr = Get-ADObject -Filter 'objectClass -eq "msWMI-Som"' -Properties "msWMI-Name","msWMI-Parm1","msWMI-Parm2" | Where {($_."msWMI-Parm2".split(";")[5..99] -join ";").toLower().replace("'","`"").replace(" =","=").replace("= ","=") -eq ($msWMIParm2.split(";")[5..99] -join ";").ToLower().replace("'","`"").replace(" =","=").replace("= ","=")}
        if ($WMIFltr -eq $null) {
            #I stole most of this from some guy online.  I don't understand most of it.
            $WMIGUID = [string]"{"+([System.Guid]::NewGuid())+"}"
            $WMIDN = "CN="+$WMIGUID+",CN=SOM,CN=WMIPolicy,CN=System,"+$defaultNamingContext
            $WMICN = $WMIGUID
            $WMIdistinguishedname = $WMIDN
            $WMIID = $WMIGUID
 
            $now = (Get-Date).ToUniversalTime()
            $msWMICreationDate = ($now.Year).ToString("0000") + ($now.Month).ToString("00") + ($now.Day).ToString("00") + ($now.Hour).ToString("00") + ($now.Minute).ToString("00") + ($now.Second).ToString("00") + "." + ($now.Millisecond * 1000).ToString("000000") + "-000"
            $msWMIName = $WMIFilter.WMIName

            #If we have a filter of the same name, we need a better name
            if (Get-ADObject -Filter 'objectClass -eq "msWMI-Som"' -properties "msWMI-Name" | Where {$_."msWMI-Name".toLower() -eq $msWMIName.ToLower()}) {$msWMIName += " vAREA52"}
            $seq = 2
            $msWMINameBase = $msWMIName
            #While (Get-ADObject -Filter 'objectClass -eq "msWMI-Som" -and msWMI-Name -eq $msWMIName') {
            While (Get-ADObject -Filter 'objectClass -eq "msWMI-Som"' -properties "msWMI-Name" | Where {$_."msWMI-Name".toLower() -eq $msWMIName.ToLower()}) {
                $msWMIName = $msWMINameBase + "r" + $seq
                $seq++
                }
            $msWMIParm1 = $WMIFilter.Description + " "

            #Now that we have all the settings we need, put them together and put it into AD
            $Attr = @{"msWMI-Name" = $msWMIName;"msWMI-Parm1" = $msWMIParm1;"msWMI-Parm2" = $msWMIParm2;"msWMI-Author" = $msWMIAuthor;"msWMI-ID"=$WMIID;"instanceType" = 4;"showInAdvancedViewOnly" = "TRUE";"distinguishedname" = $WMIdistinguishedname;"msWMI-ChangeDate" = $msWMICreationDate; "msWMI-CreationDate" = $msWMICreationDate}
            $WMIPath = ("CN=SOM,CN=WMIPolicy,CN=System,"+$defaultNamingContext)
 
            New-ADObject -name $WMICN -type "msWMI-Som" -Path $WMIPath -OtherAttributes $Attr
            Start-Sleep -Seconds (60*5)
            $WMIFltr = Get-ADObject -Filter 'objectClass -eq "msWMI-Som" -and msWMI-Name -eq $msWMIName' -Properties "msWMI-Name","msWMI-Parm1","msWMI-Parm2"
            }

        #Error handling loop for setting the GPO to use the WMI filter
        Write-Host ("Setting WMI Filter `"" + $WMIFltr."msWMI-Name" + "`"")
        do {
            try {
                Set-ADObject -Identity (Get-GPO -DisplayName $GPOName).path -Replace @{gPCWQLFilter="[$DestDomain;$($WMIFltr.Name);0]"} -EA Stop
                $success = $true
                }
            catch {$success = $false}
            } until ($success -eq $true)

        Write-Host "WMI Filter set correctly"
        
        #Forced wait just in case replication
        Start-Sleep -Seconds (60*5)

        #Error handling loop for setting the comment on the GPO
        do {
            try {
                (Get-GPO -DisplayName $GPOName).Description = "Migrated from AREA52 on $timestamp"
                $success = $true
                }
            catch {$success = $false}
            } until ($success -eq $true)

        Write-Host "GPO comment set successfully"
        #pause
        "______________________________________________________"
        }
    }

#GUWEEEEEEEEEEEEEEEEE
#------------------------------------------------------------------------
# Source File Information (DO NOT MODIFY)
# Source ID: 0e3c2774-8550-4dd2-a457-8f55790e998f
#------------------------------------------------------------------------

<#
    .NOTES
    --------------------------------------------------------------------------------
        Code generated by:  SAPIEN Technologies, Inc., PowerShell Studio 2015 v4.2.85
        Generated on:       6/16/2015 1:49 PM
        Generated by:        
        Organization:        
    --------------------------------------------------------------------------------
    .DESCRIPTION
        GUI script generated by PowerShell Studio 2015
#>
#----------------------------------------------
#region Import the Assemblies
#----------------------------------------------
[void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
#endregion Import Assemblies

#----------------------------------------------
#region Generated Form Objects
#----------------------------------------------
[System.Windows.Forms.Application]::EnableVisualStyles()
<#
$form1 = New-Object 'System.Windows.Forms.Form'
$nodepath = New-Object 'System.Windows.Forms.TextBox'
$nodename = New-Object 'System.Windows.Forms.TextBox'
$treeview1 = New-Object 'System.Windows.Forms.TreeView'
$buttonOK = New-Object 'System.Windows.Forms.Button'
$buttonCancel = New-Object 'System.Windows.Forms.Button'
$buttonQuit = New-Object 'System.Windows.Forms.Button'
$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
$label = New-Object 'System.Windows.Forms.Label'
#>
#endregion Generated Form Objects

#----------------------------------------------
# User Generated Script
#----------------------------------------------
	
	
function OnApplicationLoad {
	#Note: This function is not called in Projects
	#Note: This function runs before the form is created
	#Note: To get the script directory in the Packager use: Split-Path $hostinvocation.MyCommand.path
	#Note: To get the console output in the Packager (Windows Mode) use: $ConsoleOutput (Type: System.Collections.ArrayList)
	#Important: Form controls cannot be accessed in this function
	#TODO: Add snapins and custom code to validate the application load
		
	return $true #return true for success or false for failure
}
	
	
function OnApplicationExit {
	#Note: This function is not called in Projects
	#Note: This function runs after the form is closed
	#TODO: Add custom code to clean up and unload snapins when the application exits
		
	$script:ExitCode = 0 #Set the exit code for the Packager
}
	
$FormEvent_Load={
	$rootCN=[adsi]''
	$nodeName=$rootCN.Name
	$key="LDAP://$($rootCN.DistinguishedName)"
	$treeview1.Nodes.Add($key,$nodeName)	
}
	
	
$treeview1_NodeMouseClick=[System.Windows.Forms.TreeNodeMouseClickEventHandler]{
	$thisOU=[adsi]$_.Node.Name
	$nodename.Text = $thisOU.Name
	if( -not $_.Node.Nodes){
		$nodepath.Text=$thisOU.DistinguishedName
		$searcher=[adsisearcher]'objectClass=organizationalunit'
		$searcher.SearchRoot=$_.Node.Name
		$searcher.SearchScope='OneLevel'
		$searcher.PropertiesToLoad.Add('name')
		$OUs=$searcher.Findall() 
		<#
        foreach($ou in $OUs){
			$_.Node.Nodes.Add($ou.Path,$ou.Properties['name'])
		}
        #>
        $newOUs = @() 
        $OUs | foreach {
            $newOUs += New-Object PSObject -Property @{
                name = $_.Properties['name']
                Path = $_.Path
                }
            }
        $newOUs = $newOUs | sort name
        foreach($ou in $newOUs){
			$_.Node.Nodes.Add($ou.Path,$ou.name)
		}
	}
	$_.Node.Expand()
}
	
function Get-Ous{
	$sb={
		function Parse-Tree{
			param(
			    $CurrentNode
			)
			Write-Host "$($CurrentNode.DistinguishedName)" -fore green
			Get-ADOrganizationalUnit -Filter * -SearchScope OneLevel -SearchBase $CurrentNode.distinguishedName |
			    ForEach-Object{
			        $node=[pscustomobject]@{        
			            Name=$_.Name
			            DistinguishedName=$_.DistinguishedName
			            Children=@()
			        }
			        $CurrentNode.Children+=$node
			        Parse-Tree -CurrentNode $node
			    }
		}
	
		$root=(Get-AdDomain).DistinguishedName
		$node=[pscustomobject]@{        
			Name='Root'
			DistinguishedName=$root
			Children=@()
		}
		Import-Module ActiveDirectory
		Parse-Tree -CurrentNode $node
		$node
	}
	
}
# --End User Generated Script--
#----------------------------------------------
#region Generated Events
#----------------------------------------------
	
$Form_StateCorrection_Load=
{
	#Correct the initial state of the form to prevent the .Net maximized form issue
	$form1.WindowState = $InitialFormWindowState
}
	
$Form_Cleanup_FormClosed=
{
	#Remove all event handlers from the controls
	try
	{
		$treeview1.remove_NodeMouseClick($treeview1_NodeMouseClick)
		$form1.remove_Load($FormEvent_Load)
		$form1.remove_Load($Form_StateCorrection_Load)
		$form1.remove_FormClosed($Form_Cleanup_FormClosed)
	}
	catch [Exception]
	{ }
}
#endregion Generated Events

#----------------------------------------------
#region Generated Form Code
#----------------------------------------------
<#
$form1.SuspendLayout()
$TreeXIncrease = 100 #I did this for easier editing.  It makes the code harder to read.  Sux for you.
#
# form1
#
$form1.Controls.Add($label)
$form1.Controls.Add($nodepath)
$form1.Controls.Add($nodename)
$form1.Controls.Add($treeview1)
$form1.Controls.Add($buttonOK)
$form1.Controls.Add($buttonCancel)
$form1.AcceptButton = $buttonOK
$form1.CancelButton = $buttonCancel
$form1.FormBorderStyle = 'FixedDialog'
$form1.MaximizeBox = $False
$form1.MinimizeBox = $False
$form1.Name = "form1"
$form1.ClientSize = "" + (560 + $TreeXIncrease) + ", 389" #'622, 389'
$form1.StartPosition = 'CenterScreen'
$form1.Text = "Form"
#$form1.add_Load($FormEvent_Load)
&$FormEvent_Load | Out-Null
#
# label
#
$Label.Location = New-Object System.Drawing.Point(10,10) 
#
# nodepath
#
$nodepath.Location = "" + (225 + $TreeXIncrease) + ", 53" #'287, 38' # "" + (225 + $TreeXIncrease) + ", 53"
$nodepath.Name = "nodepath"
$nodepath.Size = '323, 20'
$nodepath.TabIndex = 3
$nodepath.ReadOnly = $true
#
# nodename
#
$nodename.Location = "" + (225 + $TreeXIncrease) + ", 27" #'287, 12' # "" + (225 + $TreeXIncrease) + ", 27"
$nodename.Name = "nodename"
$nodename.Size = '323, 20'
$nodename.TabIndex = 2
$nodename.ReadOnly = $true
#
# treeview1
#
$treeview1.Location = '12, 27'
$treeview1.Name = "treeview1"
$treeview1.Size = "" + (172 + $TreeXIncrease) + ", 334" #'172, 334'
$treeview1.TabIndex = 1
$treeview1.add_NodeMouseClick($treeview1_NodeMouseClick)
#
# buttonOK
#
$buttonOK.Anchor = 'Bottom, Right'
$buttonOK.DialogResult = 'OK'
$buttonOK.Location = '473, 354'
$buttonOK.Name = "buttonOK"
$buttonOK.Size = '75, 23'
$buttonOK.TabIndex = 0
$buttonOK.Text = "OK"
$buttonOK.UseVisualStyleBackColor = $True
#
# buttonOK
#
$buttonCancel.Anchor = 'Bottom, Right'
$buttonCancel.DialogResult = 'Cancel'
$buttonCancel.Location = '573, 354'
$buttonCancel.Name = "buttonCancel"
$buttonCancel.Size = '75, 23'
$buttonCancel.TabIndex = 0
$buttonCancel.Text = "Cancel"
$buttonCancel.UseVisualStyleBackColor = $True

$form1.ResumeLayout()
#endregion Generated Form Code

#----------------------------------------------

#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($Form_StateCorrection_Load)
#Clean up the control events
#$form1.add_FormClosed($Form_Cleanup_FormClosed)
#Show the Form
#return $form1.ShowDialog()
#>

if ($toLink -eq $true) {
    :main for ($i = 0; $i -lt $GPOsToImport.count; $i++) {
        #Be able to loop through multiple time for one GPO, so we can link it to multiple OUs
        #The loop ends when you hit Cancel #(notanymore)or hit OK without selecting an OU
        do {
            #Get GPO Name
            $GPO = $GPOsToImport[$i]
            $GPReportXML = (Get-ChildItem $GPO -Filter "gpreport.xml" | select -ExpandProperty fullname)
            $BackupGPOName = ((Get-Content $GPReportXML) -match "<Name>" | select -First 1).split(">")[1].split("<")[0]
            $GPOName = $BackupGPOName -replace "AFNETOPS","ACC"

            #Gui initialization
            #I hate how I have to redefine everythign each time.  If I don't, for some reason nodepath doesnt populate in future loops after you select an OU and hit OK.
            $form1 = New-Object 'System.Windows.Forms.Form'
            $nodepath = New-Object 'System.Windows.Forms.TextBox'
            $nodename = New-Object 'System.Windows.Forms.TextBox'
            $treeview1 = New-Object 'System.Windows.Forms.TreeView'
            $buttonOK = New-Object 'System.Windows.Forms.Button'
            $buttonCancel = New-Object 'System.Windows.Forms.Button'
            $buttonQuit = New-Object 'System.Windows.Forms.Button'
            $InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'
            $label = New-Object 'System.Windows.Forms.Label'

            $form1.SuspendLayout()
            $TreeXIncrease = 100 #I did this for easier editing.  It makes the code harder to read.  Sux for you.
            #
            # form1
            #
            $form1.Controls.Add($label)
            $form1.Controls.Add($nodepath)
            $form1.Controls.Add($nodename)
            $form1.Controls.Add($treeview1)
            $form1.Controls.Add($buttonOK)
            $form1.Controls.Add($buttonCancel)
            $form1.Controls.Add($buttonQuit)
            $form1.AcceptButton = $buttonOK
            $form1.CancelButton = $buttonCancel
            $form1.FormBorderStyle = 'FixedDialog'
            $form1.MaximizeBox = $False
            $form1.MinimizeBox = $False
            $form1.Name = "form1"
            $form1.ClientSize = "" + (560 + $TreeXIncrease) + ", 389" #'622, 389'
            $form1.StartPosition = 'CenterScreen'
            $form1.Text = "Form"
            $form1.add_Load($FormEvent_Load)
            #
            # label
            #
            $Label.Location = New-Object System.Drawing.Point(10,10) 
            $Label.Text = "Select OU to link GPO `"" + $GPOName + "`""
            $Label.Size = New-Object System.Drawing.Size(($Label.Text.length * 7),15) #280,15
            #
            # nodepath
            #
            $nodepath.Location = "" + (225 + $TreeXIncrease) + ", 53" #'287, 38' # "" + (225 + $TreeXIncrease) + ", 53"
            $nodepath.Name = "nodepath"
            $nodepath.Size = '323, 20'
            $nodepath.TabIndex = 3
            $nodepath.ReadOnly = $true
            #
            # nodename
            #
            $nodename.Location = "" + (225 + $TreeXIncrease) + ", 27" #'287, 12' # "" + (225 + $TreeXIncrease) + ", 27"
            $nodename.Name = "nodename"
            $nodename.Size = '323, 20'
            $nodename.TabIndex = 2
            $nodename.ReadOnly = $true
            #
            # treeview1
            #
            $treeview1.Location = '12, 27'
            $treeview1.Name = "treeview1"
            $treeview1.Size = "" + (172 + $TreeXIncrease) + ", 334" #'172, 334'
            $treeview1.TabIndex = 1
            $treeview1.add_NodeMouseClick($treeview1_NodeMouseClick)
            #
            # buttonOK
            #
            $buttonOK.Anchor = 'Bottom, Right'
            $buttonOK.DialogResult = 'OK'
            $buttonOK.Location = '373, 354'
            $buttonOK.Name = "buttonOK"
            $buttonOK.Size = '75, 23'
            $buttonOK.TabIndex = 0
            $buttonOK.Text = "OK"
            $buttonOK.UseVisualStyleBackColor = $True
            #
            # buttonCancel
            #
            $buttonCancel.Anchor = 'Bottom, Right'
            $buttonCancel.DialogResult = 'Cancel'
            $buttonCancel.Location = '473, 354'
            $buttonCancel.Name = "buttonCancel"
            $buttonCancel.Size = '75, 23'
            $buttonCancel.TabIndex = 0
            $buttonCancel.Text = "Next GPO"
            $buttonCancel.UseVisualStyleBackColor = $True
            #
            # buttonQuit
            #
            $buttonQuit.Anchor = 'Bottom, Right'
            $buttonQuit.DialogResult = 'Abort'
            $buttonQuit.Location = '573, 354'
            $buttonQuit.Name = "buttonQuit"
            $buttonQuit.Size = '75, 23'
            $buttonQuit.TabIndex = 0
            $buttonQuit.Text = "Quit"
            $buttonQuit.UseVisualStyleBackColor = $True

            $form1.ResumeLayout()
            #endregion Generated Form Code

            #----------------------------------------------

            #Save the initial state of the form
            $InitialFormWindowState = $form1.WindowState
            #Init the OnLoad event to correct the initial state of the form
            $form1.add_Load($Form_StateCorrection_Load)
            #Clean up the control events
            $form1.add_FormClosed($Form_Cleanup_FormClosed)
            
            #Artifact from when I tried to not remake the GUI each loop
            #$Label.Text = "Select OU to link GPO `"" + $GPOName + "`""
            #$Label.Size = New-Object System.Drawing.Size(($Label.Text.length) * 7,15) #280,15
            #$nodepath.Text = ""
            #$nodename.Text = ""

            #See if we actually want to link this GPO to a(nother) OU
            $UserOption = $form1.ShowDialog()
            #If the user hits the Quit button, exit the script
            if ($UserOption -eq "Abort") {exit}
	        if ($UserOption -eq "OK" -and $nodepath.Lines -ne "" -and $nodepath.Lines -ne $null) {
                $ou = $nodepath.Lines | select -First 1
                Write-Host ("Linking `"$GPOName`" to " + $ou)
                New-GPLink -Domain "acc.accroot.ds.af.smil.mil" -Name $GPOName -Target $ou -LinkEnabled No -Enforced No | Out-Null
                }
            } until ($UserOption -eq "Cancel")
        }
    }
#not needed since the script will exit anyway, but meh.  Useful for ISE I guess.
&$Form_Cleanup_FormClosed | Out-Null

#Display where each GPO is linked.
:main for ($i = 0; $i -lt $GPOsToImport.count; $i++) {
    #Get the GPOName
    $GPO = $GPOsToImport[$i]
    $GPReportXML = (Get-ChildItem $GPO -Filter "gpreport.xml" | select -ExpandProperty fullname)
    $BackupGPOName = ((Get-Content $GPReportXML) -match "<Name>" | select -First 1).split(">")[1].split("<")[0]
    $GPOName = $BackupGPOName -replace "AFNETOPS","ACC"

    Write-Host -f Green "`"$GPOName`" is linked to the following OUs:"
    #Get to the part in the report for where the GPO is linked
    $report = (Get-GPOReport -Name $GPOName -ReportType html).split("`n")
    :find for ($j = 0; $j -lt $report.count; $j++) {
        if ($report[$j] -like '*<th scope="col">Link Status</th>*') {break find}
        }
    $j++
    #if it's not linked anywhere, we have to handle that as a special case
    if ($report[$j].trim() -eq '<tr><td colspan="4">None</td></tr>') {
        Write-Host "I did a little sneaky on ya.  It's not linked anywhere"
        }
    #List each OU that the GPO is linked to
    else {
        for (;$report[$j] -notlike "*</table>*";$j++) {
            $report[$j].trim().split("<")[8].split(">")[1]
            }
        }
    Write-Host #Newline
    }

read-host "Press Enter to close the script."
