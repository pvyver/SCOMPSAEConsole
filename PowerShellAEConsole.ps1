################################################
# Name: AE_ConfigurationTool.ps1
# Author: Scott Murray (PFE) & Philip Van de Vyver (PFE)
#
# Use: From command prompt: powershell .\AE_ConfigurationTool.ps1
# Config: Upate Configuration in XML AE_ConfigurationTool.config 
#
# Purpose: Script connects to designated SCOM environment 2012/2016 code
# per the configuration section.  Once connected, script retreives list of 
# management packs and then lists out the display names for the user to select
# the MP they want to add to the enrichment DB.  Once an MP is selected, the
# script connects to the enrichment DB and then displays all rules and 
# monitors in the MP along with the enrichment info (if found).  The user is 
# then given the option to change the enrichment information and then commit the
# changes back to the DB.
################################################

# Change Log
# 12282012 - updated code to work on PS v3.

# Global variable declarations.  Bad programming practice?  Maybe.
# Easier? Yup.
$global:SQLServer 
$global:SQLDBName
$global:SQLQuery
$global:rms
$global:mpname
$global:rulesmonis = @()
$global:enrichment = @()
$global:GroupEnrichment = @()
$global:CustomFields = @()
$global:boolChanges = $false
$global:scom_environment
$global:default_appid
$global:default_resstate
$global:default_owner
$global:default_ticket

###########################
###   Read AE_ConfigurationTool.config XML ###
###   Configuration  ######
[Xml]$configxml = Get-Content E:\Temp\SCORCH\AE_ConfigurationTool.config
$rms = $configxml.configuration.RootManagementServerEmulator.Name
$scom_environment = $configxml.configuration.scom_environment.version 
$SQLServer = $configxml.configuration.SQLServer.Name
$SQLDBName = $configxml.configuration.SQLDBName.Name 
$SQLEnrichTableName = $configxml.configuration.SQLEnrichTable.Name 
$SQLQuery = $configxml.configuration.SQLQuery.Query 
$default_appid = $configxml.configuration.default_appid.AppID
$default_resstate = $configxml.configuration.default_resstate.ResState 
$global:default_ticket = $configxml.configuration.default_ticket.Ticket 
###   NOTE: there may be changes necessary in the ChangeEnrichment function ###
###   End Configuration  ###
############################

# Load .Net namespaces for access form controls
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

# This is the main driver of the script.  It calls all of the necessary functions to walk
# the user through selecting an MP and then setting enrichment information.
function main{

	$global:mpname = $null
	MGConnect
	DisplayRulesMonisEnrichment

	if($global:boolChanges -eq $true)
	{
		ChangeEnrichment
	}

}

function refreshView{
	GetRulesMonitors
    $global:enrichment	= getSQLData($SQLQuery)
    MergeData
    $tempcount = 1
	foreach($rm in $global:rulesmonis)
	{
		$objRuleMoniBox.Items.Add('' + $tempcount + ': ' + $rm.displayname)
		$tempcount++
	}
}

#Bulk Insert Function
    function bulkUpdate ($items)
{
            $bulkUpdateForm= New-Object System.Windows.Forms.Form 
            $bulkUpdateForm.Text = "Update Custom Field for " + $items.Count + " items?"
            $bulkUpdateForm.Size = New-Object System.Drawing.Size(500,500) 
            $bulkUpdateForm.StartPosition = "CenterScreen"

            $OKButton = New-Object System.Windows.Forms.Button
            $OKButton.Location = New-Object System.Drawing.Point(150,420)
            $OKButton.Size = New-Object System.Drawing.Size(75,23)
            $OKButton.Text = "OK"
            $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $bulkUpdateForm.AcceptButton = $OKButton
            $bulkUpdateForm.Controls.Add($OKButton)

            $CancelButton = New-Object System.Windows.Forms.Button
            $CancelButton.Location = New-Object System.Drawing.Point(250,420)
            $CancelButton.Size = New-Object System.Drawing.Size(75,23)
            $CancelButton.Text = "Cancel"
            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $bulkUpdateForm.CancelButton = $CancelButton
            $bulkUpdateForm.Controls.Add($CancelButton)

	        $objGroupLabel = New-Object System.Windows.Forms.Label
	        $objGroupLabel.Location = New-Object System.Drawing.Size(5,10)
	        $objGroupLabel.Size = New-Object System.Drawing.Size(45,20)
	        $objGroupLabel.Text = "Group"
	        $bulkUpdateForm.Controls.Add($objGroupLabel)

	        $objGroupBox = New-Object System.Windows.Forms.ComboBox
	        $objGroupBox.Location = New-Object System.Drawing.Point(50,10)
	        $objGroupBox.Size = New-Object System.Drawing.Size(400,20)
            $bulkUpdateForm.Controls.Add($objGroupBox)
            $scomGroups = Get-SCOMGroup | select DisplayName | sort-Object DisplayName
            $objGroupBox.Items.AddRange($scomGroups.DisplayName)


            #CF
	        $offset = 33
	        $start = 70
	        $objCF1Label = New-Object System.Windows.Forms.Label
	        $objCF1Label.Location = New-Object System.Drawing.Size(100,$start)
	        $objCF1Label.Size = New-Object System.Drawing.Size(30,20)
	        $objCF1Label.Text = "CF1"
	        $bulkUpdateForm.Controls.Add($objCF1Label)	

	        $objCF1Text = New-Object System.Windows.Forms.TextBox
	        $objCF1Text.Location = New-Object System.Drawing.Size(140,$start)
	        $objCF1Text.Size = New-Object System.Drawing.Size(225,20)
	        $objCF1Text.Text = ""
	        $objCF1Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	        $bulkUpdateForm.Controls.Add($objCF1Text)

            $start+=$offset
	        $objCF2Label = New-Object System.Windows.Forms.Label
	        $objCF2Label.Location = New-Object System.Drawing.Size(100,$start)
	        $objCF2Label.Size = New-Object System.Drawing.Size(30,20)
	        $objCF2Label.Text = "CF2"
	        $bulkUpdateForm.Controls.Add($objCF2Label)	

	        $objCF2Text = New-Object System.Windows.Forms.TextBox
	        $objCF2Text.Location = New-Object System.Drawing.Size(140,$start)
	        $objCF2Text.Size = New-Object System.Drawing.Size(225,20)
	        $objCF2Text.Text = ""
	        $objCF2Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	        $bulkUpdateForm.Controls.Add($objCF2Text)

            $start+=$offset
	        $objCF3Label = New-Object System.Windows.Forms.Label
	        $objCF3Label.Location = New-Object System.Drawing.Size(100,$start)
	        $objCF3Label.Size = New-Object System.Drawing.Size(30,20)
	        $objCF3Label.Text = "CF3"
	        $bulkUpdateForm.Controls.Add($objCF3Label)	

	        $objCF3Text = New-Object System.Windows.Forms.TextBox
	        $objCF3Text.Location = New-Object System.Drawing.Size(140,$start)
	        $objCF3Text.Size = New-Object System.Drawing.Size(225,20)
	        $objCF3Text.Text = ""
	        $objCF3Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	        $bulkUpdateForm.Controls.Add($objCF3Text)

            $start+=$offset
	        $objCF4Label = New-Object System.Windows.Forms.Label
	        $objCF4Label.Location = New-Object System.Drawing.Size(100,$start)
	        $objCF4Label.Size = New-Object System.Drawing.Size(30,20)
	        $objCF4Label.Text = "CF4"
	        $bulkUpdateForm.Controls.Add($objCF4Label)	

	        $objCF4Text = New-Object System.Windows.Forms.TextBox
	        $objCF4Text.Location = New-Object System.Drawing.Size(140,$start)
	        $objCF4Text.Size = New-Object System.Drawing.Size(225,20)
	        $objCF4Text.Text = ""
	        $objCF4Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	        $bulkUpdateForm.Controls.Add($objCF4Text)

            $start+=$offset
	        $objCF5Label = New-Object System.Windows.Forms.Label
	        $objCF5Label.Location = New-Object System.Drawing.Size(100,$start)
	        $objCF5Label.Size = New-Object System.Drawing.Size(30,20)
	        $objCF5Label.Text = "CF5"
	        $bulkUpdateForm.Controls.Add($objCF5Label)	

	        $objCF5Text = New-Object System.Windows.Forms.TextBox
	        $objCF5Text.Location = New-Object System.Drawing.Size(140,$start)
	        $objCF5Text.Size = New-Object System.Drawing.Size(225,20)
	        $objCF5Text.Text = ""
	        $objCF5Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	        $bulkUpdateForm.Controls.Add($objCF5Text)

            $start+=$offset
	        $objCF6Label = New-Object System.Windows.Forms.Label
	        $objCF6Label.Location = New-Object System.Drawing.Size(100,$start)
	        $objCF6Label.Size = New-Object System.Drawing.Size(30,20)
	        $objCF6Label.Text = "CF6"
	        $bulkUpdateForm.Controls.Add($objCF6Label)	

	        $objCF6Text = New-Object System.Windows.Forms.TextBox
	        $objCF6Text.Location = New-Object System.Drawing.Size(140,$start)
	        $objCF6Text.Size = New-Object System.Drawing.Size(225,20)
	        $objCF6Text.Text = ""
	        $objCF6Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	        $bulkUpdateForm.Controls.Add($objCF6Text)

            $start+=$offset
	        $objCF7Label = New-Object System.Windows.Forms.Label
	        $objCF7Label.Location = New-Object System.Drawing.Size(100,$start)
	        $objCF7Label.Size = New-Object System.Drawing.Size(30,20)
	        $objCF7Label.Text = "CF7"
	        $bulkUpdateForm.Controls.Add($objCF7Label)	

	        $objCF7Text = New-Object System.Windows.Forms.TextBox
	        $objCF7Text.Location = New-Object System.Drawing.Size(140,$start)
	        $objCF7Text.Size = New-Object System.Drawing.Size(225,20)
	        $objCF7Text.Text = ""
	        $objCF7Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	        $bulkUpdateForm.Controls.Add($objCF7Text)

            $start+=$offset
	        $objCF8Label = New-Object System.Windows.Forms.Label
	        $objCF8Label.Location = New-Object System.Drawing.Size(100,$start)
	        $objCF8Label.Size = New-Object System.Drawing.Size(30,20)
	        $objCF8Label.Text = "CF8"
	        $bulkUpdateForm.Controls.Add($objCF8Label)	

	        $objCF8Text = New-Object System.Windows.Forms.TextBox
	        $objCF8Text.Location = New-Object System.Drawing.Size(140,$start)
	        $objCF8Text.Size = New-Object System.Drawing.Size(225,20)
	        $objCF8Text.Text = ""
	        $objCF8Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	        $bulkUpdateForm.Controls.Add($objCF8Text)

            $start+=$offset
	        $objCF9Label = New-Object System.Windows.Forms.Label
	        $objCF9Label.Location = New-Object System.Drawing.Size(100,$start)
	        $objCF9Label.Size = New-Object System.Drawing.Size(30,20)
	        $objCF9Label.Text = "CF9"
	        $bulkUpdateForm.Controls.Add($objCF9Label)	

	        $objCF9Text = New-Object System.Windows.Forms.TextBox
	        $objCF9Text.Location = New-Object System.Drawing.Size(140,$start)
	        $objCF9Text.Size = New-Object System.Drawing.Size(225,20)
	        $objCF9Text.Text = ""
	        $objCF9Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	        $bulkUpdateForm.Controls.Add($objCF9Text)

            $start+=$offset
	        $objCF10Label = New-Object System.Windows.Forms.Label
	        $objCF10Label.Location = New-Object System.Drawing.Size(100,$start)
	        $objCF10Label.Size = New-Object System.Drawing.Size(35,20)
	        $objCF10Label.Text = "CF10"
	        $bulkUpdateForm.Controls.Add($objCF10Label)	

	        $objCF10Text = New-Object System.Windows.Forms.TextBox
	        $objCF10Text.Location = New-Object System.Drawing.Size(140,$start)
	        $objCF10Text.Size = New-Object System.Drawing.Size(225,20)
	        $objCF10Text.Text = ""
	        $objCF10Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	        $bulkUpdateForm.Controls.Add($objCF10Text)


            $bulkUpdateForm.Topmost = $True
            $result = $bulkUpdateForm.ShowDialog()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK)
            {
                $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
		        $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True"
		        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand	
		        $SqlCmd.Connection = $SqlConnection
		        $SqlConnection.Open()
                    foreach ($item in $items)
                        {
                            $data = $global:rulesmonis[$Item.split(":")[0] -1];
                            $data.CF1 = $objCF1Text.Text;
						    $data.CF2 = $objCF2Text.Text;
						    $data.CF3 = $objCF3Text.Text;
						    $data.CF4 = $objCF4Text.Text;
						    $data.CF5 = $objCF5Text.Text;
						    $data.CF6 = $objCF6Text.Text;
						    $data.CF7 = $objCF7Text.Text;
						    $data.CF8 = $objCF8Text.Text;
						    $data.CF9 = $objCF9Text.Text;
						    $data.CF10 = $objCF10Text.Text;
                            $data.SupportGroup = $objGroupBox.Text
                    
                    $existingRecord = getSQLData("select count(*) as GroupRecords from " + $SQLEnrichTableName + " where SupportGroup = '" + $data.SupportGroup + "' and MonitorRuleID = '" + $data.id + "'"  )
                    if ($existingRecord[1].GroupRecords -eq 0)
                    {
                        $newQuery = "Insert Into " + $SQLEnrichTableName + " (DisplayName,MonitorRuleID,ManagementPack,Type,Ticket,APPID,Owner,AlertNotes,Notes,ResolutionState,CF1,CF2,CF3,CF4,CF5,CF6,CF7,CF8,CF9,CF10,SupportGroup) values ('" + $data.DisplayName + "','" + $data.ID + "','" + $global:mpname + "','" + $data.RuleMoniType + "','" + $data.Ticket + "','" + $data.APPID + "','" + $data.Owner + "','" + $data.AlertNotes.trim() + "','" + $data.Notes.trim() + "','" + $data.ResState.trim() + "','" + $data.CF1 + "','" + $data.CF2 + "','" + $data.CF3 + "','" + $data.CF4 + "','" + $data.CF5 + "','" + $data.CF6 + "','" + $data.CF7 + "','" + $data.CF8 + "','" + $data.CF9 + "','" + $data.CF10 + "','" + $data.SupportGroup + "')"		
			            #write-host $newQuery
			            $SqlCmd.CommandText = $newQuery
			            $SqlCmd.ExecuteNonQuery()   
                    }
                    else
                    {
                        [Windows.Forms.MessageBox]::Show($data.SupportGroup  + " already added to group list of ” + $data.RuleMoniType + ": '" + $data.DisplayName + "'" ,"Add group error", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)

                    }
                    
                    }
                    $SqlConnection.Close()	
                                     
            }  
    }
    


# connects to the specified MG.
function MGConnect
{
	import-module operationsmanager
	new-scommanagementgroupconnection $rms
}


# This function retrieves the mp list from the management group and then pops 
# open a diaglog box having the user select which MP they want to set enrichment
# information for.  Only one MP can be worked with at a time.
function PickMP
{
	#get the MPs
    $managementpacks = get-scommanagementpack
################### Set up form and controls #################
	$objForm = New-Object System.Windows.Forms.Form 
	$objForm.Text = "Select Management Pack"
	$objForm.Size = New-Object System.Drawing.Size(500,200) 
	$objForm.StartPosition = "CenterScreen"

	$objForm.KeyPreview = $True
	# snags the MP name from the display and sets it to $x
	$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
		{$x=$objMPBox.Text;$objForm.Close()}})
	$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
		{$objForm.Close()}})

	$OKButton = New-Object System.Windows.Forms.Button
	$OKButton.Location = New-Object System.Drawing.Size(75,120)
	$OKButton.Size = New-Object System.Drawing.Size(75,23)
	$OKButton.Text = "OK"
	$OKButton.Add_Click({$x=$objMPBox.Text;$objForm.Close()})
	$objForm.Controls.Add($OKButton)

	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Size(150,120)
	$CancelButton.Size = New-Object System.Drawing.Size(75,23)
	$CancelButton.Text = "Cancel"
	$CancelButton.Add_Click({$objForm.Close()})
	$objForm.Controls.Add($CancelButton)

	$objLabel = New-Object System.Windows.Forms.Label
	$objLabel.Location = New-Object System.Drawing.Size(10,20) 
	$objLabel.Size = New-Object System.Drawing.Size(280,20) 
	$objLabel.Text = "Select management pack for enrichment:"
	$objForm.Controls.Add($objLabel) 

	# The combo box shows the displayname for each MP, if empty choose Name
	$objMPBox = New-Object System.Windows.Forms.ComboBox
	$objMPBox.Location = New-Object System.Drawing.Point(70,90)
	$objMPBox.Size = New-Object System.Drawing.Size(350,50)
	$managementpacks = $managementpacks | sort displayname
	foreach($mp in $managementpacks)
	{
		if(($mp.displayname -ne $null) -AND ($mp.displayname -ne ''))
		{
			$objMPBox.Items.Add($mp.displayname)
		}
	}	
	$objForm.Controls.Add($objMPBox)
	
	$objForm.Topmost = $True
	$objForm.Add_Shown({$objForm.Activate()})
	[void] $objForm.ShowDialog()
	# stores the name for the MP (not displayname) in the global var
	$global:mpname = $objMPBox.Text
}

# This function gets all of the rules and monitors associated with the MP directly
# from the specified OM environment.  It stores it in a collection of custom objects.
function GetRulesMonitors
{
		$mp = get-scommanagementpack -DisplayName $global:mpname 
		$rules = get-scomrule -managementpack $mp
		$monis = get-scommonitor -managementpack $mp

$global:rulesmonis = @()

	# Iterates through the rules and loads them into the custom object collection
	if($rules -ne $null){
		$rules | foreach{
			$myObject = New-Object PSObject
			$myObject | add-member NoteProperty Name $_.name
			$myObject | add-member NoteProperty DisplayName $_.displayname
			$myObject | add-member NoteProperty ID $_.id.tostring()
			$myObject | add-member NoteProperty Ticket ''
			$myObject | add-member NoteProperty APPID ''
			$myObject | add-member NoteProperty Owner ''
			$myObject | add-member NoteProperty Notes ''
			$myObject | add-member NoteProperty AlertNotes ''
			$myObject | add-member NoteProperty RuleMoniType 'Rule'
			$myObject | add-member NoteProperty HasChanged 'false'
			$myObject | add-member NoteProperty PrevEnrich 'false'
			$myObject | add-member NoteProperty ResState '0'
			$myObject | add-member NoteProperty CF1 ''
			$myObject | add-member NoteProperty CF2 ''
			$myObject | add-member NoteProperty CF3 ''
			$myObject | add-member NoteProperty CF4 ''
			$myObject | add-member NoteProperty CF5 ''
			$myObject | add-member NoteProperty CF6 ''
			$myObject | add-member NoteProperty CF7 ''
			$myObject | add-member NoteProperty CF8 ''
			$myObject | add-member NoteProperty CF9 ''
			$myObject | add-member NoteProperty CF10 ''
            #$supportgroup = @()
            #$myObject | Add-Member -MemberType NoteProperty -Name ArrayList -value $supportgroup
            $myObject | add-member NoteProperty SupportGroup ''
            $a.ArrayList
			$global:rulesmonis += $myObject
		}
	}
	
	# Iterates through the monitors and loads them into the custom object collection
	if($monis -ne $null) {
		$monis | foreach{
			$myObject = New-Object PSObject
			$myObject | add-member NoteProperty Name $_.name
			$myObject | add-member NoteProperty DisplayName $_.displayname
			$myObject | add-member NoteProperty ID $_.id.tostring()
			$myObject | add-member NoteProperty Ticket ''
			$myObject | add-member NoteProperty APPID ''
			$myObject | add-member NoteProperty Owner ''
			$myObject | add-member NoteProperty Notes ''
			$myObject | add-member NoteProperty AlertNotes ''
			$myObject | add-member NoteProperty RuleMoniType $_.XMLTag
			$myObject | add-member NoteProperty HasChanged 'false'
			$myObject | add-member NoteProperty PrevEnrich 'false'
			$myObject | add-member NoteProperty ResState '0'
			$myObject | add-member NoteProperty CF1 ''
			$myObject | add-member NoteProperty CF2 ''
			$myObject | add-member NoteProperty CF3 ''
			$myObject | add-member NoteProperty CF4 ''
			$myObject | add-member NoteProperty CF5 ''
			$myObject | add-member NoteProperty CF6 ''
			$myObject | add-member NoteProperty CF7 ''
			$myObject | add-member NoteProperty CF8 ''
			$myObject | add-member NoteProperty CF9 ''
			$myObject | add-member NoteProperty CF10 ''
            #$supportgroup = @()
            #$myObject | Add-Member -MemberType NoteProperty SupportGroup -value $supportgroup
            $myObject | add-member NoteProperty SupportGroup ''
			$global:rulesmonis += $myObject
		}
	}
	
	# Sorts the custom object collection containing the rules and monitors
	# to make it more consumable in the GUI.
	$global:rulesmonis = $global:rulesmonis | sort DisplayName
}

# This function connects to the enrichment database and Runs any QRY
Function getSQLData ($QRY)
{
	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True"
	$SqlConnection.Open()
	
    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
	$SqlCmd.CommandText = $QRY
	$SqlCmd.Connection = $SqlConnection

	$SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
	$SqlAdapter.SelectCommand = $SqlCmd

	$DataSet = New-Object System.Data.DataSet
	$SqlAdapter.Fill($DataSet)
	 
	$SqlConnection.Close()
	
    #return dataset
    return , $DataSet.Tables[0].rows
}

# This is the function that displays all of the rules and monitors from the selected MP along with
# all of the associated enrichment information.  This function loads the main GUI for making changes
# to the enrichment information for each of the different rules and monitors.  The use clicks the \
# edit button for each rule or monitor, makes changes, and then clicks save.  This commits teh changes
# back to the main custom object collection.  Once the user is done making changes, they click the
# commit button.  The commit button executes the logic to construct the SQL update and insert state-
# ments necessary to store the enrichment data in the database.
function DisplayRulesMonisEnrichment
{
#################### form logic ###########################

	$objForm = New-Object System.Windows.Forms.Form 
	$objForm.Text = "Alert Enrichment Configuration Tool"
	$objForm.Size = New-Object System.Drawing.Size(1010,810) 
	$objForm.StartPosition = "CenterScreen"

	$objLabel = New-Object System.Windows.Forms.Label
	$objLabel.Location = New-Object System.Drawing.Size(10,20) 
	$objLabel.Size = New-Object System.Drawing.Size(280,20) 
	$objLabel.Text = "Select management pack for enrichment:"
	$objForm.Controls.Add($objLabel) 

	# The combo box shows the displayname for each MP
	# BUG - if the displayname is not populated for the MP, it will be skipped
	#	    and no enrichment information may be set using this script.
	$objMPBox = New-Object System.Windows.Forms.ComboBox
	$objMPBox.Location = New-Object System.Drawing.Point(10,40)
	$objMPBox.Size = New-Object System.Drawing.Size(470,20)

	#get the MPs
    $managementpacks = get-scommanagementpack
	$managementpacks = $managementpacks | sort displayname
	foreach($mp in $managementpacks)
	{
		if(($mp.displayname -ne $null) -AND ($mp.displayname -ne ''))
		{
			$objMPBox.Items.Add($mp.displayname) | Out-Null
		}
        else
        {
			$objMPBox.Items.Add($mp.Name) | Out-Null
        }
	}	
	$objForm.Controls.Add($objMPBox)
    $objMPBox.Sorted = $true
	$objMPBox.Text = "Select a Management Pack"
    $objMPBox.Add_SelectedIndexChanged(
    {
            $objRuleMoniBox.Items.Clear();
            # stores the name for the MP in the global var
	        $global:mpname = $objMPBox.Text
            refreshView
            $objMPBox.Text = $global:mpname
    }
    )

    #Add Close Button to form
	$DoneButton = New-Object System.Windows.Forms.Button
	$DoneButton.Location = New-Object System.Drawing.Size(680,50)
	$DoneButton.Size = New-Object System.Drawing.Size(50,23)
	$DoneButton.Text = "Close"
	$DoneButton.Add_Click({$objForm.Close();})
	$objForm.Controls.Add($DoneButton)

	# The edit button changes all of the editable fields from ReadOnly = True to False allowing
	# edits to be made
	$GetEditButton = New-Object System.Windows.Forms.Button
	$GetEditButton.Location = New-Object System.Drawing.Size(500,50)
	$GetEditButton.Size = New-Object System.Drawing.Size(50,23)
	$GetEditButton.Text = "Edit"
	$GetEditButton.Add_Click({ 
    $objAlertNotesText.ReadOnly = $objNotesText.ReadOnly = $objOwnerText.ReadOnly = $objAPPIDText.ReadOnly = $objTicketText.ReadOnly = $objCF1Text.ReadOnly = $objCF2Text.ReadOnly = $objCF3Text.ReadOnly = $objCF4Text.ReadOnly = $objCF5Text.ReadOnly = $objCF6Text.ReadOnly = $objCF7Text.ReadOnly = $objCF8Text.ReadOnly = $objCF9Text.ReadOnly = $objCF10Text.ReadOnly = $objRESText.Readonly = $false;
    $AddGroupButton.Visible = $true
    $RemoveGroupButton.Visible = $True
    $ChangeEnrichButton.Enabled = $True
    $GetEditButton.Enabled = $false
    $objGroupBox.Enabled = $true
    })
    $objForm.Controls.Add($GetEditButton)
	
	# This is the save button that commits the GUI changes back to the custom object collection
	# for the rules and monitors
	$ChangeEnrichButton = New-Object System.Windows.Forms.Button
	$ChangeEnrichButton.Location = New-Object System.Drawing.Size(585,50)
	$ChangeEnrichButton.Size = New-Object System.Drawing.Size(50,23)
	$ChangeEnrichButton.Text = "Save"
	$ChangeEnrichButton.Add_Click({
								if($global:rulesmonis.gettype().Name -eq "Object[]"){
									$data = $global:rulesmonis[$objRuleMoniBox.Text.split(":")[0] -1];
								} else {
									$data = $global:rulesmonis
								}
								$data.Notes = $objNotesText.Text;
								$data.AlertNotes = $objAlertNotesText.Text;
								$data.Ticket = $objTicketText.Text;
								$data.Owner = $objOwnerText.Text;
								$data.APPID = $objAPPIDText.Text;
								$data.ResState = $objRESText.Text;
								$data.CF1 = $objCF1Text.Text;
								$data.CF2 = $objCF2Text.Text;
								$data.CF3 = $objCF3Text.Text;
								$data.CF4 = $objCF4Text.Text;
								$data.CF5 = $objCF5Text.Text;
								$data.CF6 = $objCF6Text.Text;
								$data.CF7 = $objCF7Text.Text;
								$data.CF8 = $objCF8Text.Text;
								$data.CF9 = $objCF9Text.Text;
								$data.CF10 = $objCF10Text.Text;
                                $data.SupportGroup = $objGroupBox.text;
                                if ($objGroupBox.text)
                                {
                                    $data.haschanged = 'true'
                                }
                                else 
                                {
                                    $output =  [System.Windows.Forms.MessageBox]::Show("Add Custom Field definition without Group definition?", "Empty Group Warning" , 4)
                                    if ($OUTPUT -eq "YES" ) 
                                          {  $data.haschanged = 'true'}
                                }
								
								if($global:rulesmonis.gettype().Name -eq "Object[]"){
									$ndx = [array]::indexof($global:rulesmonis, $data);
									$global:rulesmonis[$ndx] = $data;
								} else {
									$global:rulesmonis = $data;
								}
								$objAlertNotesText.ReadOnly = $objNotesText.ReadOnly = $objOwnerText.ReadOnly = $objAPPIDText.ReadOnly = $objTicketText.ReadOnly = $objCF1Text.ReadOnly = $objCF2Text.ReadOnly = $objCF3Text.ReadOnly = $objCF4Text.ReadOnly = $objCF5Text.ReadOnly = $objCF6Text.ReadOnly = $objCF7Text.ReadOnly = $objCF8Text.ReadOnly = $objCF9Text.ReadOnly = $objCF10Text.ReadOnly = $objRESText.Readonly = $true; 

                                ChangeEnrichment
                                $objGroupBox.Items.Clear()   
                                $objGroupBox.text = ""
                                $Groups = getSQLData("Select SupportGroup from CE_Enrichment where MonitorRuleID = '" + $data.id + "'")
                                foreach ($SupportGroup in $Groups.SupportGroup) 
                                {   
                                $Group
                                $objGroupBox.Items.Add($SupportGroup)
                                $objGroupBox.selectedindex = 0
                                } 
                                $AddGroupButton.Visible = $False
                                $RemoveGroupButton.Visible = $False
                                $ChangeEnrichButton.Enabled = $False
                                $objGroupBox.Enabled = $False
                                $GetEditButton.Enabled = $true
                                }
                                )


    $objForm.Controls.Add($ChangeEnrichButton)
    $ChangeEnrichButton.Enabled = $False

	$objMonitorRuleIDLabel = New-Object System.Windows.Forms.Label
	$objMonitorRuleIDLabel.Location = New-Object System.Drawing.Size(500,100)
	$objMonitorRuleIDLabel.Size = New-Object System.Drawing.Size(90,20)
	$objMonitorRuleIDLabel.Text = "Monitor/Rule ID"
	$objForm.Controls.Add($objMonitorRuleIDLabel)
	
	$objMonitorRuleIDText = New-Object System.Windows.Forms.TextBox
	$objMonitorRuleIDText.Location = New-Object System.Drawing.Size(600,100)
	$objMonitorRuleIDText.Size = New-Object System.Drawing.Size(375,20)
	$objMonitorRuleIDText.Text = ""
	$objMonitorRuleIDText.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objMonitorRuleIDText.ReadOnly = $true
	$objForm.Controls.Add($objMonitorRuleIDText)
	
	$objNameText = New-Object System.Windows.Forms.TextBox
	$objNameText.Location = New-Object System.Drawing.Size(600,125)
	$objNameText.Size = New-Object System.Drawing.Size(375,20)
	$objNameText.Text = ""
	$objNameText.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objNameText.ReadOnly = $true
	$objForm.Controls.Add($objNameText)
	
	$objTicketLabel = New-Object System.Windows.Forms.Label
	$objTicketLabel.Location = New-Object System.Drawing.Size(500,150)
	$objTicketLabel.Size = New-Object System.Drawing.Size(75,20)
	$objTicketLabel.Text = "Ticket Flag"
	$objForm.Controls.Add($objTicketLabel)
	
	$objTicketText = New-Object System.Windows.Forms.TextBox
	$objTicketText.Location = New-Object System.Drawing.Size(600,150)
	$objTicketText.Size = New-Object System.Drawing.Size(18,20)
	$objTicketText.Text = ""
	$objTicketText.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objTicketText.ReadOnly = $true
	$objTicketText.Add_TextChanged({
		switch ($objTicketText.Text) {
			"" {$objTicketDisplay.Text = "No Ticket"}
			"0" {$objTicketDisplay.Text = "No Ticket"}
			"1" {$objTicketDisplay.Text = "Ticket"}
			"2" {$objTicketDisplay.Text = "NOC Only"}
			"3" {$objTicketDisplay.Text = "Email Only"}
			default {$objTicketDisplay.Text = "Invalid {0-3}"}
		};
	})
	$objForm.Controls.Add($objTicketText)
	
	$objTicketDisplay = New-Object System.Windows.Forms.Label
	$objTicketDisplay.Location = New-Object System.Drawing.Size(625,150)
	$objTicketDisplay.Size = New-Object System.Drawing.Size(75,20)
	$objTicketDisplay.Text = ""
	$objForm.Controls.Add($objTicketDisplay)
	
	$objAPPIDLabel = New-Object System.Windows.Forms.Label
	$objAPPIDLabel.Location = New-Object System.Drawing.Size(500,200)
	$objAPPIDLabel.Size = New-Object System.Drawing.Size(75,20)
	$objAPPIDLabel.Text = "APP ID"
	$objForm.Controls.Add($objAPPIDLabel)

	$objAPPIDText = New-Object System.Windows.Forms.TextBox
	$objAPPIDText.Location = New-Object System.Drawing.Size(600,200)
	$objAPPIDText.Size = New-Object System.Drawing.Size(75,20)
	$objAPPIDText.Text = ""
	$objAPPIDText.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objAPPIDText.ReadOnly = $true
	$objForm.Controls.Add($objAPPIDText)
	
	$objOwnerLabel = New-Object System.Windows.Forms.Label
	$objOwnerLabel.Location = New-Object System.Drawing.Size(500,250)
	$objOwnerLabel.Size = New-Object System.Drawing.Size(75,20)
	$objOwnerLabel.Text = "Owner"
	$objForm.Controls.Add($objOwnerLabel)
	
	$objOwnerText = New-Object System.Windows.Forms.TextBox
	$objOwnerText.Location = New-Object System.Drawing.Size(600,250)
	$objOwnerText.Size = New-Object System.Drawing.Size(220,20)
	$objOwnerText.Text = ""
	$objOwnerText.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objOwnerText.ReadOnly = $true
	$objForm.Controls.Add($objOwnerText)
	
	$objRuleMonLabel = New-Object System.Windows.Forms.Label
	$objRuleMonLabel.Location = New-Object System.Drawing.Size(500,300)
	$objRuleMonLabel.Size = New-Object System.Drawing.Size(75,20)
	$objRuleMonLabel.Text = "Monitor/Rule"
	$objForm.Controls.Add($objRuleMonLabel)
	
	$objRuleMonText = New-Object System.Windows.Forms.TextBox
	$objRuleMonText.Location = New-Object System.Drawing.Size(600,300)
	$objRuleMonText.Size = New-Object System.Drawing.Size(75,20)
	$objRuleMonText.Text = ""
	$objRuleMonText.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objRuleMonText.ReadOnly = $true
	$objForm.Controls.Add($objRuleMonText)

	$objAlertNotesLabel = New-Object System.Windows.Forms.Label
	$objAlertNotesLabel.Location = New-Object System.Drawing.Size(500,350)
	$objAlertNotesLabel.Size = New-Object System.Drawing.Size(75,20)
	$objAlertNotesLabel.Text = "Alert Notes"
	$objForm.Controls.Add($objAlertNotesLabel)	
	
	$objAlertNotesText = New-Object System.Windows.Forms.TextBox
	$objAlertNotesText.Location = New-Object System.Drawing.Size(500,375)
	$objAlertNotesText.Size = New-Object System.Drawing.Size(200,50)
	$objAlertNotesText.MultiLine = $true
	$objAlertNotesText.Text = ""
	$objAlertNotesText.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objAlertNotesText.ReadOnly = $true
	$objForm.Controls.Add($objAlertNotesText)
	
	$objNotesLabel = New-Object System.Windows.Forms.Label
	$objNotesLabel.Location = New-Object System.Drawing.Size(500,450)
	$objNotesLabel.Size = New-Object System.Drawing.Size(75,20)
	$objNotesLabel.Text = "Notes"
	$objForm.Controls.Add($objNotesLabel)	
	
	$objNotesText = New-Object System.Windows.Forms.TextBox
	$objNotesText.Location = New-Object System.Drawing.Size(500,475)
	$objNotesText.Size = New-Object System.Drawing.Size(200,50)
	$objNotesText.MultiLine = $true
	$objNotesText.Text = ""
	$objNotesText.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objNotesText.ReadOnly = $true
	$objForm.Controls.Add($objNotesText)

	$objRuleMoniBox = New-Object System.Windows.Forms.ListBox
    $objRuleMoniBox.SelectionMode = "MultiExtended"
	$objRuleMoniBox.Location = New-Object System.Drawing.Point(70,90)
	$objRuleMoniBox.Size = New-Object System.Drawing.Size(400,650)
	$tempcount = 1
	foreach($rm in $global:rulesmonis)
	{
		$objRuleMoniBox.Items.Add('' + $tempcount + ': ' + $rm.displayname)
		$tempcount++
	}
	$objRuleMoniBox.Add_SelectedValueChanged({ 
								if($global:rulesmonis.gettype().Name -eq "Object[]"){
									$data = $global:rulesmonis[$objRuleMoniBox.Text.split(":")[0] -1];
								} else {
									$data = $global:rulesmonis
								}
								$objAlertNotesText.Text = $data.AlertNotes;
								$objNotesText.Text = $data.Notes;
								$objTicketText.Text = $data.Ticket;
								$objOwnerText.Text = $data.Owner;
								$objAPPIDText.Text = $data.APPID;
								$objMonitorRuleIDText.Text = $data.id;
								$objNameText.Text = $data.name;
								$objRuleMonText.Text = $data.RuleMoniType;
								$objRESText.Text = $data.ResState;
								$objCF1Text.Text = $data.CF1;
								$objCF2Text.Text = $data.CF2;
								$objCF3Text.Text = $data.CF3;
								$objCF4Text.Text = $data.CF4;
								$objCF5Text.Text = $data.CF5;
								$objCF6Text.Text = $data.CF6;
								$objCF7Text.Text = $data.CF7;
								$objCF8Text.Text = $data.CF8;
								$objCF9Text.Text = $data.CF9;
								$objCF10Text.Text = $data.CF10;
                                
                                $objGroupBox.Items.Clear()   
                                $objGroupBox.text = ""
                                #GetGroupEnrichment($data.id)
                                $Groups = getSQLData("Select SupportGroup from CE_Enrichment where MonitorRuleID = '" + $data.id + "'")
                                foreach ($Group in $Groups[1])
                                {   
                                    $objGroupBox.Items.Add($Group.SupportGroup)
                                    $objGroupBox.selectedindex = 0
                                    #$objGroupListBox.Items.Add($Group.SupportGroup)
                                    #$objGroupListBox.selectedindex = 0
                                } 
								})
	$objForm.Controls.Add($objRuleMoniBox)

    $ctxMenu = New-Object System.Windows.Forms.ContextMenu
    $ctxCreateSiteMenuItem = New-Object System.Windows.Forms.MenuItem
    $ctxCreateSiteMenuItem.Text = "Bulk Insert"        
    $ctxCreateSiteMenuItem.add_Click({ param($sender, $eargs)
            bulkUpdate $objRuleMoniBox.SelectedItems
     })
    $ctxMenu.MenuItems.AddRange(@($ctxCreateSiteMenuItem))
    $objRuleMoniBox.ContextMenu = $ctxMenu

	$objRESLabel = New-Object System.Windows.Forms.Label
	$objRESLabel.Location = New-Object System.Drawing.Size(720,300)
	$objRESLabel.Size = New-Object System.Drawing.Size(75,20)
	$objRESLabel.Text = "Res State"
	$objForm.Controls.Add($objRESLabel)
	
	$objRESText = New-Object System.Windows.Forms.TextBox
	$objRESText.Location = New-Object System.Drawing.Size(800,300)
	$objRESText.Size = New-Object System.Drawing.Size(30,20)
	$objRESText.Text = ""
	$objRESText.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objRESText.ReadOnly = $true
	$objForm.Controls.Add($objRESText)

    #Group Selection
	$objGroupLabel = New-Object System.Windows.Forms.Label
	$objGroupLabel.Location = New-Object System.Drawing.Size(500,550)
	$objGroupLabel.Size = New-Object System.Drawing.Size(50,20)
	$objGroupLabel.Text = "Groups"
	$objForm.Controls.Add($objGroupLabel)

    $objGroupBox = New-Object System.Windows.Forms.ListBox
    #$objGroupBox.SelectionMode = "MultiExtended"
	$objGroupBox.Location = New-Object System.Drawing.Point(500,575)
	$objGroupBox.Size = New-Object System.Drawing.Size(200,160)
	$objForm.Controls.Add($objGroupBox)
    $objGroupBox.Enabled = $False
    
    $objGroupBox.Add_SelectedIndexChanged(
    {
        $dataForGroup = getSQLData("Select * from CE_Enrichment where SupportGroup = '" + $objGroupBox.Text + "' and MonitorRuleID = '" + $objMonitorRuleIDText.Text  + "' order by SupportGroup")

        $objCF1Text.Text = $global:CustomFields.CF1
		$objCF2Text.Text = $dataForGroup.CF2
		$objCF3Text.Text = $dataForGroup.CF3
		$objCF4Text.Text = $dataForGroup.CF4
		$objCF5Text.Text = $dataForGroup.CF5
		$objCF6Text.Text = $dataForGroup.CF6
		$objCF7Text.Text = $dataForGroup.CF7
		$objCF8Text.Text = $dataForGroup.CF8
		$objCF9Text.Text = $dataForGroup.CF9
		$objCF10Text.Text = $dataForGroup.CF10
        }
    )

	$AddGroupButton = New-Object System.Windows.Forms.Button
	$AddGroupButton.Location = New-Object System.Drawing.Size(550,545)
	$AddGroupButton.Size = New-Object System.Drawing.Size(23,23)
	$AddGroupButton.Text = "+"
	$objForm.Controls.Add($AddGroupButton)
    $AddGroupButton.Visible = $false
    
    $AddGroupButton.Add_Click({
            $addGroupForm = New-Object System.Windows.Forms.Form 
            $addGroupForm.Text = "Add new group"
            $addGroupForm.Size = New-Object System.Drawing.Size(300,200) 
            $addGroupForm.StartPosition = "CenterScreen"

            $OKButton = New-Object System.Windows.Forms.Button
            $OKButton.Location = New-Object System.Drawing.Point(75,120)
            $OKButton.Size = New-Object System.Drawing.Size(75,23)
            $OKButton.Text = "OK"
            $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $addGroupForm.AcceptButton = $OKButton
            $addGroupForm.Controls.Add($OKButton)

            $CancelButton = New-Object System.Windows.Forms.Button
            $CancelButton.Location = New-Object System.Drawing.Point(150,120)
            $CancelButton.Size = New-Object System.Drawing.Size(75,23)
            $CancelButton.Text = "Cancel"
            $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $addGroupForm.CancelButton = $CancelButton
            $addGroupForm.Controls.Add($CancelButton)

            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10,20) 
            $label.Size = New-Object System.Drawing.Size(280,20) 
            $label.Text = "Group Name:"
            $addGroupForm.Controls.Add($label) 

            $objGroupBoxPopUp = New-Object System.Windows.Forms.ComboBox
	        $objGroupBoxPopUp.Location = New-Object System.Drawing.Point(10,40)
	        $objGroupBoxPopUp.Size = New-Object System.Drawing.Size(260,20)
            $addGroupForm.Controls.Add($objGroupBoxPopUp)
            $scomGroups = Get-SCOMGroup | select DisplayName | sort-Object DisplayName
            $objGroupBoxPopUp.Items.AddRange($scomGroups.DisplayName)

            $addGroupForm.Topmost = $True

            $addGroupForm.Add_Shown({$objGroupBoxPopUp.Select()})
            $result = $addGroupForm.ShowDialog()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK)
            {
                    if ($objGroupBox.Items.Contains($objGroupBoxPopUp.Text))
                    {
                    [Windows.Forms.MessageBox]::Show($objGroupBoxPopUp.Text  + " already added to group list”,"Add group error", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)
                    }
                    else
                    {
                    $objGroupBox.Items.Add($objGroupBoxPopUp.Text + "*")
                    $objGroupBox.Text = $objGroupBoxPopUp.Text + "*"
                    }
            }
    })

	$RemoveGroupButton = New-Object System.Windows.Forms.Button
	$RemoveGroupButton.Location = New-Object System.Drawing.Size(575,545)
	$RemoveGroupButton.Size = New-Object System.Drawing.Size(23,23)
	$RemoveGroupButton.Text = "-"
	$objForm.Controls.Add($RemoveGroupButton)
	$RemoveGroupButton.Visible = $false

    $RemoveGroupButton.Add_Click({
       if ($objGroupBox.Text -eq "")
       {
           [System.Windows.Forms.MessageBox]::Show("No group selected") 
       }
       else
       {
       $output =  [System.Windows.Forms.MessageBox]::Show("Remove Group: """ + $objGroupBox.Text + """?", "Group Removal Warning" , 4)
       if ($OUTPUT -eq "YES" ) 
            {
                        	
		            $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
		            $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True"
		            $SqlCmd = New-Object System.Data.SqlClient.SqlCommand	
		            $SqlCmd.Connection = $SqlConnection
		            $SqlConnection.Open()
                    $SqlCmd.CommandText = "delete  from CE_Enrichment where SupportGroup = '" + $objGroupBox.Text + "' and MonitorRuleID ='" + $objMonitorRuleIDText.Text + "'"
		            $SqlCmd.ExecuteNonQuery()
                    $SqlConnection.Close()	
	
                    $objGroupBox.Items.Clear()   
                    $objGroupBox.text = ""
                    $Groups = getSQLData("Select SupportGroup from CE_Enrichment where MonitorRuleID = '" + $objMonitorRuleIDText.Text + "'")
                    #GetGroupEnrichment($objMonitorRuleIDText.Text)

                    foreach ($Group in $Groups[1])
                    {   
                        $objGroupBox.Items.Add($Group.SupportGroup)
                    } 
                    $objAlertNotesText.ReadOnly = $objNotesText.ReadOnly = $objOwnerText.ReadOnly = $objAPPIDText.ReadOnly = $objTicketText.ReadOnly = $objCF1Text.ReadOnly = $objCF2Text.ReadOnly = $objCF3Text.ReadOnly = $objCF4Text.ReadOnly = $objCF5Text.ReadOnly = $objCF6Text.ReadOnly = $objCF7Text.ReadOnly = $objCF8Text.ReadOnly = $objCF9Text.ReadOnly = $objCF10Text.ReadOnly = $objRESText.Readonly = $true;
                    $objAlertNotesText.Text = $objNotesText.Text = $objOwnerText.Text = $objAPPIDText.Text = $objTicketText.Text = $objCF1Text.Text = $objCF2Text.Text = $objCF3Text.Text = $objCF4Text.Text = $objCF5Text.Text = $objCF6Text.Text = $objCF7Text.Text = $objCF8Text.Text = $objCF9Text.Text = $objCF10Text.Text = $objRESText.Text = "";
                    $AddGroupButton.Visible = $False
                    $RemoveGroupButton.Visible = $False
                    $ChangeEnrichButton.Enabled = $False
                    $objGroupBox.Enabled = $False
                    $GetEditButton.Enabled = $true
                    if ($objGroupBox.items.Count  -gt 0)
                    {
                      $objGroupBox.selectedindex = 0  
                    }
            } 
        }
           
    })

#CF
	$offset = 33
	$start = 375
	$objCF1Label = New-Object System.Windows.Forms.Label
	$objCF1Label.Location = New-Object System.Drawing.Size(720,$start)
	$objCF1Label.Size = New-Object System.Drawing.Size(30,20)
	$objCF1Label.Text = "CF1"
	$objForm.Controls.Add($objCF1Label)	
	
	$objCF1Text = New-Object System.Windows.Forms.TextBox
	$objCF1Text.Location = New-Object System.Drawing.Size(755,$start)
	$objCF1Text.Size = New-Object System.Drawing.Size(225,20)
	$objCF1Text.Text = ""
	$objCF1Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objCF1Text.ReadOnly = $true
	$objForm.Controls.Add($objCF1Text)
	
	$start+=$offset
	$objCF2Label = New-Object System.Windows.Forms.Label
	$objCF2Label.Location = New-Object System.Drawing.Size(720,$start)
	$objCF2Label.Size = New-Object System.Drawing.Size(30,20)
	$objCF2Label.Text = "CF2"
	$objForm.Controls.Add($objCF2Label)	
	
	$objCF2Text = New-Object System.Windows.Forms.TextBox
	$objCF2Text.Location = New-Object System.Drawing.Size(755,$start)
	$objCF2Text.Size = New-Object System.Drawing.Size(225,20)
	$objCF2Text.Text = ""
	$objCF2Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objCF2Text.ReadOnly = $true
	$objForm.Controls.Add($objCF2Text)
	
	$start+=$offset
	$objCF3Label = New-Object System.Windows.Forms.Label
	$objCF3Label.Location = New-Object System.Drawing.Size(720,$start)
	$objCF3Label.Size = New-Object System.Drawing.Size(30,20)
	$objCF3Label.Text = "CF3"
	$objForm.Controls.Add($objCF3Label)	
	
	$objCF3Text = New-Object System.Windows.Forms.TextBox
	$objCF3Text.Location = New-Object System.Drawing.Size(755,$start)
	$objCF3Text.Size = New-Object System.Drawing.Size(225,20)
	$objCF3Text.Text = ""
	$objCF3Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objCF3Text.ReadOnly = $true
	$objForm.Controls.Add($objCF3Text)
	
	$start+=$offset
	$objCF4Label = New-Object System.Windows.Forms.Label
	$objCF4Label.Location = New-Object System.Drawing.Size(720,$start)
	$objCF4Label.Size = New-Object System.Drawing.Size(30,20)
	$objCF4Label.Text = "CF4"
	$objForm.Controls.Add($objCF4Label)	
	
	$objCF4Text = New-Object System.Windows.Forms.TextBox
	$objCF4Text.Location = New-Object System.Drawing.Size(755,$start)
	$objCF4Text.Size = New-Object System.Drawing.Size(225,20)
	$objCF4Text.Text = ""
	$objCF4Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objCF4Text.ReadOnly = $true
	$objForm.Controls.Add($objCF4Text)
	
	$start+=$offset
	$objCF5Label = New-Object System.Windows.Forms.Label
	$objCF5Label.Location = New-Object System.Drawing.Size(720,$start)
	$objCF5Label.Size = New-Object System.Drawing.Size(30,20)
	$objCF5Label.Text = "CF5"
	$objForm.Controls.Add($objCF5Label)	
	
	$objCF5Text = New-Object System.Windows.Forms.TextBox
	$objCF5Text.Location = New-Object System.Drawing.Size(755,$start)
	$objCF5Text.Size = New-Object System.Drawing.Size(225,20)
	$objCF5Text.Text = ""
	$objCF5Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objCF5Text.ReadOnly = $true
	$objForm.Controls.Add($objCF5Text)
	
	$start+=$offset
	$objCF6Label = New-Object System.Windows.Forms.Label
	$objCF6Label.Location = New-Object System.Drawing.Size(720,$start)
	$objCF6Label.Size = New-Object System.Drawing.Size(30,20)
	$objCF6Label.Text = "CF6"
	$objForm.Controls.Add($objCF6Label)	
	
	$objCF6Text = New-Object System.Windows.Forms.TextBox
	$objCF6Text.Location = New-Object System.Drawing.Size(755,$start)
	$objCF6Text.Size = New-Object System.Drawing.Size(225,20)
	$objCF6Text.Text = ""
	$objCF6Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objCF6Text.ReadOnly = $true
	$objForm.Controls.Add($objCF6Text)
	
	$start+=$offset
	$objCF7Label = New-Object System.Windows.Forms.Label
	$objCF7Label.Location = New-Object System.Drawing.Size(720,$start)
	$objCF7Label.Size = New-Object System.Drawing.Size(30,20)
	$objCF7Label.Text = "CF7"
	$objForm.Controls.Add($objCF7Label)	
	
	$objCF7Text = New-Object System.Windows.Forms.TextBox
	$objCF7Text.Location = New-Object System.Drawing.Size(755,$start)
	$objCF7Text.Size = New-Object System.Drawing.Size(225,20)
	$objCF7Text.Text = ""
	$objCF7Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objCF7Text.ReadOnly = $true
	$objForm.Controls.Add($objCF7Text)
	
	$start+=$offset
	$objCF8Label = New-Object System.Windows.Forms.Label
	$objCF8Label.Location = New-Object System.Drawing.Size(720,$start)
	$objCF8Label.Size = New-Object System.Drawing.Size(30,20)
	$objCF8Label.Text = "CF8"
	$objForm.Controls.Add($objCF8Label)	
	
	$objCF8Text = New-Object System.Windows.Forms.TextBox
	$objCF8Text.Location = New-Object System.Drawing.Size(755,$start)
	$objCF8Text.Size = New-Object System.Drawing.Size(225,20)
	$objCF8Text.Text = ""
	$objCF8Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objCF8Text.ReadOnly = $true
	$objForm.Controls.Add($objCF8Text)
	
	$start+=$offset
	$objCF9Label = New-Object System.Windows.Forms.Label
	$objCF9Label.Location = New-Object System.Drawing.Size(720,$start)
	$objCF9Label.Size = New-Object System.Drawing.Size(30,20)
	$objCF9Label.Text = "CF9"
	$objForm.Controls.Add($objCF9Label)	
	
	$objCF9Text = New-Object System.Windows.Forms.TextBox
	$objCF9Text.Location = New-Object System.Drawing.Size(755,$start)
	$objCF9Text.Size = New-Object System.Drawing.Size(225,20)
	$objCF9Text.Text = ""
	$objCF9Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objCF9Text.ReadOnly = $true
	$objForm.Controls.Add($objCF9Text)
	
	$start+=$offset
	$objCF10Label = New-Object System.Windows.Forms.Label
	$objCF10Label.Location = New-Object System.Drawing.Size(720,$start)
	$objCF10Label.Size = New-Object System.Drawing.Size(35,20)
	$objCF10Label.Text = "CF10"
	$objForm.Controls.Add($objCF10Label)	
	
	$objCF10Text = New-Object System.Windows.Forms.TextBox
	$objCF10Text.Location = New-Object System.Drawing.Size(755,$start)
	$objCF10Text.Size = New-Object System.Drawing.Size(225,20)
	$objCF10Text.Text = ""
	$objCF10Text.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
	$objCF10Text.ReadOnly = $true
	$objForm.Controls.Add($objCF10Text)
	
#CF End
	$objForm.Topmost = $True
	$objForm.Add_Shown({$objForm.Activate()})
	[void] $objForm.ShowDialog()
	
}

# This function merges the data from the rules and monitors with the enrichment data
# in the custom object collection
function MergeData
{
	foreach($e in $global:enrichment[1]){
		foreach($rm in $global:rulesmonis) {
			if($rm.id.trim() -eq $e.monitorruleid.tostring().trim() ) {
				$rm.ticket = $e.ticket
				$rm.APPID = $e.APPID
				$rm.Notes = $e.Notes
				$rm.AlertNotes = $e.AlertNotes
				$rm.Owner = $e.Owner
				$rm.PrevEnrich = 'true'
				$rm.ResState = $e.ResolutionState
				$rm.CF1 = $e.CF1
				$rm.CF2 = $e.CF2
				$rm.CF3 = $e.CF3
				$rm.CF4 = $e.CF4
				$rm.CF5 = $e.CF5
				$rm.CF6 = $e.CF6
				$rm.CF7 = $e.CF7
				$rm.CF8 = $e.CF8
				$rm.CF9 = $e.CF9
				$rm.CF10 = $e.CF10
                #$rm.SupportGroup = $e.SupportGroup
			}
		}
	}
}

# This function constructs the SQL statements for inserting and modifying the enrichment information
# This function is executed post the commit button being pressed and the GUI has closed.
function ChangeEnrichment
{
	# This section dumps out debug info the console.
	write-host "Here are the enrichment changes"
	$toenrich = $global:rulesmonis | where{$_.haschanged -eq 'true'}
	write-host $toenrich.count
	write-host $toenrich.size
	
	if($toenrich -ne $null)
	{
		$global:rulesmonis | where{$_.haschanged -eq 'true'} | ft *
		
		$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
		$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True"
		$SqlCmd = New-Object System.Data.SqlClient.SqlCommand	
		$SqlCmd.Connection = $SqlConnection
		$SqlConnection.Open()
	
###########################
###   MAKE CHANGES HERE ###
###   Update and Insert statements may need to be updated ###
###   The update and insert statements match the table created with the provided SQL table create script ###

        $toenrich | foreach{
            if ($_.SupportGroup.contains("*") -eq $true) 
            {
            $_.PrevEnrich = $false
            $_.SupportGroup = $_.SupportGroup.replace("*","")
            }
			if($_.APPID -eq "") { $_.APPID = $default_appid }
			if($_.Owner -eq "") { $_.Owner = $default_owner }		
			if($_.ResState -eq "") { $_.ResState = $default_resstate }
			if($_.ticket -eq "") { $_.ticket = $default_ticket }
			$_.Notes = $_.Notes.trim() + " Updated: " + (Get-Date).tostring()
			if($_.PrevEnrich -eq 'true')
			{
				$newQuery = "update " + $SQLEnrichTableName + " set Displayname = '" + $_.Displayname + "', ManagementPack = '" + $global:mpname + "', Type = '" + $_.RuleMoniType + "', Ticket = '" + $_.Ticket + "', APPID = '" + $_.APPID + "', Owner = '" + $_.Owner + "', AlertNotes = '" + $_.AlertNotes.trim() + "', Notes = '" + $_.Notes.trim() + "', ResolutionState ='" + $_.ResState.trim() + "', CF1 = '" + $_.CF1 + "', CF2 = '" + $_.CF2 + "', CF3 = '" + $_.CF3 + "', CF4 = '" + $_.CF4 + "', CF5 = '" + $_.CF5 + "', CF6 = '" + $_.CF6 + "', CF7 = '" + $_.CF7 + "', CF8 = '" + $_.CF8 + "', CF9 = '" + $_.CF9 + "', CF10 = '" + $_.CF10 + "' where MonitorRuleID = '" + $_.id + "' and SupportGroup = '" + $_.SupportGroup + "'"
            Write-Host $newQuery
			}

			else
			{
				$newQuery = "Insert Into " + $SQLEnrichTableName + " (DisplayName,MonitorRuleID,ManagementPack,Type,Ticket,APPID,Owner,AlertNotes,Notes,ResolutionState,CF1,CF2,CF3,CF4,CF5,CF6,CF7,CF8,CF9,CF10,SupportGroup) values ('" + $_.DisplayName + "','" + $_.ID + "','" + $global:mpname + "','" + $_.RuleMoniType + "','" + $_.Ticket + "','" + $_.APPID + "','" + $_.Owner + "','" + $_.AlertNotes.trim() + "','" + $_.Notes.trim() + "','" + $_.ResState.trim() + "','" + $_.CF1 + "','" + $_.CF2 + "','" + $_.CF3 + "','" + $_.CF4 + "','" + $_.CF5 + "','" + $_.CF6 + "','" + $_.CF7 + "','" + $_.CF8 + "','" + $_.CF9 + "','" + $_.CF10 + "','" + $_.SupportGroup + "')"		
			}	
###########################
			$newQuery
			$SqlCmd.CommandText = $newQuery
			$SqlCmd.ExecuteNonQuery()
		}		
		$SqlConnection.Close()	
	}
}

# Execute the main driver function for the script

main

