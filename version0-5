<#
This script is to rename computers in bulk. Each time it runs, it will prompt for admin credentials. 

Input:
Input must be provided in the form of a CSV file. The first column should be the current name of the
computer. The second column should be the name you want to apply to the computer. There should be no
header row in the CSV file. 

Actions:
This script renames any provided computers that are connected to the network. It will restart the
target computer immediately. 

Output:
As output a CSV will be provided that has the following information: 
    -Previous Host Name
    -New Host Name
    -Rename Successful (True or False)
    -Reason for Failure (If applicable)
Additional output can be produced by piping the $RenamingInformationTable to various output commands
(Format-Table, Convert-ToHTML, Out-GridView, Write-Host, Out-File, etc.). 

Version 0.5
This version successfully renames computers that are network-connected and recognize the supplied 
administrator credentials, but does not successfully output a report. 
#>
$AdminCredentials = Get-Credential  # Must be administrator credential
$i = 1  # Define counter variable for progress bar. 

# Define collection of computers to rename
$Computers = Import-Csv "ComputersToRename.csv" -Header OldName, NewName

# Create Table
$RenamingInformationTable = New-Object system.Data.DataTable "Renaming Information"

#Create Columns for Table
$OldName = New-Object system.Data.DataColumn "Previous Name",([string])
$NewName = New-Object system.Data.DataColumn "New Name",([string])
$SuccessStatus = New-Object system.Data.DataColumn "Renaming Successful",([bool])
$FailureReason = New-Object system.Data.DataColumn "Failure Reason",([string])

# Add Columns to Table
$RenamingInformationTable.columns.Add($OldName)
$RenamingInformationTable.columns.Add($NewName)
$RenamingInformationTable.Columns.Add($SuccessStatus)
$RenamingInformationTable.Columns.Add($FailureReason)

ForEach ($Computer in $Computers) 
{ #ForEach loop 1
    # Display progress bar for scan
    Write-Progress -Activity "Renaming specified machines..." -Status "Machine $i of $($Computers.Count)" -PercentComplete (($i / $Computers.Count) * 100)

    If(Test-Connection -ComputerName $Computer.OldName -Quiet)
    { #If statement 1
        # Attempt to rename the current $Computer. 
        $RenameVarTemp = Rename-Computer -ComputerName $Computer.OldName -NewName $Computer.NewName -DomainCredential $AdminCredentials -Force -Restart -ErrorAction SilentlyContinue

        # Set the value of $SuccessMessage
        $SuccessMessage = $RenameVarTemp.HasSucceeded

        If($SuccessMessage-eq $true) {$FailureMessage = "N/A"}
        Else {$FailureMessage = $Error[0].ToString()}
    } #If statement 1

    Else 
    { #Else statement 2
        $SuccessMessage = $false
        $FailureMessage = "Machine offline"
    } #Else statement 2
    
    # Create a row. 
    $Row = $RenamingInformationTable.NewRow()

    # Enter data into the row. 
    $Row.OldName = $Computer.OldName
    $Row.NewName  =$Computer.NewName
    $Row.SuccessStatus = $SuccessMessage
    $Row.FailureReason = $FailureMessage

    # Add the Row to the Table
    $RenamingInformationTable.Rows.Add($Row)

    #Increment iteration count
    $i++
} # Close ForEach loop 1

<# Section for formatting and outputting data. #>
$output = read-host "Output Type ((T)able, (G)ridview, (H)TML, default List)"

# Table
If ($output -match "T"){ $RenamingInformationTable | Format-Table }

#HTML
ElseIf ($output -match "H")
{ 
    $htmlfile = ".\LogonActivity.html"

    # HTML Styles
    $style = "<style>"
    $style = $style + "BODY{background-color:#F2F2F2;}"
    $style = $style + "TABLE{border-width: 1px;border-style: solid;border-color: black;}"
    $style = $style + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:#BDBDBD}"
    $style = $style + "TD{border-width: 1px;padding: 5px;border-style: solid;border-color: black;background-color:#D8D8D8}"
    $style = $style + "</style>"
    
    $RenamingInformationTable | Select-Object OldName, NewName, SuccessStatus, FailureReason | ConvertTo-HTML -Title "Lab Ping Response Report" -head $style -body "<h2>Logon Activity:</h2>" | Out-File $htmlfile
    Invoke-Expression $htmlfile
}

# Grid View
ElseIf ($output -match "G") { $RenamingInformationTable | Out-GridView -Title "Ping Responsiveness" }

# Default output, returns the table object in list form by default
Else { $RenamingInformationTable }
