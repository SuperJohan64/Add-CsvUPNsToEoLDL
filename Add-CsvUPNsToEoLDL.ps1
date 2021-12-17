# Turns on Transcription to log all activity.
$TimeStamp = Get-Date -Format "yyyyMMddTHHmmssffff"
$LogFile = "$PSScriptRoot\Logs\$TimeStamp.txt"
Start-Transcript -Path $LogFile

# A small function to end the transcription and close the script
function EndTranscription {
    Stop-Transcript
    Invoke-Item $LogFile
    Try {Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop}
    Catch {}
}

# Connects to Exchange Online using the current user's session.
Try {Connect-ExchangeOnline -UserPrincipalName (Get-ADUser $env:USERNAME).UserPrincipalName -ErrorAction Stop}
Catch {
    Write-Host "Could not connect to Exchange Online using the current user's session."
    Try {Connect-ExchangeOnline -ErrorAction Stop}
    Catch {
        Write-Host "Could not connect to Exchange Online. Closing script."
        EndTranscription
        Exit
    }    
}

# Prompts the user for an AAD Group's Object ID then gets the Group's detials.
$EOLGroupIdentity = Read-Host "Enter an EOL Distribution Group's Identity (name, alias, email address, GUID)"
Try {$EOLGroup = Get-DistributionGroup -Identity $EOLGroupIdentity -ErrorAction Stop}
Catch {
    EndTranscription
    Exit
}

# Displays Group Detials and other instructions to the user.
Write-Host "Group Details:"
Write-Host "`nObject ID:" $EOLGroup.ObjectId
Write-Host "DisplayName:" $EOLGroup.DisplayName
Write-Host "Description:" $EOLGroup.Description
Write-Host "PrimarySmtpAddress:" $EOLGroup.PrimarySmtpAddress
Write-Host "`nProvide a Csv file containing a column titled 'UserPrincipalName' containing a list of UserPrincipalNames from AD or AAD."
PAUSE
Write-Host ""

# Launches an open file dialog window from .Net to select the CSV File.
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = $PSScriptRoot
$OpenFileDialog.filter = "CSV (*.csv)| *.csv"
$OpenFileDialog.ShowDialog() | Out-Null
$OpenFileDialog.filename
$CsvFilePath = $OpenFileDialog.filename

# Imports the data from the CSV file.
Try {$CsvData = Import-Csv -Path $CsvFilePath -ErrorAction Stop}
Catch {
    EndTranscription
    Exit
}

# A loop that parses each user in the CSV file and adds them to the AAD Group.
foreach ($CsvUser in $CsvData) {
    Write-Host "Adding" $CsvUser.UserPrincipalName "to" $EOLGroup.PrimarySmtpAddress
    Add-DistributionGroupMember -Identity $EOLGroup.PrimarySmtpAddress -Member $CsvUser.UserPrincipalName
}

EndTranscription
