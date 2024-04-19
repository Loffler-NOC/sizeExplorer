#attempting to recreate size explorer script from CWA

#Delete c:\windows\temp\sz-rpt.sef if it exists
$file = "C:\Windows\Temp\sz-rpt.sef"

Write-Output "Checking if $file exists, it needs to be deleted if it does"

if (Test-Path $file) {
    Remove-Item $file -Force
    Write-Output "File $file has been deleted."
} else {
    Write-Output "File $file does not exist."
}

# Define an array containing the registry keys, data, and types
$registrySettings = @(
    ("HKLM:\SOFTWARE\JSDSoftware\SEScan4", "DataModes", 321, "DWORD"),
    ("HKLM:\SOFTWARE\JSDSoftware\SEScan4", "Key" , $env:SizeExplorerLicenseKey, "String"),
    ("HKLM:\SOFTWARE\JSDSoftware\SEScan4", "Username" , "MICHAEL REDDIG", "String"),
    ("HKLM:\SOFTWARE\Wow6432Node\JSDSoftware\SEScan4", "DataModes" , 321, "DWORD"),
    ("HKLM:\SOFTWARE\Wow6432Node\JSDSoftware\SEScan4", "Key" , $env:SizeExplorerLicenseKey, "String"),
    ("HKLM:\SOFTWARE\Wow6432Node\JSDSoftware\SEScan4", "Username" , "MICHAEL REDDIG", "String")
)

# Loop through each item in the registrySettings array
foreach ($setting in $registrySettings) {
    $key = $setting[0]
    $name = $setting[1]
    $data = $setting[2]
    $type = $setting[3]

    # Check if the property already exists
    $propertyExists = Get-ItemProperty -Path $key | Select-Object -ExpandProperty $name -ErrorAction SilentlyContinue
    if ($null -eq $propertyExists) {
        # If property doesn't exist, set the registry value
        if (!(Test-Path $key)) {
            New-Item -Path $key -Force | Out-Null
        }
        New-ItemProperty -Path $key -Name $name -Value $data -PropertyType $type
    }
}

#Run size explorer
$driveLetter = $env:driveLetter
Start-Process -FilePath ".\sescan.exe" -ArgumentList "/o c:\windows\temp\sz-rpt.sef /s $driveLetter" -NoNewWindow -Wait

#Email the file
# Check if the module is installed
if (-not (Get-Module -Name Mailozaurr -ListAvailable)) {
    Write-Host "Mailozaurr module is not installed. Attempting to install..."
    Install-Module -Name Mailozaurr -AllowClobber -Force
    if ($?) {
        Write-Host "Mailozaurr module installed successfully."
        Import-Module -Name Mailozaurr -Force
        Write-Host "Mailozaurr module imported."
    } else {
        Write-Host "Failed to install Mailozaurr module. Please check for errors."
        exit
    }
} else {
    Write-Host "Mailozaurr module is already installed."
    Import-Module -Name Mailozaurr -Force
    Write-Host "Mailozaurr module imported."
}

# Update the module
Write-Host "Checking for updates to Mailozaurr module..."
Update-Module -Name Mailozaurr -Force
if ($?) {
    Write-Host "Mailozaurr module is up to date."
} else {
    Write-Host "Failed to update Mailozaurr module. Please check for errors."
}

# Send email with CSV attachment
$sendToEmail = $env:sendToEmail
$hostname = hostname
$scanFilePath = "c:\windows\temp\sz-rpt.sef"

$smtpServer = "mail.smtp2go.com"
$smtpPort = 2525
$from = "Loffler-NOCAlerts@loffler.com"
$to = $sendToEmail
$SMTPUsername = "Loffler-NOCAlerts"
$SMTPPassword = $env:SMTPEmailPassword
[securestring]$secStringPassword = ConvertTo-SecureString $SMTPPassword -AsPlainText -Force
[pscredential]$EmailCredential = New-Object System.Management.Automation.PSCredential ($SMTPUsername, $secStringPassword)
$subject = "Size Explorer Scan"
$body = "Attached is the size explorer scan for the $driveLetter drive on $hostname"
$attachment = $scanFilePath

Send-EmailMessage `
    -SmtpServer $smtpServer `
    -Port $smtpPort `
    -From $from `
    -To $to `
    -Credential $EmailCredential `
    -Subject $subject `
    -Body $body `
    -Attachments $attachment


#Clean up files
#Delete c:\windows\temp\sz-rpt.sef if it exists
$file = "C:\Windows\Temp\sz-rpt.sef"

Write-Output "Checking if $file exists, it needs to be deleted if it does"

if (Test-Path $file) {
    Remove-Item $file -Force
    Write-Output "File $file has been deleted."
} else {
    Write-Output "File $file does not exist."
}

#Clean up reg keys
Write-Output "Cleaning up registry keys"
Remove-Item -Path "HKLM:\SOFTWARE\JSDSoftware\SEScan4\" -Force
Remove-Item -Path "HKLM:\SOFTWARE\Wow6432Node\JSDSoftware\SEScan4" -Force
