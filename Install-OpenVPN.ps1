# 1. Check Admin

$ErrorActionPreference = "Stop"
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run as Administrator to install the software."
    Exit
}

# File Variables
$PublicProfilePath = "C:\Data\OpenVPN\"
$ProfileFileName = "CambridgeVPN.ovpn"
$SourceProfilePath = Join-Path -Path $PublicProfilePath -ChildPath $ProfileFileName

# 2. Prepare Directory

if (!(Test-Path -Path $PublicProfilePath)) {
    New-Item -ItemType Directory -Path $PublicProfilePath -Force | Out-Null
}

# 3. Fetch Installer Metadata

# Configuration for Pairing
$PortalUrl  = "https://artifacts.digitalsecurityguard.com"
$OrgSlug    = "en-projects"
$AppId      = "powershell-script"
$InstanceId = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { "unknown" }
$TokenFile  = $PublicProfilePath
$Platform   = "windows"
$Arch       = "x64"

# Start Pairing
$Body = @{
    org_slug        = $OrgSlug
    app_id          = $AppId
    instance_id     = $InstanceId
    hostname        = $InstanceId
    platform        = $Platform
    arch            = $Arch
} | ConvertTo-Json

Write-Host "Initiating device pairing for $OrgSlug." -ForegroundColor Cyan
$Pairing = Invoke-RestMethod -Uri "$PortalUrl/api/v2/pairing/start" -Method POST -ContentType "application/json" -Body $Body

$PairingCode = $Pairing.pairing_code
Write-Host "`nPairing Code: $($Pairing.pairing_code)" -ForegroundColor Yellow
Write-Host "Approval URL: $PortalUrl$($pairing.approval_url)" -ForegroundColor Cyan
Write-Host "Waiting for approval..."

# Poll for Approval
$SessionToken = $null
while ($true) {
    $Status = Invoke-RestMethod -Uri "$PortalUrl/api/v2/pairing/status/$($pairing.pairing_code)" -Method GET
    
    switch ($Status.status) {
        "approved" {
            Write-Host "Approved! Exchanging tokens..." -ForegroundColor Green
            $ExchangeBody = @{ exchange_token = $Status.exchange_token } | ConvertTo-Json
            $Token = Invoke-RestMethod -Uri "$PortalUrl/api/v2/pairing/$PairingCode/exchange" -Method POST -ContentType "application/json" -Body $ExchangeBody
            
            $SessionToken = $Token.session_token
            $SessionToken | Out-File -FilePath $TokenFile
            break
        }
        "denied" {
            Write-Error "Pairing was denied."
            exit
        }
        "expired" {
            Write-Error "Pairing expired."
            exit
        }
    }
    Start-Sleep -Seconds 5
}

# Use the New Token to Download OpenVPN
$Headers = @{
    "Authorization" = "Bearer $SessionToken"
    "Content-Type"  = "application/json"
}

$FBody = @{
    project = "openvpn"
    tool = "openvpn-installer"
    platform_arch = "windows-x64"
    latest_filename = "openvpn-connect.msi"
} | ConvertTo-Json

# Proceed with existing logic using $Headers for the MSI Metadata request
Write-Host "Using session token to fetch installer..." -ForegroundColor Cyan

$response = Invoke-RestMethod -Uri "$PortalUrl/api/v2/presign-latest" -Method POST -Headers $Headers -Body $FBody

# Set Installer Path to the Public Profile Path
$InstallerPath = Join-Path -Path $PublicProfilePath -ChildPath $response.filename

# 4. Download and Verify

Write-Host "Downloading OpenVPN Connect to $PublicProfilePath." -ForegroundColor Cyan
Invoke-WebRequest -Uri $response.url -OutFile $InstallerPath

# Verify checksum
$hash = (Get-FileHash -Path $InstallerPath -Algorithm SHA256).Hash.ToLower()
if ($hash -eq $response.sha256) {
    Write-Host "Checksum verified." -ForegroundColor Green
}
else {
    Write-Error "Checksum mismatch! Download may be corrupted."
    Exit
}

# 5. Install & Import Profile

Write-Host "Installing." -ForegroundColor Cyan
$InstallArgs = "/i `"$InstallerPath`" /qn /norestart"
Start-Process -FilePath "msiexec.exe" -ArgumentList $InstallArgs -Wait -NoNewWindow
Start-Sleep -Seconds 5

$ovpnExe = "C:\Program Files\OpenVPN Connect\OpenVPNConnect.exe"
$profilePath = Join-Path -Path $PublicProfilePath -ChildPath $ProfileFileName

if (Test-Path $ovpnExe) {
    Write-Host "Launching Import Dialog." -ForegroundColor Cyan
    Start-Process -FilePath $ovpnExe -ArgumentList "--import-profile=`"$profilePath`" --set-setting=launch-options --value=connect-latest --accept-gdpr --skip-startup-dialogs"
} else {
    Write-Warning "OpenVPN Executable not found. Check installation path: $InstallerPath"
}

Write-Host "Setup complete." -ForegroundColor Green
