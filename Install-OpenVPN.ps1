param(
  [Parameter(Mandatory=$false)]
  [string]$ProfileFileName
)

if (-not $ProfileFileNam) {
  $ProfileFileNam = Read-Host "Be sure the OVPN profile is downloaded in 'C:\Data\OpenVPN\'.`nEnter the '.ovpn' profile name to continue. e.g. 'CompanyVPN.ovpn'"
    }
$PublicProfilePath = "C:\Data\OpenVPN\"
$SourceProfilePath = Join-Path -Path $PSScriptRoot -ChildPath $ProfileFileName

# 1. Check Admin

if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run as Administrator to install the software."
    Exit
}

# 2. Prepare Directory

if (!(Test-Path -Path $PublicProfilePath)) {
    New-Item -ItemType Directory -Path $PublicProfilePath -Force | Out-Null
}

# Copy the OVPN profile to the public path if it exists locally
if (Test-Path $SourceProfilePath) {
    Copy-Item -Path $SourceProfilePath -Destination $PublicProfilePath -Force
}

# 3. Fetch Installer Metadata

$token = "apt_Npg1WhCax9WmcktHOrbC2vYH_9PrrCl2jUuJGrzWsgY"
$baseUrl = "https://artifacts.digitalsecurityguard.com"

$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type" = "application/json"
}

$body = @{
    project = "openvpn"
    tool = "openvpn-installer"
    platform_arch = "windows-x64"
    latest_filename = "openvpn-connect.msi"
} | ConvertTo-Json

Write-Host "Fetching installer." -ForegroundColor Cyan
$response = Invoke-RestMethod -Uri "$baseUrl/dl/latest/openvpn/openvpn-installer/windows-x64/openvpn-connect.msi" -Method POST -Headers $headers -Body $body

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
    Write-Host "Launching Import Dialog..." -ForegroundColor Cyan
    Start-Process -FilePath $ovpnExe -ArgumentList "--import-profile=`"$profilePath`" --set-setting=launch-options --value=connect-latest --accept-gdpr --skip-startup-dialogs"
} else {
    Write-Warning "OpenVPN Executable not found. Check installation path: $InstallerPath"
}

Write-Host "Setup complete." -ForegroundColor Green