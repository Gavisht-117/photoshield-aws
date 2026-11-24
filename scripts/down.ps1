# Deletes everything Terraform created (stops all charges). Terraform lists what
# it will delete and asks you to type "yes" first.
# Run:  powershell -ExecutionPolicy Bypass -File scripts\down.ps1
$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$tf   = 'E:\tools\terraform\terraform.exe'
$env:TF_PLUGIN_CACHE_DIR = 'E:\tools\tf-plugin-cache'
$env:TMP  = 'E:\tools\tmp'
$env:TEMP = 'E:\tools\tmp'

Push-Location (Join-Path $root 'terraform')
try {
    & $tf destroy
    if ($LASTEXITCODE -ne 0) { throw 'terraform destroy failed or was cancelled' }
}
finally {
    Pop-Location
}
