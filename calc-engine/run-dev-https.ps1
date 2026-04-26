$ErrorActionPreference = "Stop"

$certDir = Join-Path $env:USERPROFILE ".office-addin-dev-certs"
$certFile = Join-Path $certDir "localhost.crt"
$keyFile = Join-Path $certDir "localhost.key"

if (-not (Test-Path $certFile) -or -not (Test-Path $keyFile)) {
    throw "Office dev certificates were not found. Run 'npm run certs' from excel-addin first."
}

python -m uvicorn app.main:app --host localhost --port 8000 --ssl-keyfile $keyFile --ssl-certfile $certFile
