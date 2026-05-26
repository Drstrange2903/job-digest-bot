$workflow = Join-Path $PSScriptRoot "..\.github\workflows\job-digest.yml"
$content = Get-Content -LiteralPath $workflow -Raw
$content = $content -replace 'cron: "0 \*/2 \* \* \*"', 'cron: "*/10 * * * *"'
$content = $content -replace 'Production mode: every 2 hours\.', 'Test mode: every 10 minutes. After email is confirmed, change to: 0 */2 * * *'
Set-Content -LiteralPath $workflow -Value $content -NoNewline
Write-Host "Updated GitHub Actions schedule to every 10 minutes."
