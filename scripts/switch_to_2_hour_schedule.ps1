$workflow = Join-Path $PSScriptRoot "..\.github\workflows\job-digest.yml"
$content = Get-Content -LiteralPath $workflow -Raw
$content = $content -replace 'cron: "\*/10 \* \* \* \*"', 'cron: "0 */2 * * *"'
$content = $content -replace 'Test mode: every 10 minutes\. After email is confirmed, change to: 0 \*/2 \* \* \*', 'Production mode: every 2 hours.'
Set-Content -LiteralPath $workflow -Value $content -NoNewline
Write-Host "Updated GitHub Actions schedule to every 2 hours."
