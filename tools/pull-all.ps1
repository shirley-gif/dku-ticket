$projects = @(
  "projects\dku_ticket_web_chat",
  "projects\lib_ticket_api",
  "projects\dkul_systems_ticket",
  "projects\dku_ticket_core"
)
foreach ($p in $projects) {
  Write-Host "== PULL $p =="
  Push-Location $p
  clasp pull
  Pop-Location
}
