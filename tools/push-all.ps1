$projects = @(
  "projects\dku_ticket_web_chat",
  "projects\lib_ticket_api",
  "projects\dkul_systems_ticket",
  "projects\dku_ticket_core"
)
foreach ($p in $projects) {
  Write-Host "== PUSH $p =="
  Push-Location $p
  clasp push -f
  Pop-Location
}
