$base = "https://minot.2t3.app"
$urls = @(
  "/api/waivers/counts",
  "/api/waivers/pending",
  "/api/waivers/asset/M-9999",
  "/waivers",
  "/waivers/manifest.webmanifest",
  "/waivers/card/M-9999/print?noprint=1"
)
foreach ($u in $urls) {
  try {
    $r = Invoke-WebRequest -Uri ($base + $u) -Method GET -TimeoutSec 15 -UseBasicParsing
    $body = $r.Content
    if ($body.Length -gt 200) { $body = $body.Substring(0, 200) + "..." }
    Write-Host ("OK " + $u + " -> " + $r.StatusCode + " " + $body)
  } catch {
    Write-Host ("FAIL " + $u + " -> " + $_.Exception.Message)
  }
}
