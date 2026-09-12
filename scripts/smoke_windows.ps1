$ErrorActionPreference = "Stop"

uv run ruff check .
if ($LASTEXITCODE -ne 0) { throw "Ruff falhou; corrigir antes de continuar." }
$previousGuiRequirement = $env:CNAB_NOX_REQUIRE_GUI
try {
    $env:CNAB_NOX_REQUIRE_GUI = "1"
    uv run pytest
    if ($LASTEXITCODE -ne 0) { throw "Testes falharam; corrigir e retestar." }
} finally {
    $env:CNAB_NOX_REQUIRE_GUI = $previousGuiRequirement
}
$previousImportSmoke = $env:CNAB_NOX_SMOKE
try {
    $env:CNAB_NOX_SMOKE = "1"
    uv run python scripts/launch_gui.py
    if ($LASTEXITCODE -ne 0) { throw "Smoke de importacao falhou." }
} finally {
    $env:CNAB_NOX_SMOKE = $previousImportSmoke
}
Write-Host "Testes automatizados concluidos; skip POSIX nao valida ACL Windows."
Write-Host "Abra dist\GeradorCNABNOX\GeradorCNABNOX.exe e execute docs/HOMOLOGACAO_WINDOWS.md."
