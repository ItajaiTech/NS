# Script rápido para usar o Ollama localmente (sem VS Code)
# Uso: .\usar-ollama-rapido.ps1 "sua pergunta aqui"

 param (
    [Parameter(Mandatory=$true, Position=0)]
    [string]$Prompt
)

$Url = "http://localhost:11434"
$Model = "qwen3-coder-next:cloud"

Write-Host "Prompt: $Prompt" -ForegroundColor Cyan
Write-Host "Enviando para Ollama local..." -ForegroundColor Yellow

$body = @{
    model = $Model
    prompt = $Prompt
    stream = $false
} | ConvertTo-Json

try {
    $result = Invoke-WebRequest -Uri "$Url/api/generate" -Method POST -Body $body -ContentType "application/json" -UseBasicParsing
    $output = $result.Content | ConvertFrom-Json
    Write-Host "" -ForegroundColor Green
    Write-Host "Resposta:" -ForegroundColor Green
    Write-Host $output.response -ForegroundColor White
} catch {
    Write-Host "Erro: $_" -ForegroundColor Red
}
