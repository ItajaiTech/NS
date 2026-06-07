# Configuração de Extensão VS Code para Ollama Local
# Execute este script como administrador para configurar

Write-Host "=== Configurando VS Code para Ollama Local ===" -ForegroundColor Cyan
Write-Host ""

# Verificar se VS Code está instalado
$codePath = Get-Command code -ErrorAction SilentlyContinue
if ($null -eq $codePath) {
    Write-Host "❌ VS Code não encontrado no PATH" -ForegroundColor Red
    Write-Host "Instale o VS Code em: https://code.visualstudio.com/" -ForegroundColor Yellow
    exit 1
}
Write-Host "✅ VS Code encontrado" -ForegroundColor Green

# Verificar Ollama
$ollamaPath = Get-Command ollama -ErrorAction SilentlyContinue
if ($null -eq $ollamaPath) {
    Write-Host "❌ Ollama não encontrado" -ForegroundColor Red
    exit 1
}
Write-Host "✅ Ollama encontrado" -ForegroundColor Green

# Criar diretório de configuração
$configDir = "$env:APPDATA\Code\User\globalStorage\llm.ai-assistant"
if (-not (Test-Path $configDir)) {
    New-Item -Path $configDir -ItemType Directory -Force | Out-Null
    Write-Host "✅ Diretório de configuração criado" -ForegroundColor Green
}

# Criar arquivo de configuração
$configFile = "$configDir\llm-config.json"
$configContent = @{
    "provider" = "ollama"
    "url" = "http://localhost:11434"
    "model" = "qwen3-coder-next:cloud"
} | ConvertTo-Json -Depth 3

$configContent | Set-Content -Path $configFile -Encoding UTF8
Write-Host "✅ Configuração gravada em: $configFile" -ForegroundColor Green

Write-Host ""
Write-Host "=== Instruções ===" -ForegroundColor Cyan
Write-Host "1. Abra VS Code" -ForegroundColor White
Write-Host "2. Pressione Ctrl+Shift+P para abror o palette" -ForegroundColor White
Write-Host "3. Digite 'LLM: Configure' e selecione 'LLM: Configure Provider'" -ForegroundColor White
Write-Host "4. Escolha 'ollama'" -ForegroundColor White
Write-Host "5. Configure a URL como 'http://localhost:11434'" -ForegroundColor White
Write-Host ""
Write-Host "Ou instale manualmente a extensão 'LLM: AI Assistant' na VS Code" -ForegroundColor Yellow
Write-Host "Extension ID: llm.ai-assistant" -ForegroundColor Yellow
Write-Host ""
Write-Host "Configuration completed successfully!" -ForegroundColor Green
