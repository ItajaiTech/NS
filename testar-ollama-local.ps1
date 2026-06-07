# Teste de conexão com Ollama local
function Test-OllamaConnection {
    param(
        [string]$Url = "http://localhost:11434",
        [string]$Model = "qwen3-coder-next:cloud"
    )
    
    Write-Host "Testando conexão com Ollama em $Url..." -ForegroundColor Cyan
    
    try {
        $response = Invoke-WebRequest -Uri "$Url/api/version" -Method GET -UseBasicParsing
        Write-Host "✅ Conectado ao Ollama!" -ForegroundColor Green
        Write-Host "Versão: $($response.Content)" -ForegroundColor Yellow
        
        # Testar geração de texto
        Write-Host "`nTestando geração de texto com o modelo '$Model'..." -ForegroundColor Cyan
        $body = @{
            model = $Model
            prompt = "Olá, responda apenas 'Olá, sou rápido e local!'"
            stream = $false
        } | ConvertTo-Json
        
        $result = Invoke-WebRequest -Uri "$Url/api/generate" -Method POST -Body $body -ContentType "application/json" -UseBasicParsing
        $output = $result.Content | ConvertFrom-Json
        Write-Host "Resultado: $($output.response)" -ForegroundColor Green
        
    } catch {
        Write-Host "❌ Erro: $_" -ForegroundColor Red
    }
}

# Executar teste
Test-OllamaConnection
