
# Instalação do módulo ImportExcel (executar apenas se ainda não estiver instalado)
# Install-Module -Name ImportExcel -Scope CurrentUser

Import-Module ImportExcel

# Caminho do arquivo XLSX de entrada
$xlsxPath = "lansweeper_export.xlsx"

if (!(Test-Path $xlsxPath)) {
    Write-Host "Arquivo '$xlsxPath' não encontrado." -ForegroundColor Red
    exit
}

# Lê dados da planilha original incluindo campos da consulta SQL
$rows = Import-Excel -Path $xlsxPath | Select-Object AssetID, AssetName, AssetTypename, Manufacturer, Model, Statename, PurchaseDate, SerialNumber, custom7, Comments

# Lista de status permitidos para OK durante bipagem
$statusPermitidos = @("Broken", "Active", "Reservado", "Stock", "Old", "In repair")

# Lista de status a serem considerados na verificação final
$statusRecon = @("Broken", "Reservado", "Stock", "Old", "In repair")

Write-Host "Total de equipamentos carregados: $($rows.Count)" -ForegroundColor Cyan
Write-Host "Digite ou escaneie o AssetName ou SerialNumber (ou 'sair' para encerrar):"

# Lista para armazenar log
$log = @()

# Lista para rastrear entradas já bipadas (normalizadas)
$entradasRegistradas = @()

# Função para normalizar AssetName ou SerialNumber
$normalize = { param($n) ($n -replace '^\s*NT-','' -replace '\s','').ToLower() }

while ($true) {
    $entrada = Read-Host "Entrada"
    if ($entrada -eq "sair") { break }

    $entradaLimpa = & $normalize $entrada

    # Verifica se já foi bipado
    if ($entradasRegistradas -contains $entradaLimpa) {
        Write-Host "Entrada já registrada anteriormente. Ignorando..." -ForegroundColor Yellow
        continue
    }

    # Procura equipamento por AssetName OU SerialNumber
    $equipamento = $rows | Where-Object {
        (& $normalize $_.AssetName) -eq $entradaLimpa -or (& $normalize $_.SerialNumber) -eq $entradaLimpa
    } | Select-Object -First 1

    if ($equipamento) {
        if ($statusPermitidos -contains $equipamento.Statename) {
            Write-Host "OK - $($equipamento.AssetName) | Serial: $($equipamento.SerialNumber) | Status: $($equipamento.Statename)" -ForegroundColor Green
            $resultado = "OK"
        } elseif ($statusRecon -contains $equipamento.Statename) {
            Write-Host "Encontrado - $($equipamento.AssetName) | Serial: $($equipamento.SerialNumber) | Status: $($equipamento.Statename)" -ForegroundColor Cyan
            $resultado = "Bipado (status não OK)"
        } else {
            Write-Host "Encontrado - $($equipamento.AssetName) | Serial: $($equipamento.SerialNumber) | Status: $($equipamento.Statename) (não permitido)" -ForegroundColor Yellow
            $resultado = "Não permitido"
        }

        $log += [PSCustomObject]@{
            AssetID      = $equipamento.AssetID
            AssetName    = $equipamento.AssetName
            AssetTypename= $equipamento.AssetTypename
            Manufacturer = $equipamento.Manufacturer
            Model        = $equipamento.Model
            Statename    = $equipamento.Statename
            PurchaseDate = $equipamento.PurchaseDate
            SerialNumber = $equipamento.SerialNumber
            Custom7      = $equipamento.custom7
            Comments     = $equipamento.Comments
            Resultado    = $resultado
            DataHora     = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }

        $entradasRegistradas += $entradaLimpa
    } else {
        Write-Host "Equipamento não encontrado" -ForegroundColor Red
        $log += [PSCustomObject]@{
            AssetID      = "N/A"
            AssetName    = $entrada
            AssetTypename= "N/A"
            Manufacturer = "N/A"
            Model        = "N/A"
            Statename    = "N/A"
            PurchaseDate = "N/A"
            SerialNumber = "N/A"
            Custom7      = "N/A"
            Comments     = "N/A"
            Resultado    = "Não encontrado"
            DataHora     = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }

        $entradasRegistradas += $entradaLimpa
    }
}

# Data atual para o nome do arquivo
$today = Get-Date -Format "yyyy-MM-dd"

# Exporta log para XLSX com data no nome
$xlsxLog = "stock_$today.xlsx"
$log | Export-Excel -Path $xlsxLog -AutoSize -BoldTopRow
Write-Host "Arquivo de log gerado: $xlsxLog" -ForegroundColor Cyan

# ===== Verificação final de estoque =====
$relevantesPlanilha = $rows | Where-Object { $_.Statename -in $statusRecon }

# Lista de entradas bipadas normalizadas (AssetName ou SerialNumber)
$bipadosNorm = $log | ForEach-Object {
    @(& $normalize $_.AssetName, "&" $normalize $_.SerialNumber)
} | Select-Object -Unique


# Verifica se algum ativo relevante não foi bipado
$faltando = $relevantesPlanilha | Where-Object {
    $assetNorm = & $normalize $_.AssetName
    $serialNorm = & $normalize $_.SerialNumber
    ($assetNorm -notin $bipadosNorm) -and ($serialNorm -notin $bipadosNorm)
}

if ($faltando.Count -eq 0) {
    Write-Host "Finalizado sem intercorrencia" -ForegroundColor Green
} else {
    Write-Host "Erro! favor rever o estoque" -ForegroundColor Red
}
