<#
.SYNOPSIS
    Cria um ambiente ficticio (homologacao x producao) para testar o Deploy Gate
    sem precisar de nenhum servidor real.

.DESCRIPTION
    Gera em demo\ambiente\:
      homologacao\   binarios "testados"
      producao\      binarios em uso
      saida\         onde relatorios, logs e pastas de conferencia sao gravados
    e um config.psd1 apontando para elas.

    Situacoes criadas (uma de cada, para o relatorio mostrar todos os casos):
      - arquivo identico nos dois lados
      - arquivo novo que ainda nao foi para producao
      - arquivo recompilado na homologacao (mais novo que a producao)
      - arquivo alterado direto na producao (alarme)
      - arquivo que so existe na producao
      - dois lotes de compilacao separados por algumas horas

.EXAMPLE
    .\demo\New-DemoEnvironment.ps1
    .\src\DeployGate.ps1 -Config .\demo\ambiente\config.psd1
#>
[CmdletBinding()]
param(
    [string] $Raiz = (Join-Path $PSScriptRoot "ambiente"),
    [switch] $Recriar
)

$ErrorActionPreference = "Stop"

if (Test-Path -LiteralPath $Raiz) {
    if (-not $Recriar) {
        Write-Host "Ja existe: $Raiz  (use -Recriar para apagar e gerar de novo)" -ForegroundColor Yellow
        exit 0
    }
    Remove-Item -LiteralPath $Raiz -Recurse -Force
}

$homolog = Join-Path $Raiz "homologacao"
$prod    = Join-Path $Raiz "producao"
$saida   = Join-Path $Raiz "saida"
foreach ($d in @($homolog, $prod, $saida)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }

function New-Binario {
    param([string]$Pasta, [string]$Nome, [string]$Conteudo, [datetime]$Data)
    $caminho = Join-Path $Pasta $Nome
    $dir = Split-Path -Parent $caminho
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [IO.File]::WriteAllText($caminho, $Conteudo)
    (Get-Item -LiteralPath $caminho).LastWriteTime = $Data
}

$base     = (Get-Date).Date.AddDays(-3).AddHours(9)   # ha 3 dias, 09:00
$lote1    = $base                                      # lote da manha
$lote2    = $base.AddHours(5)                          # lote da tarde
$antigo   = $base.AddDays(-30)

$sep = [IO.Path]::DirectorySeparatorChar

# identicos
foreach ($n in @("core.dll", "util.dll", "relatorios${sep}base.dll")) {
    New-Binario $homolog $n "v1 $n" $antigo
    New-Binario $prod    $n "v1 $n" $antigo
}

# lote 1: tres arquivos compilados juntos de manha
New-Binario $homolog "vendas.dll"                 "v2 vendas"   $lote1
New-Binario $prod    "vendas.dll"                 "v1 vendas"   $antigo
New-Binario $homolog "estoque.dll"                "v2 estoque"  $lote1.AddMinutes(3)
New-Binario $prod    "estoque.dll"                "v1 estoque"  $antigo
New-Binario $homolog "relatorios${sep}novo_painel.dll" "v1 painel" $lote1.AddMinutes(6)

# lote 2: compilado a tarde
New-Binario $homolog "financeiro.exe"             "v3 financeiro" $lote2
New-Binario $prod    "financeiro.exe"             "v2 financeiro" $antigo
New-Binario $homolog "integracao${sep}api.dll"    "v1 api"        $lote2.AddMinutes(2)

# alarme: alguem mexeu direto na producao depois da homologacao
New-Binario $homolog "fiscal.dll"                 "v4 fiscal"         $antigo
New-Binario $prod    "fiscal.dll"                 "v4 fiscal hotfix"  $base.AddDays(1)

# so na producao (orfao)
New-Binario $prod    "legado.ocx"                 "v1 legado"   $antigo

# lixo que deve ser ignorado
New-Binario $homolog "build.log"                  "log qualquer" $lote1
New-Binario $homolog "~temp.tmp"                  "temp"         $lote1

$cfg = @"
@{
    Origem      = '$homolog'
    Destino     = '$prod'
    PastaSaida  = '$saida'
    MinutosLote = 15
}
"@
[IO.File]::WriteAllText((Join-Path $Raiz "config.psd1"), $cfg)

Write-Host "Ambiente de demonstracao criado em: $Raiz" -ForegroundColor Green
Write-Host ""
Write-Host "Proximo passo:" -ForegroundColor Cyan
Write-Host "  .\src\DeployGate.ps1 -Config `"$(Join-Path $Raiz 'config.psd1')`"" -ForegroundColor Cyan
