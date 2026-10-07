<#
.SYNOPSIS
    Testes de ponta a ponta do Deploy Gate, sem dependencias externas.

.DESCRIPTION
    Cada teste cria um ambiente ficticio novo (demo\New-DemoEnvironment.ps1),
    roda o DeployGate.ps1 como processo separado, do mesmo jeito que os .bat
    fazem, e confere arquivos, logs e pastas geradas.

    Uso:  pwsh -File tests/Run-Tests.ps1
    Sai com codigo 1 se algum teste falhar (para uso em CI).
#>
$ErrorActionPreference = "Stop"

$raizRepo = Split-Path -Parent $PSScriptRoot
$script   = Join-Path $raizRepo "src/DeployGate.ps1"
$geraDemo = Join-Path $raizRepo "demo/New-DemoEnvironment.ps1"
$exe      = (Get-Process -Id $PID).Path          # pwsh ou powershell.exe, o que estiver rodando
$tmpBase  = Join-Path ([IO.Path]::GetTempPath()) ("deploygate-tests-" + [guid]::NewGuid().ToString("N").Substring(0,8))

$script:falhas = 0
$script:total  = 0

function Invoke-Gate {
    param([string[]]$Argumentos)
    $saida = & $exe -NoProfile -ExecutionPolicy Bypass -File $script @Argumentos 2>&1 | Out-String
    return [pscustomobject]@{ Codigo = $LASTEXITCODE; Texto = $saida }
}

function New-Ambiente {
    $raiz = Join-Path $tmpBase ([guid]::NewGuid().ToString("N").Substring(0,8))
    & $exe -NoProfile -ExecutionPolicy Bypass -File $geraDemo -Raiz $raiz | Out-Null
    return [pscustomobject]@{
        Raiz    = $raiz
        Config  = Join-Path $raiz "config.psd1"
        Homolog = Join-Path $raiz "homologacao"
        Prod    = Join-Path $raiz "producao"
        Saida   = Join-Path $raiz "saida"
    }
}

function Get-Sha1([string]$Caminho) { (Get-FileHash -LiteralPath $Caminho -Algorithm SHA1).Hash }

function New-Decisoes {
    # monta um decisoes.csv igual ao baixado pelo HTML
    param($Amb, [hashtable]$Decisoes)
    $linhas = @("Arquivo;Status;VersaoOrigem;HashOrigem;DECISAO;CHAMADO;OBSERVACAO")
    foreach ($k in $Decisoes.Keys) {
        $h = Get-Sha1 (Join-Path $Amb.Homolog $k)
        $linhas += "$k;?;-;$h;$($Decisoes[$k]);CH-1;"
    }
    $arq = Join-Path $Amb.Raiz "decisoes.csv"
    [IO.File]::WriteAllLines($arq, $linhas)
    return $arq
}

function Assert-True {
    param([bool]$Condicao, [string]$Mensagem)
    if (-not $Condicao) { throw $Mensagem }
}

function Test-Caso {
    param([string]$Nome, [scriptblock]$Corpo)
    $script:total++
    try {
        & $Corpo
        Write-Host "  [ok]    $Nome" -ForegroundColor Green
    } catch {
        $script:falhas++
        Write-Host "  [FALHA] $Nome" -ForegroundColor Red
        Write-Host "          $($_.Exception.Message)" -ForegroundColor Red
    }
}

$sep = [IO.Path]::DirectorySeparatorChar
Write-Host "Deploy Gate - testes ($exe)" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------- comparacao

Test-Caso "classifica cada situacao corretamente" {
    $a = New-Ambiente
    $r = Invoke-Gate @("-Config", $a.Config, "-Silencioso", "-Diag", "-SemExcel")
    Assert-True ($r.Codigo -eq 0) "codigo de saida $($r.Codigo)`n$($r.Texto)"
    $diag = Import-Csv (Get-ChildItem (Join-Path $a.Saida "diagnostico_*.csv")).FullName -Delimiter ';'
    $st = @{}; foreach ($l in $diag) { $st[$l.Arquivo.Replace('\','/')] = $l.Status }
    Assert-True ($st["core.dll"]                   -eq "IGUAL")             "core.dll: $($st['core.dll'])"
    Assert-True ($st["vendas.dll"]                 -eq "ORIGEM_MAIS_NOVA")  "vendas.dll: $($st['vendas.dll'])"
    Assert-True ($st["relatorios/novo_painel.dll"] -eq "FALTA_NO_DESTINO")  "novo_painel: $($st['relatorios/novo_painel.dll'])"
    Assert-True ($st["fiscal.dll"]                 -eq "DESTINO_MAIS_NOVO") "fiscal.dll: $($st['fiscal.dll'])"
    Assert-True ($st["legado.ocx"]                 -eq "SO_NO_DESTINO")     "legado.ocx: $($st['legado.ocx'])"
}

Test-Caso "ignora arquivos temporarios e logs" {
    $a = New-Ambiente
    Invoke-Gate @("-Config", $a.Config, "-Silencioso", "-Diag", "-SemExcel") | Out-Null
    $diag = Import-Csv (Get-ChildItem (Join-Path $a.Saida "diagnostico_*.csv")).FullName -Delimiter ';'
    Assert-True (-not ($diag.Arquivo -contains "build.log"))  "build.log entrou na comparacao"
    Assert-True (-not ($diag.Arquivo -contains "~temp.tmp"))  "~temp.tmp entrou na comparacao"
}

Test-Caso "agrupa pendencias em lotes de compilacao" {
    $a = New-Ambiente
    $r = Invoke-Gate @("-Config", $a.Config, "-Silencioso", "-SemExcel")
    Assert-True ($r.Texto -match "Aguardando decisao : 5 em 2 lote") "esperava 5 pendencias em 2 lotes`n$($r.Texto)"
    $plan = Import-Csv (Get-ChildItem (Join-Path $a.Saida "decisoes_planilha_*.csv")).FullName -Delimiter ';'
    Assert-True (@($plan.LOTE | Sort-Object -Unique).Count -eq 2) "planilha nao tem 2 lotes"
}

Test-Caso "sem configuracao, falha com mensagem clara" {
    $r = Invoke-Gate @("-Silencioso", "-Config", (Join-Path $tmpBase "nao-existe.psd1"))
    Assert-True ($r.Codigo -ne 0) "deveria falhar"
    Assert-True ($r.Texto -match "nao encontrado") "mensagem inesperada: $($r.Texto)"
}

# ---------------------------------------------------------------- separacao

Test-Caso "separa so os autorizados e registra quem esta em teste" {
    $a = New-Ambiente
    $dec = New-Decisoes $a @{ "vendas.dll" = "AUTORIZADO"; "financeiro.exe" = "EM TESTE" }
    $r = Invoke-Gate @("-Config", $a.Config, "-Aplicar", $dec, "-ParaStaging", "-Responsavel", "Teste")
    Assert-True ($r.Codigo -eq 0) "codigo $($r.Codigo)`n$($r.Texto)"
    $st = (Get-ChildItem (Join-Path $a.Saida "publicar_*") -Directory)[0].FullName
    Assert-True (Test-Path (Join-Path $st "vendas.dll"))          "vendas.dll nao foi separado"
    Assert-True (-not (Test-Path (Join-Path $st "financeiro.exe"))) "financeiro.exe nao deveria ser separado"
    $obs = Import-Csv (Join-Path $a.Saida "em_observacao.csv") -Delimiter ';'
    Assert-True ($obs.Arquivo -contains "financeiro.exe") "financeiro.exe nao foi para observacao"
    # producao intacta
    Assert-True ((Get-Content (Join-Path $a.Prod "vendas.dll") -Raw) -eq "v1 vendas") "producao foi alterada na separacao"
}

Test-Caso "recusa arquivo recompilado depois da decisao" {
    $a = New-Ambiente
    $dec = New-Decisoes $a @{ "vendas.dll" = "AUTORIZADO" }
    [IO.File]::WriteAllText((Join-Path $a.Homolog "vendas.dll"), "v3 vendas recompilado")
    $r = Invoke-Gate @("-Config", $a.Config, "-Aplicar", $dec, "-ParaStaging", "-Responsavel", "Teste")
    Assert-True ($r.Texto -match "RECUSADO\s+vendas.dll") "nao recusou`n$($r.Texto)"
}

Test-Caso "recusa sobrescrever producao mais nova, mesmo com CSV editado a mao" {
    $a = New-Ambiente
    $dec = New-Decisoes $a @{ "fiscal.dll" = "AUTORIZADO" }
    $r = Invoke-Gate @("-Config", $a.Config, "-Aplicar", $dec, "-Responsavel", "Teste")
    Assert-True ($r.Texto -match "RECUSADO\s+fiscal.dll") "nao recusou o alarme`n$($r.Texto)"
    Assert-True ((Get-Content (Join-Path $a.Prod "fiscal.dll") -Raw) -eq "v4 fiscal hotfix") "hotfix da producao foi sobrescrito"
}

Test-Caso "arquivo em teste sai da fila e volta quando recompilado" {
    $a = New-Ambiente
    $dec = New-Decisoes $a @{ "financeiro.exe" = "EM TESTE" }
    Invoke-Gate @("-Config", $a.Config, "-Aplicar", $dec, "-ParaStaging", "-Responsavel", "Teste") | Out-Null
    $r1 = Invoke-Gate @("-Config", $a.Config, "-Silencioso", "-SemExcel")
    Assert-True ($r1.Texto -match "Em observacao\s+: 1") "nao entrou em observacao`n$($r1.Texto)"
    [IO.File]::WriteAllText((Join-Path $a.Homolog "financeiro.exe"), "v4 financeiro")
    $r2 = Invoke-Gate @("-Config", $a.Config, "-Silencioso", "-SemExcel")
    Assert-True ($r2.Texto -match "Em observacao\s+: 0") "nao saiu da observacao apos recompilar`n$($r2.Texto)"
}

# ---------------------------------------------------------------- publicacao

Test-Caso "publica o conferido, recusa adulterado e ignora intruso" {
    $a = New-Ambiente
    $dec = New-Decisoes $a @{ "vendas.dll" = "AUTORIZADO"; "estoque.dll" = "AUTORIZADO" }
    Invoke-Gate @("-Config", $a.Config, "-Aplicar", $dec, "-ParaStaging", "-Responsavel", "Ana") | Out-Null
    $st = (Get-ChildItem (Join-Path $a.Saida "publicar_*") -Directory)[0].FullName

    [IO.File]::WriteAllText((Join-Path $st "estoque.dll"), "adulterado")
    [IO.File]::WriteAllText((Join-Path $st "intruso.dll"), "ninguem autorizou")

    $r = Invoke-Gate @("-Config", $a.Config, "-PublicarStaging", $st, "-Responsavel", "Bruno")
    Assert-True ($r.Codigo -eq 0) "codigo $($r.Codigo)`n$($r.Texto)"
    Assert-True ((Get-Content (Join-Path $a.Prod "vendas.dll") -Raw)  -eq "v2 vendas")  "vendas.dll nao foi publicado"
    Assert-True ((Get-Content (Join-Path $a.Prod "estoque.dll") -Raw) -eq "v1 estoque") "estoque.dll adulterado chegou na producao"
    Assert-True (-not (Test-Path (Join-Path $a.Prod "intruso.dll")))                    "intruso chegou na producao"
    Assert-True (Test-Path (Join-Path $st "_PUBLICADO.txt"))                            "pasta nao foi marcada como publicada"

    $log = Import-Csv (Join-Path $a.Saida "publicacoes.csv") -Delimiter ';'
    $v = $log | Where-Object { $_.Arquivo -eq "vendas.dll" }
    Assert-True ($v.Resultado -eq "PUBLICADO")       "log sem PUBLICADO para vendas.dll"
    Assert-True ($v.Autorizador -eq "Ana")           "autorizador errado no log"
    Assert-True ($v.Detalhe -match "Bruno")          "conferente nao aparece no log"
}

Test-Caso "publicacao recusa se a producao mudou depois da separacao" {
    $a = New-Ambiente
    $dec = New-Decisoes $a @{ "vendas.dll" = "AUTORIZADO" }
    Invoke-Gate @("-Config", $a.Config, "-Aplicar", $dec, "-ParaStaging", "-Responsavel", "Ana") | Out-Null
    $st = (Get-ChildItem (Join-Path $a.Saida "publicar_*") -Directory)[0].FullName
    $p = Join-Path $a.Prod "vendas.dll"
    [IO.File]::WriteAllText($p, "hotfix urgente direto na producao")
    (Get-Item $p).LastWriteTime = (Get-Date)
    $r = Invoke-Gate @("-Config", $a.Config, "-PublicarStaging", $st, "-Responsavel", "Bruno")
    Assert-True ((Get-Content $p -Raw) -eq "hotfix urgente direto na producao") "hotfix foi sobrescrito`n$($r.Texto)"
}

Test-Caso "simulacao nao altera nada" {
    $a = New-Ambiente
    $dec = New-Decisoes $a @{ "vendas.dll" = "AUTORIZADO" }
    Invoke-Gate @("-Config", $a.Config, "-Aplicar", $dec, "-Responsavel", "Ana", "-Simular") | Out-Null
    Assert-True ((Get-Content (Join-Path $a.Prod "vendas.dll") -Raw) -eq "v1 vendas") "simulacao alterou a producao"
}

# ---------------------------------------------------------------- resumo
Remove-Item -LiteralPath $tmpBase -Recurse -Force -ErrorAction SilentlyContinue
Write-Host ""
if ($script:falhas -gt 0) {
    Write-Host "$($script:falhas) de $($script:total) teste(s) falharam." -ForegroundColor Red
    exit 1
}
Write-Host "Todos os $($script:total) testes passaram." -ForegroundColor Green
exit 0
