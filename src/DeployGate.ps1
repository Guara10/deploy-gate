<#
.SYNOPSIS
    Deploy Gate - controle de publicacao de binarios da homologacao para a producao.

.DESCRIPTION
    Compara a pasta de HOMOLOGACAO (origem) com a pasta de PRODUCAO (destino)
    por hash SHA1, agrupa o que esta pendente por LOTE DE COMPILACAO e gera:
      - um relatorio HTML interativo (busca, filtros e botoes por arquivo)
      - uma planilha XLSX com lista suspensa colorida (ou CSV, sem Excel)

    A publicacao acontece em duas etapas, sempre com conferencia humana:
      1) os arquivos autorizados sao COPIADOS para uma pasta de conferencia
         (staging) junto com um _manifesto.csv que guarda o hash autorizado;
      2) depois de conferida, a pasta e publicada na producao. O hash de cada
         arquivo e revalidado e tudo e registrado em publicacoes.csv.

.FLUXO DE USO
    1) .\DeployGate.ps1
       -> gera relatorio_*.html e decisoes_planilha_*.xlsx (ou .csv)

    2) Quem e responsavel decide, de um dos dois jeitos:
       a) No HTML: botao verde (AUTORIZADO) ou vermelho (EM TESTE) e
          "Baixar decisoes" no rodape (gera decisoes.csv).
       b) Na planilha: coluna DECISAO.

    3) .\DeployGate.ps1 -AplicarUltima
       AUTORIZADO -> copia para a pasta de conferencia
       EM TESTE   -> nao copia; registra em em_observacao.csv

    4) .\DeployGate.ps1 -PublicarUltima
       -> revalida o hash e publica na producao

.OBSERVACAO AUTOMATICA
    Um arquivo marcado EM TESTE sai da fila de pendencias enquanto o binario
    nao mudar. Se for recompilado, o hash muda, ele sai da observacao sozinho e
    volta para a fila. O relatorio mostra ha quantos dias cada um esta parado.

.PROTECOES
    - Comparacao por HASH (SHA1). Data serve para AGRUPAR, nunca para decidir.
    - Ao aplicar, o hash e reconferido: build recompilado apos a autorizacao e RECUSADO.
    - Na publicacao, o hash e conferido de novo contra o manifesto.
    - Arquivos com a producao mais nova ou divergente da homologacao NUNCA entram
      na lista de autorizacao. Sao alarme e exigem analise manual.
    - Arquivos largados na pasta de conferencia fora do manifesto sao ignorados.

.CONFIGURACAO
    Copie config.example.psd1 para config.psd1 (mesma pasta do script) e ajuste
    os caminhos. Parametros passados na linha de comando tem prioridade.

.EXAMPLE
    .\DeployGate.ps1
.EXAMPLE
    .\DeployGate.ps1 -Origem "D:\homolog\bin" -Destino "D:\prod\bin" -PastaSaida "D:\deploy"
.EXAMPLE
    .\DeployGate.ps1 -DataInicio "01/07/2026" -DataFim "31/07/2026" -MinutosLote 30
.EXAMPLE
    .\DeployGate.ps1 -Aplicar ".\decisoes.csv" -ParaStaging -Simular
.EXAMPLE
    .\DeployGate.ps1 -PublicarStaging "D:\deploy\publicar_20260831_1530" -Responsavel "Fulano"
#>

[CmdletBinding()]
param(
    [string]   $Origem        = "",          # homologacao
    [string]   $Destino       = "",          # producao
    [string]   $PastaSaida    = "",          # relatorios, logs e pastas de conferencia
    [string]   $LogPublicacao = "",          # padrao: <PastaSaida>\publicacoes.csv
    [string]   $ArqObservacao = "",          # padrao: <PastaSaida>\em_observacao.csv
    [string]   $Config        = "",          # padrao: config.psd1 ao lado do script
    [string[]] $Extensoes     = @(),
    [string[]] $Ignorar       = @(),
    [int]      $MinutosLote   = 0,
    [string]   $DataInicio    = "",
    [string]   $DataFim       = "",
    [string]   $Aplicar       = "",
    [switch]   $AplicarUltima,               # acha sozinho o decisoes.csv mais recente baixado
    [switch]   $PublicarUltima,              # acha sozinho a ultima pasta de conferencia nao publicada
    [switch]   $ParaStaging,                 # com -Aplicar: separa numa pasta de conferencia em vez de ir direto ao destino
    [string]   $PastaStaging  = "",          # opcional: caminho da pasta de conferencia
    [string]   $PublicarStaging = "",        # 2a etapa: publica uma pasta de conferencia ja conferida
    [string]   $Responsavel   = "",          # quem autorizou/conferiu (evita a pergunta interativa)
    [switch]   $Confirmar,                   # responde "sim" as confirmacoes (uso em automacao)
    [switch]   $Simular,
    [switch]   $SemExcel,
    [switch]   $Diag,                        # gera CSV com tudo que foi lido, para diagnostico
    [switch]   $Silencioso
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Web

# ================================================================ CONFIGURACAO

$Usuario = [Environment]::UserName
$Maquina = [Environment]::MachineName

if ($Config -eq "") { $Config = Join-Path $PSScriptRoot "config.psd1" }
$cfg = @{}
if (Test-Path -LiteralPath $Config) {
    $cfg = Import-PowerShellDataFile -LiteralPath $Config
} elseif ($PSBoundParameters.ContainsKey("Config")) {
    Write-Host "ERRO: arquivo de configuracao nao encontrado -> $Config" -ForegroundColor Red; exit 1
}

function Get-Valor {
    param($Atual, [string]$Chave, $Padrao)
    if ($Atual -is [array]) { if ($Atual.Count -gt 0) { return $Atual } }
    elseif ($Atual -and "$Atual" -ne "" -and "$Atual" -ne "0") { return $Atual }
    if ($cfg.ContainsKey($Chave) -and $null -ne $cfg[$Chave]) { return $cfg[$Chave] }
    return $Padrao
}

$Origem        = Get-Valor $Origem        "Origem"        ""
$Destino       = Get-Valor $Destino       "Destino"       ""
$PastaSaida    = Get-Valor $PastaSaida    "PastaSaida"    ""
$Extensoes     = @(Get-Valor $Extensoes   "Extensoes"     @())
$Ignorar       = @(Get-Valor $Ignorar     "Ignorar"       @("*.tmp","*.bak","*.log","~*","*.pdb"))
$MinutosLote   = [int](Get-Valor $MinutosLote "MinutosLote" 15)

foreach ($par in @(@("Origem",$Origem), @("Destino",$Destino), @("PastaSaida",$PastaSaida))) {
    if ([string]::IsNullOrWhiteSpace($par[1])) {
        Write-Host "ERRO: '$($par[0])' nao definido." -ForegroundColor Red
        Write-Host "Copie config.example.psd1 para config.psd1 e ajuste os caminhos," -ForegroundColor Red
        Write-Host "ou passe -Origem, -Destino e -PastaSaida na linha de comando." -ForegroundColor Red
        exit 1
    }
}
# caminhos absolutos: o resto do script recorta caminhos relativos a partir deles
function Resolve-Caminho([string]$c) {
    $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($c).TrimEnd('\','/')
}
$Origem     = Resolve-Caminho $Origem
$Destino    = Resolve-Caminho $Destino
$PastaSaida = Resolve-Caminho $PastaSaida

$LogPublicacao = Get-Valor $LogPublicacao "LogPublicacao" (Join-Path $PastaSaida "publicacoes.csv")
$ArqObservacao = Get-Valor $ArqObservacao "ArqObservacao" (Join-Path $PastaSaida "em_observacao.csv")

function Read-Confirmacao {
    param([string]$Pergunta, [string]$Esperado = '^[Ss]')
    if ($Confirmar) { Write-Host "$Pergunta -> sim (-Confirmar)" -ForegroundColor DarkGray; return $true }
    return ((Read-Host $Pergunta) -match $Esperado)
}

function Read-Responsavel {
    param([string]$Pergunta)
    if ($Responsavel -ne "") { return $Responsavel }
    return (Read-Host $Pergunta)
}

# ==================================================================== COMUNS

function Test-Caminhos {
    foreach ($p in @($Origem, $Destino)) {
        if (-not (Test-Path -LiteralPath $p)) {
            Write-Host "ERRO: caminho nao encontrado -> $p" -ForegroundColor Red
            Write-Host "Verifique se o disco esta mapeado para ESTE usuario." -ForegroundColor Red
            exit 1
        }
    }
}

function ConvertTo-Data {
    param([string]$Texto,[switch]$FimDoDia)
    if ([string]::IsNullOrWhiteSpace($Texto)) { return $null }
    $d = [datetime]::MinValue
    $ptbr = [Globalization.CultureInfo]::GetCultureInfo("pt-BR")
    $inv  = [Globalization.CultureInfo]::InvariantCulture
    $est  = [Globalization.DateTimeStyles]::None
    if (-not [datetime]::TryParse($Texto,$ptbr,$est,[ref]$d)) {
        if (-not [datetime]::TryParse($Texto,$inv,$est,[ref]$d)) {
            Write-Host "ERRO: data invalida -> '$Texto'. Use dd/MM/aaaa." -ForegroundColor Red; exit 1
        }
    }
    if ($FimDoDia -and $d.TimeOfDay.TotalSeconds -eq 0) { $d = $d.Date.AddDays(1).AddSeconds(-1) }
    return $d
}

function Test-NoPeriodo {
    param($Item,$Ini,$Fim)
    if ($null -eq $Ini -and $null -eq $Fim) { return $true }
    $ref = if ($Item.DataOrigemRaw) { $Item.DataOrigemRaw } else { $Item.DataDestinoRaw }
    if ($null -eq $ref) { return $false }
    if ($Ini -and $ref -lt $Ini) { return $false }
    if ($Fim -and $ref -gt $Fim) { return $false }
    return $true
}

function Get-Inventario {
    param([string]$Raiz)
    $itens = Get-ChildItem -LiteralPath $Raiz -Recurse -File -ErrorAction SilentlyContinue
    if ($Extensoes.Count -gt 0) { $itens = $itens | Where-Object { $Extensoes -contains $_.Extension.ToLower() } }
    foreach ($padrao in $Ignorar) { $itens = $itens | Where-Object { $_.Name -notlike $padrao } }

    $mapa = @{}
    foreach ($f in $itens) {
        $rel = $f.FullName.Substring($Raiz.Length).TrimStart('\','/')
        try   { $hash = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA1).Hash }
        catch { $hash = "ERRO_LEITURA" }
        $versao = $null
        if ($f.Extension -match '^\.(exe|dll|ocx)$') { try { $versao = $f.VersionInfo.FileVersion } catch { } }
        $mapa[$rel.ToLower()] = [pscustomobject]@{
            Relativo=$rel; Caminho=$f.FullName; Tamanho=$f.Length
            Data=$f.LastWriteTime; Hash=$hash; Versao=$versao
        }
    }
    return $mapa
}

function Compare-Pastas {
    Write-Host "Lendo $Origem ..."  -ForegroundColor Cyan
    $mapaOrigem = Get-Inventario -Raiz $Origem
    Write-Host "Lendo $Destino ..." -ForegroundColor Cyan
    $mapaDestino = Get-Inventario -Raiz $Destino

    $lista  = New-Object System.Collections.Generic.List[object]
    $chaves = ($mapaOrigem.Keys + $mapaDestino.Keys) | Sort-Object -Unique

    foreach ($k in $chaves) {
        $m = $mapaOrigem[$k]; $n = $mapaDestino[$k]
        if     ($m -and -not $n)     { $status="FALTA_NO_DESTINO"; $sev=1 }
        elseif ($n -and -not $m)     { $status="SO_NO_DESTINO";    $sev=3 }
        elseif ($m.Hash -eq $n.Hash) { $status="IGUAL";      $sev=9 }
        else {
            $delta = ($m.Data - $n.Data).TotalMinutes
            if     ($delta -gt  1) { $status="ORIGEM_MAIS_NOVA";     $sev=2 }
            elseif ($delta -lt -1) { $status="DESTINO_MAIS_NOVO";     $sev=0 }
            else                   { $status="DIVERGENTE_MESMA_DATA"; $sev=0 }
        }
        $lista.Add([pscustomobject]@{
            Arquivo    = if ($m) { $m.Relativo } else { $n.Relativo }
            Status     = $status
            Sev        = $sev
            DataOrigemRaw   = if ($m) { $m.Data } else { $null }
            DataDestinoRaw   = if ($n) { $n.Data } else { $null }
            DataOrigem      = if ($m) { $m.Data.ToString("dd/MM/yyyy HH:mm:ss") } else { "-" }
            DataDestino      = if ($n) { $n.Data.ToString("dd/MM/yyyy HH:mm:ss") } else { "-" }
            VersaoOrigem    = if ($m -and $m.Versao) { $m.Versao } else { "-" }
            VersaoDestino    = if ($n -and $n.Versao) { $n.Versao } else { "-" }
            TamanhoOrigem   = if ($m) { $m.Tamanho } else { $null }
            HashOrigem      = if ($m) { $m.Hash } else { "" }
            OrigemFull = if ($m) { $m.Caminho } else { $null }
            Lote       = ""
            DiasObs    = $null
            ObsQuem    = ""
            ObsChamado = ""
        })
    }
    return $lista
}

function Group-Lotes {
    param($Pendentes,[int]$Minutos)
    $ord = @($Pendentes | Sort-Object DataOrigemRaw)
    $lotes = New-Object System.Collections.Generic.List[object]
    $atual = $null
    foreach ($item in $ord) {
        if ($null -eq $atual -or ($item.DataOrigemRaw - $atual.Fim).TotalMinutes -gt $Minutos) {
            $atual = [pscustomobject]@{
                Numero=$lotes.Count+1; Inicio=$item.DataOrigemRaw; Fim=$item.DataOrigemRaw
                Itens=(New-Object System.Collections.Generic.List[object])
            }
            $lotes.Add($atual)
        }
        $atual.Fim = $item.DataOrigemRaw
        $atual.Itens.Add($item)
    }
    foreach ($l in $lotes) {
        $rot = "LOTE {0:00} - {1}" -f $l.Numero, $l.Inicio.ToString("dd/MM/yyyy HH:mm")
        foreach ($i in $l.Itens) { $i.Lote = $rot }
    }
    return $lotes
}

function Import-CsvAuto {
    param([string]$Caminho)
    $fs = [IO.File]::OpenRead($Caminho)
    $b = New-Object byte[] 3
    [void]$fs.Read($b,0,3); $fs.Close()
    $temBom = ($b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)
    $enc = if ($temBom) { "UTF8" } else { "Default" }
    return @(Import-Csv -LiteralPath $Caminho -Delimiter ';' -Encoding $enc)
}

function Import-Planilha {
    param([string]$Caminho)
    $ext = [IO.Path]::GetExtension($Caminho).ToLower()
    if ($ext -in @(".csv",".txt")) { return Import-CsvAuto $Caminho }

    $excel = $null; $wb = $null
    try   { $excel = New-Object -ComObject Excel.Application }
    catch { Write-Host "ERRO: Excel nao instalado nesta maquina; nao da para ler $ext." -ForegroundColor Red; exit 1 }

    try {
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($Caminho, 0, $true)
        $ws = $wb.Worksheets.Item(1)
        $ur = $ws.UsedRange
        $nl = $ur.Rows.Count; $nc = $ur.Columns.Count
        $val = $ur.Value2

        $cab = @(); for ($c=1; $c -le $nc; $c++) { $cab += ("" + $val.GetValue(1,$c)).Trim() }
        $saida = New-Object System.Collections.Generic.List[object]
        for ($r=2; $r -le $nl; $r++) {
            $o = New-Object psobject
            for ($c=1; $c -le $nc; $c++) {
                if ($cab[$c-1] -ne "") {
                    $o | Add-Member -NotePropertyName $cab[$c-1] -NotePropertyValue ("" + $val.GetValue($r,$c)).Trim() -Force
                }
            }
            if (("" + $o.Arquivo).Trim() -ne "") { $saida.Add($o) }
        }
        return @($saida)
    }
    finally {
        if ($wb)    { $wb.Close($false) }
        if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    }
}

function Format-LinhaCsv {
    # campos entre aspas: texto livre (motivo, mensagem de erro) pode conter ';'
    param([object[]]$Campos)
    ($Campos | ForEach-Object { '"' + ("" + $_).Replace('"','""') + '"' }) -join ';'
}

function Write-LogPublicacao {
    param([string]$Arquivo,[string]$Status,[string]$Versao,[string]$Chamado,
          [string]$Autorizador,[string]$Resultado,[string]$Detalhe)
    $dir = Split-Path -Parent $LogPublicacao
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $LogPublicacao)) {
        "DataHora;Publicou;Maquina;Arquivo;Status;Versao;Chamado;Autorizador;Resultado;Detalhe" |
            Out-File -LiteralPath $LogPublicacao -Encoding UTF8
    }
    Format-LinhaCsv @((Get-Date -Format "dd/MM/yyyy HH:mm:ss"),
        $Usuario,$Maquina,$Arquivo,$Status,$Versao,$Chamado,$Autorizador,$Resultado,$Detalhe) |
        Out-File -LiteralPath $LogPublicacao -Append -Encoding UTF8
}

function Get-Observacoes {
    $mapa = @{}
    if (-not (Test-Path -LiteralPath $ArqObservacao)) { return $mapa }
    foreach ($l in (Import-CsvAuto $ArqObservacao)) {
        $k = ("" + $l.Arquivo).ToLower() + "|" + ("" + $l.HashOrigem).ToUpper()
        $mapa[$k] = $l
    }
    return $mapa
}

function Add-Observacao {
    param([string]$Arquivo,[string]$HashOrigem,[string]$Quem,[string]$Chamado,[string]$Motivo)
    $dir = Split-Path -Parent $ArqObservacao
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $ArqObservacao)) {
        "Arquivo;HashOrigem;DataDecisao;Quem;Chamado;Motivo" | Out-File -LiteralPath $ArqObservacao -Encoding UTF8
    }
    $ja = Get-Observacoes
    if ($ja.ContainsKey($Arquivo.ToLower() + "|" + $HashOrigem.ToUpper())) { return }   # mantem a data original
    Format-LinhaCsv @($Arquivo,$HashOrigem,(Get-Date -Format "dd/MM/yyyy HH:mm:ss"),$Quem,$Chamado,$Motivo) |
        Out-File -LiteralPath $ArqObservacao -Append -Encoding UTF8
}

function Test-Alarme {
    # Verdadeiro quando a producao tem um arquivo DIFERENTE e mais novo (ou da
    # mesma data) que o candidato. Publicar por cima apagaria uma alteracao que
    # so existe na producao. A checagem e refeita aqui, e nao so no HTML, porque
    # o CSV de decisoes pode ser editado a mao.
    param([string]$Candidato, [string]$NaProducao)
    if (-not (Test-Path -LiteralPath $NaProducao)) { return $false }
    $c = Get-Item -LiteralPath $Candidato
    $p = Get-Item -LiteralPath $NaProducao
    if ((Get-FileHash -LiteralPath $c.FullName -Algorithm SHA1).Hash -eq
        (Get-FileHash -LiteralPath $p.FullName -Algorithm SHA1).Hash) { return $false }
    return (($c.LastWriteTime - $p.LastWriteTime).TotalMinutes -le 1)
}

# ================================================ RESOLUCAO AUTOMATICA

if ($AplicarUltima) {
    $casa   = [Environment]::GetFolderPath("UserProfile")
    $pastas = @((Join-Path $casa "Downloads"),(Join-Path $casa "Desktop"),(Join-Path $casa "Documents"),$PastaSaida) |
              Where-Object { Test-Path -LiteralPath $_ }
    $cand = @()
    foreach ($p in $pastas) {
        $cand += Get-ChildItem -LiteralPath $p -File -ErrorAction SilentlyContinue |
                 Where-Object { $_.Name -like "decisoes*.csv" -or $_.Name -like "decisoes_planilha_*.xlsx" -or $_.Name -like "decisoes_planilha_*.csv" }
    }
    $escolha = $cand | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    # o CSV baixado do HTML tem prioridade sobre a planilha gerada no passo 1
    $doHtml = @($cand | Where-Object { $_.Name -like "decisoes*.csv" } | Sort-Object LastWriteTime -Descending)
    if ($doHtml.Count -gt 0) { $escolha = $doHtml[0] }

    if (-not $escolha) {
        Write-Host "Nao achei nenhum arquivo de decisoes." -ForegroundColor Red
        Write-Host "No HTML, marque os programas e clique em 'Baixar decisoes' no rodape." -ForegroundColor Red
        if (-not $Confirmar) { Read-Host "Enter para sair" | Out-Null }; exit 1
    }
    if ($cand.Count -gt 1) {
        Write-Host "Arquivos de decisao encontrados ($($cand.Count)):" -ForegroundColor DarkGray
        foreach ($x in ($cand | Sort-Object LastWriteTime -Descending | Select-Object -First 5)) {
            Write-Host ("   {0}  {1}" -f $x.LastWriteTime.ToString("dd/MM HH:mm"), $x.FullName) -ForegroundColor DarkGray
        }
    }
    Write-Host "Arquivo de decisoes encontrado:" -ForegroundColor Cyan
    Write-Host "  $($escolha.FullName)" -ForegroundColor White
    Write-Host "  baixado em $($escolha.LastWriteTime.ToString('dd/MM/yyyy HH:mm'))" -ForegroundColor DarkGray
    if (-not (Read-Confirmacao "E este mesmo? (S/N)")) { Write-Host "Cancelado."; exit 0 }
    $Aplicar = $escolha.FullName
    $ParaStaging = $true
}

if ($PublicarUltima) {
    $cands = @(Get-ChildItem -LiteralPath $PastaSaida -Directory -ErrorAction SilentlyContinue |
               Where-Object { $_.Name -like "publicar_*" -and -not (Test-Path -LiteralPath (Join-Path $_.FullName "_PUBLICADO.txt")) })
    $escolha = $cands | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $escolha) {
        Write-Host "Nenhuma pasta de conferencia pendente em $PastaSaida." -ForegroundColor Yellow
        if (-not $Confirmar) { Read-Host "Enter para sair" | Out-Null }; exit 0
    }
    $qtd = @(Get-ChildItem -LiteralPath $escolha.FullName -Recurse -File | Where-Object { $_.Name -notlike "_*" }).Count
    Write-Host "Pasta de conferencia pendente:" -ForegroundColor Cyan
    Write-Host "  $($escolha.FullName)  ($qtd arquivo(s))" -ForegroundColor White
    if ($cands.Count -gt 1) { Write-Host "  (existem $($cands.Count) pastas pendentes; esta e a mais recente)" -ForegroundColor Yellow }
    if (-not (Read-Confirmacao "Publicar esta pasta na producao? (S/N)")) { Write-Host "Cancelado."; exit 0 }
    $PublicarStaging = $escolha.FullName
}

# ================================================ MODO: PUBLICAR STAGING (2a etapa)

if ($PublicarStaging -ne "") {

    if (-not (Test-Path -LiteralPath $PublicarStaging)) {
        Write-Host "ERRO: pasta nao encontrada -> $PublicarStaging" -ForegroundColor Red; exit 1
    }
    Test-Caminhos

    $manif = Join-Path $PublicarStaging "_manifesto.csv"
    $marca = Join-Path $PublicarStaging "_PUBLICADO.txt"

    if (-not (Test-Path -LiteralPath $manif)) {
        Write-Host "ERRO: _manifesto.csv nao encontrado nesta pasta." -ForegroundColor Red
        Write-Host "Sem manifesto nao da para saber o que foi autorizado. Nao publico." -ForegroundColor Red
        exit 1
    }
    if (Test-Path -LiteralPath $marca) {
        Write-Host "Esta pasta JA FOI PUBLICADA:" -ForegroundColor Yellow
        Get-Content -LiteralPath $marca | ForEach-Object { Write-Host "  $_" -ForegroundColor Yellow }
        # republicar exige confirmacao explicita, mesmo com -Confirmar
        $r = Read-Host "Publicar novamente mesmo assim? (digite SIM)"
        if ($r -cne "SIM") { Write-Host "Cancelado."; exit 0 }
    }

    $itens = Import-CsvAuto $manif
    Write-Host "Manifesto: $($itens.Count) arquivo(s) autorizado(s)." -ForegroundColor Cyan

    # arquivos soltos na pasta que NAO estao no manifesto = ninguem autorizou
    $noDisco = @(Get-ChildItem -LiteralPath $PublicarStaging -Recurse -File |
                 Where-Object { $_.Name -notlike "_*" } |
                 ForEach-Object { $_.FullName.Substring($PublicarStaging.Length).TrimStart('\','/').ToLower() })
    $noManif = @($itens | ForEach-Object { ("" + $_.Arquivo).ToLower() })
    $intrusos = @($noDisco | Where-Object { $noManif -notcontains $_ })
    if ($intrusos.Count -gt 0) {
        Write-Host ""
        Write-Host "ATENCAO: $($intrusos.Count) arquivo(s) na pasta SEM autorizacao no manifesto:" -ForegroundColor Red
        $intrusos | Select-Object -First 15 | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
        Write-Host "Estes NAO serao publicados. Alguem largou arquivo aqui por fora do processo." -ForegroundColor Red
    }

    $quem = Read-Responsavel "Quem conferiu esta pasta"
    if ([string]::IsNullOrWhiteSpace($quem)) { Write-Host "Cancelado: responsavel obrigatorio."; exit 1 }

    $ok=0; $rec=0; $erro=0
    if ($Simular) { Write-Host "`n*** SIMULACAO - nada sera copiado ***`n" -ForegroundColor Yellow }

    foreach ($i in $itens) {
        $orig = Join-Path $PublicarStaging $i.Arquivo
        $alvo = Join-Path $Destino         $i.Arquivo

        if (-not (Test-Path -LiteralPath $orig)) {
            Write-Host "  FALTANDO  $($i.Arquivo) -> sumiu da pasta de conferencia" -ForegroundColor Red
            Write-LogPublicacao $i.Arquivo ("" + $i.Status) ("" + $i.VersaoOrigem) ("" + $i.Chamado) ("" + $i.Autorizador) "RECUSADO" "ausente no staging | conferido por $quem"
            $rec++; continue
        }
        $h = (Get-FileHash -LiteralPath $orig -Algorithm SHA1).Hash
        if ($h -ne ("" + $i.HashOrigem).ToUpper()) {
            Write-Host "  ALTERADO  $($i.Arquivo) -> conteudo mudou depois da separacao" -ForegroundColor Red
            Write-LogPublicacao $i.Arquivo ("" + $i.Status) ("" + $i.VersaoOrigem) ("" + $i.Chamado) ("" + $i.Autorizador) "RECUSADO" "hash divergente no staging | conferido por $quem"
            $rec++; continue
        }
        if (Test-Alarme $orig $alvo) {
            Write-Host "  ALARME    $($i.Arquivo) -> producao mudou depois da separacao" -ForegroundColor Red
            Write-LogPublicacao $i.Arquivo ("" + $i.Status) ("" + $i.VersaoOrigem) ("" + $i.Chamado) ("" + $i.Autorizador) "RECUSADO" "alarme: producao mais nova | conferido por $quem"
            $rec++; continue
        }
        if ($Simular) { Write-Host "  COPIARIA  $($i.Arquivo)" -ForegroundColor Yellow; $ok++; continue }
        try {
            $pa = Split-Path -Parent $alvo
            if (-not (Test-Path -LiteralPath $pa)) { New-Item -ItemType Directory -Path $pa -Force | Out-Null }
            Copy-Item -LiteralPath $orig -Destination $alvo -Force
            Write-Host "  PUBLICADO $($i.Arquivo)" -ForegroundColor Green
            Write-LogPublicacao $i.Arquivo ("" + $i.Status) ("" + $i.VersaoOrigem) ("" + $i.Chamado) ("" + $i.Autorizador) "PUBLICADO" "via staging | conferido por $quem"
            $ok++
        } catch {
            Write-Host "  ERRO      $($i.Arquivo) -> $($_.Exception.Message)" -ForegroundColor Red
            Write-LogPublicacao $i.Arquivo ("" + $i.Status) ("" + $i.VersaoOrigem) ("" + $i.Chamado) ("" + $i.Autorizador) "ERRO" $_.Exception.Message
            $erro++
        }
    }

    if (-not $Simular -and $ok -gt 0) {
        "Publicado em $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') por $Usuario ($Maquina). Conferido por: $quem. Arquivos: $ok." |
            Out-File -LiteralPath $marca -Encoding UTF8 -Append
    }

    Write-Host ""
    Write-Host "Publicados : $ok"  -ForegroundColor Green
    Write-Host "Recusados  : $rec" -ForegroundColor Red
    Write-Host "Erros      : $erro" -ForegroundColor Red
    if (-not $Simular) { Write-Host "Log: $LogPublicacao" -ForegroundColor Cyan }
    exit 0
}

# ============================================================ MODO: APLICAR

if ($Aplicar -ne "") {

    if (-not (Test-Path -LiteralPath $Aplicar)) {
        Write-Host "ERRO: arquivo nao encontrado -> $Aplicar" -ForegroundColor Red; exit 1
    }
    Test-Caminhos

    $linhas = Import-Planilha $Aplicar
    if ($linhas.Count -eq 0) { Write-Host "Arquivo sem linhas."; exit 0 }

    $colunas = $linhas[0].PSObject.Properties.Name
    foreach ($c in @("Arquivo","HashOrigem")) {
        if ($colunas -notcontains $c) {
            Write-Host "ERRO: coluna '$c' nao encontrada. Nao altere o cabecalho da planilha." -ForegroundColor Red; exit 1
        }
    }
    $temDecisao = ($colunas -contains "DECISAO") -or ($colunas -contains "AUTORIZADO")

    if (-not $temDecisao) {
        Write-Host "ERRO: nao achei a coluna DECISAO (nem AUTORIZADO)." -ForegroundColor Red; exit 1
    }

    # quantas decisoes existem de fato neste arquivo?
    $qtdMarcadas = @($linhas | Where-Object {
        $v = ("" + $_.DECISAO); if ([string]::IsNullOrWhiteSpace($v)) { $v = "" + $_.AUTORIZADO }
        -not [string]::IsNullOrWhiteSpace($v)
    }).Count

    Write-Host "Arquivo lido : $Aplicar" -ForegroundColor Cyan
    Write-Host "Linhas       : $($linhas.Count)" -ForegroundColor Cyan
    Write-Host "Com decisao  : $qtdMarcadas" -ForegroundColor Cyan

    if ($qtdMarcadas -eq 0) {
        Write-Host ""
        Write-Host "NENHUMA DECISAO PREENCHIDA NESTE ARQUIVO." -ForegroundColor Red
        Write-Host "Nada sera copiado e nenhuma pasta sera criada." -ForegroundColor Red
        Write-Host ""
        Write-Host "Se voce marcou os botoes no HTML, faltou clicar em 'Baixar decisoes'" -ForegroundColor Yellow
        Write-Host "no rodape da pagina antes de rodar este passo. As marcacoes ficam" -ForegroundColor Yellow
        Write-Host "so na tela ate voce baixar o arquivo." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Se voce usou a planilha, preencha a coluna DECISAO e salve." -ForegroundColor Yellow
        exit 1
    }

    $autorizador = Read-Responsavel "Quem autorizou esta lista (responsavel pela liberacao)"
    if ([string]::IsNullOrWhiteSpace($autorizador)) { Write-Host "Cancelado: responsavel obrigatorio."; exit 1 }

    # destino: producao direto, ou pasta de conferencia (staging)
    $pastaSt = ""
    if ($ParaStaging -or $PastaStaging -ne "") {
        $pastaSt = if ($PastaStaging -ne "") { $PastaStaging }
                   else { Join-Path $PastaSaida ("publicar_" + (Get-Date -Format "yyyyMMdd_HHmm")) }
        if (-not $Simular -and -not (Test-Path -LiteralPath $pastaSt)) {
            New-Item -ItemType Directory -Path $pastaSt -Force | Out-Null
        }
        $manifSt = Join-Path $pastaSt "_manifesto.csv"
        if (-not $Simular) {
            "Arquivo;HashOrigem;VersaoOrigem;Status;Chamado;Autorizador;DataPreparo" |
                Out-File -LiteralPath $manifSt -Encoding UTF8
        }
        Write-Host "Separando em: $pastaSt" -ForegroundColor Cyan
    }
    $destinoReal = if ($pastaSt -ne "") { $pastaSt } else { $Destino }

    $ok=0; $rec=0; $obs=0; $pulou=0; $erro=0
    if ($Simular) { Write-Host "`n*** MODO SIMULACAO - nada sera copiado ***`n" -ForegroundColor Yellow }

    foreach ($l in $linhas) {
        $d = ("" + $l.DECISAO)
        if ([string]::IsNullOrWhiteSpace($d)) { $d = "" + $l.AUTORIZADO }
        $d = $d.Trim().ToUpper()

        $ehAutorizado = $d -in @("AUTORIZADO","S","SIM","Y","1")
        $ehTeste      = $d -in @("EM TESTE","EM_TESTE","TESTE","OBSERVACAO","EM OBSERVACAO")

        if ($ehTeste) {
            if (-not $Simular) {
                Add-Observacao $l.Arquivo $l.HashOrigem $autorizador ("" + $l.CHAMADO) ("" + $l.OBSERVACAO)
            }
            Write-Host "  EM TESTE  $($l.Arquivo)" -ForegroundColor Yellow
            $obs++; continue
        }
        if (-not $ehAutorizado) { $pulou++; continue }

        $origemFull = Join-Path $Origem       $l.Arquivo
        $alvoFull   = Join-Path $destinoReal  $l.Arquivo

        if (-not (Test-Path -LiteralPath $origemFull)) {
            Write-Host "  RECUSADO  $($l.Arquivo) -> sumiu da homologacao" -ForegroundColor Red
            Write-LogPublicacao $l.Arquivo "?" ("" + $l.VersaoOrigem) ("" + $l.CHAMADO) $autorizador "RECUSADO" "origem inexistente"
            $rec++; continue
        }

        $hashAgora = (Get-FileHash -LiteralPath $origemFull -Algorithm SHA1).Hash
        if ($hashAgora -ne ("" + $l.HashOrigem).ToUpper()) {
            Write-Host "  RECUSADO  $($l.Arquivo) -> recompilado apos a autorizacao" -ForegroundColor Red
            Write-LogPublicacao $l.Arquivo ("" + $l.Status) ("" + $l.VersaoOrigem) ("" + $l.CHAMADO) $autorizador "RECUSADO" "hash divergente"
            $rec++; continue
        }

        if (Test-Alarme $origemFull (Join-Path $Destino $l.Arquivo)) {
            Write-Host "  RECUSADO  $($l.Arquivo) -> producao mais nova ou divergente (alarme)" -ForegroundColor Red
            Write-LogPublicacao $l.Arquivo ("" + $l.Status) ("" + $l.VersaoOrigem) ("" + $l.CHAMADO) $autorizador "RECUSADO" "alarme: producao mais nova"
            $rec++; continue
        }

        if ($pastaSt -eq "" -and (Test-Path -LiteralPath $alvoFull)) {
            if ((Get-FileHash -LiteralPath $alvoFull -Algorithm SHA1).Hash -eq $hashAgora) {
                Write-Host "  JA OK     $($l.Arquivo)" -ForegroundColor DarkGray; $pulou++; continue
            }
        }

        if ($Simular) { Write-Host "  COPIARIA  $($l.Arquivo)" -ForegroundColor Yellow; $ok++; continue }

        try {
            $pastaAlvo = Split-Path -Parent $alvoFull
            if (-not (Test-Path -LiteralPath $pastaAlvo)) { New-Item -ItemType Directory -Path $pastaAlvo -Force | Out-Null }
            Copy-Item -LiteralPath $origemFull -Destination $alvoFull -Force
            if ($pastaSt -ne "") {
                Write-Host "  SEPARADO  $($l.Arquivo)" -ForegroundColor Green
                Format-LinhaCsv @($l.Arquivo, $hashAgora, ("" + $l.VersaoOrigem),
                    ("" + $l.Status), ("" + $l.CHAMADO), $autorizador, (Get-Date -Format "dd/MM/yyyy HH:mm:ss")) |
                    Out-File -LiteralPath $manifSt -Append -Encoding UTF8
            } else {
                Write-Host "  PUBLICADO $($l.Arquivo)" -ForegroundColor Green
                Write-LogPublicacao $l.Arquivo ("" + $l.Status) ("" + $l.VersaoOrigem) ("" + $l.CHAMADO) $autorizador "PUBLICADO" ""
            }
            $ok++
        } catch {
            Write-Host "  ERRO      $($l.Arquivo) -> $($_.Exception.Message)" -ForegroundColor Red
            Write-LogPublicacao $l.Arquivo ("" + $l.Status) ("" + $l.VersaoOrigem) ("" + $l.CHAMADO) $autorizador "ERRO" $_.Exception.Message
            $erro++
        }
    }

    Write-Host ""
    if ($pastaSt -ne "") { Write-Host "Separados       : $ok" -ForegroundColor Green }
    else                 { Write-Host "Publicados      : $ok" -ForegroundColor Green }
    Write-Host "Em observacao   : $obs"   -ForegroundColor Yellow
    Write-Host "Recusados       : $rec"   -ForegroundColor Red
    Write-Host "Sem decisao     : $pulou" -ForegroundColor DarkGray
    Write-Host "Erros           : $erro"  -ForegroundColor Red

    if ($pastaSt -ne "" -and -not $Simular -and $ok -gt 0) {
        $lista = Join-Path $pastaSt "_CONFERIR.txt"
        $txt = New-Object System.Collections.Generic.List[string]
        $txt.Add("CONFERENCIA DE PUBLICACAO")
        $txt.Add("Preparado em $(Get-Date -Format 'dd/MM/yyyy HH:mm') por $Usuario")
        $txt.Add("Autorizado por: $autorizador")
        $txt.Add("Arquivos separados: $ok")
        $txt.Add("")
        foreach ($m in (Import-CsvAuto (Join-Path $pastaSt "_manifesto.csv"))) {
            $txt.Add(("  [ ] {0}   ({1})   chamado: {2}" -f $m.Arquivo, $m.VersaoOrigem, $m.Chamado))
        }
        $txt.Add("")
        $txt.Add("Depois de conferir, publique com:")
        $txt.Add("  .\DeployGate.ps1 -PublicarStaging `"$pastaSt`"")
        $txt | Out-File -LiteralPath $lista -Encoding UTF8

        Write-Host ""
        Write-Host "Pasta de conferencia: $pastaSt" -ForegroundColor Cyan
        Write-Host "Lista impressa em   : $lista"   -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Depois de conferir, rode:" -ForegroundColor Yellow
        Write-Host "  .\DeployGate.ps1 -PublicarStaging `"$pastaSt`"" -ForegroundColor Yellow
    }
    elseif (-not $Simular) {
        Write-Host "Log        : $LogPublicacao"  -ForegroundColor Cyan
        if ($obs -gt 0) { Write-Host "Observacao : $ArqObservacao" -ForegroundColor Cyan }
    }
    exit 0
}

# =========================================================== MODO: COMPARAR

Test-Caminhos

$ini = ConvertTo-Data $DataInicio
$fim = ConvertTo-Data $DataFim -FimDoDia
if ($ini -and $fim -and $ini -gt $fim) { Write-Host "ERRO: DataInicio maior que DataFim." -ForegroundColor Red; exit 1 }
$temFiltro = ($null -ne $ini -or $null -ne $fim)

$resultado = Compare-Pastas

# ---- marca o que esta em observacao (casa arquivo + hash; recompilou, sai da observacao)
$mapaObs = Get-Observacoes
foreach ($r in $resultado) {
    if ($r.Status -notin @("FALTA_NO_DESTINO","ORIGEM_MAIS_NOVA")) { continue }
    $k = $r.Arquivo.ToLower() + "|" + ("" + $r.HashOrigem).ToUpper()
    if ($mapaObs.ContainsKey($k)) {
        $o = $mapaObs[$k]
        $r.Status     = "EM_OBSERVACAO"
        $r.Sev        = 4
        $r.ObsQuem    = "" + $o.Quem
        $r.ObsChamado = "" + $o.Chamado
        $dt = ConvertTo-Data ("" + $o.DataDecisao)
        if ($dt) { $r.DiasObs = [int]((Get-Date) - $dt).TotalDays }
    }
}

$totPendentes = @($resultado | Where-Object { $_.Status -in @("FALTA_NO_DESTINO","ORIGEM_MAIS_NOVA") }).Count
$totAlarmes   = @($resultado | Where-Object { $_.Status -in @("DESTINO_MAIS_NOVO","DIVERGENTE_MESMA_DATA") }).Count
$totOrfaos    = @($resultado | Where-Object { $_.Status -eq "SO_NO_DESTINO" }).Count
$totObs       = @($resultado | Where-Object { $_.Status -eq "EM_OBSERVACAO" }).Count

$emFoco = @($resultado | Where-Object { Test-NoPeriodo $_ $ini $fim })

if ($Diag) {
    if (-not (Test-Path -LiteralPath $PastaSaida)) { New-Item -ItemType Directory -Path $PastaSaida -Force | Out-Null }
    $diagArq = Join-Path $PastaSaida ("diagnostico_" + (Get-Date -Format "yyyyMMdd_HHmm") + ".csv")
    $resultado | Select-Object Arquivo,Status,DataOrigem,DataDestino,VersaoOrigem,VersaoDestino,TamanhoOrigem,HashOrigem |
        Export-Csv -LiteralPath $diagArq -Delimiter ';' -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "DIAGNOSTICO" -ForegroundColor Magenta
    Write-Host "  Total de arquivos comparados: $($resultado.Count)" -ForegroundColor Magenta
    $resultado | Group-Object Status | Sort-Object Count -Descending | ForEach-Object {
        Write-Host ("  {0,-22} {1}" -f $_.Name, $_.Count) -ForegroundColor Magenta
    }
    Write-Host "  Apos filtro de periodo: $($emFoco.Count)" -ForegroundColor Magenta
    Write-Host "  Detalhe completo em: $diagArq" -ForegroundColor Magenta
    Write-Host ""
}

$pendentes = @($emFoco | Where-Object { $_.Status -in @("FALTA_NO_DESTINO","ORIGEM_MAIS_NOVA") })
$alarmes   = @($emFoco | Where-Object { $_.Status -in @("DESTINO_MAIS_NOVO","DIVERGENTE_MESMA_DATA") })
$orfaos    = @($emFoco | Where-Object { $_.Status -eq "SO_NO_DESTINO" })
$observ    = @($emFoco | Where-Object { $_.Status -eq "EM_OBSERVACAO" })
$iguais    = @($emFoco | Where-Object { $_.Status -eq "IGUAL" })

$foraDoFiltro = $totPendentes - $pendentes.Count
$lotes = Group-Lotes -Pendentes $pendentes -Minutos $MinutosLote

$rotuloPeriodo = if (-not $temFiltro) { "todo o historico" }
    elseif ($ini -and $fim) { "{0} a {1}" -f $ini.ToString("dd/MM/yyyy"), $fim.ToString("dd/MM/yyyy") }
    elseif ($ini)           { "a partir de {0}" -f $ini.ToString("dd/MM/yyyy") }
    else                    { "ate {0}" -f $fim.ToString("dd/MM/yyyy") }

if (-not (Test-Path -LiteralPath $PastaSaida)) { New-Item -ItemType Directory -Path $PastaSaida -Force | Out-Null }
$carimbo = Get-Date -Format "yyyyMMdd_HHmm"
$sufixo  = if ($temFiltro) {
    "_{0}_a_{1}" -f $(if ($ini) { $ini.ToString("ddMMyyyy") } else { "inicio" }),
                    $(if ($fim) { $fim.ToString("ddMMyyyy") } else { "hoje" })
} else { "" }
$baseNome  = "decisoes_planilha_$carimbo$sufixo"
$htmlSaida = Join-Path $PastaSaida "relatorio_$carimbo$sufixo.html"
$xlsxSaida = Join-Path $PastaSaida "$baseNome.xlsx"
$csvSaida  = Join-Path $PastaSaida "$baseNome.csv"

# ---------------------------------------------------- dados para planilha/HTML
$paraDecisao = New-Object System.Collections.Generic.List[object]
foreach ($lote in $lotes) {
    foreach ($it in ($lote.Itens | Sort-Object Arquivo)) {
        $paraDecisao.Add([pscustomobject]@{
            LOTE=$it.Lote; QtdLote=$lote.Itens.Count; Arquivo=$it.Arquivo; Status=$it.Status
            DataOrigem=$it.DataOrigem; DataDestino=$it.DataDestino; VersaoOrigem=$it.VersaoOrigem; VersaoDestino=$it.VersaoDestino
            TamanhoOrigem=$it.TamanhoOrigem; HashOrigem=$it.HashOrigem; DECISAO=""; CHAMADO=""; OBSERVACAO=""
        })
    }
}

# ---------------------------------------------------- gera XLSX (ou CSV)
$arqGerado = ""
$usouExcel = $false

$ehWindows = ($PSVersionTable.PSEdition -eq "Desktop") -or $IsWindows
if ($paraDecisao.Count -gt 0 -and -not $SemExcel -and $ehWindows) {
    $excel = $null; $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = "Autorizacao"

        $cols = @("LOTE","Arquivo","Status","DataOrigem","DataDestino","VersaoOrigem","VersaoDestino","TamanhoOrigem","HashOrigem","DECISAO","CHAMADO","OBSERVACAO")
        $nl = $paraDecisao.Count + 1
        $nc = $cols.Count
        $arr = New-Object 'object[,]' $nl,$nc
        for ($c=0; $c -lt $nc; $c++) { $arr[0,$c] = $cols[$c] }
        for ($r=0; $r -lt $paraDecisao.Count; $r++) {
            $item = $paraDecisao[$r]
            for ($c=0; $c -lt $nc; $c++) { $arr[$r+1,$c] = "" + $item.($cols[$c]) }
        }
        $ws.Range($ws.Cells(1,1), $ws.Cells($nl,$nc)).Value2 = $arr

        # cabecalho
        $hdr = $ws.Range($ws.Cells(1,1), $ws.Cells(1,$nc))
        $hdr.Interior.Color = 55 + 71*256 + 79*65536
        $hdr.Font.Color = 16777215
        $hdr.Font.Bold  = $true
        $ws.Rows.Item(1).RowHeight = 20

        $ws.Activate() | Out-Null
        $ws.Application.ActiveWindow.SplitRow = 1
        $ws.Application.ActiveWindow.FreezePanes = $true
        $hdr.AutoFilter() | Out-Null

        $colDec = 10
        $rngDec = $ws.Range($ws.Cells(2,$colDec), $ws.Cells($nl,$colDec))
        $rngDec.Validation.Delete()
        $rngDec.Validation.Add(3,1,1,"AUTORIZADO,EM TESTE") | Out-Null
        $rngDec.Validation.InCellDropdown = $true
        $rngDec.Validation.InputTitle   = "Decisao"
        $rngDec.Validation.InputMessage = "AUTORIZADO = pode ir para producao   |   EM TESTE = segura, fica em observacao"
        $rngDec.HorizontalAlignment = -4108

        $fcOk = $rngDec.FormatConditions.Add(1,3,'="AUTORIZADO"')
        $fcOk.Interior.Color = 198 + 239*256 + 206*65536
        $fcOk.Font.Color     = 0   + 97*256  + 0*65536
        $fcOk.Font.Bold      = $true
        $fcNo = $rngDec.FormatConditions.Add(1,3,'="EM TESTE"')
        $fcNo.Interior.Color = 255 + 199*256 + 206*65536
        $fcNo.Font.Color     = 6   + 0*256   + 156*65536
        $fcNo.Font.Bold      = $true

        # colunas tecnicas discretas
        $ws.Columns.Item(9).Font.Color = 12500670          # HashOrigem cinza
        $ws.Columns.Item(9).ColumnWidth = 12
        $ws.Columns.Item(1).ColumnWidth = 24
        $ws.Columns.Item(2).ColumnWidth = 38
        $ws.Columns.Item(3).ColumnWidth = 20
        $ws.Columns.Item(10).ColumnWidth = 16
        $ws.Columns.Item(11).ColumnWidth = 14
        $ws.Columns.Item(12).ColumnWidth = 30
        $ws.Range($ws.Cells(1,4), $ws.Cells($nl,8)).EntireColumn.AutoFit() | Out-Null

        if (Test-Path -LiteralPath $xlsxSaida) { Remove-Item -LiteralPath $xlsxSaida -Force }
        $wb.SaveAs($xlsxSaida, 51)
        $arqGerado = $xlsxSaida; $usouExcel = $true
    }
    catch {
        Write-Host "Aviso: nao consegui gerar XLSX ($($_.Exception.Message)). Gerando CSV." -ForegroundColor Yellow
    }
    finally {
        if ($wb)    { $wb.Close($false) }
        if ($excel) { $excel.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

if ($paraDecisao.Count -gt 0 -and -not $usouExcel) {
    $paraDecisao | Export-Csv -LiteralPath $csvSaida -Delimiter ';' -NoTypeInformation -Encoding UTF8
    $arqGerado = $csvSaida
}

# ---------------------------------------------------- HTML interativo
function HE([string]$s) { [System.Web.HttpUtility]::HtmlEncode($s) }

$paraHtml = @($emFoco | Sort-Object Sev, Arquivo | ForEach-Object {
    [pscustomobject]@{
        arq  = $_.Arquivo
        st   = $_.Status
        lote = $_.Lote
        dm   = $_.DataOrigem
        dn   = $_.DataDestino
        vm   = $_.VersaoOrigem
        vn   = $_.VersaoDestino
        hash = $_.HashOrigem
        dias = $_.DiasObs
        quem = $_.ObsQuem
    }
})
$json = if ($paraHtml.Count -gt 0) { $paraHtml | ConvertTo-Json -Compress -Depth 3 } else { "[]" }
if ($paraHtml.Count -eq 1) { $json = "[$json]" }

$avisoFiltro = ""
if ($temFiltro -and $foraDoFiltro -gt 0) {
    $avisoFiltro = "<div class='alerta'><b>Aten&ccedil;&atilde;o:</b> h&aacute; <b>$foraDoFiltro</b> pend&ecirc;ncia(s) <b>fora</b> do per&iacute;odo filtrado. " +
                   "Sem filtro o total &eacute; <b>$totPendentes</b> pendente(s), <b>$totObs</b> em observa&ccedil;&atilde;o, <b>$totAlarmes</b> alarme(s) e <b>$totOrfaos</b> s&oacute; na produ&ccedil;&atilde;o.</div>"
}

$cabecalho = @"
<!DOCTYPE html><html lang="pt-br"><head><meta charset="utf-8">
<title>Deploy Gate - homologacao x producao</title>
<style>
 *{box-sizing:border-box}
 body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;margin:0;padding:24px 24px 90px;color:#222;background:#fafafa}
 h1{font-size:20px;margin:0 0 2px}
 .sub{color:#666;margin-bottom:16px}
 .cards{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap}
 .card{border:1px solid #ccc;border-radius:6px;padding:10px 16px;min-width:140px;background:#fff}
 .card b{display:block;font-size:26px}
 .alerta{background:#fff3e0;border:1px solid #ffb74d;border-left:5px solid #ef6c00;border-radius:6px;padding:12px;margin-bottom:16px}
 .box{background:#eceff1;border:1px solid #cfd8dc;border-radius:6px;padding:12px;margin-bottom:16px}
 code{background:#fff;padding:2px 5px;border:1px solid #ddd;border-radius:3px}
 .ctrl{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:12px;position:sticky;top:0;background:#fafafa;padding:10px 0;z-index:5;border-bottom:1px solid #e0e0e0}
 #busca{flex:1;min-width:220px;padding:8px 12px;border:1px solid #bbb;border-radius:5px;font-size:14px}
 .chip{padding:6px 12px;border:1px solid #bbb;border-radius:16px;background:#fff;cursor:pointer;font-size:12px;user-select:none}
 .chip.on{background:#37474f;color:#fff;border-color:#37474f}
 table{border-collapse:collapse;width:100%;background:#fff}
 th{background:#37474f;color:#fff;padding:7px;text-align:left;font-size:12px;position:sticky;top:56px;cursor:pointer}
 td{border-bottom:1px solid #e0e0e0;padding:5px 7px;vertical-align:middle}
 tr.dOK{background:#e8f5e9} tr.dNO{background:#ffebee}
 .btn{border:1px solid #bbb;background:#fff;border-radius:5px;width:30px;height:26px;cursor:pointer;font-size:14px;line-height:1;margin-right:3px}
 .btn.ok.on{background:#2e7d32;border-color:#2e7d32;color:#fff}
 .btn.no.on{background:#c62828;border-color:#c62828;color:#fff}
 .btn:disabled{opacity:.25;cursor:not-allowed}
 .tag{font-size:11px;padding:2px 6px;border-radius:3px;white-space:nowrap}
 .t-FALTA_NO_DESTINO{background:#ffe0b2} .t-ORIGEM_MAIS_NOVA{background:#fff9c4}
 .t-DESTINO_MAIS_NOVO,.t-DIVERGENTE_MESMA_DATA{background:#ffcdd2}
 .t-SO_NO_DESTINO{background:#e1bee7} .t-EM_OBSERVACAO{background:#b3e5fc}
 .rodape{position:fixed;left:0;right:0;bottom:0;background:#37474f;color:#fff;padding:10px 24px;display:flex;gap:14px;align-items:center;flex-wrap:wrap}
 .rodape button{padding:8px 16px;border:0;border-radius:5px;background:#43a047;color:#fff;font-size:13px;cursor:pointer}
 .rodape button.sec{background:#78909c}
 .vazio{padding:24px;text-align:center;color:#888;background:#fff}
</style></head><body>
<h1>Deploy Gate &mdash; homologa&ccedil;&atilde;o &times; produ&ccedil;&atilde;o</h1>
<div class="sub">$(HE $Origem) &nbsp;x&nbsp; $(HE $Destino) &mdash; $(Get-Date -Format 'dd/MM/yyyy HH:mm') por $Usuario em $Maquina &mdash; per&iacute;odo: <b>$(HE $rotuloPeriodo)</b></div>
$avisoFiltro
<div class="cards">
 <div class="card" style="background:#ffe0b2"><b>$($pendentes.Count)</b>aguardando decis&atilde;o</div>
 <div class="card" style="background:#b3e5fc"><b>$($observ.Count)</b>em observa&ccedil;&atilde;o</div>
 <div class="card" style="background:#ffcdd2"><b>$($alarmes.Count)</b>alarmes</div>
 <div class="card" style="background:#e1bee7"><b>$($orfaos.Count)</b>s&oacute; na produ&ccedil;&atilde;o</div>
 <div class="card"><b>$($iguais.Count)</b>id&ecirc;nticos</div>
</div>
<div class="box">
Clique em <b>&#10003;</b> para autorizar ou <b>&#9209;</b> para deixar em teste. Depois clique em
<b>Baixar decis&otilde;es</b> no rodap&eacute; e d&ecirc; duplo-clique em
<code>2 - Separar autorizados.bat</code>.<br>
Os programas autorizados s&atilde;o <b>copiados</b> para uma pasta de confer&ecirc;ncia
(o original continua na homologa&ccedil;&atilde;o, nada &eacute; movido nem apagado). Depois de conferir, use
<code>3 - Publicar conferidos.bat</code>.<br>
Tamb&eacute;m d&aacute; para decidir pela planilha: <code>$(HE (Split-Path -Leaf $arqGerado))</code>
</div>
<div class="ctrl">
 <input id="busca" placeholder="Pesquisar programa, lote, vers&atilde;o..." autocomplete="off">
 <span class="chip on" data-f="PEND">Pend&ecirc;ncias</span>
 <span class="chip" data-f="FALTA_NO_DESTINO">Falta na produ&ccedil;&atilde;o</span>
 <span class="chip" data-f="ORIGEM_MAIS_NOVA">Desatualizado</span>
 <span class="chip" data-f="EM_OBSERVACAO">Em observa&ccedil;&atilde;o</span>
 <span class="chip" data-f="ALARME">Alarmes</span>
 <span class="chip" data-f="SO_NO_DESTINO">S&oacute; na produ&ccedil;&atilde;o</span>
 <span class="chip" data-f="IGUAL">Id&ecirc;nticos</span>
 <span class="chip" data-f="TODOS">Tudo</span>
 <span class="chip" data-f="MARCADOS">Marcados</span>
</div>
<table id="tab"><thead><tr>
 <th style="width:80px">Decis&atilde;o</th><th data-s="arq">Arquivo</th><th data-s="st">Situa&ccedil;&atilde;o</th>
 <th data-s="lote">Lote</th><th data-s="dm">Data homolog.</th><th data-s="dn">Data produ&ccedil;&atilde;o</th><th data-s="vm">Vers&atilde;o homolog.</th>
</tr></thead><tbody id="corpo"></tbody></table>
<div id="vazio" class="vazio" style="display:none">Nenhum registro para este filtro.</div>
<div class="rodape">
 <span id="resumo"></span>
 <button onclick="baixar()">Baixar decis&otilde;es (CSV)</button>
 <button class="sec" onclick="marcarVisiveis('OK')">Autorizar vis&iacute;veis</button>
 <button class="sec" onclick="limpar()">Limpar</button>
</div>
<script>
var DADOS = $json;
</script>
"@

$js = @'
<script>
(function(){
  var DEC = {}, filtro = "PEND", ordem = null, asc = true;
  var ALARME = ["DESTINO_MAIS_NOVO","DIVERGENTE_MESMA_DATA"];
  var PEND   = ["FALTA_NO_DESTINO","ORIGEM_MAIS_NOVA"];

  DADOS.forEach(function(d, i){ d._i = i; });

  function podeDecidir(st){ return PEND.indexOf(st) >= 0 || st === "EM_OBSERVACAO"; }

  function esc(s){
    if (s === null || s === undefined) return "";
    return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;")
                    .replace(/>/g,"&gt;").replace(/"/g,"&quot;");
  }

  function visiveis(){
    var q = document.getElementById("busca").value.toLowerCase().trim();
    return DADOS.filter(function(d){
      if (filtro === "PEND"){
        if (PEND.indexOf(d.st) < 0 && d.st !== "EM_OBSERVACAO" && ALARME.indexOf(d.st) < 0 && d.st !== "SO_NO_DESTINO") return false;
      } else if (filtro === "ALARME"){
        if (ALARME.indexOf(d.st) < 0) return false;
      } else if (filtro === "MARCADOS"){
        if (!DEC[d._i]) return false;
      } else if (filtro !== "TODOS"){
        if (d.st !== filtro) return false;
      }
      if (!q) return true;
      var alvo = (d.arq + " " + d.st + " " + (d.lote||"") + " " + (d.vm||"") + " " + (d.quem||"")).toLowerCase();
      return alvo.indexOf(q) >= 0;
    });
  }

  function render(){
    var lista = visiveis();
    if (ordem){
      lista = lista.slice().sort(function(a,b){
        var x = (a[ordem] === null || a[ordem] === undefined) ? "" : String(a[ordem]);
        var y = (b[ordem] === null || b[ordem] === undefined) ? "" : String(b[ordem]);
        return asc ? x.localeCompare(y) : y.localeCompare(x);
      });
    }
    var h = [];
    lista.forEach(function(d){
      var m = DEC[d._i] || "";
      var cls = m === "OK" ? "dOK" : (m === "NO" ? "dNO" : "");
      var trava = podeDecidir(d.st) ? "" : "disabled";
      var extra = "";
      if (d.st === "EM_OBSERVACAO" && d.dias !== null && d.dias !== undefined){
        extra = " <small>(" + d.dias + "d" + (d.quem ? ", " + esc(d.quem) : "") + ")</small>";
      }
      h.push('<tr class="' + cls + '" data-i="' + d._i + '">'
        + '<td><button class="btn ok ' + (m === "OK" ? "on" : "") + '" data-v="OK" ' + trava + ' title="Autorizar">&#10003;</button>'
        + '<button class="btn no ' + (m === "NO" ? "on" : "") + '" data-v="NO" ' + trava + ' title="Deixar em teste">&#9209;</button></td>'
        + '<td>' + esc(d.arq) + '</td>'
        + '<td><span class="tag t-' + d.st + '">' + d.st + '</span>' + extra + '</td>'
        + '<td>' + esc(d.lote || "-") + '</td>'
        + '<td>' + esc(d.dm) + '</td><td>' + esc(d.dn) + '</td><td>' + esc(d.vm) + '</td></tr>');
    });
    document.getElementById("corpo").innerHTML = h.join("");
    document.getElementById("vazio").style.display = lista.length ? "none" : "block";

    var nOk = 0, nNo = 0;
    for (var k in DEC){ if (DEC[k] === "OK") nOk++; else if (DEC[k] === "NO") nNo++; }
    document.getElementById("resumo").innerHTML =
      lista.length + " de " + DADOS.length + " exibidos &nbsp;|&nbsp; <b>" + nOk +
      "</b> autorizados &nbsp;|&nbsp; <b>" + nNo + "</b> em teste";
  }

  document.getElementById("corpo").addEventListener("click", function(ev){
    var b = ev.target.closest ? ev.target.closest("button.btn") : null;
    if (!b || b.disabled) return;
    var i = b.parentNode.parentNode.getAttribute("data-i");
    var v = b.getAttribute("data-v");
    if (DEC[i] === v) delete DEC[i]; else DEC[i] = v;
    render();
  });

  document.getElementById("busca").addEventListener("input", render);

  document.querySelectorAll(".chip").forEach(function(c){
    c.addEventListener("click", function(){
      document.querySelectorAll(".chip").forEach(function(x){ x.classList.remove("on"); });
      c.classList.add("on");
      filtro = c.getAttribute("data-f");
      render();
    });
  });

  document.querySelectorAll("th[data-s]").forEach(function(t){
    t.addEventListener("click", function(){
      var k = t.getAttribute("data-s");
      if (ordem === k) asc = !asc; else { ordem = k; asc = true; }
      render();
    });
  });

  window.marcarVisiveis = function(v){
    visiveis().forEach(function(d){ if (podeDecidir(d.st)) DEC[d._i] = v; });
    render();
  };
  window.limpar = function(){ DEC = {}; render(); };

  window.baixar = function(){
    var linhas = ["Arquivo;Status;VersaoOrigem;HashOrigem;DECISAO;CHAMADO;OBSERVACAO"], n = 0;
    DADOS.forEach(function(d){
      var m = DEC[d._i];
      if (!m) return;
      n++;
      linhas.push([d.arq, d.st, d.vm, d.hash, (m === "OK" ? "AUTORIZADO" : "EM TESTE"), "", ""].join(";"));
    });
    if (!n){ alert("Nenhuma decisao marcada. Clique nos botoes verde ou vermelho antes de baixar."); return; }
    var blob = new Blob(["\ufeff" + linhas.join("\r\n")], {type: "text/csv;charset=utf-8;"});
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "decisoes.csv";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    alert(n + " decisao(oes) baixada(s) em decisoes.csv.\n\nAgora rode: 2 - Separar autorizados.bat");
  };

  render();
})();
</script>
'@

$rodapeHtml = @"
$(if ($alarmes.Count -gt 0) {
"<div class='alerta' style='margin-top:20px'><b>Sobre os alarmes:</b> arquivos com a produ&ccedil;&atilde;o mais nova ou divergente da homologa&ccedil;&atilde;o n&atilde;o podem ser autorizados por aqui.
Significa que houve compila&ccedil;&atilde;o direta em produ&ccedil;&atilde;o e publicar por cima apagaria c&oacute;digo que s&oacute; existe na produ&ccedil;&atilde;o. Trate um a um.</div>"
})
</body></html>
"@

($cabecalho + $js + $rodapeHtml) | Out-File -LiteralPath $htmlSaida -Encoding UTF8

# ---------------------------------------------------- resumo
Write-Host ""
if ($temFiltro) { Write-Host "Periodo filtrado: $rotuloPeriodo" -ForegroundColor Cyan }
Write-Host "Aguardando decisao : $($pendentes.Count) em $($lotes.Count) lote(s)" -ForegroundColor Yellow
Write-Host "Em observacao      : $($observ.Count)" -ForegroundColor Cyan
Write-Host "Alarmes            : $($alarmes.Count)" -ForegroundColor Red
Write-Host "So na producao     : $($orfaos.Count)"  -ForegroundColor Magenta
Write-Host "Identicos          : $($iguais.Count)"  -ForegroundColor Green

$velhos = @($observ | Where-Object { $_.DiasObs -ne $null -and $_.DiasObs -gt 30 })
if ($velhos.Count -gt 0) {
    Write-Host ""
    Write-Host "$($velhos.Count) programa(s) em observacao ha mais de 30 dias:" -ForegroundColor Yellow
    foreach ($v in ($velhos | Sort-Object -Property DiasObs -Descending | Select-Object -First 10)) {
        Write-Host ("  {0,4}d  {1}" -f $v.DiasObs, $v.Arquivo) -ForegroundColor DarkYellow
    }
}
if ($temFiltro -and $foraDoFiltro -gt 0) {
    Write-Host ""
    Write-Host "ATENCAO: $foraDoFiltro pendencia(s) FORA do periodo nao aparecem acima." -ForegroundColor Red
    Write-Host "Total sem filtro: $totPendentes pendente(s), $totObs em observacao, $totAlarmes alarme(s), $totOrfaos so na producao." -ForegroundColor Red
}
Write-Host ""
Write-Host "Relatorio : $htmlSaida" -ForegroundColor Cyan
if ($arqGerado -ne "") { Write-Host "Planilha  : $arqGerado" -ForegroundColor Cyan }
if (-not $Silencioso) { Start-Process $htmlSaida }
