# Copie este arquivo para config.psd1 (na mesma pasta) e ajuste os caminhos.
# config.psd1 fica fora do Git (.gitignore): caminhos reais de servidor nao vao para o repositorio.
@{
    # pasta com os binarios ja testados (homologacao)
    Origem      = 'D:\ambientes\homologacao\bin'

    # pasta de onde os clientes/servidores recebem os binarios (producao)
    Destino     = 'D:\ambientes\producao\bin'

    # relatorios, logs e pastas de conferencia
    PastaSaida  = 'D:\ambientes\_deploy'

    # arquivos compilados com ate N minutos de diferenca formam um mesmo lote
    MinutosLote = 15

    # opcional: so estas extensoes (vazio = todas)
    Extensoes   = @()

    # nunca comparar
    Ignorar     = @('*.tmp', '*.bak', '*.log', '~*', '*.pdb')
}
