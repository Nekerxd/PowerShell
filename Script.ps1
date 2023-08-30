# -------------------------------- VARIÁVEIS --------------------------------
# Telegram
$botToken = "tokenBot"
$chatId = "chatID"

# Paths de arquivos
$arquivoExcel = "$pwd\Excel.xlsx"
$tabelaGerada = "$pwd\Tabela.png"

# Inicia a Lista de Criptomoedas 
$criptomoedas = New-Object System.Collections.Generic.List[System.String]

# -------------------------------- FUNÇÕES DO SISTEMA --------------------------------

# Consulta a API e retorna a última mensagem recebida
function ConsultarMensagemAPI() {

    $urlUpdate = "https://api.telegram.org/bot${botToken}/getUpdates"

    $bodyUpdate = @{
        "offset" = -1
        "limit" = 1
        "timeout" = $null
    } | ConvertTo-Json

    $respostaUpdate = Invoke-RestMethod -Uri $urlUpdate -Method POST -ContentType 'application/json' -Body $bodyUpdate
    
    return $respostaUpdate.result.message.text
}

# Envia uma mensagem para o usuário no Telegram
function EnviaMensagem() {
    param (
        [string]$mensagem = "Feito!"
    )
    
    $urlMensagem = "https://api.telegram.org/bot$botToken/sendMessage"
    
    $bodyMensagem = @{
        "text" = $mensagem
        "chat_id" = $chatId
        "parse_mode" = "HTML"
    } | ConvertTo-Json

    Invoke-RestMethod -Uri $urlMensagem -Method POST -ContentType 'application/json' -Body $bodyMensagem | Out-Null
}

# Converte a lista de criptomoedas para uma mensagem válida.
function EnviaListaCripto() {
    param (
        [array]$listaCriptomoedas
    )

    $lista = [String]::Join(", ", $listaCriptomoedas)
    $listaAtual = "<b>Lista atual: ${lista}</b>"

    EnviaMensagem $listaAtual
}

# Realiza as operações de adição e remoção de criptomoedas
function AtualizaLista() {
    
    $dadosMensagem = $ultimaMensagem.Split()

    if ($dadosMensagem[0] -eq 'Add' -and $criptomoedas.Count -lt 10){
        
        $criptomoedas.Add($dadosMensagem[1])
        EnviaMensagem("<b>Criptomoeda adicionada!</b>")
        
    } elseif ($dadosMensagem[0] -eq 'Remove'){

        $criptomoedas.Remove($dadosMensagem[1])
        EnviaMensagem("<b>Criptomoeda removida!</b>")

    } elseif ($dadosMensagem[0] -eq 'Mostrar'){

        EnviaListaCripto $criptomoedas

    } else {

        EnviaMensagem("Comando Inválido ou Limite Atingido")
    }
}

# -------------------------------- MONITORA NOVAS MENSAGENS --------------------------------

# Loop infinito para realizar as consultas.
$ultimaMensagem = ConsultarMensagemAPI
while ($true) {
    
    # Consultar a mensagem a cada 5 segundos.
    $mensagemAtual = ConsultarMensagemAPI
    if ($mensagemAtual -ne $ultimaMensagem) {
        $ultimaMensagem = $mensagemAtual
        AtualizaLista
    }
    
    # -------------------------------- INICIA A APLICAÇÃO EXCEL --------------------------------
    
    # Consultar criptomoedas a cada 5 minutos.
    if ((Get-Date).Minute % 5 -eq 0 -and $criptomoedas.Count -ne 0) {

        EnviaMensagem("<b>Executando busca dos dados, aguarde um minuto.</b>")

        # Garantir acesso ao recursos avançados de interface do usuário 
        Add-Type -AssemblyName System.Windows.Forms

        # Inicia o aplicativo Excel e abre o arquivo especificado.
        $excel = New-Object -ComObject Excel.Application
        $workbook = $excel.Workbooks.Open($arquivoExcel)
        $sheet = $workbook.Worksheets.Item(1)

        # -------------------------------- LIMPA A TABELA GERADA ANTERIORMENTE --------------------------------

        # Define as células já utilizadas na planilha
        $celulasUsadas = $sheet.UsedRange

        # Verifica se há linhas a serem removidas com exceção do cabeçalho.
        if ($celulasUsadas.Rows.Count -gt 1) {
            # Define o intervalo para excluir as linhas.
            $primeraLinha = $celulasUsadas.Row + 1
            $ultimaLinha = $celulasUsadas.Row + $celulasUsadas.Rows.Count - 1
            
            # Excluir as linhas da tabela.
            $sheet.Rows("${primeraLinha}:${ultimaLinha}").EntireRow.Delete() | Out-Null
        }

        # -------------------------------- CONSULTA E EXTRAI OS DADOS DAS CRIPTOMOEDAS --------------------------------

        # Início a partir da segunda linha p/ manter o cabeçalho.
        $row = 2

        # Para cada moeda na lista de criptomoedas, consultar e atualizar os dados.
        foreach ($criptomoeda in $criptomoedas) {

            # Consulta os dados de uma determinada moeda
            $respostaCripto = Invoke-RestMethod -Uri "https://api.coingecko.com/api/v3/simple/price?ids=${criptomoeda}&vs_currencies=BRL&include_market_cap=true&include_24hr_vol=true&include_24hr_change=true&include_last_updated_at=true&precision=2"
            
            # Extrai os dados da resposta
            $valorAtual = $respostaCripto.${criptomoeda}.brl
            $flutuacao24h = ($respostaCripto.${criptomoeda}.brl_24h_change) * 0.01
            $volume24h = $respostaCripto.${criptomoeda}.brl_24h_vol
            $capMercado = $respostaCripto.${criptomoeda}.brl_market_cap
            $ultimaData = $respostaCripto.${criptomoeda}.last_updated_at

            #Converte Timestamp para Horário de Brasília.
            $dataUnix = [System.DateTimeOffset]::FromUnixTimeSeconds($ultimaData).DateTime.AddHours(-3)
            $dataAtualizacao = $dataUnix.ToString("MM/dd/yyyy HH:mm")

            # -------------------------------- PASSA O VALOR DADOS PARA A TABELA EXCEL --------------------------------

            # Defina os dados que deseja inserir na tabela.
            $dadosRespostaCripto = @(
                @($criptomoeda, $valorAtual, $flutuacao24h, $volume24h, $capMercado, $dataAtualizacao)
            )

            # Para cada dado retornado pela API, inserir em uma coluna.
            $col = 1
            foreach ($dado in $dadosRespostaCripto) {
                $sheet.Cells.Item($row, $col) = $dado
                $col++
            }

            # -------------------------------- ESTILIZA AS CÉLULAS COM OS VALORES --------------------------------
            
            # Formatando Células
            $sheet.Range("A${row}:F${row}").Font.Bold = $true
            $sheet.Range("B${row}").NumberFormat = "R$ #.##0,00"
            $sheet.Range("C${row}").NumberFormat = "0,00%"
            $sheet.Range("D${row}:E${row}").NumberFormat = "R$ #.##0,00"
            $sheet.Range("F${row}").NumberFormat = "dd/mm/aaaa hh:mm"

            $row++
        }

        # Salva as alterações
        $workbook.Save()
        
        # -------------------------------- GERA A IMAGEM QUE SERÁ ENVIADA AO IMGUR --------------------------------
        
        # Seleciona as células da tabela
        $row--
        $rangeTabela = "A1:F${row}"
        $tabela = $sheet.Range($rangeTabela)

        # Copia a tabela para a área de transferência como imagem 
        $tabela.CopyPicture(1, 2) | Out-Null

        # Objeto para acessar os dados do print
        $dataObject = [System.Windows.Forms.Clipboard]::GetDataObject()

        # Verifica se a área de transferência contém uma imagem
        if ($dataObject -and $dataObject.ContainsImage()) {
            # Pega a imagem da área de transferência
            $imagemTabela = $dataObject.GetImage()
            
            # Salva a imagem como png
            $imagemTabela.Save($tabelaGerada, [System.Drawing.Imaging.ImageFormat]::Png)
            
            Write-Host "A imagem da tabela foi gerada com sucesso em $tabelaGerada."
        } else {
            Write-Host "Não foi possível obter uma imagem da tabela da área de transferência."
        }

        # -------------------------------- ENCERRA A EXCEL APPLICATION --------------------------------

        # Fecha o Excel e limpa da memória
        $workbook.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($tabela) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

        # Limpa as variáveis
        Remove-Variable excel, workbook, sheet, tabela

        # -------------------------------- API IMGUR PARA UPLOAD DA TABELA GERADA --------------------------------
        
        # Variáveis
        $clientId = "clientID"
        $urlUpload = "https://api.imgur.com/3/image"
        
        # Converte o arquivo em Bytes
        $imgBytes = [System.IO.File]::ReadAllBytes($tabelaGerada)
        
        # Converte os bytes da imagem para Base64
        $imgBase64 = [System.Convert]::ToBase64String($imgBytes)
        
        # Define o header da requisição
        $headerUpload= @{
            Authorization = "Client-ID $clientId"
        }
        
        #Define o body para a requisição
        $paramsUpload = @{ "image" = $imgBase64 } | ConvertTo-Json
        
        # Requisição para a API do Imgur
        $dadosUpload = Invoke-RestMethod $urlUpload -Method 'POST' -Headers $headerUpload -ContentType 'application/json' -Body $paramsUpload
        
        # -------------------------------- API TELEGRAM ENVIA TABELA FINALIZADA --------------------------------

        $mensagem = "<b><a href='$($dadosUpload.data.link)'>Tabela de Criptomoedas Atualizadas</a></b>"

        EnviaMensagem $mensagem

        Start-Sleep -Seconds 60
    }

    Start-Sleep -Seconds 5
}
