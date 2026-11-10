Option Explicit

' Macro para planilhas de estoque no formato Morada Nova.
'
' Estrutura esperada no arquivo .xlsm:
'   Aba "Cadastro de Clientes" -> cadastro de clientes (col A = codigo, col B = nome)
'   Aba "Planilha1"            -> dados de saida/estoque a processar
'
' Se a aba "Cadastro de Clientes" nao existir, busca um CSV na mesma pasta cujo nome
' contenha "Cadastro" e "Cliente". Se a aba "Planilha1" nao existir,
' usa a aba ativa.
'
' O macro:
' - Preenche a coluna I com o codigo do cliente (primeira ocorrencia de cada pedido).
' - Preenche a coluna I com o codigo da mercadoria na linha de titulo "SAIDA DE ... (...)".
'   Usa a aba "Cadastro de Itens" (ou um CSV na mesma pasta com "Cadastro" e "Item"/"Itens"
'   no nome, ex.: Cadastro de Itens.csv) com a mesma logica de busca usada para o cliente.
'   Mercadorias sem correspondencia confiavel (ou ambiguas, com mais de um item candidato)
'   ficam com a coluna I em branco e um alerta na coluna K. Ajustes manuais/excecoes ficam
'   na funcao CriarMapaMercadorias().
' - Gera o resumo do pedido na coluna J (primeira ocorrencia).
'   Pedidos pagos com BOLETO NAO recebem resumo (so um alerta na coluna K), para nao
'   serem lancados no Contas a Receber.
'   Unidade: KG para CAROCO e OLEO (ou titulo com a palavra "KG"); SACOS para o resto.
' - Deixa a coluna I em branco quando nao encontra correspondencia com confianca suficiente.
' - Grava alertas na coluna K para pedidos com clientes diferentes ou preco atipico.
' - Diferencas de ate R$ 5,00 no preco unitario nao geram alerta.

Private Const TOLERANCIA_PRECO_UNITARIO As Double = 5#

Public Sub GerarResumoPedidos_ComCodigoCliente()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim numPedido As Variant
    Dim cliente As String
    Dim qtd As Double
    Dim valorUnitario As Double
    Dim formaPagamento As String
    Dim mercadoriaAtual As String
    Dim dataAtual As String
    Dim unidade As String
    Dim dictPedidos As Object
    Dim linhasPrimeiraOcorrencia As Object
    Dim clientesPorPedido As Object
    Dim precoPorMercadoria As Object
    Dim precoPadraoPorMercadoria As Object
    Dim item As String
    Dim resumo As String
    Dim tituloLinha As String
    Dim regex As Object
    Dim matches As Object
    Dim dictCodigosExatos As Object
    Dim dictIndiceTokens As Object
    Dim nomesBase As Variant
    Dim codigosBase As Variant
    Dim nomesNormBase As Variant
    Dim cacheCodigos As Object
    Dim codigoCliente As String
    Dim chaveMercadoria As String
    Dim chavePreco As String
    Dim alerta As String
    Dim totalAlertas As Long
    Dim mapaMercadorias As Object
    Dim codigosItens As Variant
    Dim nomesNormItens As Variant
    Dim itensCarregados As Boolean
    Dim cacheItens As Object
    Dim totalMercadoriasSemCodigo As Long
    Dim pedidosBoleto As Object
    Dim totalBoleto As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Planilha1")
    On Error GoTo 0
    If ws Is Nothing Then Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Set dictPedidos = CreateObject("Scripting.Dictionary")
    Set linhasPrimeiraOcorrencia = CreateObject("Scripting.Dictionary")
    Set clientesPorPedido = CreateObject("Scripting.Dictionary")
    Set precoPorMercadoria = CreateObject("Scripting.Dictionary")
    Set precoPadraoPorMercadoria = CreateObject("Scripting.Dictionary")
    Set cacheCodigos = CreateObject("Scripting.Dictionary")
    Set pedidosBoleto = CreateObject("Scripting.Dictionary")
    Set regex = CreateObject("VBScript.RegExp")

    If Not CarregarCadastroClientes(dictCodigosExatos, dictIndiceTokens, nomesBase, codigosBase, nomesNormBase) Then
        MsgBox "Nao encontrei cadastro de clientes." & vbCrLf & _
               "Verifique se existe uma aba chamada 'Cadastro de Clientes' neste arquivo" & vbCrLf & _
               "ou um arquivo CSV com 'Cadastro' e 'Cliente' no nome na mesma pasta.", vbExclamation
        Exit Sub
    End If

    Set mapaMercadorias = CriarMapaMercadorias()
    Set cacheItens = CreateObject("Scripting.Dictionary")
    itensCarregados = CarregarCadastroItens(codigosItens, nomesNormItens)

    With regex
        .Global = False
        .IgnoreCase = True
        .Pattern = "SA[IÍ]DA DE (.+?) \([^\)]*\)\s*(\d{2}/\d{2}/\d{2,4})"
    End With

    ' Comeca na linha 1 (nao na 2) porque a primeira linha de "SAIDA DE ..." de cada
    ' aba costuma estar na linha 1, e agora tambem recebe o codigo da mercadoria/alertas.
    ws.Range("I1:K" & lastRow).ClearContents

    For i = 2 To lastRow
        numPedido = Trim(ws.Cells(i, "A").Value)

        If numPedido <> "" And IsNumeric(numPedido) Then
            cliente = Trim(ws.Cells(i, "B").Value)
            qtd = ws.Cells(i, "D").Value
            valorUnitario = ws.Cells(i, "F").Value
            formaPagamento = Trim(ws.Cells(i, "H").Value)

            mercadoriaAtual = ""
            dataAtual = ""
            For j = i To 1 Step -1
                If ws.Cells(j, "A").Font.Bold = True Then
                    tituloLinha = Trim(ws.Cells(j, "A").Value)
                    If regex.Test(tituloLinha) Then
                        Set matches = regex.Execute(tituloLinha)
                        mercadoriaAtual = Trim(matches(0).SubMatches(0))
                        dataAtual = Trim(matches(0).SubMatches(1))
                        Exit For
                    End If
                End If
            Next j

            unidade = UnidadeDaMercadoria(mercadoriaAtual)

            If Not dictPedidos.exists(numPedido) Then
                dictPedidos.Add numPedido, ""
                linhasPrimeiraOcorrencia.Add numPedido, i
                clientesPorPedido.Add numPedido, NormalizarNomeCliente(cliente)
            End If

            ' Qualquer linha do pedido com BOLETO marca o pedido inteiro (nao gera resumo).
            If EhBoleto(formaPagamento) Then
                If Not pedidosBoleto.exists(numPedido) Then pedidosBoleto.Add numPedido, True
            End If

            If clientesPorPedido.exists(numPedido) Then
                If clientesPorPedido(numPedido) <> "" And NormalizarNomeCliente(cliente) <> "" Then
                    If clientesPorPedido(numPedido) <> NormalizarNomeCliente(cliente) Then
                        AdicionarAlerta ws.Cells(i, "K"), "Pedido " & numPedido & " com cliente diferente da primeira ocorrencia"
                        AdicionarAlerta ws.Cells(linhasPrimeiraOcorrencia(numPedido), "K"), "Pedido " & numPedido & " tambem aparece com outro cliente"
                    End If
                End If
            End If

            If Not EhVendaAntecipada(formaPagamento) And mercadoriaAtual <> "" And IsNumeric(valorUnitario) And valorUnitario > 0 Then
                chaveMercadoria = NormalizarNomeCliente(mercadoriaAtual)
                chavePreco = Format(CDbl(valorUnitario), "0.00")
                RegistrarPrecoMercadoria precoPorMercadoria, chaveMercadoria, chavePreco
            End If

            item = qtd & " " & unidade & " DE " & mercadoriaAtual & " A R$ " & Format(valorUnitario, "#,##0.00")

            If dictPedidos(numPedido) = "" Then
                dictPedidos(numPedido) = item
            Else
                dictPedidos(numPedido) = dictPedidos(numPedido) & " - " & item
            End If
        End If
    Next i

    DefinirPrecoPadraoPorMercadoria precoPorMercadoria, precoPadraoPorMercadoria

    For i = 2 To lastRow
        numPedido = Trim(ws.Cells(i, "A").Value)
        If numPedido <> "" And IsNumeric(numPedido) Then
            formaPagamento = Trim(ws.Cells(i, "H").Value)
            valorUnitario = ws.Cells(i, "F").Value
            mercadoriaAtual = ObterMercadoriaDaLinha(ws, i, regex)

            If Not EhVendaAntecipada(formaPagamento) And mercadoriaAtual <> "" And IsNumeric(valorUnitario) And valorUnitario > 0 Then
                chaveMercadoria = NormalizarNomeCliente(mercadoriaAtual)
                chavePreco = Format(CDbl(valorUnitario), "0.00")
                If precoPadraoPorMercadoria.exists(chaveMercadoria) Then
                    If Abs(CDbl(valorUnitario) - CDbl(precoPadraoPorMercadoria(chaveMercadoria))) > TOLERANCIA_PRECO_UNITARIO Then
                        AdicionarAlerta ws.Cells(i, "K"), "Preco fora do padrao de " & mercadoriaAtual & ": R$ " & Format(CDbl(valorUnitario), "#,##0.00") & " | padrao R$ " & Format(CDbl(precoPadraoPorMercadoria(chaveMercadoria)), "#,##0.00")
                    End If
                End If
            End If
        End If
    Next i

    For i = 2 To lastRow
        numPedido = Trim(ws.Cells(i, "A").Value)
        If numPedido <> "" And IsNumeric(numPedido) Then
            cliente = ws.Cells(i, "B").Value
            codigoCliente = EncontrarCodigoCliente(cliente, dictCodigosExatos, dictIndiceTokens, nomesBase, codigosBase, nomesNormBase, cacheCodigos)
            ws.Cells(i, "I").Value = codigoCliente
        End If
    Next i

    totalMercadoriasSemCodigo = PreencherCodigosMercadoria(ws, lastRow, regex, mapaMercadorias, itensCarregados, codigosItens, nomesNormItens, cacheItens)

    For Each numPedido In dictPedidos.keys
        i = linhasPrimeiraOcorrencia(numPedido)
        cliente = ws.Cells(i, "B").Value
        formaPagamento = ws.Cells(i, "H").Value

        For j = i To 1 Step -1
            If ws.Cells(j, "A").Font.Bold = True Then
                tituloLinha = Trim(ws.Cells(j, "A").Value)
                If regex.Test(tituloLinha) Then
                    Set matches = regex.Execute(tituloLinha)
                    mercadoriaAtual = Trim(matches(0).SubMatches(0))
                    dataAtual = Trim(matches(0).SubMatches(1))
                    Exit For
                End If
            End If
        Next j

        If pedidosBoleto.exists(numPedido) Then
            ' BOLETO nao entra no Contas a Receber: sem resumo na coluna J, para
            ' nao ser lancado nem pelo robo nem manualmente.
            AdicionarAlerta ws.Cells(i, "K"), "Pedido " & numPedido & " pago com BOLETO - NAO LANCAR (resumo nao gerado)"
            totalBoleto = totalBoleto + 1
        Else
            resumo = cliente & " - " & numPedido & " - " & dataAtual & " - " & dictPedidos(numPedido) & " - " & formaPagamento
            ws.Cells(i, "J").Value = resumo
        End If
    Next numPedido

    totalAlertas = ContarAlertas(ws, lastRow)
    alerta = "Codigos dos clientes e resumo dos pedidos gerados com sucesso!"
    If totalBoleto > 0 Then alerta = alerta & vbCrLf & vbCrLf & "Pedidos com BOLETO (sem resumo, nao lancar): " & totalBoleto
    If totalMercadoriasSemCodigo > 0 Then alerta = alerta & vbCrLf & vbCrLf & "Mercadorias sem codigo encontrado: " & totalMercadoriasSemCodigo & " (veja a coluna K nas linhas de titulo 'SAIDA DE ...')"
    If totalAlertas > 0 Then alerta = alerta & vbCrLf & vbCrLf & "Alertas encontrados na coluna K: " & totalAlertas
    MsgBox alerta, vbInformation
End Sub

Private Function ObterMercadoriaDaLinha(ByVal ws As Worksheet, ByVal linha As Long, ByVal regex As Object) As String
    Dim j As Long
    Dim tituloLinha As String
    Dim matches As Object

    For j = linha To 1 Step -1
        If ws.Cells(j, "A").Font.Bold = True Then
            tituloLinha = Trim(ws.Cells(j, "A").Value)
            If regex.Test(tituloLinha) Then
                Set matches = regex.Execute(tituloLinha)
                ObterMercadoriaDaLinha = Trim(matches(0).SubMatches(0))
                Exit Function
            End If
        End If
    Next j
End Function

Private Sub RegistrarPrecoMercadoria(ByVal precoPorMercadoria As Object, ByVal mercadoria As String, ByVal preco As String)
    Dim dictPrecos As Object

    If mercadoria = "" Or preco = "" Then Exit Sub

    If precoPorMercadoria.exists(mercadoria) Then
        Set dictPrecos = precoPorMercadoria(mercadoria)
    Else
        Set dictPrecos = CreateObject("Scripting.Dictionary")
        precoPorMercadoria.Add mercadoria, dictPrecos
    End If

    If dictPrecos.exists(preco) Then
        dictPrecos(preco) = CLng(dictPrecos(preco)) + 1
    Else
        dictPrecos.Add preco, 1
    End If
End Sub

Private Sub DefinirPrecoPadraoPorMercadoria(ByVal precoPorMercadoria As Object, ByVal precoPadraoPorMercadoria As Object)
    Dim mercadoria As Variant
    Dim preco As Variant
    Dim dictPrecos As Object
    Dim melhorPreco As String
    Dim melhorQtd As Long

    For Each mercadoria In precoPorMercadoria.keys
        Set dictPrecos = precoPorMercadoria(mercadoria)
        melhorPreco = ""
        melhorQtd = 0

        For Each preco In dictPrecos.keys
            If CLng(dictPrecos(preco)) > melhorQtd Then
                melhorQtd = CLng(dictPrecos(preco))
                melhorPreco = CStr(preco)
            End If
        Next preco

        If melhorQtd >= 2 Then precoPadraoPorMercadoria.Add CStr(mercadoria), melhorPreco
    Next mercadoria
End Sub

Private Function EhVendaAntecipada(ByVal formaPagamento As String) As Boolean
    Dim formaNorm As String
    formaNorm = NormalizarNomeCliente(formaPagamento)
    EhVendaAntecipada = (InStr(1, formaNorm, "VENDA ANTECIPADA", vbTextCompare) > 0)
End Function

Private Function EhBoleto(ByVal formaPagamento As String) As Boolean
    EhBoleto = (InStr(1, NormalizarNomeCliente(formaPagamento), "BOLET", vbTextCompare) > 0)
End Function

' Mercadorias vendidas por KG (o resto e vendido em SACOS). Pelo historico de 2026,
' so CAROCO (~R$ 1,80/kg) e OLEO (~R$ 8,00/kg); tambem vale qualquer titulo com a
' palavra "KG" solta (ex.: "KG OLEO", "CAROCO KG") - "MILHO MOIDO 40KG" continua SACOS.
' Para incluir outro produto, acrescente a palavra em Case abaixo.
Private Function UnidadeDaMercadoria(ByVal mercadoria As String) As String
    Dim palavras As Variant
    Dim i As Long

    UnidadeDaMercadoria = "SACOS"
    palavras = Split(NormalizarNomeMercadoria(mercadoria), " ")
    For i = LBound(palavras) To UBound(palavras)
        Select Case CStr(palavras(i))
            Case "CAROCO", "OLEO", "KG"
                UnidadeDaMercadoria = "KG"
                Exit Function
        End Select
    Next i
End Function

Private Sub AdicionarAlerta(ByVal celula As Range, ByVal texto As String)
    If Trim(CStr(celula.Value)) = "" Then
        celula.Value = texto
    ElseIf InStr(1, CStr(celula.Value), texto, vbTextCompare) = 0 Then
        celula.Value = celula.Value & " | " & texto
    End If
End Sub

Private Function ContarAlertas(ByVal ws As Worksheet, ByVal lastRow As Long) As Long
    Dim i As Long
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, "K").Value)) <> "" Then ContarAlertas = ContarAlertas + 1
    Next i
End Function

Private Function CarregarCadastroClientes(ByRef dictCodigosExatos As Object, ByRef dictIndiceTokens As Object, ByRef nomesBase As Variant, ByRef codigosBase As Variant, ByRef nomesNormBase As Variant) As Boolean
    On Error GoTo Falha

    Dim linhas As Collection
    Dim i As Long
    Dim nome As String
    Dim nomeNorm As String
    Dim codigo As String

    Set dictCodigosExatos = CreateObject("Scripting.Dictionary")
    Set dictIndiceTokens = CreateObject("Scripting.Dictionary")
    Set linhas = New Collection

    Dim wsBase As Worksheet
    On Error Resume Next
    Set wsBase = ThisWorkbook.Sheets("Cadastro de Clientes")
    On Error GoTo Falha

    If Not wsBase Is Nothing Then
        Dim lastRowBase As Long
        Dim startRow As Long
        lastRowBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
        If IsNumeric(wsBase.Cells(1, "A").Value) Then
            startRow = 1
        Else
            startRow = 2
        End If
        For i = startRow To lastRowBase
            codigo = Trim(CStr(wsBase.Cells(i, "A").Value))
            nome = Trim(CStr(wsBase.Cells(i, "B").Value))
            If codigo <> "" And nome <> "" Then
                nomeNorm = NormalizarNomeCliente(nome)
                If NomeCadastroValido(nomeNorm) Then linhas.Add Array(codigo, nome, nomeNorm)
            End If
        Next i
    Else
        Dim arquivoCsv As String
        Dim linha As String
        Dim campos As Variant
        Dim f As Integer

        arquivoCsv = EncontrarArquivoCadastro()
        If arquivoCsv = "" Then Exit Function

        f = FreeFile
        Open arquivoCsv For Input As #f
        Do Until EOF(f)
            Line Input #f, linha
            campos = SepararCsvPontoVirgula(linha)
            If IsArray(campos) Then
                If UBound(campos) >= 1 Then
                    codigo = Trim(CStr(campos(0)))
                    nome = Trim(CStr(campos(1)))
                    If codigo <> "" And nome <> "" Then
                        nomeNorm = NormalizarNomeCliente(nome)
                        If NomeCadastroValido(nomeNorm) Then linhas.Add Array(codigo, nome, nomeNorm)
                    End If
                End If
            End If
        Loop
        Close #f
    End If

    If linhas.Count = 0 Then Exit Function

    ReDim codigosBase(1 To linhas.Count, 1 To 1)
    ReDim nomesBase(1 To linhas.Count, 1 To 1)
    ReDim nomesNormBase(1 To linhas.Count, 1 To 1)

    For i = 1 To linhas.Count
        codigo = CStr(linhas(i)(0))
        nome = CStr(linhas(i)(1))
        nomeNorm = CStr(linhas(i)(2))

        codigosBase(i, 1) = codigo
        nomesBase(i, 1) = nome
        nomesNormBase(i, 1) = nomeNorm

        If Not dictCodigosExatos.exists(nomeNorm) Then dictCodigosExatos.Add nomeNorm, codigo
        AdicionarNomeAoIndiceTokens dictIndiceTokens, nomeNorm, i
    Next i

    CarregarCadastroClientes = True
    Exit Function

Falha:
    On Error Resume Next
    CarregarCadastroClientes = False
End Function

Private Function EncontrarArquivoCadastro() As String
    Dim pasta As String
    Dim arquivo As String
    Dim melhorArquivo As String
    Dim melhorData As Date

    pasta = ThisWorkbook.Path
    If pasta = "" Then pasta = CurDir$

    arquivo = Dir(pasta & "\*.csv")
    Do While arquivo <> ""
        If InStr(1, arquivo, "Cadastro", vbTextCompare) > 0 And InStr(1, arquivo, "Cliente", vbTextCompare) > 0 Then
            If FileDateTime(pasta & "\" & arquivo) > melhorData Then
                melhorData = FileDateTime(pasta & "\" & arquivo)
                melhorArquivo = pasta & "\" & arquivo
            End If
        End If
        arquivo = Dir()
    Loop

    EncontrarArquivoCadastro = melhorArquivo
End Function

Private Function SepararCsvPontoVirgula(ByVal linha As String) As Variant
    Dim campos() As String
    Dim atual As String
    Dim i As Long
    Dim ch As String
    Dim dentroAspas As Boolean
    Dim idx As Long

    ReDim campos(0 To 0)
    idx = 0

    For i = 1 To Len(linha)
        ch = Mid$(linha, i, 1)
        If ch = """" Then
            dentroAspas = Not dentroAspas
        ElseIf ch = ";" And Not dentroAspas Then
            campos(idx) = atual
            atual = ""
            idx = idx + 1
            ReDim Preserve campos(0 To idx)
        Else
            atual = atual & ch
        End If
    Next i

    campos(idx) = atual
    SepararCsvPontoVirgula = campos
End Function

Private Function EncontrarCodigoCliente(ByVal nomeDigitado As String, ByVal dictCodigosExatos As Object, ByVal dictIndiceTokens As Object, ByVal nomesBase As Variant, ByVal codigosBase As Variant, ByVal nomesNormBase As Variant, ByVal cacheCodigos As Object) As String
    On Error GoTo Falha

    Dim nomeNorm As String
    Dim candidatos As Object
    Dim chave As Variant
    Dim i As Long
    Dim score As Double
    Dim melhorScore As Double
    Dim melhorCodigo As String
    Dim avaliouCandidato As Boolean

    nomeNorm = NormalizarNomeCliente(nomeDigitado)
    If nomeNorm = "" Then Exit Function

    If cacheCodigos.exists(nomeNorm) Then
        EncontrarCodigoCliente = cacheCodigos(nomeNorm)
        Exit Function
    End If

    ' Existem varios clientes cadastrados com o nome "FRANCISCO DE ASSIS BEZERRA"
    ' (codigos 005799, 004680, 001358, 000518). O de Morada Nova e sempre o 518.
    If nomeNorm = "FRANCISCO DE ASSIS BEZERRA" Then
        EncontrarCodigoCliente = "518"
        cacheCodigos.Add nomeNorm, EncontrarCodigoCliente
        Exit Function
    End If

    If dictCodigosExatos.exists(nomeNorm) Then
        EncontrarCodigoCliente = dictCodigosExatos(nomeNorm)
        cacheCodigos.Add nomeNorm, EncontrarCodigoCliente
        Exit Function
    End If

    Set candidatos = ObterCandidatosPorTokens(nomeNorm, dictIndiceTokens)
    For Each chave In candidatos.keys
        i = CLng(chave)
        avaliouCandidato = True
        score = PontuarSimilaridadeObvia(nomeNorm, CStr(nomesNormBase(i, 1)))
        If score > melhorScore Then
            melhorScore = score
            melhorCodigo = CStr(codigosBase(i, 1))
        End If
    Next chave

    If Not avaliouCandidato Or melhorScore < 80 Then
        For i = LBound(nomesNormBase, 1) To UBound(nomesNormBase, 1)
            score = PontuarSimilaridadeObvia(nomeNorm, CStr(nomesNormBase(i, 1)))
            If score > melhorScore Then
                melhorScore = score
                melhorCodigo = CStr(codigosBase(i, 1))
            End If
        Next i
    End If

    If melhorScore >= 80 Then EncontrarCodigoCliente = melhorCodigo
    cacheCodigos.Add nomeNorm, EncontrarCodigoCliente
    Exit Function

Falha:
    EncontrarCodigoCliente = ""
End Function

Private Function PontuarSimilaridadeObvia(ByVal a As String, ByVal b As String) As Double
    If a = "" Or b = "" Then Exit Function
    If a = b Then
        PontuarSimilaridadeObvia = 100
        Exit Function
    End If

    Dim palavrasA As Variant
    Dim palavrasB As Variant
    Dim i As Long, j As Long
    Dim tokenA As String
    Dim tokenB As String
    Dim pesoA As Double
    Dim somaA As Double
    Dim somaB As Double
    Dim comunsPesados As Double
    Dim comunsRelevantes As Long
    Dim melhorPeso As Double

    palavrasA = Split(a, " ")
    palavrasB = Split(b, " ")

    For i = LBound(palavrasA) To UBound(palavrasA)
        tokenA = CStr(palavrasA(i))
        If TokenClienteRelevante(tokenA) Then
            pesoA = 1#
        ElseIf Len(tokenA) >= 4 Then
            pesoA = 0.3
        Else
            pesoA = 0#
        End If
        somaA = somaA + pesoA
        If pesoA > 0 Then
            melhorPeso = 0#
            For j = LBound(palavrasB) To UBound(palavrasB)
                tokenB = CStr(palavrasB(j))
                If tokenA = tokenB Then
                    melhorPeso = pesoA
                    Exit For
                ElseIf pesoA = 1# And Len(tokenA) >= 4 And Len(tokenB) >= 4 And _
                       Abs(Len(tokenA) - Len(tokenB)) <= 1 Then
                    If LevenshteinDistancia(tokenA, tokenB) = 1 Then
                        If 0.85 > melhorPeso Then melhorPeso = 0.85
                    End If
                End If
            Next j
            If melhorPeso > 0 Then
                comunsPesados = comunsPesados + melhorPeso
                If melhorPeso >= 0.85 Then comunsRelevantes = comunsRelevantes + 1
            End If
        End If
    Next i

    For j = LBound(palavrasB) To UBound(palavrasB)
        tokenB = CStr(palavrasB(j))
        If TokenClienteRelevante(tokenB) Then
            somaB = somaB + 1#
        ElseIf Len(tokenB) >= 4 Then
            somaB = somaB + 0.3
        End If
    Next j

    If comunsRelevantes < 1 Then Exit Function
    If somaA = 0 Or somaB = 0 Then Exit Function

    PontuarSimilaridadeObvia = (0.7 * (comunsPesados / somaA) + 0.3 * (comunsPesados / somaB)) * 100
End Function

Private Function LevenshteinDistancia(ByVal s As String, ByVal t As String) As Long
    Dim m As Long, n As Long, i As Long, j As Long, custo As Long
    m = Len(s): n = Len(t)
    If m = 0 Then LevenshteinDistancia = n: Exit Function
    If n = 0 Then LevenshteinDistancia = m: Exit Function
    Dim d() As Long
    ReDim d(0 To m, 0 To n)
    For i = 0 To m: d(i, 0) = i: Next i
    For j = 0 To n: d(0, j) = j: Next j
    For j = 1 To n
        For i = 1 To m
            If Mid(s, i, 1) = Mid(t, j, 1) Then custo = 0 Else custo = 1
            d(i, j) = d(i - 1, j - 1) + custo
            If d(i - 1, j) + 1 < d(i, j) Then d(i, j) = d(i - 1, j) + 1
            If d(i, j - 1) + 1 < d(i, j) Then d(i, j) = d(i, j - 1) + 1
        Next i
    Next j
    LevenshteinDistancia = d(m, n)
End Function

Private Sub AdicionarNomeAoIndiceTokens(ByVal dictIndiceTokens As Object, ByVal nomeNorm As String, ByVal indice As Long)
    On Error Resume Next

    Dim palavras As Variant
    Dim token As String
    Dim i As Long
    Dim vistos As Object
    Dim col As Collection

    Set vistos = CreateObject("Scripting.Dictionary")
    palavras = Split(nomeNorm, " ")

    For i = LBound(palavras) To UBound(palavras)
        token = CStr(palavras(i))
        If TokenClienteRelevante(token) Then
            If Not vistos.exists(token) Then
                vistos.Add token, True
                If dictIndiceTokens.exists(token) Then
                    Set col = dictIndiceTokens(token)
                Else
                    Set col = New Collection
                    dictIndiceTokens.Add token, col
                End If
                col.Add indice
            End If
        End If
    Next i
End Sub

Private Function ObterCandidatosPorTokens(ByVal nomeNorm As String, ByVal dictIndiceTokens As Object) As Object
    Dim candidatos As Object
    Dim palavras As Variant
    Dim token As String
    Dim i As Long
    Dim indice As String
    Dim col As Collection
    Dim item As Variant

    Set candidatos = CreateObject("Scripting.Dictionary")
    palavras = Split(nomeNorm, " ")

    For i = LBound(palavras) To UBound(palavras)
        token = CStr(palavras(i))
        If TokenClienteRelevante(token) Then
            If dictIndiceTokens.exists(token) Then
                Set col = dictIndiceTokens(token)
                For Each item In col
                    indice = CStr(item)
                    If Not candidatos.exists(indice) Then candidatos.Add indice, True
                Next item
            End If
        End If
    Next i

    Set ObterCandidatosPorTokens = candidatos
End Function

Private Function NomeCadastroValido(ByVal nomeNorm As String) As Boolean
    If nomeNorm = "" Then Exit Function
    If InStr(1, nomeNorm, "NAO USAR", vbTextCompare) > 0 Then Exit Function
    If InStr(1, nomeNorm, "NAO UTILIZAR", vbTextCompare) > 0 Then Exit Function
    If InStr(1, nomeNorm, "CANCELAD", vbTextCompare) > 0 Then Exit Function
    NomeCadastroValido = True
End Function

Private Function TokenClienteRelevante(ByVal token As String) As Boolean
    token = Trim$(token)
    If Len(token) < 4 Then Exit Function

    Select Case token
        Case "JOSE", "JOAO", "MARIA", "ANA", "ANTONIO", "ANTONIA", "FRANCISCO", "FRANCISCA", "RAIMUNDO", "RAIMUNDA", "SILVA", "SOUSA", "SOUZA", "OLIVEIRA", "PEREIRA", "RODRIGUES", "COMERCIO", "TRANSPORTES", "AGROPECUARIA"
            TokenClienteRelevante = False
        Case Else
            TokenClienteRelevante = True
    End Select
End Function

Private Function NormalizarNomeCliente(ByVal texto As String) As String
    texto = UCase$(Trim$(texto))
    texto = Replace(texto, ChrW$(&HC1), "A")
    texto = Replace(texto, ChrW$(&HC0), "A")
    texto = Replace(texto, ChrW$(&HC2), "A")
    texto = Replace(texto, ChrW$(&HC3), "A")
    texto = Replace(texto, ChrW$(&HC9), "E")
    texto = Replace(texto, ChrW$(&HCA), "E")
    texto = Replace(texto, ChrW$(&HCD), "I")
    texto = Replace(texto, ChrW$(&HD3), "O")
    texto = Replace(texto, ChrW$(&HD4), "O")
    texto = Replace(texto, ChrW$(&HD5), "O")
    texto = Replace(texto, ChrW$(&HDA), "U")
    texto = Replace(texto, ChrW$(&HDC), "U")
    texto = Replace(texto, ChrW$(&HC7), "C")
    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "À", "A")
    texto = Replace(texto, "Â", "A")
    texto = Replace(texto, "Ã", "A")
    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "Ê", "E")
    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ô", "O")
    texto = Replace(texto, "Õ", "O")
    texto = Replace(texto, "Ú", "U")
    texto = Replace(texto, "Ç", "C")
    texto = Replace(texto, ".", " ")
    texto = Replace(texto, ",", " ")
    texto = Replace(texto, "-", " ")
    texto = Replace(texto, "/", " ")
    texto = Replace(texto, "\", " ")
    texto = Replace(texto, "(", " ")
    texto = Replace(texto, ")", " ")
    texto = Replace(texto, "&", " E ")
    texto = " " & texto & " "
    texto = Replace(texto, " LTDA ", " ")
    texto = Replace(texto, " ME ", " ")
    texto = Replace(texto, " EPP ", " ")
    texto = Replace(texto, " EIRELI ", " ")
    texto = Replace(texto, " S A ", " ")
    texto = Replace(texto, " SA ", " ")
    Do While InStr(texto, "  ") > 0
        texto = Replace(texto, "  ", " ")
    Loop
    NormalizarNomeCliente = Trim$(texto)
End Function

Private Function MaximoLong(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaximoLong = a Else MaximoLong = b
End Function

' ================================================================
' CODIGO DA MERCADORIA (coluna I nas linhas de titulo "SAIDA DE ...")
' ================================================================

' Preenche ws.Cells(linhaTitulo, "I") com o codigo da mercadoria em toda linha
' em negrito que bata com o padrao "SAIDA DE <mercadoria> (...)". Retorna a
' quantidade de titulos que ficaram sem codigo (para reportar ao usuario).
Private Function PreencherCodigosMercadoria(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal regex As Object, ByVal mapaMercadorias As Object, ByVal itensCarregados As Boolean, ByVal codigosItens As Variant, ByVal nomesNormItens As Variant, ByVal cacheItens As Object) As Long
    Dim j As Long
    Dim tituloLinha As String
    Dim matches As Object
    Dim mercadoria As String
    Dim nomeNorm As String
    Dim codigo As String
    Dim mensagemAmbiguidade As String
    Dim semCodigo As Long

    For j = 1 To lastRow
        If ws.Cells(j, "A").Font.Bold = True Then
            tituloLinha = Trim(ws.Cells(j, "A").Value)
            If regex.Test(tituloLinha) Then
                Set matches = regex.Execute(tituloLinha)
                mercadoria = Trim(matches(0).SubMatches(0))
                nomeNorm = NormalizarNomeMercadoria(mercadoria)

                codigo = ""
                mensagemAmbiguidade = ""

                If mapaMercadorias.exists(nomeNorm) Then
                    codigo = CStr(mapaMercadorias(nomeNorm))
                ElseIf itensCarregados Then
                    codigo = EncontrarCodigoItem(nomeNorm, codigosItens, nomesNormItens, cacheItens, mensagemAmbiguidade)
                End If

                If codigo <> "" Then
                    ws.Cells(j, "I").Value = codigo
                Else
                    semCodigo = semCodigo + 1
                    If mensagemAmbiguidade <> "" Then
                        AdicionarAlerta ws.Cells(j, "K"), "Mercadoria '" & mercadoria & "': " & mensagemAmbiguidade
                    Else
                        AdicionarAlerta ws.Cells(j, "K"), "Codigo da mercadoria nao encontrado para '" & mercadoria & "'"
                    End If
                End If
            End If
        End If
    Next j

    PreencherCodigosMercadoria = semCodigo
End Function

' Excecoes/mapeamentos manuais: nome da mercadoria (ja normalizado por
' NormalizarNomeMercadoria) -> codigo do item. Usado quando o nome do
' titulo nao bate (ou bate de forma ambigua) com o Cadastro de Itens.
' Para adicionar um novo mapeamento: mapa.Add "<NOME NORMALIZADO>", "<CODIGO>"
Private Function CriarMapaMercadorias() As Object
    Dim mapa As Object
    Set mapa = CreateObject("Scripting.Dictionary")

    ' Confirmados manualmente pelo usuario (2026-09-18):
    ' ATENCAO: estes 3 codigos NAO batem com o que esta hoje em "Cadastro de Itens.csv"
    ' (la o codigo 21 aparece como "TORTA DE CAROCO DE ALGODAO COM CASCA", o 15 como
    ' "RACAO BOVINO LACTACAO 22% ..." e o 1191 nem consta no arquivo, que vai ate ~1012).
    ' Ficaram assim porque foi o que foi informado - confirme se o CSV esta desatualizado
    ' antes de rodar em producao.
    ' Codigos com 10 digitos (zero a esquerda), no mesmo formato usado pelo
    ' Cadastro de Itens.csv, para a coluna I ficar consistente independente
    ' de a mercadoria ter sido resolvida por mapeamento manual ou por busca
    ' automatica no cadastro.
    mapa.Add "RESIDUO COMUM", "0000000021"
    mapa.Add "RESIDUO MOIDO", "0000001191"
    mapa.Add "RESIDUO MOIDA", "0000001191"
    mapa.Add "BALANCEADA", "0000000015"
    mapa.Add "RACAO BALANCEADA", "0000000015"

    ' Confirmados manualmente pelo usuario (2026-09-18, segunda rodada):
    mapa.Add "SOJA", "0000000004"          ' = CASCA DE SOJA
    mapa.Add "SORGO", "0000000022"         ' = SORGO TRITURADO - padrao e triturado
    mapa.Add "CAROCO", "0000000012"        ' = CAROCO DE ALGODAO
    mapa.Add "OLEO", "0000000002"          ' = OLEO DE ALGODAO BRUTO
    mapa.Add "XEREM", "0000000013"         ' = MILHO TRITURADO, mesmo produto
    mapa.Add "MILHO XEREM", "0000000013"   ' idem

    ' Confirmados manualmente pelo usuario (2026-09-18, terceira rodada) - racoes de
    ' suino, que no cadastro comecam com "RACAO" (ex.: "RACAO SUINO CRESCIMENTO"):
    mapa.Add "SUINO TERMINACAO", "0000001283"
    mapa.Add "SUINO CRESCIMENTO", "0000001798"
    mapa.Add "SUINO INICIAL", "0000001811"
    mapa.Add "SUINO LACTACAO", "0000008156"
    mapa.Add "SUINO GESTACAO", "0000003783"

    ' "BOVIGOLDE" (titulo) x "BOVIGOLD" (cadastro, codigo 594) - so difere por 1 letra,
    ' por isso precisa de mapeamento manual (o comparador de itens usa igualdade exata,
    ' de propósito - ver comentario em TokenPresenteEmColecao).
    mapa.Add "BOVIGOLDE", "0000000594"
    mapa.Add "SAL BOVIGOLDE", "0000000594"

    ' Bloqueados de proposito (ficam com codigo em branco + alerta na coluna K em vez
    ' de deixar o comparador escolher sozinho): titulos "CRESCIMENTO" e "INICIAL" sem
    ' prefixo (provavelmente SUINO CRESCIMENTO / SUINO INICIAL digitados incompletos)
    ' bateriam por engano com RACAO CAIPIRA CRESCIMENTO (0000000041) e RACAO FRANGO
    ' INICIAL (0000000847), que sao produtos de outra linha (aves), nao suino.
    mapa.Add "CRESCIMENTO", ""
    mapa.Add "INICIAL", ""

    ' Ainda SEM codigo (nao encontrados em Cadastro de Itens.csv - confirme o codigo,
    ' preencha com 10 digitos (com zeros a esquerda, ex.: "0000000123") para ficar no
    ' mesmo formato do restante da coluna I, e descomente a linha correspondente:
    ' mapa.Add "TORTA COMUM", ""
    ' mapa.Add "TORTA MOIDA", ""
    ' mapa.Add "TORTA MOIDO", ""
    ' mapa.Add "TERMINACAO", ""
    ' mapa.Add "SAL 35", ""
    ' mapa.Add "PRE PARTO", ""
    ' mapa.Add "POSTURA", ""
    ' mapa.Add "FRANGO POSTURA", ""

    ' AMBIGUOS no Cadastro de Itens (mais de um item com as mesmas palavras) -
    ' escolha o codigo certo e descomente:
    ' mapa.Add "DDGS", ""      ' candidatos: 0000000015, 0000000471, 0000000488 (racoes que citam DDGS na descricao)

    Set CriarMapaMercadorias = mapa
End Function

' Carrega o Cadastro de Itens (aba "Cadastro de Itens" ou CSV "Cadastro de Itens.csv"
' na mesma pasta), retornando os arrays paralelos codigosItens/nomesNormItens.
Private Function CarregarCadastroItens(ByRef codigosItens As Variant, ByRef nomesNormItens As Variant) As Boolean
    On Error GoTo Falha

    Dim linhas As Collection
    Dim i As Long
    Dim nome As String
    Dim nomeNorm As String
    Dim codigo As String

    Set linhas = New Collection

    Dim wsBase As Worksheet
    On Error Resume Next
    Set wsBase = ThisWorkbook.Sheets("Cadastro de Itens")
    On Error GoTo Falha

    If Not wsBase Is Nothing Then
        Dim lastRowBase As Long
        Dim startRow As Long
        lastRowBase = wsBase.Cells(wsBase.Rows.Count, "A").End(xlUp).Row
        If IsNumeric(wsBase.Cells(1, "A").Value) Then
            startRow = 1
        Else
            startRow = 2
        End If
        For i = startRow To lastRowBase
            codigo = Trim(CStr(wsBase.Cells(i, "A").Value))
            nome = Trim(CStr(wsBase.Cells(i, "B").Value))
            If codigo <> "" And nome <> "" Then
                nomeNorm = NormalizarNomeMercadoria(nome)
                If nomeNorm <> "" Then linhas.Add Array(codigo, nomeNorm)
            End If
        Next i
    Else
        Dim arquivoCsv As String
        Dim linha As String
        Dim campos As Variant
        Dim f As Integer

        arquivoCsv = EncontrarArquivoCadastroItens()
        If arquivoCsv = "" Then Exit Function

        f = FreeFile
        Open arquivoCsv For Input As #f
        Do Until EOF(f)
            Line Input #f, linha
            campos = SepararCsvPontoVirgula(linha)
            If IsArray(campos) Then
                If UBound(campos) >= 1 Then
                    codigo = Trim(CStr(campos(0)))
                    nome = Trim(CStr(campos(1)))
                    If codigo <> "" And nome <> "" Then
                        nomeNorm = NormalizarNomeMercadoria(nome)
                        If nomeNorm <> "" Then linhas.Add Array(codigo, nomeNorm)
                    End If
                End If
            End If
        Loop
        Close #f
    End If

    If linhas.Count = 0 Then Exit Function

    ReDim codigosItens(1 To linhas.Count, 1 To 1)
    ReDim nomesNormItens(1 To linhas.Count, 1 To 1)

    For i = 1 To linhas.Count
        codigosItens(i, 1) = linhas(i)(0)
        nomesNormItens(i, 1) = linhas(i)(1)
    Next i

    CarregarCadastroItens = True
    Exit Function

Falha:
    On Error Resume Next
    CarregarCadastroItens = False
End Function

Private Function EncontrarArquivoCadastroItens() As String
    Dim pasta As String
    Dim arquivo As String
    Dim melhorArquivo As String
    Dim melhorData As Date

    pasta = ThisWorkbook.Path
    If pasta = "" Then pasta = CurDir$

    arquivo = Dir(pasta & "\*.csv")
    Do While arquivo <> ""
        If InStr(1, arquivo, "Cadastro", vbTextCompare) > 0 And InStr(1, arquivo, "Iten", vbTextCompare) > 0 Then
            If FileDateTime(pasta & "\" & arquivo) > melhorData Then
                melhorData = FileDateTime(pasta & "\" & arquivo)
                melhorArquivo = pasta & "\" & arquivo
            End If
        End If
        arquivo = Dir()
    Loop

    EncontrarArquivoCadastroItens = melhorArquivo
End Function

' Busca o codigo do item cujo nome normalizado contem TODAS as palavras
' relevantes de nomeNorm (a mercadoria extraida do titulo), preferindo o
' item com menos palavras "extras" (mais proximo do nome procurado).
' Se houver mais de um item empatado no melhor resultado, NAO escolhe
' nenhum (evita cravar o codigo errado) e devolve a lista em mensagemAmbiguidade.
Private Function EncontrarCodigoItem(ByVal nomeNorm As String, ByVal codigosItens As Variant, ByVal nomesNormItens As Variant, ByVal cacheItens As Object, ByRef mensagemAmbiguidade As String) As String
    On Error GoTo Falha

    Dim tokensTitulo As Collection
    Dim tokensItem As Collection
    Dim tokenTit As Variant
    Dim i As Long
    Dim casados As Long
    Dim relevantesTitulo As Long
    Dim extras As Long
    Dim melhorExtras As Long
    Dim melhorCodigo As String
    Dim empatados As String
    Dim qtdMelhores As Long

    mensagemAmbiguidade = ""
    EncontrarCodigoItem = ""

    If nomeNorm = "" Then Exit Function

    If cacheItens.exists(nomeNorm) Then
        EncontrarCodigoItem = cacheItens(nomeNorm)
        Exit Function
    End If

    Set tokensTitulo = TokensRelevantesMercadoria(nomeNorm)
    relevantesTitulo = tokensTitulo.Count
    If relevantesTitulo = 0 Then Exit Function

    melhorExtras = -1
    qtdMelhores = 0

    For i = LBound(nomesNormItens, 1) To UBound(nomesNormItens, 1)
        Set tokensItem = TokensRelevantesMercadoria(CStr(nomesNormItens(i, 1)))

        casados = 0
        For Each tokenTit In tokensTitulo
            If TokenPresenteEmColecao(CStr(tokenTit), tokensItem) Then casados = casados + 1
        Next tokenTit

        If casados = relevantesTitulo Then
            extras = tokensItem.Count - casados
            If melhorExtras = -1 Or extras < melhorExtras Then
                melhorExtras = extras
                melhorCodigo = CStr(codigosItens(i, 1))
                qtdMelhores = 1
                empatados = melhorCodigo
            ElseIf extras = melhorExtras Then
                qtdMelhores = qtdMelhores + 1
                empatados = empatados & ", " & CStr(codigosItens(i, 1))
            End If
        End If
    Next i

    If qtdMelhores = 1 Then
        EncontrarCodigoItem = melhorCodigo
    ElseIf qtdMelhores > 1 Then
        mensagemAmbiguidade = "varios itens do cadastro correspondem (codigos: " & empatados & ")"
    End If

    cacheItens.Add nomeNorm, EncontrarCodigoItem
    Exit Function

Falha:
    EncontrarCodigoItem = ""
End Function

' Comparacao EXATA (sem tolerancia Levenshtein). Testado contra o cadastro real:
' com tolerancia de 1 letra, palavras curtas do titulo davam falso positivo em
' itens sem nenhuma relacao (ex.: "SOJA" casava com "SODA CAUSTICA", "CASCA"
' com "CALCA DEVOLVIDA", "POSTURA" com "MAQUINA DE COSTURA"). Sinonimos
' conhecidos (MOIDO/TRITURADO, GRAO/GRAOS) ja sao tratados em
' NormalizarNomeMercadoria antes de chegar aqui.
Private Function TokenPresenteEmColecao(ByVal token As String, ByVal colecao As Collection) As Boolean
    Dim item As Variant
    For Each item In colecao
        If token = CStr(item) Then
            TokenPresenteEmColecao = True
            Exit Function
        End If
    Next item
End Function

Private Function TokensRelevantesMercadoria(ByVal nomeNorm As String) As Collection
    Dim brutos As Variant
    Dim i As Long
    Dim col As New Collection

    brutos = Split(nomeNorm, " ")
    For i = LBound(brutos) To UBound(brutos)
        If CStr(brutos(i)) <> "" And Not EhStopwordMercadoria(CStr(brutos(i))) Then
            col.Add CStr(brutos(i))
        End If
    Next i

    Set TokensRelevantesMercadoria = col
End Function

Private Function EhStopwordMercadoria(ByVal token As String) As Boolean
    Select Case token
        Case "DE", "DA", "DO", "DAS", "DOS", "EM", "COM", "SEM", "E", "P", "PARA", "A", "O", "AO", "AOS"
            EhStopwordMercadoria = True
    End Select
End Function

' Normaliza o nome da mercadoria para comparacao: maiusculas, sem acento,
' pontuacao vira espaco, espacos duplicados colapsados. Tambem unifica
' sinonimos comuns entre o texto dos titulos e o Cadastro de Itens
' (ex.: a planilha usa "MOIDO/MOIDA", o cadastro usa "TRITURADO/TRITURADA").
Private Function NormalizarNomeMercadoria(ByVal texto As String) As String
    Dim palavras As Variant
    Dim i As Long

    texto = UCase$(Trim$(texto))
    texto = Replace(texto, ChrW$(&HC1), "A")
    texto = Replace(texto, ChrW$(&HC0), "A")
    texto = Replace(texto, ChrW$(&HC2), "A")
    texto = Replace(texto, ChrW$(&HC3), "A")
    texto = Replace(texto, ChrW$(&HC9), "E")
    texto = Replace(texto, ChrW$(&HCA), "E")
    texto = Replace(texto, ChrW$(&HCD), "I")
    texto = Replace(texto, ChrW$(&HD3), "O")
    texto = Replace(texto, ChrW$(&HD4), "O")
    texto = Replace(texto, ChrW$(&HD5), "O")
    texto = Replace(texto, ChrW$(&HDA), "U")
    texto = Replace(texto, ChrW$(&HDC), "U")
    texto = Replace(texto, ChrW$(&HC7), "C")
    texto = Replace(texto, "Á", "A")
    texto = Replace(texto, "À", "A")
    texto = Replace(texto, "Â", "A")
    texto = Replace(texto, "Ã", "A")
    texto = Replace(texto, "É", "E")
    texto = Replace(texto, "Ê", "E")
    texto = Replace(texto, "Í", "I")
    texto = Replace(texto, "Ó", "O")
    texto = Replace(texto, "Ô", "O")
    texto = Replace(texto, "Õ", "O")
    texto = Replace(texto, "Ú", "U")
    texto = Replace(texto, "Ç", "C")
    texto = Replace(texto, ".", " ")
    texto = Replace(texto, ",", " ")
    texto = Replace(texto, "-", " ")
    texto = Replace(texto, "/", " ")
    texto = Replace(texto, "\", " ")
    texto = Replace(texto, "(", " ")
    texto = Replace(texto, ")", " ")
    texto = Replace(texto, "+", " ")
    texto = Replace(texto, "*", " ")
    texto = Replace(texto, "%", " ")
    texto = Replace(texto, "&", " E ")

    Do While InStr(texto, "  ") > 0
        texto = Replace(texto, "  ", " ")
    Loop
    texto = Trim$(texto)

    If texto = "" Then
        NormalizarNomeMercadoria = ""
        Exit Function
    End If

    palavras = Split(texto, " ")
    For i = LBound(palavras) To UBound(palavras)
        Select Case CStr(palavras(i))
            Case "TRITURADO": palavras(i) = "MOIDO"
            Case "TRITURADA": palavras(i) = "MOIDA"
            Case "GRAOS": palavras(i) = "GRAO"
        End Select
    Next i

    NormalizarNomeMercadoria = Trim$(Join(palavras, " "))
End Function
