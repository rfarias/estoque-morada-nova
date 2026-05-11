Option Explicit

' Macro independente para planilhas de estoque no formato Morada Nova.
' Uso:
' 1. Coloque a planilha de estoque na mesma pasta do arquivo CSV de cadastro de clientes.
' 2. No Excel, pressione ALT+F11 > Inserir > Modulo.
' 3. Cole este codigo no modulo.
' 4. Execute GerarResumoPedidos_ComCodigoCliente.
'
' O macro:
' - Le o primeiro CSV da pasta cujo nome contenha "Cadastro" e "Cliente".
' - Usa a coluna A do CSV como codigo e a coluna B como nome.
' - Preenche a coluna I com o codigo do cliente somente na primeira ocorrencia de cada pedido.
' - Gera o resumo do pedido na primeira ocorrencia de cada pedido, na coluna J.
' - Quando nao encontra cliente com confianca suficiente, deixa a coluna I em branco.
' - Grava alertas na coluna K quando detectar pedido com clientes diferentes ou preco bem fora do padrao.
' - Diferencas de ate R$ 3,00 no preco unitario nao geram alerta.

Private Const TOLERANCIA_PRECO_UNITARIO As Double = 3#

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

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Set dictPedidos = CreateObject("Scripting.Dictionary")
    Set linhasPrimeiraOcorrencia = CreateObject("Scripting.Dictionary")
    Set clientesPorPedido = CreateObject("Scripting.Dictionary")
    Set precoPorMercadoria = CreateObject("Scripting.Dictionary")
    Set precoPadraoPorMercadoria = CreateObject("Scripting.Dictionary")
    Set cacheCodigos = CreateObject("Scripting.Dictionary")
    Set regex = CreateObject("VBScript.RegExp")

    If Not CarregarCadastroClientes(dictCodigosExatos, dictIndiceTokens, nomesBase, codigosBase, nomesNormBase) Then
        MsgBox "Nao encontrei o CSV de cadastro de clientes na mesma pasta da planilha.", vbExclamation
        Exit Sub
    End If

    With regex
        .Global = False
        .IgnoreCase = True
        .Pattern = "SA[IÍ]DA DE (.+?) \([^\)]*\)\s*(\d{2}/\d{2}/\d{2,4})"
    End With

    ws.Range("I2:K" & lastRow).ClearContents

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

            If UCase(mercadoriaAtual) = "CAROÇO" Or UCase(mercadoriaAtual) = "CAROCO" Then
                unidade = "KG"
            Else
                unidade = "SACOS"
            End If

            If Not dictPedidos.exists(numPedido) Then
                dictPedidos.Add numPedido, ""
                linhasPrimeiraOcorrencia.Add numPedido, i
                clientesPorPedido.Add numPedido, NormalizarNomeCliente(cliente)
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

    For Each numPedido In dictPedidos.keys
        i = linhasPrimeiraOcorrencia(numPedido)
        cliente = ws.Cells(i, "B").Value
        formaPagamento = ws.Cells(i, "H").Value
        codigoCliente = EncontrarCodigoCliente(cliente, dictCodigosExatos, dictIndiceTokens, nomesBase, codigosBase, nomesNormBase, cacheCodigos)
        ws.Cells(i, "I").Value = codigoCliente

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

        resumo = cliente & " - " & numPedido & " - " & dataAtual & " - " & dictPedidos(numPedido) & " - " & formaPagamento
        ws.Cells(i, "J").Value = resumo
    Next numPedido

    totalAlertas = ContarAlertas(ws, lastRow)
    alerta = "Codigos dos clientes e resumo dos pedidos gerados com sucesso!"
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

    Dim arquivoCsv As String
    Dim linhas As Collection
    Dim linha As String
    Dim campos As Variant
    Dim f As Integer
    Dim i As Long
    Dim nome As String
    Dim nomeNorm As String
    Dim codigo As String

    arquivoCsv = EncontrarArquivoCadastro()
    If arquivoCsv = "" Then Exit Function

    Set dictCodigosExatos = CreateObject("Scripting.Dictionary")
    Set dictIndiceTokens = CreateObject("Scripting.Dictionary")
    Set linhas = New Collection

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
    Close #f
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

    nomeNorm = NormalizarNomeCliente(nomeDigitado)
    If nomeNorm = "" Then Exit Function

    If cacheCodigos.exists(nomeNorm) Then
        EncontrarCodigoCliente = cacheCodigos(nomeNorm)
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
        score = PontuarSimilaridadeObvia(nomeNorm, CStr(nomesNormBase(i, 1)))
        If score > melhorScore Then
            melhorScore = score
            melhorCodigo = CStr(codigosBase(i, 1))
        End If
    Next chave

    If melhorScore >= 80 Then EncontrarCodigoCliente = melhorCodigo
    cacheCodigos.Add nomeNorm, EncontrarCodigoCliente
    Exit Function

Falha:
    EncontrarCodigoCliente = ""
End Function

Private Function PontuarSimilaridadeObvia(ByVal a As String, ByVal b As String) As Double
    Dim palavrasA As Variant
    Dim palavrasB As Variant
    Dim i As Long, j As Long
    Dim token As String
    Dim baseToken As String
    Dim relevantesA As Long
    Dim relevantesB As Long
    Dim comuns As Long

    If a = "" Or b = "" Then Exit Function
    If a = b Then
        PontuarSimilaridadeObvia = 100
        Exit Function
    End If

    palavrasA = Split(a, " ")
    palavrasB = Split(b, " ")

    For i = LBound(palavrasA) To UBound(palavrasA)
        token = CStr(palavrasA(i))
        If TokenClienteRelevante(token) Then
            relevantesA = relevantesA + 1
            For j = LBound(palavrasB) To UBound(palavrasB)
                baseToken = CStr(palavrasB(j))
                If token = baseToken Then
                    comuns = comuns + 1
                    Exit For
                End If
            Next j
        End If
    Next i

    For j = LBound(palavrasB) To UBound(palavrasB)
        If TokenClienteRelevante(CStr(palavrasB(j))) Then relevantesB = relevantesB + 1
    Next j

    If comuns < 2 Then Exit Function
    If relevantesA = 0 Or relevantesB = 0 Then Exit Function

    PontuarSimilaridadeObvia = (comuns * 100#) / MaximoLong(relevantesA, relevantesB)
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
