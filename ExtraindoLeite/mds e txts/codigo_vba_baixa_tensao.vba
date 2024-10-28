Sub mainBT()

    Calculo_Corrente
    DescobrirMetodo
    CorAgrup
    CapacidadeCorrente
    CurtoCircuito
    QuedaTensao
    M_E
    NeT
    'Application.Wait (Now + TimeValue("00:00:02"))
    'VerificacaoTamanhoCabo
    
End Sub
Sub DescobrirMetodo()
    Dim wsCalculosBT As Worksheet
    Dim wsMetodoReferencia As Worksheet
    Dim wsCorrecaoAgrup As Worksheet
    Dim wsCorrecaoTemp As Worksheet
    Dim lastRowCalculosBT As Long
    Dim lastRowMetodoReferencia0 As Long
    Dim lastRowMetodoReferencia1 As Long
    Dim lastRowCorrecaoTemp As Long
    Dim i As Long, j As Long, p As Long
    Dim palavraM As String, palavraL As String
    Dim valorEncontrado As Variant
    Dim metodoEncontrado As Boolean
    Dim ValorN As Variant, valorO As Variant, valorI As Variant
    Dim valorD As Variant
    Dim texto As String
    
    'Descobrindo Métodos de Referência
    
    ' Defina as planilhas de trabalho
    Set wsCalculosBT = ThisWorkbook.Sheets("Cálculos BT")
    Set wsMetodoReferencia = ThisWorkbook.Sheets("MR_NBR5410")
    Set wsCorrecaoTemp = ThisWorkbook.Sheets("Correcao_Temp")
    Set wsCorrecaoAgrup = ThisWorkbook.Sheets("Cor_Agrup")
    
    ' Encontre a última linha com dados em ambas as planilhas
    lastRowCalculosBT = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "D").End(xlUp).Row ' literalmente a ultima linha com dados
    lastRowMetodoReferencia0 = wsMetodoReferencia.Cells(wsMetodoReferencia.Rows.Count, "A").End(xlUp).Row
    lastRowMetodoReferencia1 = wsMetodoReferencia.Cells(wsMetodoReferencia.Rows.Count, "B").End(xlUp).Row
    lastRowCorrecaoTemp = 30
    
    ' Percorrendo as células nas colunas M, N da planilha "Cálculos BT"
    For i = 9 To lastRowCalculosBT
        'Tipo de cabo
        palavraM = Trim(wsCalculosBT.Cells(i, "N").Value)
        'Maneira de Instalar
        palavraN = Trim(wsCalculosBT.Cells(i, "O").Value)
        metodoEncontrado = False
        
        ' Verificando se as células das colunas M e N estão preenchidas
        # se as palavras nao sao vazias entao
        If palavraN <> "" And palavraM <> "" Then
            isolacao = Trim(wsCalculosBT.Cells(i, "L").Value)
            temp = Trim(wsCalculosBT.Cells(i, "Q").Value)
            
            If isolacao = "" Or temp = "" Then
                wsCalculosBT.Cells(i, "K").Interior.Color = RGB(255, 0, 0)
                wsCalculosBT.Cells(i, "Q").Interior.Color = RGB(255, 0, 0)
                MsgBox "Preencha todos os campos!"
                GoTo ProximoRegistro
            End If
            
            'Limpar a formatação das células
            wsCalculosBT.Cells(i, "K").Interior.ColorIndex = xlNone
            wsCalculosBT.Cells(i, "Q").Interior.ColorIndex = xlNone
        End If
        
        ' Percorra as células na coluna B da planilha MR_NBR5410
        For j = 2 To lastRowMetodoReferencia0 ' Começa da linha 2 para ignorar cabeçalhos
            If InStr(1, wsMetodoReferencia.Cells(j, "A").Value, palavraM, vbTextCompare) > 0 _
                And InStr(1, wsMetodoReferencia.Cells(j, "B").Value, palavraN, vbTextCompare) > 0 Then
                
                ' Se todas as palavras forem encontradas na célula alvo
                valorEncontrado = wsMetodoReferencia.Cells(j, "B").Offset(0, 1).Value ' Valor à direita de B na planilha de referência
                wsCalculosBT.Cells(i, "P").Value = valorEncontrado ' Insira o valor na coluna O da planilha "Cálculos BT"
                texto = Replace(wsCalculosBT.Cells(i, "P").Value, vbLf, " ")
                wsCalculosBT.Cells(i, "P").Value = texto
                metodoEncontrado = True
                Exit For ' Não é necessário procurar mais nesta planilha de referência
            End If
        Next j
        
        ' Exibir mensagem de erro detalhada caso o método não seja encontrado
        If Not metodoEncontrado Then
           MsgBox "Método não encontrado. Verifique as condições!"
           Exit Sub
        End If
        

    ////////////////////////////////



        'Fazendo a Correção pelo Fator de Temperatura
        
        ' Verificar valores em coluna O da planilha "Cálculos BT"
        'Método de Referência
        valorO = wsCalculosBT.Cells(i, "P").Value
        'Temperatura
        temp = wsCalculosBT.Cells(i, "Q").Value
        'Isolação
        isolacao = wsCalculosBT.Cells(i, "K").Value
        
        If valorO = "B1" Or valorO = "B2" Or valorO = "C" Or valorO = "E" Or valorO = "F" Then
            For j = 3 To 17 ' Começa da linha 3 para ignorar cabeçalhos
                If wsCorrecaoTemp.Cells(j, "D").Value = temp Then
                    valorD = wsCorrecaoTemp.Cells(j, "D").Offset(0, 1).Value ' Valor de isolacao, 1 coluna a direita
                    If isolacao = "LSHF/A" Or isolacao = "PVC" Then
                        wsCalculosBT.Cells(i, "R").Value = valorD
                    ElseIf isolacao = "EPR/HEPR" Or isolacao = "XLPE" Then
                        wsCalculosBT.Cells(i, "R").Value = wsCorrecaoTemp.Cells(j, "D").Offset(0, 2).Value ' Segunda célula à direita de D na planilha de correção
                    End If
                    Exit For ' Não é necessário procurar mais neste range
                End If
            Next j
        End If

        If valorO = "A1" Or valorO = "A2" Then
            For p = 18 To 32 ' Começa da linha 3 para ignorar cabeçalhos
                If wsCorrecaoTemp.Cells(p, "D").Value = temp Then
                    valorD = wsCorrecaoTemp.Cells(p, "D").Offset(0, 1).Value ' Valor à direita de D na planilha de correção
                    If isolacao = "LSHF/A" Or isolacao = "PVC" Then
                        wsCalculosBT.Cells(i, "R").Value = valorD
                    ElseIf isolacao = "EPR/HEPR" Or isolacao = "XLPE" Then
                        wsCalculosBT.Cells(i, "R").Value = wsCorrecaoTemp.Cells(p, "D").Offset(0, 2).Value ' Segunda célula à direita de D na planilha de correção
                    End If
                    Exit For ' Não é necessário procurar mais neste range
                End If
            Next p
        End If


ProximoRegistro:
        ' Limpar a formatação das células vazias
        wsCalculosBT.Cells(i, "J").Interior.ColorIndex = xlNone
        wsCalculosBT.Cells(i, "Q").Interior.ColorIndex = xlNone
    Next i
    
    ' Limpar objetos
    Set wsCalculosBT = Nothing
    Set wsMetodoReferencia = Nothing
    Set wsCorrecaoTemp = Nothing
    
    

End Sub
Sub CorAgrup()
    Dim wsCalculosBT As Worksheet
    Dim wsCorAgrup As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim foundCell As Range
    Dim valorEncontrado As Variant
    Dim ValorN As String
    Dim valorQ As String
    Dim valorM As String
    Dim wsCamadas As Worksheet
    
    'Fazendo a Correção pelo Agrupamento dos Cabos
    
    Set wsCalculosBT = ThisWorkbook.Sheets("Cálculos BT")
    Set wsCorAgrup = ThisWorkbook.Sheets("Cor_Agrup")
    lastRow = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "S").End(xlUp).Row
    For i = 9 To lastRow
        camadas = wsCalculosBT.Cells(i, "S").Value
        metodo = wsCalculosBT.Cells(i, "P").Value
        circuitos = wsCalculosBT.Cells(i, "T").Value
        maneira_instalar = wsCalculosBT.Cells(i, "O").Value
        If camadas = 1 Then
        ' Verificação se o circuito está preenchido
            If circuitos = "" Then
                wsCalculosBT.Cells(i, "T").Interior.Color = RGB(255, 0, 0) ' Pinte a célula de vermelho
                MsgBox "Preencha todos os campos!"
                GoTo ProximoRegistro
            Else
                wsCalculosBT.Cells(i, "T").Interior.ColorIndex = xlNone ' Limpe a formatação da célula
            End If
        
            Set foundCell = wsCorAgrup.Rows(2).Find(What:=circuitos, LookIn:=xlValues, LookAt:=xlWhole)
            'Descobrindo o fator de correção
            If metodo = "A1" Or metodo = "A2" Or metodo = "B1" Or metodo = "B2" Or metodo = "E" Or metodo = "F" And maneira_instalar = "bandeja não-perfurada" Then
                If Not foundCell Is Nothing Then
                    valorEncontrado = foundCell.Offset(1, 0).Value ' 1 coluna  a baixo
                    wsCalculosBT.Cells(i, "V").Value = valorEncontrado
                End If
            End If
            If metodo = "C" Then
                If Not foundCell Is Nothing Then
                    valorEncontrado = foundCell.Offset(2, 0).Value
                    wsCalculosBT.Cells(i, "V").Value = valorEncontrado
                End If
            End If
            If metodo = "E" Or metodo = "F" Then
                If Not foundCell Is Nothing And maneira_instalar = "bandeja perfurada" Then
                    valorEncontrado = foundCell.Offset(4, 0).Value
                    wsCalculosBT.Cells(i, "V").Value = valorEncontrado
                End If
                If Not foundCell Is Nothing And maneira_instalar = "leito" Then
                    valorEncontrado = foundCell.Offset(5, 0).Value
                    wsCalculosBT.Cells(i, "V").Value = valorEncontrado
                End If
            End If
        End If
        If camadas = 2 Or camadas = 3 Then
            Set wsCamadas = ThisWorkbook.Sheets("Fator Camada")
            Dim quant_camadas As Range
            Dim linha_quant_camadas As Long
            Dim quant_circuitos As Range
            Dim coluna_quant_circuitos As Long
            
            Set quant_camadas = wsCamadas.Columns("B").Find(What:=camadas)
    
            ' Verifica se o valor foi encontrado
            If Not quant_camadas Is Nothing Then
                linha_quant_camadas = quant_camadas.Row
            End If
            If circuitos < 9 Then
                Set quant_circuitos = wsCamadas.Rows(2).Find(What:=circuitos)
                If Not quant_circuitos Is Nothing Then
                    coluna_quant_circuitos = quant_circuitos.Column
                End If
            End If
            
            If circuitos = 9 Or circuitos > 9 Then
                 coluna_quant_circuitos = 10
            End If
            Fator = wsCamadas.Cells(linha_quant_camadas, coluna_quant_circuitos).Value
            wsCalculosBT.Cells(i, "V").Value = Fator
        End If
    Next i
ProximoRegistro:
        ' Limpando a formatação das células
        wsCalculosBT.Cells(i, "S").Interior.ColorIndex = xlNone
End Sub
Sub CurtoCircuito()

'Dimensionamento pelo Método da Corrente de Curto Circuito

Dim wsCalculosBT As Worksheet
Dim wsCapCorrente As Worksheet
Dim wsCapCorrente2 As Worksheet
Dim corrente_curto As String
Dim tempo_op_prot As String
Dim ValorN As String, valorI As String
Dim i As Integer

'Definindo a pasta de trabalho
Set wsCalculosBT = ThisWorkbook.Sheets("Cálculos BT")
Set wsCapCorrente = ThisWorkbook.Sheets("Cap_Corrente")
Set wsCapCorrente2 = ThisWorkbook.Sheets("Cap_Corrente2")

Worksheets("Cálculos BT").Unprotect Password:="dvs*ito" 'Desprotege a planilha para a macro funcionar


'Ultima linha preenchida da planilha
LastRowCalcBT = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "J").End(xlUp).Row
num_cond = 1
limite_min = wsCalculosBT.Cells(5, 6).Value
    
    For i = 9 To LastRowCalcBT

        corrente_curto = wsCalculosBT.Cells(i, "L").Value
        tempo_op_prot = wsCalculosBT.Cells(i, "M").Value
        'metodo
        metodo = wsCalculosBT.Cells(i, "P").Value
        'isolacao
        valorI = wsCalculosBT.Cells(i, "K").Value
        
        num_cond = 1
        If corrente_curto <> "" And tempo_op_prot <> "" Then
            bitola = (1000 * corrente_curto * Sqr(tempo_op_prot)) / 143
            If (metodo = "A1" Or metodo = "A2" Or metodo = "B1" Or metodo = "B2" Or metodo = "C" Or metodo = "D") And (valorI = "PVC" Or valorI = "LSHF/A") Then
                For j = 7 To 30
                    If bitola > wsCapCorrente.Cells(j, "A").Value And bitola < wsCapCorrente.Cells(j + 1, "A").Value Then
                        bitola_escolhida = wsCapCorrente.Cells(j + 1, "A").Value
                        If bitola_escolhida < limite_min Then
                            wsCalculosBT.Cells(i, "AD").Value = limite_min
                            wsCalculosBT.Cells(i, "AC").Value = num_cond
                        Else
                            wsCalculosBT.Cells(i, "AD").Value = bitola_escolhida
                            wsCalculosBT.Cells(i, "AC").Value = num_cond
                        End If
                        Exit For
                    End If
                Next j
            End If
        
            If (metodo = "A1" Or metodo = "A2" Or metodo = "B1" Or metodo = "B2" Or metodo = "C" Or metodo = "D") And (valorI = "EPR/HEPR" Or valorI = "XLPE") Then
                For j = 56 To 79
                    If bitola > wsCapCorrente.Cells(j, "A").Value And bitola < wsCapCorrente.Cells(j + 1, "A").Value Then
                        bitola_escolhida = wsCapCorrente.Cells(j + 1, "A").Value
                        If bitola_escolhida < limite_min Then
                            wsCalculosBT.Cells(i, "AD").Value = limite_min
                            wsCalculosBT.Cells(i, "AC").Value = num_cond
                        Else
                            wsCalculosBT.Cells(i, "AD").Value = bitola_escolhida
                            wsCalculosBT.Cells(i, "AC").Value = num_cond
                        End If
                        Exit For
                    End If
                Next j
            End If
            
            If (metodo = "E" Or metodo = "F") And (valorI = "EPR/HEPR" Or valorI = "XLPE") Then
                For j = 7 To 30
                    If bitola > wsCapCorrente2.Cells(j, "A").Value And bitola < wsCapCorrente2.Cells(j + 1, "A").Value Then
                        bitola_escolhida = wsCapCorrente2.Cells(j + 1, "A").Value
                        If bitola_escolhida < limite_min Then
                            wsCalculosBT.Cells(i, "AD").Value = limite_min
                            wsCalculosBT.Cells(i, "AC").Value = num_cond
                        Else
                            wsCalculosBT.Cells(i, "AD").Value = wsCapCorrente2.Cells(j + 1, "A").Value
                            wsCalculosBT.Cells(i, "AC").Value = num_cond
                        End If
                        Exit For
                    End If
                Next j
            End If
                
            If (metodo = "E" Or metodo = "F") And (valorI = "PVC" Or valorI = "LSHF/A") Then
                For j = 38 To 61
                    If bitola > wsCapCorrente2.Cells(j, "A").Value And bitola < wsCapCorrente2.Cells(j + 1, "A").Value Then
                        bitola_escolhida = wsCapCorrente2.Cells(j + 1, "A").Value
                        If bitola_escolhida < limite_min Then
                            wsCalculosBT.Cells(i, "AD").Value = limite_min
                            wsCalculosBT.Cells(i, "AC").Value = num_cond
                        Else
                            wsCalculosBT.Cells(i, "AD").Value = wsCapCorrente2.Cells(j + 1, "A").Value
                            wsCalculosBT.Cells(i, "AC").Value = num_cond
                        End If
                        Exit For
                    End If
                Next j
            End If
            
        Else
            MsgBox "Preencha todos os campos para realizar essa operação!"
        End If
    Next i
End Sub
Sub QuedaTensao()

Dim wsCalculosBT As Worksheet
Dim wsQuedaTensao As Worksheet
Dim condutores As String
Dim QuedaTensao As Variant
Dim i As Integer


'Definindo as pastas de trabalho

Set wsCalculosBT = ThisWorkbook.Sheets("Cálculos BT")
Set wsQuedaTensao = ThisWorkbook.Sheets("Queda_Tensao")
limite = wsCalculosBT.Cells(6, 6).Value
bitola_limite = wsCalculosBT.Cells(4, 6).Value
bitola_limite_min = wsCalculosBT.Cells(5, 6).Value

'sen = Sqr(1 - fp ^ 2)

LastRowCalcBT = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "D").End(xlUp).Row

For i = 9 To LastRowCalcBT

     condutores = wsCalculosBT.Cells(i, "I").Value
     comprimento = wsCalculosBT.Cells(i, "J").Value
     fp = wsCalculosBT.Cells(i, "F").Value
     corrente = wsCalculosBT.Cells(i, "W").Value
     tensao = wsCalculosBT.Cells(i, "D").Value
     tipocabo = wsCalculosBT.Cells(i, "N").Value
     sistema = wsCalculosBT.Cells(i, "G").Value
     sen = Sqr(1 - fp ^ 2)
     num_cond = 1
     
Calculo_Queda1:
   
    If condutores = "2" And sistema = "FF" Then
        For j = 3 To 20
            secao = wsQuedaTensao.Cells(j, "A").Value
            Rca = wsQuedaTensao.Cells(j, "B").Value
            Xl = wsQuedaTensao.Cells(j, "C").Value
            QuedaTensao = ((2 * (((Rca * fp) + (Xl * sen))) * (corrente / num_cond) * (comprimento / 1000)) / tensao)
            
            If QuedaTensao <= limite Then
                If secao < bitola_limite_min Then
                    wsCalculosBT.Cells(i, "AG").Value = bitola_limite_min
                    wsCalculosBT.Cells(i, "AF").Value = num_cond
                    Exit For
                End If
            End If
    
            If QuedaTensao <= limite Then
                If secao > bitola_limite Then
                    num_cond = num_cond + 1
                    GoTo Calculo_Queda1
                Else
                    wsCalculosBT.Cells(i, "AG").Value = secao
                    wsCalculosBT.Cells(i, "AF").Value = num_cond
                    Exit For
                End If
            End If
            If secao >= bitola_limite Then
                num_cond = num_cond + 1
                GoTo Calculo_Queda1
                Exit For
            End If
        Next j
    End If
    
    num_cond1 = 1
    
Calculo_Queda2:
    
    If condutores = "3" And (sistema = "3F + N" Or sistema = "3F + T" Or sistema = "3F") And tipocabo = "Cabos unipolares em trifólio " Then
        For j = 3 To 20
            Rca = wsQuedaTensao.Cells(j, "F").Value
            Xl = wsQuedaTensao.Cells(j, "G").Value
            secao = wsQuedaTensao.Cells(j, "A").Value
            
            QuedaTensao = ((Sqr(3) * (((Rca * fp) + (Xl * sen))) * (corrente / num_cond1) * (comprimento / 1000)) / tensao)
            
            If QuedaTensao <= limite Then
                If secao < bitola_limite_min Then
                    wsCalculosBT.Cells(i, "AG").Value = bitola_limite_min
                    wsCalculosBT.Cells(i, "AF").Value = num_cond
                    Exit For
                End If
            End If
             
            If QuedaTensao <= limite Then
                If secao > bitola_limite Then
                    num_cond1 = num_cond1 + 1
                    GoTo Calculo_Queda2
                Else
                    wsCalculosBT.Cells(i, "AG").Value = secao
                    wsCalculosBT.Cells(i, "AF").Value = num_cond1
                    Exit For
                End If
            End If
            If secao >= bitola_limite Then
                num_cond1 = num_cond1 + 1
                GoTo Calculo_Queda2
                Exit For
            End If
        Next j
    End If
    
    num_cond2 = 1
        
Calculo_Queda3:

    If condutores = "3" And (sistema = "3F + N" Or sistema = "3F + T" Or sistema = "3F") And tipocabo <> "Cabos unipolares em trifólio " Then
        For j = 3 To 20
            secao = wsQuedaTensao.Cells(j, "A").Value
            Rca = wsQuedaTensao.Cells(j, "D").Value
            Xl = wsQuedaTensao.Cells(j, "E").Value
            
            QuedaTensao = ((Sqr(3) * (((Rca * fp) + (Xl * sen))) * (corrente / num_cond2) * (comprimento / 1000)) / tensao)
            
            If QuedaTensao <= limite Then
                If secao < bitola_limite_min Then
                    wsCalculosBT.Cells(i, "AG").Value = bitola_limite_min
                    wsCalculosBT.Cells(i, "AF").Value = num_cond
                    Exit For
                End If
            End If
             
            If QuedaTensao <= limite Then
                If secao > bitola_limite Then
                    num_cond2 = num_cond2 + 1
                    GoTo Calculo_Queda3
                Else
                    wsCalculosBT.Cells(i, "AG").Value = secao
                    wsCalculosBT.Cells(i, "AF").Value = num_cond2
                    Exit For
                End If
            End If
            
            If secao >= bitola_limite Then
                num_cond2 = num_cond2 + 1
                GoTo Calculo_Queda3
                Exit For
            End If
        Next j
    End If
    
    num_cond3 = 1
    
Calculo_Queda4:

    If sistema = "CC" Then
        For j = 3 To 20
            secao = wsQuedaTensao.Cells(j, "A").Value
            Rca = wsQuedaTensao.Cells(j, "D").Value
            Xl = wsQuedaTensao.Cells(j, "E").Value
            
            QuedaTensao = (2 * Rca * (corrente / num_cond3) * (comprimento / 1000)) / tensao
            
            If QuedaTensao <= limite Then
                If secao < bitola_limite_min Then
                    wsCalculosBT.Cells(i, "AG").Value = bitola_limite_min
                    wsCalculosBT.Cells(i, "AF").Value = num_cond
                    Exit For
                End If
            End If
             
            If QuedaTensao <= limite Then
                If secao > bitola_limite Then
                    num_cond3 = num_cond3 + 1
                    GoTo Calculo_Queda4
                Else
                    wsCalculosBT.Cells(i, "AG").Value = secao
                    wsCalculosBT.Cells(i, "AF").Value = num_cond3
                    Exit For
                End If
            End If
            
            If secao >= bitola_limite Then
                num_cond3 = num_cond3 + 1
                GoTo Calculo_Queda4
                Exit For
            End If
        Next j
    End If
Next i

End Sub
Sub CapacidadeCorrente()

    Dim wsCalculosBT As Worksheet
    Dim wsCapCorrente As Worksheet
    Dim wsCapCorrente2 As Worksheet
    Dim LastRowCalcBT As Long
    Dim num_cond As Integer
    Dim num_cond1 As Integer
    Dim num_cond2 As Integer
    Dim num_cond3 As Integer
    Dim i As Long
    Dim j As Long
    Dim q As Long
    Dim p As Long
    Dim coluna_aux As Integer
    Dim celulalimite As Long
    Dim limite As Double
    Dim limite_min As Double
    Dim metodo As String
    Dim isolacao As String
    Dim condutores As String
    Dim corrente As Double
    Dim tipocabo As String
    Dim material As String
    Dim nova_corrente As Double
    Dim secao_limite As Double
    Dim secao_encontrada As Double

    ' Definindo as planilhas
    Set wsCalculosBT = ThisWorkbook.Sheets("Cálculos BT")
    Set wsCapCorrente = ThisWorkbook.Sheets("Cap_Corrente")
    Set wsCapCorrente2 = ThisWorkbook.Sheets("Cap_Corrente2")
    
    ' Encontrando a última linha com dados na coluna H da planilha "Cálculos BT"
    LastRowCalcBT = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "H").End(xlUp).Row
    limite = wsCalculosBT.Cells(4, 6).Value
    limite_min = wsCalculosBT.Cells(5, 6).Value
    
    ' Loop para processar cada linha
    For i = 9 To LastRowCalcBT
    
        metodo = wsCalculosBT.Cells(i, "P").Value
        isolacao = wsCalculosBT.Cells(i, "K").Value
        condutores = wsCalculosBT.Cells(i, "I").Value
        corrente = wsCalculosBT.Cells(i, "W").Value
        tipocabo = wsCalculosBT.Cells(i, "N").Value
        material = wsCalculosBT.Cells(i, "U").Value
        If metodo <> "" And isolacao <> "" And condutores <> "" And corrente > 0 Then
        num_cond = 1

Calculo1:
            ' Verificação do tipo de isolação e método de instalação
            If (isolacao = "LSHF/A" Or isolacao = "PVC") And (metodo = "A1" Or metodo = "A2" Or metodo = "B1" Or metodo = "B2" Or metodo = "C" Or metodo = "D") Then
                ' Encontrando a coluna correta para busca
                For j = 2 To 13
                    If metodo = wsCapCorrente.Cells(3, j).Value And condutores = wsCapCorrente.Cells(5, j).Value Then
                        coluna_aux = j
                        Exit For
                    End If
                Next j
                
                ' Encontrando a célula correspondente à bitola limite
                If material = "Cobre" Then
                    For q = 7 To 30
                        If limite = wsCapCorrente.Cells(q, "A").Value Then
                            celulalimite = q
                            Exit For
                        End If
                    Next q
                ElseIf material = "Alumínio" Then
                    For q = 32 To 47
                        If limite = wsCapCorrente.Cells(q, "A").Value Then
                            celulalimite = q
                            Exit For
                        End If
                    Next q
                End If
                
                secao_limite = wsCapCorrente.Cells(celulalimite, "A").Value
                
                ' Processando para material "Cobre"
                If material = "Cobre" Then
                    For p = 7 To 30
                        nova_corrente = corrente / num_cond
                        If nova_corrente > wsCapCorrente.Cells(p, coluna_aux).Value And nova_corrente <= wsCapCorrente.Cells(p + 1, coluna_aux).Value Then
                            secao_encontrada = wsCapCorrente.Cells(p + 1, "A").Value
                            If secao_encontrada > secao_limite Then
                                num_cond = num_cond + 1
                                GoTo Calculo1
                            ElseIf secao_encontrada < limite_min Then
                                wsCalculosBT.Cells(i, "AA").Value = limite_min
                                wsCalculosBT.Cells(i, "Z").Value = num_cond
                            Else
                                wsCalculosBT.Cells(i, "AA").Value = secao_encontrada
                                wsCalculosBT.Cells(i, "Z").Value = num_cond
                            End If
                            Exit For
                        End If
                    Next p
                ElseIf material = "Alumínio" Then
                    ' Processando para material "Alumínio"
                    For p = 32 To 47
                        nova_corrente = corrente / num_cond
                        If nova_corrente > wsCapCorrente.Cells(p, coluna_aux).Value And nova_corrente <= wsCapCorrente.Cells(p + 1, coluna_aux).Value Then
                            secao_encontrada = wsCapCorrente.Cells(p + 1, "A").Value
                            If secao_encontrada > secao_limite Then
                                num_cond = num_cond + 1
                                GoTo Calculo1
                            ElseIf secao_encontrada < limite_min Then
                                wsCalculosBT.Cells(i, "AA").Value = limite_min
                                wsCalculosBT.Cells(i, "Z").Value = num_cond
                            Else
                                wsCalculosBT.Cells(i, "AA").Value = secao_encontrada
                                wsCalculosBT.Cells(i, "Z").Value = num_cond
                            End If
                            Exit For
                        End If
                    Next p
                End If
            End If
            ' Calculo2
            num_cond1 = 1
Calculo2:
            If (isolacao = "EPR/HEPR" Or isolacao = "XLPE") And (metodo = "A1" Or metodo = "A2" Or metodo = "B1" Or metodo = "B2" Or metodo = "C" Or metodo = "D") Then
                ' Encontrando a coluna correta para fazer a busca
                For j = 2 To 13
                    If metodo = wsCapCorrente.Cells(52, j).Value And condutores = wsCapCorrente.Cells(54, j).Value Then
                        coluna_aux = j
                        Exit For
                    End If
                Next j

                ' Encontrando a célula correspondente à bitola limite
                If material = "Cobre" Then
                    For q = 56 To 79
                        If limite = wsCapCorrente.Cells(q, "A").Value Then
                            celulalimite = q
                            Exit For
                        End If
                    Next q
                ElseIf material = "Alumínio" Then
                    For q = 81 To 96
                        If limite = wsCapCorrente.Cells(q, "A").Value Then
                            celulalimite = q
                            Exit For
                        End If
                    Next q
                End If

                secao_limite = wsCapCorrente.Cells(celulalimite, "A").Value

                ' Processamento para material "Cobre"
                If material = "Cobre" Then
                    For p = 56 To 79 ' Encontrar na planilha de tabelas a bitola correspondente ao método de instalação e corrente calculada
                        nova_corrente = corrente / num_cond1
                        If nova_corrente > wsCapCorrente.Cells(p, coluna_aux).Value And nova_corrente <= wsCapCorrente.Cells(p + 1, coluna_aux).Value Then
                            secao_encontrada = wsCapCorrente.Cells(p + 1, "A").Value
                            If secao_encontrada > secao_limite Then
                                num_cond1 = num_cond1 + 1
                                GoTo Calculo2
                            ElseIf secao_encontrada < limite_min Then
                                wsCalculosBT.Cells(i, "AA").Value = limite_min
                                wsCalculosBT.Cells(i, "Z").Value = num_cond1
                            Else
                                wsCalculosBT.Cells(i, "AA").Value = secao_encontrada
                                wsCalculosBT.Cells(i, "Z").Value = num_cond1
                                Exit For
                            End If
                            Exit For
                        End If
                        If nova_corrente > wsCapCorrente.Cells(79, coluna_aux).Value Then
                            num_cond1 = num_cond1 + 1
                            GoTo Calculo2
                        End If
                    Next p
                ElseIf material = "Alumínio" Then
                    ' Processando para material "Alumínio"
                    For p = 81 To 96
                        nova_corrente = corrente / num_cond1
                        If nova_corrente > wsCapCorrente.Cells(p, coluna_aux).Value And nova_corrente <= wsCapCorrente.Cells(p + 1, coluna_aux).Value Then
                            secao_encontrada = wsCapCorrente.Cells(p + 1, "A").Value
                            If secao_encontrada > secao_limite Then
                                num_cond1 = num_cond1 + 1
                                GoTo Calculo2
                            ElseIf secao_encontrada < limite_min Then
                                wsCalculosBT.Cells(i, "AA").Value = limite_min
                                wsCalculosBT.Cells(i, "Z").Value = num_cond1
                            Else
                                wsCalculosBT.Cells(i, "AA").Value = secao_encontrada
                                wsCalculosBT.Cells(i, "Z").Value = num_cond1
                            End If
                            Exit For
                        End If
                    Next p
                End If
            End If
            ' Calculo3
            num_cond2 = 1
Calculo3:
            If (isolacao = "LSHF/A" Or isolacao = "PVC") And (metodo = "E" Or metodo = "F") Then
                ' Encontrando a coluna correta para fazer a busca
                For j = 2 To 8
                    If metodo = wsCapCorrente2.Cells(4, j).Value And condutores = wsCapCorrente2.Cells(5, j).Value Then
                        coluna_aux = j
                        Exit For
                    End If
                Next j

                ' Encontrando a célula correspondente à bitola limite
                If material = "Cobre" Then
                    For q = 7 To 30
                        If limite = wsCapCorrente2.Cells(q, "A").Value Then
                            celulalimite = q
                            Exit For
                        End If
                    Next q
                ElseIf material = "Alumínio" Then
                    For q = 32 To 47
                        If limite = wsCapCorrente2.Cells(q, "A").Value Then
                            celulalimite = q
                            Exit For
                        End If
                    Next q
                End If

                secao_limite = wsCapCorrente2.Cells(celulalimite, "A").Value

                ' Processamento para material "Cobre"
                If material = "Cobre" Then
                    For p = 7 To 30 ' Encontrar na planilha de tabelas a bitola correspondente ao método de instalação e corrente calculada
                        nova_corrente = corrente / num_cond2
                        If nova_corrente > wsCapCorrente2.Cells(p, coluna_aux).Value And nova_corrente <= wsCapCorrente2.Cells(p + 1, coluna_aux).Value Then
                            secao_encontrada = wsCapCorrente2.Cells(p + 1, "A").Value
                            If secao_encontrada > secao_limite Then
                                num_cond2 = num_cond2 + 1
                                GoTo Calculo3
                            ElseIf secao_encontrada < limite_min Then
                                wsCalculosBT.Cells(i, "AA").Value = limite_min
                                wsCalculosBT.Cells(i, "Z").Value = num_cond2
                            Else
                                wsCalculosBT.Cells(i, "AA").Value = secao_encontrada
                                wsCalculosBT.Cells(i, "Z").Value = num_cond2
                                Exit For
                            End If
                            Exit For
                        End If
                        If nova_corrente > wsCapCorrente2.Cells(30, coluna_aux).Value Then
                            num_cond2 = num_cond2 + 1
                            GoTo Calculo3
                        End If
                    Next p
                ElseIf material = "Alumínio" Then
                    ' Processando para material "Alumínio"
                    For p = 32 To 47
                        nova_corrente = corrente / num_cond2
                        If nova_corrente > wsCapCorrente2.Cells(p, coluna_aux).Value And nova_corrente <= wsCapCorrente2.Cells(p + 1, coluna_aux).Value Then
                            secao_encontrada = wsCapCorrente2.Cells(p + 1, "A").Value
                            If secao_encontrada > secao_limite Then
                                num_cond2 = num_cond2 + 1
                                GoTo Calculo3
                            ElseIf secao_encontrada < limite_min Then
                                wsCalculosBT.Cells(i, "AA").Value = limite_min
                                wsCalculosBT.Cells(i, "Z").Value = num_cond2
                            Else
                                wsCalculosBT.Cells(i, "AA").Value = secao_encontrada
                                wsCalculosBT.Cells(i, "Z").Value = num_cond2
                            End If
                            Exit For
                        End If
                    Next p
                End If
            End If
            ' Calculo4
            num_cond3 = 1
Calculo4:
            If (isolacao = "EPR/HEPR" Or isolacao = "XLPE") And (metodo = "E" Or metodo = "F") Then
                ' Encontrando a coluna correta para fazer a busca
                For j = 2 To 8
                    If metodo = wsCapCorrente2.Cells(52, j).Value And condutores = wsCapCorrente2.Cells(53, j).Value Then
                        coluna_aux = j
                        Exit For
                    End If
                Next j

                ' Encontrando a célula correspondente à bitola limite
                If material = "Cobre" Then
                    For q = 55 To 78
                        If limite = wsCapCorrente2.Cells(q, "A").Value Then
                            celulalimite = q
                            Exit For
                        End If
                    Next q
                ElseIf material = "Alumínio" Then
                    For q = 80 To 95
                        If limite = wsCapCorrente2.Cells(q, "A").Value Then
                            celulalimite = q
                            Exit For
                        End If
                    Next q
                End If

                secao_limite = wsCapCorrente2.Cells(celulalimite, "A").Value

                ' Processamento para material "Cobre"
                If material = "Cobre" Then
                    For p = 55 To 78 ' Encontrar na planilha de tabelas a bitola correspondente ao método de instalação e corrente calculada
                        nova_corrente = corrente / num_cond3
                        If nova_corrente > wsCapCorrente2.Cells(p, coluna_aux).Value And nova_corrente <= wsCapCorrente2.Cells(p + 1, coluna_aux).Value Then
                            secao_encontrada = wsCapCorrente2.Cells(p + 1, "A").Value
                            If secao_encontrada > secao_limite Then
                                num_cond3 = num_cond3 + 1
                                GoTo Calculo4
                            ElseIf secao_encontrada < limite_min Then
                                wsCalculosBT.Cells(i, "AA").Value = limite_min
                                wsCalculosBT.Cells(i, "Z").Value = num_cond3
                            Else
                                wsCalculosBT.Cells(i, "AA").Value = secao_encontrada
                                wsCalculosBT.Cells(i, "Z").Value = num_cond3
                                Exit For
                            End If
                            Exit For
                        End If
                        If nova_corrente > wsCapCorrente2.Cells(78, coluna_aux).Value Then
                            num_cond3 = num_cond3 + 1
                            GoTo Calculo4
                        End If
                    Next p
                ElseIf material = "Alumínio" Then
                    ' Processando para material "Alumínio"
                    For p = 80 To 95
                        nova_corrente = corrente / num_cond3
                        If nova_corrente > wsCapCorrente2.Cells(p, coluna_aux).Value And nova_corrente <= wsCapCorrente2.Cells(p + 1, coluna_aux).Value Then
                            secao_encontrada = wsCapCorrente2.Cells(p + 1, "A").Value
                            If secao_encontrada > secao_limite Then
                                num_cond3 = num_cond3 + 1
                                GoTo Calculo4
                            ElseIf secao_encontrada < limite_min Then
                                wsCalculosBT.Cells(i, "AA").Value = limite_min
                                wsCalculosBT.Cells(i, "Z").Value = num_cond3
                            Else
                                wsCalculosBT.Cells(i, "AA").Value = secao_encontrada
                                wsCalculosBT.Cells(i, "Z").Value = num_cond3
                            End If
                            Exit For
                            If nova_corrente > wsCapCorrente2.Cells(95, coluna_aux).Value Then
                                num_cond3 = num_cond3 + 1
                                GoTo Calculo4
                            End If
                        End If
                    Next p
                End If
            End If
        End If
    Next i
End Sub


Sub NeT()

    Dim wsCalculosBT As Worksheet
    Set wsCalculosBT = ThisWorkbook.Sheets("Cálculos BT")
    Dim i As Integer

    
    ' Ultima linha da seção nominal
    LastRowCalcBT = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "Z").End(xlUp).Row
    
    For i = 9 To LastRowCalcBT
        
        sistema = wsCalculosBT.Cells(i, "G").Value
        secao = wsCalculosBT.Cells(i, "AM").Value
        
        If sistema = "" Or sistema = "3F" Or sistema = "FF" Or sistema = "CC" Then
            wsCalculosBT.Cells(i, "AI").Value = "-"
            wsCalculosBT.Cells(i, "AJ").Value = "-"
        End If
        If sistema = "3F + T" And secao <> "" Then
            If secao <= 16 Then
                wsCalculosBT.Cells(i, "AI").Value = secao
                wsCalculosBT.Cells(i, "AJ").Value = "-"
            ElseIf secao > 16 And secao <= 35 Then
                wsCalculosBT.Cells(i, "AI").Value = 16
                wsCalculosBT.Cells(i, "AJ").Value = "-"
            ElseIf secao > 35 Then
                wsCalculosBT.Cells(i, "AI").Value = secao / 2
                wsCalculosBT.Cells(i, "AJ").Value = "-"
            End If
        End If
        
        If sistema = "3F + N" And secao <> "" Then
            If secao <= 25 Then
                wsCalculosBT.Cells(i, "AI").Value = "-"
                wsCalculosBT.Cells(i, "AJ").Value = secao
            ElseIf secao = 35 Or secao = 50 Then
                wsCalculosBT.Cells(i, "AI").Value = "-"
                wsCalculosBT.Cells(i, "AJ").Value = 25
            ElseIf secao = 70 Then
                wsCalculosBT.Cells(i, "AI").Value = "-"
                wsCalculosBT.Cells(i, "AJ").Value = 35
            ElseIf secao = 95 Then
                wsCalculosBT.Cells(i, "AI").Value = "-"
                wsCalculosBT.Cells(i, "AJ").Value = 50
            ElseIf secao = 120 Or secao = 150 Then
                wsCalculosBT.Cells(i, "AI").Value = "-"
                wsCalculosBT.Cells(i, "AJ").Value = 70
            ElseIf secao = 185 Then
                wsCalculosBT.Cells(i, "AI").Value = "-"
                wsCalculosBT.Cells(i, "AJ").Value = 95
            ElseIf secao = 240 Then
                wsCalculosBT.Cells(i, "AI").Value = "-"
                wsCalculosBT.Cells(i, "AJ").Value = 120
            ElseIf secao = 300 Then
                wsCalculosBT.Cells(i, "AI").Value = "-"
                wsCalculosBT.Cells(i, "AJ").Value = 150
            ElseIf secao = 400 Then
                wsCalculosBT.Cells(i, "AH").Value = "-"
                wsCalculosBT.Cells(i, "AI").Value = 185
            ElseIf secao > 400 Then
                wsCalculosBT.Cells(i, "AI").Value = "-"
                wsCalculosBT.Cells(i, "AJ").Value = secao / 2
            End If
        End If
        
    Next i
    
    
    
End Sub
Sub Calculo_Corrente()

    Dim wsCalculosBT As Worksheet
    Set wsCalculosBT = ThisWorkbook.Sheets("Cálculos BT")
    Dim i As Integer
    
    LastRowCalcBT = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "D").End(xlUp).Row
    
    For i = 9 To LastRowCalcBT
    
        tensao = wsCalculosBT.Cells(i, "D").Value
        potencia = wsCalculosBT.Cells(i, "E").Value
        fp = wsCalculosBT.Cells(i, "F").Value
        sistema = wsCalculosBT.Cells(i, "G").Value
        
        If sistema = "3F + N" Or sistema = "3F + T" Or sistema = "3F" Then
            corrente = (potencia * 1000) / (Sqr(3) * tensao * fp)
            wsCalculosBT.Cells(i, "H").Value = corrente
        ElseIf sistema = "CC" Then
            corrente = (potencia * 1000) / tensao
            wsCalculosBT.Cells(i, "H").Value = corrente
        ElseIf sistema = "FF" Then
            corrente = (potencia * 1000) / (tensao * fp)
            wsCalculosBT.Cells(i, "H").Value = corrente
        End If
        
    Next i
 
End Sub
Sub M_E()

Dim wsCalculosBT As Worksheet
Set wsCalculosBT = ThisWorkbook.Sheets("Cálculos BT")
Dim i As Integer

LastRowCalcBT = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "D").End(xlUp).Row

For i = 9 To LastRowCalcBT

    'quantidade
    quantidade_cap_corrente = wsCalculosBT.Cells(i, "Z").Value
    'secao
    cap_corrente = wsCalculosBT.Cells(i, "AA").Value
    'secao
    curto = wsCalculosBT.Cells(i, "AD").Value
    'quantidade
    quantidade_curto = wsCalculosBT.Cells(i, "AC").Value
    'quantidade
    quantidade_queda_tensao = wsCalculosBT.Cells(i, "AF").Value
    'secao
    queda_tensao = wsCalculosBT.Cells(i, "AG").Value

    mult_cap_corrente = quantidade_cap_corrente * cap_corrente
    mult_curto = quantidade_curto * curto
    mult_queda_tensao = quantidade_queda_tensao * queda_tensao

    If mult_cap_corrente >= mult_curto And mult_cap_corrente >= mult_queda_tensao Then
    
        wsCalculosBT.Cells(i, "AL").Value = quantidade_cap_corrente
        wsCalculosBT.Cells(i, "AM").Value = cap_corrente
        
    ElseIf mult_curto >= mult_cap_corrente And mult_curto >= mult_queda_tensao Then
   
        wsCalculosBT.Cells(i, "AL").Value = quantidade_curto
        wsCalculosBT.Cells(i, "AM").Value = curto
        
    ElseIf mult_queda_tensao >= mult_cap_corrente And mult_queda_tensao >= mult_curto Then
   
        wsCalculosBT.Cells(i, "AL").Value = quantidade_queda_tensao
        wsCalculosBT.Cells(i, "AM").Value = queda_tensao
        
    End If
    
Next i
    
End Sub
Sub VerificacaoTamanhoCabo()

    Dim ws As Worksheet
    Dim i As Integer

    Set ws = ThisWorkbook.Sheets("Cálculos BT")

    LastRowCalcBT = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row

For i = 7 To LastRowCalcBT

    Tipo_Cabo = ws.Cells(i, "N").Value
    secao = ws.Cells(i, "AL").Value
    
    If (Tipo_Cabo = "Cabo multipolar" Or Tipo_Cabo = "Cabo multipolar em eletroduto de seção circular" Or Tipo_Cabo = "Cabo multipolar em eletroduto aparente de seção não circular") And secao > 70 Then
        MsgBox "Cuidado! Não é recomendado utilizar cabos multipolares para seções nominais maiores que 70 mm²."
    End If
Next i
End Sub