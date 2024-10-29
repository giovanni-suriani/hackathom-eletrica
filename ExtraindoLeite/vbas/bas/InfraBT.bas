Attribute VB_Name = "InfraBT"
Sub Verificacao()
    Dim resposta As VbMsgBoxResult
    
    resposta = MsgBox("O dimensionamento está sendo realizado para um eletrocentro?", vbYesNo, "Pergunta")
    
    If resposta = vbYes Then
        Call DimLeitos1
    Else
        Call DimLeitos2
    End If
End Sub

Sub DiametroExterno()
    Dim wsInfraBT As Worksheet
    Dim wsAuxInfra As Worksheet
    Dim i As Integer
    Set wsInfraBT = ThisWorkbook.Sheets("Dim Infra BT")
    Set wsAuxInfra = ThisWorkbook.Sheets("Aux_Infra")
    
    ultima_linha = wsInfraBT.Cells(wsInfraBT.Rows.Count, "D").End(xlUp).Row
    
    For i = 5 To ultima_linha
        tipocabo = wsInfraBT.Cells(i, "E").Value
        sistema = wsInfraBT.Cells(i, "F").Value
        secao = wsInfraBT.Cells(i, "C").Value
        If ((tipocabo = "Cabo multipolar" Or tipocabo = "Cabo multipolar em eletroduto de seção circular" Or tipocabo = "Cabo multipolar em eletroduto aparente de seção não circular") And (sistema = "3F + N" Or sistema = "3F + T" Or sistema = "3F")) Then
            For j = 3 To 17
                If secao = wsAuxInfra.Cells(j, "L").Value Then
                    wsInfraBT.Cells(i, "G").Value = wsAuxInfra.Cells(j, "M").Value
                    Exit For
                End If
            Next j
        End If
        If ((tipocabo = "Cabo multipolar" Or tipocabo = "Cabo multipolar em eletroduto de seção circular" Or tipocabo = "Cabo multipolar em eletroduto aparente de seção não circular") And sistema = "FF") Then
            For p = 3 To 17
                If secao = wsAuxInfra.Cells(p, "H").Value Then
                    wsInfraBT.Cells(i, "G").Value = wsAuxInfra.Cells(p, "I").Value
                    Exit For
                End If
            Next p
        End If
        If tipocabo = "Cabos unipolares em trifólio " Or tipocabo = "Condutores isolados ou cabos unipolares em eletroduto de seção circular" Or tipocabo = "Condutores isolados ou cabos unipolares em eletroduto aparente de seção não circular" Or tipocabo = "Condutores isolados ou cabos unipolares em eletroduto de seção circular" Or tipocabo = "Cabos unipolares" Or tipocabo = "Cabos unipolares em trifólio" Then
            For q = 3 To 17
                If secao = wsAuxInfra.Cells(q, "D").Value Then
                    wsInfraBT.Cells(i, "G").Value = wsAuxInfra.Cells(q, "E").Value
                    Exit For
                End If
            Next q
        End If
    Next i
End Sub
Sub LarguraOcupada()
    Dim wsInfraBT As Worksheet
    Dim wsCalcBT
    Dim tipocabo As String
    Dim i As Integer
    Set wsInfraBT = ThisWorkbook.Sheets("Dim Infra BT")
    Set wsCalcBT = ThisWorkbook.Sheets("Cálculos BT")
    Dim taxa As Integer
    
    taxa = InputBox("Qual a taxa de ocupação percentual você deseja que o leito tenha?")
    
    ultima_celula = wsInfraBT.Cells(wsInfraBT.Rows.Count, "D").End(xlUp).Row
    
    For i = 5 To ultima_celula
    
        tipocabo = wsInfraBT.Cells(i, "E").Value
        larocupada = wsInfraBT.Cells(i, "I").Value
        sistema = wsInfraBT.Cells(i, "F").Value
        quantcabos = wsInfraBT.Cells(i, "B").Value
        diametro = wsInfraBT.Cells(i, "G").Value
        
        taxa_p = taxa / 100
        
        If diametro <> "" Then
            De = diametro / 2
        End If
        If (tipocabo = "Cabo multipolar" Or tipocabo = "Cabo multipolar em eletroduto de seção circular" Or tipocabo = "Cabo multipolar em eletroduto aparente de seção não circular") And (sistema = "FF" Or sistema = "3F + N" Or sistema = "3F + T" Or sistema = "3F") Then
            wsInfraBT.Cells(i, "I").Value = (diametro * quantcabos + ((quantcabos + 1) * De)) * (1 + (1 - taxa_p))
        End If
        If (tipocabo = "Condutores isolados ou cabos unipolares em eletroduto de seção circular" Or tipocabo = "Condutores isolados ou cabos unipolares em eletroduto aparente de seção não circular" Or tipocabo = "Condutores isolados ou cabos unipolares em eletroduto de seção circular" Or tipocabo = "Cabos unipolares") And (sistema = "3F + N" Or sistema = "3F + T" Or sistema = "3F") Then
            wsInfraBT.Cells(i, "I").Value = ((quantcabos * diametro * 3) + ((quantcabos + 1) * De)) * (1 + (1 - taxa_p))
        End If
        If tipocabo = "Cabos unipolares em trifólio " And (sistema = "3F + N" Or sistema = "3F + T" Or sistema = "3F") Then
            wsInfraBT.Cells(i, "I").Value = ((quantcabos * 2 * diametro + ((quantcabos + 1) * De)) * (1 + (1 - taxa_p)))
        End If
    Next i
    
End Sub
Sub DimLeitos1()
    Dim wsInfraBT As Worksheet
    Dim x As Variant
    Dim x1() As Double ' Declarando x1 como um array
    Dim x2() As Double ' Declarando x2 como um array
    Dim y As Integer
    Dim proximoMaior As Integer
    Dim i As Integer
    Dim j As Integer
    Dim trecho1 As Variant
    Dim trecho2 As Variant
    Dim ultimaLinha1 As Long ' Adicionei a declaração da variável ultimaLinha1
    
    ' Definir a planilha de destino
    Set wsInfraBT = ThisWorkbook.Sheets("Dim Infra BT")
    Set wsCalculosBT = ThisWorkbook.Sheets("Cálculos BT")
    ultimaLinha1 = wsInfraBT.Cells(wsInfraBT.Rows.Count, "A").End(xlUp).Row
    
    taxa = InputBox("Qual a taxa de ocupação percentual você deseja que o leito tenha?")
    taxa_p = taxa / 100
    ' Definir o vetor x fora do loop
    x = Array(200, 400, 600)
    ' Loop para percorrer as colunas de trecho
    For i = 12 To 30
        ' Obter o valor do trecho na linha 1
        trecho1 = wsInfraBT.Cells(1, i).Value
        
        ' Verificar se trecho1 é numérico
        If IsNumeric(trecho1) Then
            ' Loop para percorrer as linhas de trecho
            For j = 5 To ultimaLinha1
                altura = (wsInfraBT.Cells(j, "H").Value) - 25
                trecho2 = wsInfraBT.Cells(j, "AG").Value
                ' Verificar se trecho2 é numérico
                If IsNumeric(trecho2) Then
                    ' Se os valores de trecho1 e trecho2 forem iguais, obter o valor de y
                    If trecho1 = trecho2 Then
                        y = wsInfraBT.Cells(3, i).Value
                        ' Redimensionar os arrays x1 e x2 para o mesmo tamanho de x
                        ReDim x1(LBound(x) To UBound(x))
                        ReDim x2(LBound(x) To UBound(x))
                        
                        ' Inicializar o valor de proximoMaior fora do loop
                        proximoMaior = 0
                        
                        ' Percorrer o vetor x
                        For q = LBound(x) To UBound(x)
                            x1(q) = (x(q) * altura) * taxa_p
                            If x1(q) > y Then
                                ' Se for maior, atribuir o valor atual de x a proximoMaior e sair do loop
                                proximoMaior = x(q)
                                Exit For
                            End If
                        Next q
                        
                        ' Verificar se encontrou um valor maior que y
                        If proximoMaior > 0 Then
                            ' Se encontrou, escrever o valor encontrado na coluna "AE"
                            wsInfraBT.Cells(j, "AH").Value = "1 nível de " & proximoMaior
                        Else
                            leito = 0
                            For p1 = LBound(x) To UBound(x)
                                x1(p1) = (x(p1) * altura) * taxa_p
                                For p2 = LBound(x) To UBound(x)
                                    x2(p2) = (x(p2) * altura) * taxa_p
                                    If x1(p1) + x2(p2) > y Then
                                        leito = x(p1) + x(p2)
                                        wsInfraBT.Cells(j, "AH").Value = "1 nível de " & x(p1) & " + " & "1 nível de " & x(p2)
                                        Exit For
                                    End If
                                Next p2
                            Next p1
                            If leito = 0 Then
                                wsInfraBT.Cells(j, "AH").Value = "Limite atingido. Redimensione o trecho!"
                            End If
                        End If
                    End If
                End If
            Next j
        End If
    Next i
End Sub
