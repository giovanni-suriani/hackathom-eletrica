Attribute VB_Name = "Módulo7"
Sub ListaCabosMT()
    Dim wsLista As Worksheet
    Dim wsCalculosMT As Worksheet
    Dim lastRow As Long
    Dim i As Integer
    Dim colunas As Variant
    
    Set wsLista = ThisWorkbook.Sheets("Lista de Cabos")
    Set wsCalculosMT = ThisWorkbook.Sheets("Cálculos MT")
    
    colunas = Array("C", "I", "H", "M", "AD", "AE", "L")
    
    For i = 0 To UBound(colunas)
        lastRow = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, colunas(i)).End(xlUp).Row
        wsCalculosMT.Range(colunas(i) & "10:" & colunas(i) & lastRow).Copy
        wsLista.Cells(2, i + 1).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
    Next i
    
    lastRowAG = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AG").End(xlUp).Row
    lastRowAH = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AH").End(xlUp).Row
    'neutro
    wsCalculosMT.Range("AH10:AH" & lastRowAH).Copy
    wsLista.Range("I2").PasteSpecial Paste:=xlPasteValues
    'terra
    wsCalculosMT.Range("AG10:AG" & lastRowAG).Copy
    wsLista.Range("K2").PasteSpecial Paste:=xlPasteValues
End Sub
Sub ComprimentoTotal()

Dim wsLista As Worksheet
Dim wsCalculosMT As Worksheet
Dim i As Integer
Dim j As Integer

Set wsLista = ThisWorkbook.Sheets("Lista de Cabos")
Set wsCalculosMT = ThisWorkbook.Sheets("Cálculos MT")
Set wsCalculosBT = ThisWorkbook.Sheets("Cálculos BT")

LastRowE = wsLista.Cells(wsLista.Rows.Count, "E").End(xlUp).Row
lastRowM = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "M").End(xlUp).Row
ultimalinhaM = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "M").End(xlUp).Row
LastRowH = wsLista.Cells(wsLista.Rows.Count, "H").End(xlUp).Row
LastRowG = wsLista.Cells(wsLista.Rows.Count, "G").End(xlUp).Row

For i = 10 To lastRowM
    tipocabo = wsCalculosMT.Cells(i, "M").Value
    If tipocabo = "3 Cabos Unipolares Em Plano Separados S=2D" Or tipocabo = "3 Cabos Unipolares Justapostos Em Trifólio" Then
        a = 3
    End If
    If tipocabo = "Cabo Tripolar" Then
        a = 1
    End If
    
    For j = 2 To LastRowE
        cabosporfase = wsLista.Cells(j, "E").Value
        compcir = wsLista.Cells(j, "G").Value
        wsLista.Cells(j, "H").Value = compcir * a * cabosporfase
    Next j
Next i

For p = 9 To ultimalinhaM
    tipocabo2 = wsCalculosBT.Cells(j, "M").Value
    sistema = wsCalculosBT.Cells(j, "F").Value
    
    If (tipocabo2 = "Condutores isolados ou cabos unipolares em eletroduto de seção circular" Or tipocabo2 = "Condutores isolados ou cabos unipolares em eletroduto aparente de seção não circular" Or tipocabo2 = "Condutores isolados ou cabos unipolares em eletroduto de seção circular" Or tipocabo2 = "Cabos unipolares" Or tipocabo2 = "Cabos unipolares em trifólio") And (sistema = "3F" Or sistema = "3F + N" Or sistema = "3F + T") Then
        a = 3
    End If
    
    If (tipocabo2 = "Cabo multipolar" Or tipocabo2 = "Cabo multipolar em eletroduto de seção circular" Or tipocabo2 = "Cabo multipolar em eletroduto aparente de seção não circular" Or tipocabo2 = "Cabo multipolar em eletroduto aparente de seção não circular") Then
        a = 1
    End If
    
    If (tipocabo2 = "Condutores isolados ou cabos unipolares em eletroduto de seção circular" Or tipocabo2 = "Condutores isolados ou cabos unipolares em eletroduto aparente de seção não circular" Or tipocabo2 = "Condutores isolados ou cabos unipolares em eletroduto de seção circular" Or tipocabo2 = "Cabos unipolares" Or tipocabo2 = "Cabos unipolares em trifólio") And (sistema = "FF") Then
        a = 2
    End If
    
    For q = 2 To LastRowG
        cabosporfase = wsLista.Cells(q, "E").Value
        compcir = wsLista.Cells(q, "G").Value
        wsLista.Cells(q, "H").Value = compcir * a * cabosporfase
    Next q
Next p
        
End Sub
Sub ComprimentoNeutroeTerraMT()

Dim wsLista As Worksheet
Dim wsCalculosMT As Worksheet
Dim UltimaLinhaCompNeutro As Long
Dim UltimaLinhaCompTerra As Long
Dim i As Long, j As Long
Dim CompNeutro As Variant, CompTerra As Variant
Dim cabosporfase As Variant, compcirc As Variant


Set wsLista = ThisWorkbook.Sheets("Lista de Cabos")
Set wsCalculosMT = ThisWorkbook.Sheets("Cálculos MT")


UltimaLinhaCompNeutro = wsLista.Cells(wsLista.Rows.Count, "I").End(xlUp).Row

For i = 2 To UltimaLinhaCompNeutro
    CompNeutro = wsLista.Cells(i, "I").Value
    cabosporfase = wsLista.Cells(i, "E").Value
    compcirc = wsLista.Cells(i, "G").Value
    If CompNeutro <> "-" Then
        wsLista.Cells(i, "J").Value = cabosporfase * compcirc
    Else
        wsLista.Cells(i, "J").Value = "-"
    End If
Next i

UltimaLinhaCompTerra = wsLista.Cells(wsLista.Rows.Count, "K").End(xlUp).Row

For j = 2 To UltimaLinhaCompTerra
    cabosporfase = wsLista.Cells(j, "E").Value
    compcirc = wsLista.Cells(j, "G").Value
    SecTerra = wsLista.Cells(j, "K").Value
    If SecTerra <> "-" Then
        wsLista.Cells(j, "L").Value = cabosporfase * compcirc
    Else
        wsLista.Cells(j, "L").Value = "-"
    End If
Next j

End Sub
Sub ComprimentoNeutroeTerraBT()
Dim wsCalculosBT As Worksheet
Dim wsLista As Worksheet
Dim i As Long
Set wsCalculosBT = ThisWorkbook.Sheets("Cálculos BT")
Set wsLista = ThisWorkbook.Sheets("Lista de Cabos")
UltimalinhaJ = wsLista.Cells(wsLista.Rows.Count, "J").End(xlUp).Row
UltimalinhaI = wsLista.Cells(wsLista.Rows.Count, "I").End(xlUp).Row
For i = UltimalinhaJ To UltimalinhaI
    cabosporfase = wsLista.Cells(i, "E").Value
    compcirc = wsLista.Cells(i, "G").Value
    SecNeutro = wsLista.Cells(i, "I").Value
    If SecNeutro <> "-" Then
        wsLista.Cells(i, "J") = cabosporfase * compcir
    Else
        wsLista.Cells(i, "J") = "-"
    End If
Next i
    
For j = UltimalinhaJ To UltimalinhaI
    cabosporfase = wsLista.Cells(j, "E").Value
    compcirc = wsLista.Cells(j, "G").Value
    SecTerra = wsLista.Cells(j, "K").Value
    If SecTerra <> "-" Then
        wsLista.Cells(j, "L") = cabosporfase * compcir
    Else
        wsLista.Cells(j, "L") = "-"
    End If
Next j
    
End Sub

Sub ConcatenarColunas()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Definir a planilha onde os dados estão localizados
    Set ws = ThisWorkbook.Sheets("Lista de Cabos") ' Altere "Nome_da_sua_planilha" para o nome correto da sua planilha
    
    ' Encontrar a última linha preenchida na coluna A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop através das linhas da segunda à última preenchida
    For i = 2 To lastRow
        ' Concatenar os valores das colunas A, B, C, D e F
        ws.Cells(i, "M").Value = ws.Cells(i, "A").Value & "-" & _
                                 ws.Cells(i, "B").Value & "-" & _
                                 ws.Cells(i, "C").Value & "-" & _
                                 ws.Cells(i, "D").Value & "-" & _
                                 ws.Cells(i, "F").Value
    Next i
End Sub
Sub ConcatenarERemoverDuplicatas()
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim lastRowOrigem As Long
    Dim i As Long
    Dim concatValue As String
    
    ' Definir as planilhas origem e destino
    Set wsOrigem = ThisWorkbook.Sheets("Lista de Cabos") ' Altere para o nome correto da sua planilha de origem
    Set wsDestino = ThisWorkbook.Sheets("PlanilhaAux") ' Altere para o nome correto da sua planilha de destino
    
    ' Encontrar a última linha preenchida na coluna A da planilha origem
    lastRowOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row
    
    ' Loop para concatenar valores e exibi-los na coluna A da planilha destino, separados por hífen
    For i = 2 To lastRowOrigem
        ' Concatenar os valores das colunas A, B, C, D e F da planilha origem, separados por hífen
        concatValue = wsOrigem.Cells(i, "A").Value & "-" & _
                      wsOrigem.Cells(i, "B").Value & "-" & _
                      wsOrigem.Cells(i, "C").Value & "-" & _
                      wsOrigem.Cells(i, "D").Value & "-" & _
                      wsOrigem.Cells(i, "F").Value
        
        ' Exibir o valor concatenado na coluna A da planilha destino
        wsDestino.Cells(i, "A").Value = concatValue
    Next i
    
    ' Remover as duplicatas na coluna A da planilha destino
    wsDestino.Columns("A:A").RemoveDuplicates Columns:=1, Header:=xlNo
End Sub
Sub TransportarValores()
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim lastRowOrigem As Long
    Dim i As Long
    
    ' Definir as planilhas origem e destino
    Set wsOrigem = ThisWorkbook.Sheets("PlanilhaAux") ' Altere para o nome correto da sua planilha de origem
    Set wsDestino = ThisWorkbook.Sheets("Lista de Materiais") ' Altere para o nome correto da sua planilha de destino
    
    ' Encontrar a última linha preenchida na coluna B da planilha origem
    lastRowOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "B").End(xlUp).Row
    
    ' Copiar os valores das colunas B, C, D, E, F, G e H da planilha origem para as colunas A, B, C, D, E, F e G da planilha destino
    For i = 2 To lastRowOrigem
        wsDestino.Cells(i, "A").Value = wsOrigem.Cells(i, "B").Value
        wsDestino.Cells(i, "B").Value = wsOrigem.Cells(i, "C").Value
        wsDestino.Cells(i, "C").Value = wsOrigem.Cells(i, "D").Value
        wsDestino.Cells(i, "D").Value = wsOrigem.Cells(i, "E").Value
        wsDestino.Cells(i, "E").Value = wsOrigem.Cells(i, "F").Value
        wsDestino.Cells(i, "F").Value = wsOrigem.Cells(i, "G").Value
        wsDestino.Cells(i, "G").Value = wsOrigem.Cells(i, "H").Value
    Next i
End Sub
Sub ListaCabosBT()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim i As Long
    Dim lastRow As Long
    
    Set ws1 = ThisWorkbook.Sheets("Cálculos BT")
    Set ws2 = ThisWorkbook.Sheets("Lista de Cabos")
    
    lastRow = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row + 1 ' Incrementa 1 para ir para a próxima linha vazia
    
    For i = 9 To ws1.Cells(Rows.Count, "C").End(xlUp).Row
        ws2.Cells(lastRow, "A").Value = ws1.Cells(i, "C").Value / 1000
        ws2.Cells(lastRow, "B").Value = "0,6/1"
        lastRow = lastRow + 1 ' Incrementa a variável lastRow para a próxima linha na planilha "Lista de Cabos"
    Next i
    
    ultimaLinhaC = ws2.Cells(ws2.Rows.Count, "C").End(xlUp).Row + 1

    For i = 9 To ws1.Cells(ws1.Rows.Count, "J").End(xlUp).Row
        ws2.Cells(ultimaLinhaC, "C").Value = ws1.Cells(i, "J").Value
        ws2.Cells(ultimaLinhaC, "D").Value = ws1.Cells(i, "M").Value
        ws2.Cells(ultimaLinhaC, "E").Value = ws1.Cells(i, "AI").Value
        ws2.Cells(ultimaLinhaC, "F").Value = ws1.Cells(i, "AJ").Value
        ws2.Cells(ultimaLinhaC, "G").Value = ws1.Cells(i, "I").Value
        ws2.Cells(ultimaLinhaC, "I").Value = ws1.Cells(i, "AG").Value
        ws2.Cells(ultimaLinhaC, "K").Value = ws1.Cells(i, "AF").Value
        ultimaLinhaC = ultimaLinhaC + 1
    Next i
    
    
End Sub
Sub mainlista()
    ListaCabosMT
    ListaCabosBT
    ComprimentoTotal
    ComprimentoNeutroeTerraMT
    'ComprimentoNeutroeTerraBT
    ConcatenarColunas
    ConcatenarERemoverDuplicatas
End Sub
Sub ApagarValoresColunas()
    Dim planilha As Worksheet
    Dim coluna As Range
    
    ' Altere "Planilha1" para o nome da sua planilha
    Set planilha = ThisWorkbook.Sheets("Lista de Cabos")
    
    ' Loop para percorrer as colunas de A até M a partir da linha 2
    For Each coluna In planilha.Range("A2:M" & planilha.Cells(planilha.Rows.Count, "A").End(xlUp).Row).Columns
        ' Apaga os valores na coluna atual
        coluna.ClearContents
    Next coluna
End Sub


