Attribute VB_Name = "LimpaCampos_BT"
Sub LimparCampos()

    Dim wsCalculosBT As Worksheet
    
    Set wsCalculosBT = ThisWorkbook.Sheets("Cálculos BT")

    'LastRowG = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "G").End(xlUp).Row
    'LastRowO = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "O").End(xlUp).Row
    'LastRowQ = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "Q").End(xlUp).Row
    'LastRowS = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "S").End(xlUp).Row
    'LastRowW = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "W").End(xlUp).Row
    'LastRowX = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "X").End(xlUp).Row
    LastRowZ = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "Z").End(xlUp).Row
    LastRowAA = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "AA").End(xlUp).Row
    'LastRowAB = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "AB").End(xlUp).Row
    LastRowAC = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "AC").End(xlUp).Row
    LastRowAD = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "AD").End(xlUp).Row
    'LastRowAE = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "AE").End(xlUp).Row
    LastRowAF = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "AF").End(xlUp).Row
    lastRowAG = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "AG").End(xlUp).Row
    'lastRowAH = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "AH").End(xlUp).Row
    LastRowAI = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "AI").End(xlUp).Row
    LastRowAJ = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "AJ").End(xlUp).Row
    LastRowAL = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "AL").End(xlUp).Row
    LastRowAM = wsCalculosBT.Cells(wsCalculosBT.Rows.Count, "AM").End(xlUp).Row
    
    'wsCalculosBT.Range("G9:G" & LastRowG).ClearContents
    'wsCalculosBT.Range("O9:O" & LastRowO).ClearContents
    'wsCalculosBT.Range("Q9:Q" & LastRowQ).ClearContents
    'wsCalculosBT.Range("S9:S" & LastRowS).ClearContents
    'wsCalculosBT.Range("W9:W" & LastRowW).ClearContents
    'wsCalculosBT.Range("X9:X" & LastRowX).ClearContents
    wsCalculosBT.Range("Z9:Z" & LastRowZ).ClearContents
    wsCalculosBT.Range("AA9:AA" & LastRowAA).ClearContents
    'wsCalculosBT.Range("AB9:AB" & LastRowAB).ClearContents
    wsCalculosBT.Range("AC9:AC" & LastRowAC).ClearContents
    wsCalculosBT.Range("AD9:AD" & LastRowAD).ClearContents
    'wsCalculosBT.Range("AE9:AE" & LastRowAE).ClearContents
    wsCalculosBT.Range("AF9:AF" & LastRowAF).ClearContents
    wsCalculosBT.Range("AG9:AG" & lastRowAG).ClearContents
    'wsCalculosBT.Range("AH9:AH" & lastRowAH).ClearContents
    wsCalculosBT.Range("AI9:AI" & LastRowAI).ClearContents
    wsCalculosBT.Range("AJ9:AJ" & LastRowAJ).ClearContents
    wsCalculosBT.Range("AL9:AL" & LastRowAJ).ClearContents
    wsCalculosBT.Range("AM9:AM" & LastRowAJ).ClearContents
End Sub

