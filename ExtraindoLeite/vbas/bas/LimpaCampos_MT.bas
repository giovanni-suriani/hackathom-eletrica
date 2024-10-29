Attribute VB_Name = "LimpaCampos_MT"
Sub LimparCamposMT()
Dim wsCalculosMT As Worksheet
Set wsCalculosMT = ThisWorkbook.Sheets("Cálculos MT")

LastRowA = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "A").End(xlUp).Row
LastRowB = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "B").End(xlUp).Row
LastRowC = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "C").End(xlUp).Row
LastRowD = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "D").End(xlUp).Row
LastRowE = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "E").End(xlUp).Row
LastRowF = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "F").End(xlUp).Row
LastRowH = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "H").End(xlUp).Row
LastRowI = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "I").End(xlUp).Row
LastRowJ = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "J").End(xlUp).Row
LastRowK = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "K").End(xlUp).Row
LastRowL = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "L").End(xlUp).Row
lastRowM = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "M").End(xlUp).Row
LastRowN = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "N").End(xlUp).Row
LastRowP = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "P").End(xlUp).Row
LastRowR = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "R").End(xlUp).Row
LastRowU = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "U").End(xlUp).Row
LastRowV = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "V").End(xlUp).Row
LastRowX = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "X").End(xlUp).Row
LastRowY = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "Y").End(xlUp).Row
LastRowZ = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "Z").End(xlUp).Row
LastRowAA = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AA").End(xlUp).Row
LastRowAB = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AB").End(xlUp).Row
LastRowAC = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AC").End(xlUp).Row
LastRowAD = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AD").End(xlUp).Row
LastRowAE = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AE").End(xlUp).Row
LastRowAF = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AF").End(xlUp).Row
lastRowAG = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AG").End(xlUp).Row
lastRowAH = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AH").End(xlUp).Row

wsCalculosMT.Range("A10:A" & LastRowA).ClearContents
wsCalculosMT.Range("B10:B" & LastRowB).ClearContents
wsCalculosMT.Range("C10:C" & LastRowC).ClearContents
wsCalculosMT.Range("D10:D" & LastRowD).ClearContents
wsCalculosMT.Range("E10:E" & LastRowE).ClearContents
wsCalculosMT.Range("F10:F" & LastRowF).ClearContents
wsCalculosMT.Range("H10:H" & LastRowH).ClearContents
wsCalculosMT.Range("I10:I" & LastRowI).ClearContents
wsCalculosMT.Range("J10:J" & LastRowJ).ClearContents
wsCalculosMT.Range("K10:K" & LastRowK).ClearContents
wsCalculosMT.Range("L10:L" & LastRowL).ClearContents
wsCalculosMT.Range("M10:M" & lastRowM).ClearContents
wsCalculosMT.Range("N10:N" & LastRowN).ClearContents
wsCalculosMT.Range("P10:P" & LastRowP).ClearContents
wsCalculosMT.Range("R10:R" & LastRowR).ClearContents
wsCalculosMT.Range("U10:U" & LastRowU).ClearContents
wsCalculosMT.Range("V10:V" & LastRowV).ClearContents
wsCalculosMT.Range("X10:X" & LastRowX).ClearContents
wsCalculosMT.Range("Y10:Y" & LastRowY).ClearContents
wsCalculosMT.Range("Z10:Z" & LastRowZ).ClearContents
wsCalculosMT.Range("AA10:AA" & LastRowAA).ClearContents
wsCalculosMT.Range("AB10:AB" & LastRowAB).ClearContents
wsCalculosMT.Range("AC10:AC" & LastRowAC).ClearContents
wsCalculosMT.Range("AD10:AD" & LastRowAD).ClearContents
wsCalculosMT.Range("AE10:AE" & LastRowAE).ClearContents
wsCalculosMT.Range("AF10:AF" & LastRowAF).ClearContents
wsCalculosMT.Range("AG10:AG" & lastRowAG).ClearContents
wsCalculosMT.Range("AH10:AH" & lastRowAH).ClearContents


End Sub

Sub LimparMetodos()

Dim wsCalculosMT As Worksheet
Set wsCalculosMT = ThisWorkbook.Sheets("Cálculos MT")

LastRowU = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "U").End(xlUp).Row
LastRowV = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "V").End(xlUp).Row
LastRowX = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "X").End(xlUp).Row
LastRowY = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "Y").End(xlUp).Row
LastRowZ = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "Z").End(xlUp).Row
LastRowAA = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AA").End(xlUp).Row
LastRowAB = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AB").End(xlUp).Row
LastRowAC = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AC").End(xlUp).Row
LastRowAD = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AD").End(xlUp).Row
LastRowAE = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AE").End(xlUp).Row
LastRowAF = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AF").End(xlUp).Row
lastRowAG = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AG").End(xlUp).Row
lastRowAH = wsCalculosMT.Cells(wsCalculosMT.Rows.Count, "AH").End(xlUp).Row


wsCalculosMT.Range("U10:U" & LastRowU).ClearContents
wsCalculosMT.Range("V10:V" & LastRowV).ClearContents
wsCalculosMT.Range("X10:X" & LastRowX).ClearContents
wsCalculosMT.Range("Y10:Y" & LastRowY).ClearContents
wsCalculosMT.Range("Z10:Z" & LastRowZ).ClearContents
wsCalculosMT.Range("AA10:AA" & LastRowAA).ClearContents
wsCalculosMT.Range("AB10:AB" & LastRowAB).ClearContents
wsCalculosMT.Range("AC10:AC" & LastRowAC).ClearContents
wsCalculosMT.Range("AD10:AD" & LastRowAD).ClearContents
wsCalculosMT.Range("AE10:AE" & LastRowAE).ClearContents
wsCalculosMT.Range("AF10:AF" & LastRowAF).ClearContents
wsCalculosMT.Range("AG10:AG" & lastRowAG).ClearContents
wsCalculosMT.Range("AH10:AH" & lastRowAH).ClearContents

End Sub

