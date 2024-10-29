Attribute VB_Name = "LimpaCampos_BT_MT"
Sub LimparCamposMT2()

    Application.ScreenUpdating = False

    Dim wsCalculosMT As Worksheet
    Set wsCalculosMT = ThisWorkbook.Sheets("Cálculos MT")

    Range("A10:F10").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Range("H10:N10").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Range("P10").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Range("R10").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Range("U10:AH10").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Range("A10").Select
    
End Sub

Sub LimparCamposBT2()

    Application.ScreenUpdating = False

    Dim wsCalculosBT As Worksheet
    
    Set wsCalculosBT = ThisWorkbook.Sheets("Cálculos BT")

    Range("A9:V9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Range("Z9:AM9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    
    Range("A9").Select

End Sub
