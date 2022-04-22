Attribute VB_Name = "Module1"
Sub CleanUpData()

    Dim i As Integer
    i = 1
    
    Do While i <= Worksheets.Count
        
        Worksheets(i).Select
        
        AddHeaders
        Headers
        
        
        i = i + 1
    Loop
    

End Sub




Sub AddHeaders()
Attribute AddHeaders.VB_Description = "This macro places headers on the worksheets"
Attribute AddHeaders.VB_ProcData.VB_Invoke_Func = " \n14"
'
' AddHeaders Macro
' This macro places headers on the worksheets
'

'
    ActiveWindow.SmallScroll Down:=-9
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Region"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Category"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Jan"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Feb"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Mar"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Total"
    Range("A2").Select
End Sub





