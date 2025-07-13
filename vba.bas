Attribute VB_Name = "Module2"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveCell.FormulaR1C1 = "=RC[-2]*RC[-1]"
End Sub

Sub 乱数設置()
    Dim rStart
    Dim cStart
    
    rStart = ActiveCell.Row
    cStart = ActiveCell.Column
    
    Randomize
    
    For rCount = 0 To 4
        For cCount = 0 To 4
            Cells(rStart + rCount, cStart + cCount).Value = Int(25 * Rnd + 1)
        Next cCount
    Next rCount
End Sub
