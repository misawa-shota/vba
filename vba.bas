Attribute VB_Name = "Module2"
Sub テスト()
    Dim sum
    
    sum = Range("B2").Value + Range("C2").Value
    
    Debug.Print "合計値", sum
    
    Range("D2").Value = sum
End Sub
