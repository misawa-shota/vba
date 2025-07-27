Attribute VB_Name = "Module2"
Option Explicit
Dim num As Integer


Sub テスト()
  
  Dim ws As Worksheet
  
  Set ws = Worksheets("Sheet2")
  ws.Range("A2").Value = " テスト "
  
End Sub

