Attribute VB_Name = "Module2"
Option Explicit
Dim num As Integer


Sub テスト()
  
  Dim s1 As String
  Dim s2 As String
  Dim s3 As String
  Dim s4 As String
  
  s1 = "こんにちは 。"
  s2 = " お元気ですか ？ "
  
  s3 = s1 & s2
  s4 = s1 + s2
  
  Debug.Print s3
  Debug.Print s4
  
End Sub

