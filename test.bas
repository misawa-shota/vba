Attribute VB_Name = "Module2"
Option Explicit
Dim num As Integer


Sub テスト()

  Dim newTitle As String
  newTitle = InputBox(" 追加したい項目名を入力して下さい ")
  
  If Not IsEmpty(newTitle) Then
    Dim i As Integer
    
    For i = 1 To 100
    
      ' 項目を追加する処理
      If IsEmpty(Cells(6, i).Value) Then
      
        '  新規項目の列を追加する処理
        Cells(6, i).Value = newTitle
        Range(Cells(7, i - 1), Cells(35, i - 1)).AutoFill Destination:=Range(Cells(7, i - 1), Cells(35, i)), Type:=xlFillDefault
        
        '  テーブルの枠線の指定
        Range(Cells(7, i), Cells(35, i)).Borders(xlEdgeLeft).Weight = xlThin
        Cells(6, i).Borders(xlEdgeTop).Weight = xlMedium
        Cells(6, i).Borders(xlEdgeRight).Weight = xlMedium
        Cells(6, i).Borders(xlEdgeLeft).Weight = xlThin
        
        '  新規追した加項目のセル内の配置の指定
        Cells(6, i).HorizontalAlignment = xlCenter
        Cells(6, i).VerticalAlignment = xlCenter
        
        ' 新規追加した列の幅を調整する
        Dim length As Integer
        length = Len(newTitle)

        If length > 7 Then
          Columns(i).AutoFit
        End If
        
        ' 新規追加した列のデータ入力範囲内のデータを空にする処理（オートフィルで隣のデータをコピーするため）
        Range(Cells(8, i), Cells(35, i)).Value = ""
        
        Exit For
      End If
      
    Next i
  End If
  
End Sub


