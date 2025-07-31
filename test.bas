Attribute VB_Name = "Module2"
Option Explicit
Dim num As Integer


Sub テスト()

  ' 追加したい項目を入力するボックスの表示
  Dim newTitle As String
  newTitle = InputBox(" 追加したい項目名を入力して下さい ")
  
  ' ボックスに項目名が入力された時のみ以下の処理を実行
  If Not IsEmpty(newTitle) Then
    
    ' 商品のワークシートにのみ以下の処理を繰り返し実行
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
      ws.Select
      Select Case ws.Name
        Case "グラフ", "写真", "計算用シート（始まり）", "計算用シート（終わり）"
          ' 処理対象外なので、処理なし
          
        Case Else
          ' 商品のシートにのみ以下の処理を実行
          Dim i As Long
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
      End Select
    Next ws
  End If
  
End Sub

  Sub ワークシート取得()
  
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
      Select Case ws.Name
        Case "グラフ", "写真", "計算用シート（始まり）", "計算用シート（終わり）"
        
        Case Else
          Dim i As Long
          For i = 1 To 6
            If IsEmpty(Cells(3, i)) Then
              ws.Cells(3, i) = " テスト "
            End If
          Next i
        End Select
    Next ws
    
  End Sub

  


