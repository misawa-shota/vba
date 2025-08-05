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
        Case "写真"
           ' 処理対象外なので、処理なし
           
        Case "グラフ"
        ' グラフシートに新しい項目を追加する処理
        Dim graphSheet As Worksheet
        Set graphSheet = Worksheets("グラフ")

        Dim n As Long
        For n = 1 To 100

          If IsEmpty(graphSheet.Cells(6, n).Value) Then
            graphSheet.Cells(6, n).Value = newTitle
            graphSheet.Range(Cells(7, n - 1), Cells(7, n - 1)).AutoFill Destination:=graphSheet.Range(Cells(7, n - 1), Cells(7, n)), Type:=xlFillDefault

            graphSheet.Cells(6, n).Borders(xlEdgeTop).Weight = xlThin
            graphSheet.Cells(6, n).Borders(xlEdgeRight).Weight = xlThin
            graphSheet.Cells(6, n).Borders(xlEdgeBottom).LineStyle = xlDouble

            graphSheet.Cells(6, n).HorizontalAlignment = xlCenter
            graphSheet.Cells(6, n).VerticalAlignment = xlCenter

            Dim stringLength As Integer
            stringLength = Len(newTitle)
            If stringLength > 7 Then
              graphSheet.Columns(n).AutoFit
            End If
            Exit For
          End If
        Next n
             
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

  Sub グラフ作成()
  
     Dim chartSheet As Worksheet
     Set chartSheet = Worksheets("グラフ")
     
     Dim chart As chart
     Set chart = ActiveChart
     
     With chartSheet.Shapes.AddChart2.chart
      .HasTitle = True
      .ChartTitle.Text = "異物の種類と発生件数"
      .ChartType = xlColumnClustered
      .SetSourceData Range(Cells(6, "B"), Cells(7, "K"))
      
      With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "発生件数"
            .AxisTitle.Orientation = xlVertical
      End With
    End With
    
  End Sub

