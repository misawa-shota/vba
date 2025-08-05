Attribute VB_Name = "Module2"
Option Explicit
Dim num As Integer


Sub �e�X�g()

  ' �ǉ����������ڂ���͂���{�b�N�X�̕\��
  Dim newTitle As String
  newTitle = InputBox(" �ǉ����������ږ�����͂��ĉ����� ")
  
  ' �{�b�N�X�ɍ��ږ������͂��ꂽ���݈̂ȉ��̏��������s
  If Not IsEmpty(newTitle) Then
    
    ' ���i�̃��[�N�V�[�g�ɂ݈̂ȉ��̏������J��Ԃ����s
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
      ws.Select
      Select Case ws.Name
        Case "�ʐ^"
           ' �����ΏۊO�Ȃ̂ŁA�����Ȃ�
           
        Case "�O���t"
        ' �O���t�V�[�g�ɐV�������ڂ�ǉ����鏈��
        Dim graphSheet As Worksheet
        Set graphSheet = Worksheets("�O���t")

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
          ' ���i�̃V�[�g�ɂ݈̂ȉ��̏��������s
          Dim i As Long
          For i = 1 To 100
          
            ' ���ڂ�ǉ����鏈��
            If IsEmpty(Cells(6, i).Value) Then
            
              '  �V�K���ڂ̗��ǉ����鏈��
              Cells(6, i).Value = newTitle
              Range(Cells(7, i - 1), Cells(35, i - 1)).AutoFill Destination:=Range(Cells(7, i - 1), Cells(35, i)), Type:=xlFillDefault
              
              '  �e�[�u���̘g���̎w��
              Range(Cells(7, i), Cells(35, i)).Borders(xlEdgeLeft).Weight = xlThin
              Cells(6, i).Borders(xlEdgeTop).Weight = xlMedium
              Cells(6, i).Borders(xlEdgeRight).Weight = xlMedium
              Cells(6, i).Borders(xlEdgeLeft).Weight = xlThin
              
              '  �V�K�ǂ��������ڂ̃Z�����̔z�u�̎w��
              Cells(6, i).HorizontalAlignment = xlCenter
              Cells(6, i).VerticalAlignment = xlCenter
              
              ' �V�K�ǉ�������̕��𒲐�����
              Dim length As Integer
              length = Len(newTitle)
      
              If length > 7 Then
                Columns(i).AutoFit
              End If
              
               ' �V�K�ǉ�������̃f�[�^���͔͈͓��̃f�[�^����ɂ��鏈���i�I�[�g�t�B���ŗׂ̃f�[�^���R�s�[���邽�߁j
              Range(Cells(8, i), Cells(35, i)).Value = ""
              
              Exit For
            End If
            
          Next i
      End Select
    Next ws
    
    
    
    
    
  End If
End Sub

  Sub �O���t�쐬()
  
     Dim chartSheet As Worksheet
     Set chartSheet = Worksheets("�O���t")
     
     Dim chart As chart
     Set chart = ActiveChart
     
     With chartSheet.Shapes.AddChart2.chart
      .HasTitle = True
      .ChartTitle.Text = "�ٕ��̎�ނƔ�������"
      .ChartType = xlColumnClustered
      .SetSourceData Range(Cells(6, "B"), Cells(7, "K"))
      
      With .Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "��������"
            .AxisTitle.Orientation = xlVertical
      End With
    End With
    
  End Sub

