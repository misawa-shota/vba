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
        Case "�O���t", "�ʐ^", "�v�Z�p�V�[�g�i�n�܂�j", "�v�Z�p�V�[�g�i�I���j"
          ' �����ΏۊO�Ȃ̂ŁA�����Ȃ�
          
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

  Sub ���[�N�V�[�g�擾()
  
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
      Select Case ws.Name
        Case "�O���t", "�ʐ^", "�v�Z�p�V�[�g�i�n�܂�j", "�v�Z�p�V�[�g�i�I���j"
        
        Case Else
          Dim i As Long
          For i = 1 To 6
            If IsEmpty(Cells(3, i)) Then
              ws.Cells(3, i) = " �e�X�g "
            End If
          Next i
        End Select
    Next ws
    
  End Sub

  


