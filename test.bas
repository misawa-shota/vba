Attribute VB_Name = "Module2"
Option Explicit
Dim num As Integer


Sub �e�X�g()

  Dim newTitle As String
  newTitle = InputBox(" �ǉ����������ږ�����͂��ĉ����� ")
  
  If Not IsEmpty(newTitle) Then
    Dim i As Integer
    
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
  End If
  
End Sub


