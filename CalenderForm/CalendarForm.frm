VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalendarForm 
   Caption         =   "���t�I��"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4200
   OleObjectBlob   =   "CalendarForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "CalendarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Enum DateColor

    Saturday = vbBlue
    
    Sunday = vbRed
    
    Weekdays = vbBlack
    

End Enum
'�t�H�[�����\������Ă���Ԃ͗L��
Private DayLabelArray(1 To 37) As CalendarEvents





Private Sub CB_Close_Click()

    MsgBox "�L�����Z������܂����B" & vbCrLf & "�I�����܂��B"
    Unload Me
    
    End
    
End Sub

Private Sub CB_Month_Change()

    Call DateSet

End Sub

Private Sub CB_OK_Click()

    If IsDate(TB_Date.Value) Then
    
        'iDate���O���[�o���ϐ��Ƃ��āA�W�����W���[����Public�Ő錾���Ă���
        iDate = TB_Date.Value
        
        Unload Me
        
        Else
        
        MsgBox "���t��I�����Ă��������B"
    
    End If


End Sub

Private Sub CB_Today_Click()

    CB_Year.Text = Year(Date)
    CB_Month.Text = Month(Date)
    TB_Date.Text = Date
    

End Sub

Private Sub CB_Year_Change()
Dim i As Long
Dim iYear As Long

    Call DateSet
      
End Sub

Private Sub SpinButton1_SpinDown()
    
    Select Case CB_Month.Text
    
        Case 12
        
            CB_Month.Text = 1
            
            CB_Year.Text = CB_Year.Text + 1
            
        Case Else
               
            CB_Month.Text = CB_Month.Text + 1
            
    End Select
    
    
    
End Sub

Private Sub SpinButton1_SpinUp()

    Select Case CB_Month.Text
    
        Case 1
        
            CB_Month.Text = 12
            
            CB_Year.Text = CB_Year.Text - 1
            
        
        Case Else
               
            CB_Month.Text = CB_Month.Text - 1
        
    End Select

    
    
End Sub



Private Sub UserForm_Initialize()

Dim i As Long

'------- �R���{�{�b�N�X�ݒ� -------
    For i = -1 To 1
    
        CB_Year.AddItem Year(Date) + i
        
    Next i
    
    For i = 1 To 12
    
        CB_Month.AddItem i
        
    Next i
    
    CB_Year.Text = Year(Date)
    CB_Month.Text = Month(Date)
    
    TB_Date.Text = Date
    
'------- ���t���x���ݒ� -------

    For i = 1 To 37
    
        '�C�x���g�p���x���N���X�𐶐�
        Set DayLabelArray(i) = New CalendarEvents
        '���x�����Z�b�g
        DayLabelArray(i).SetLabel Controls("Label" & i), i
    
    Next i
    
    

End Sub



Private Sub DateSet()
Dim i As Long
Dim j As Long



    If CB_Year = "" Or CB_Month = "" Then Exit Sub
    
    
'------- ���x�������� -------
    
    For i = 1 To 37
    
        Controls("Label" & i).Caption = ""
        Controls("Label" & i).Visible = False
        Controls("Label" & i).BackColor = Me.BackColor
        
        
        Select Case i Mod 7
        
            Case 6
            
                Controls("Label" & i).ForeColor = DateColor.Saturday
            
            Case 0
            
                Controls("Label" & i).ForeColor = DateColor.Sunday
            
            Case Else
            
                Controls("Label" & i).ForeColor = DateColor.Weekdays
        
        End Select
        
    
    Next i
    
    
'------- ���t���x���Z�b�g -------
Dim iYear As Long
Dim iMonth As Long
Dim FirstWeekDay As Long
Dim EndDay As Long
  
    iYear = CB_Year
    iMonth = CB_Month
    
    '���̌���1���̗j���ԍ����擾�i���j�n�܂�j
    FirstWeekDay = Weekday(DateSerial(iYear, iMonth, 1), vbMonday)
    
    '���̌��̍ŏI�����擾�i������1���O�j
    EndDay = Day(DateAdd("d", -1, DateAdd("m", 1, DateSerial(iYear, iMonth, 1))))
    
    
    For i = 1 To EndDay
    
        '���x���ԍ��Ɏn�܂�̗j���𑫂���1�������Ǝn�܂�̃��x���ɂȂ�
        Controls("Label" & i + FirstWeekDay - 1).Caption = i
        Controls("Label" & i + FirstWeekDay - 1).Visible = True
        
        '�����Ȃ�F��t����
        If DateSerial(iYear, iMonth, i) = Date Then
        
            Controls("Label" & i + FirstWeekDay - 1).BackColor = RGB(200, 200, 200)
            
        End If
        
        
        '�j���ɐF��t����
        j = 1
        
        Do
        
            If DateSerial(iYear, iMonth, i) = ThisWorkbook.Sheets("�x�����X�g").Cells(j, 1).Value Then
            
                Controls("Label" & i + FirstWeekDay - 1).ForeColor = DateColor.Sunday
            
            End If
            
            j = j + 1
        
        Loop While ThisWorkbook.Sheets("�x�����X�g").Cells(j, 1) <> ""
        
    Next i
        
    TB_Date.SetFocus
    

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
    
        MsgBox "�L�����Z���{�^���ŏI�����Ă��������B"
        
        Cancel = True
    
    End If
    


End Sub
