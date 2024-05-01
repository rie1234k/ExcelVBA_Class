VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ControlEvents_SampleForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3480
   OleObjectBlob   =   "ControlEvents_SampleForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "ControlEvents_SampleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ControlArray() As ControlEvents  'Form�\�����̊ԁA�L���ɂ���

Private Sub CB_Close_Click()
    
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()

Dim myControl As Object

    For Each myControl In Me.Controls
        
        '�R���g���[�����Ƃɐݒ�
        Select Case True
        
            Case myControl.Name Like "TextBox*"
            
                
                '�z��̐ݒ�
                If isEmptyArray(ControlArray) Then
                    
                    ReDim ControlArray(0)
                
                Else
                
                    ReDim Preserve ControlArray(UBound(ControlArray) + 1)
                
                End If
                
                
                '�R���g���[�����ƂɁA�C���X�^���X��
                Set ControlArray(UBound(ControlArray)) = New ControlEvents
                
                '�R���g���[���^�C�v�ɍ��킹�āA�R���g���[����ݒ�
                ControlArray(UBound(ControlArray)).SetControl_TextBox myControl, myControl.Name, Me
                
                myControl.IMEMode = fmIMEModeOff
                
            
            Case myControl.Name Like "CommandButton*"
            
                '�z��̐ݒ�
                If isEmptyArray(ControlArray) Then
                
                    ReDim ControlArray(0)
                
                Else
                
                    ReDim Preserve ControlArray(UBound(ControlArray) + 1)
                
                End If
                
                
                '�R���g���[�����ƂɁA�C���X�^���X��
                Set ControlArray(UBound(ControlArray)) = New ControlEvents
                
                '�R���g���[���^�C�v�ɍ��킹�āA�R���g���[����ݒ�
                ControlArray(UBound(ControlArray)).SetControl_CommandButton myControl, myControl.Name, Me

      
                        
        End Select
        
        
   Next myControl
   
    

End Sub



Public Function isEmptyArray(v() As ControlEvents) As Boolean

Dim tmp As Long


On Error GoTo ErrUbound:

    tmp = UBound(v)
    isEmptyArray = False
    
    Exit Function


ErrUbound:

    If Err.Number <> 9 Then
        
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
    End If
    
    isEmptyArray = True


End Function
