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

Private ClassControlCollection As Collection
Private Sub CB_Close_Click()
    
    Unload Me
    
End Sub



Private Sub UserForm_Initialize()

Dim myControl As msforms.Control
Dim ClassControl As ControlEvents


    Set ClassControlCollection = New Collection
    

    For Each myControl In Me.Controls
        
        '���i���Ƃɐݒ�
        Select Case True
        
            Case myControl.Name Like "TextBox*"
                
                
                '�C���X�^���X��
                Set ClassControl = New ControlEvents
                
                '�^�C�v�ɍ��킹�āA���i��ݒ�
                Set ClassControl.SetControl_TextBox = myControl
                
                '�R���N�V�����ɒǉ�
                ClassControlCollection.Add ClassControl
                
                myControl.IMEMode = fmIMEModeOff
                
            
            Case myControl.Name Like "CommandButton*"
            
                
                Set ClassControl = New ControlEvents
                
                Set ClassControl.SetControl_CommandButton = myControl
                 
                ClassControlCollection.Add ClassControl
                        
        End Select
             
   Next myControl
     

End Sub
