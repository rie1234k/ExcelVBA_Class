VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_ProgressBar 
   Caption         =   "������"
   ClientHeight    =   1485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4320
   OleObjectBlob   =   "UF_ProgressBar.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UF_ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    
'------- �v���O���X�o�[�ݒ� -------

    With Label1
        
        .Width = 180
        .Height = 20
        .Top = 20
        .Left = 20
    
    End With
    
    With L_Gauge
        
        .Width = 0
        .Height = 20
        .Top = 20
        .Left = 20
        .BackColor = rgbDodgerBlue  '�Q�[�W�̐F
        
    End With
    
    
'------- ��ʕ\���ʒu�ݒ� -------

'�G�N�Z�����\������Ă����ʂ̐^�񒆂ɕ\������i�}���`�f�B�X�v���C�Ή��j

Dim WinData(3) As Long

    With ActiveWindow

        WinData(0) = .Top
        WinData(1) = .Left
        WinData(2) = .Width
        WinData(3) = .Height

    End With

Dim FormData(3) As Long

    With Me

        FormData(2) = .Width
        FormData(3) = .Height


    End With

    FormData(0) = WinData(0) + ((WinData(3) - FormData(3)) / 2)

    FormData(1) = WinData(1) + ((WinData(2) - FormData(2)) / 2)

    With Me

        .StartUpPosition = 0
        .Top = FormData(0)
        .Left = FormData(1)

    End With

End Sub
