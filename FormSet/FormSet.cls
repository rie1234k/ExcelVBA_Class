VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'���[�U�[�t�H�[���̐ݒ������N���X���W���[��


Private myform_ As Object

'���[�U�[�t�H�[�����擾����
Public Property Get TargetForm() As Object

    Set TargetForm = myform_

End Property

'���[�U�[�t�H�[����ݒ肷��
Public Property Set TargetForm(ByVal NewForm As Object)

    Set myform_ = NewForm

End Property



'�G�N�Z�����\������Ă���E�B���h�E�̐^�񒆂ɕ\������i�f���A���f�B�X�v���C�Ή��p�j
Public Sub PositionCenterSet()

Dim WinData(3) As Long

    With ActiveWindow
        
        WinData(0) = .Top
        WinData(1) = .Left
        WinData(2) = .Width
        WinData(3) = .Height
    
    
    End With
    
    Dim FormData(3) As Long
    
    With myform_
    
        FormData(2) = .Width
        FormData(3) = .Height
    
    
    End With
    
    FormData(0) = WinData(0) + ((WinData(3) - FormData(3)) / 2)
    
    FormData(1) = WinData(1) + ((WinData(2) - FormData(2)) / 2)
    
    With myform_
        
        
        .StartUpPosition = 0
        .Top = FormData(0)
        .Left = FormData(1)
         
    
    End With
    

End Sub

'�G�N�Z�����\������Ă���E�B���h�E�̍����ɕ\������i�f���A���f�B�X�v���C�Ή��p�j
Public Sub PositionLeftSet()

Dim WinData(3) As Long

    With ActiveWindow
        
        WinData(0) = .Top
        WinData(1) = .Left
        WinData(2) = .Width
        WinData(3) = .Height
    
    
    End With
    
    Dim FormData(3) As Long
    
    With myform_
    
        FormData(2) = .Width
        FormData(3) = .Height
    
    
    End With
    
    FormData(0) = WinData(0) + ((WinData(3) - FormData(3)) / 2)
    
    FormData(1) = WinData(1) + ((WinData(2) - FormData(2)) / 4)
    
    With myform_
        
        
        .StartUpPosition = 0
        .Top = FormData(0)
        .Left = FormData(1)
         
    
    End With
    

End Sub
'�G�N�Z�����\������Ă���E�B���h�E�̉E���ɕ\������i�f���A���f�B�X�v���C�Ή��p�j
Public Sub PositionRightSet()

Dim WinData(3) As Long

    With ActiveWindow
        
        WinData(0) = .Top
        WinData(1) = .Left
        WinData(2) = .Width
        WinData(3) = .Height
    
    
    End With
    
    Dim FormData(3) As Long
    
    With myform_
    
        FormData(2) = .Width
        FormData(3) = .Height
    
    
    End With
    
    FormData(0) = WinData(0) + ((WinData(3) - FormData(3)) / 2)
    
    FormData(1) = WinData(1) + ((WinData(2) - FormData(2)) / 2) + ((WinData(2) - FormData(2)) / 8)
    
    With myform_
        
        
        .StartUpPosition = 0
        .Top = FormData(0)
        .Left = FormData(1)
         
    
    End With
    

End Sub

'�t�H�[���̃f�U�C����ݒ肷��

'���΁@rgbHoneydew

Public Sub FormatSet()
Dim obj As Object

    On Error Resume Next
    
    myform_.Font.Name = "Meiryo UI"
    myform_.BackColor = rgbGhostWhite
    
    For Each obj In myform_.Controls
    
        obj.Font.Name = "Meiryo UI"
        obj.Font.Size = 11
        
        If obj.Name = "CB_Cancel" Then
        
            obj.BackColor = rgbMistyRose
        
        End If
        
        If obj.Name = "CB_OK" Then
            
            obj.BackColor = rgbLightBlue
        
        End If
        
        
        If TypeName(obj) = "Label" Or TypeName(obj) = "Frame" Or TypeName(obj) = "CheckBox" Or TypeName(obj) = "OptionButton" Then
        
            obj.BackColor = rgbGhostWhite
        
        End If
        
        
    
    Next obj
    
    

End Sub

