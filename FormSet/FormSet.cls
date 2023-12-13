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

'ユーザーフォームの設定をするクラスモジュール


Private myform_ As Object

'ユーザーフォームを取得する
Public Property Get TargetForm() As Object

    Set TargetForm = myform_

End Property

'ユーザーフォームを設定する
Public Property Set TargetForm(ByVal NewForm As Object)

    Set myform_ = NewForm

End Property



'エクセルが表示されているウィンドウの真ん中に表示する（デュアルディスプレイ対応用）
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

'エクセルが表示されているウィンドウの左寄りに表示する（デュアルディスプレイ対応用）
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
'エクセルが表示されているウィンドウの右寄りに表示する（デュアルディスプレイ対応用）
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

'フォームのデザインを設定する

'薄緑　rgbHoneydew

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


