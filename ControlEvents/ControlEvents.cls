VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents ControlSet_TextBox As MSForms.TextBox 'ここをコントロール部品に合わせてラベルやテキストボックスに変更
Attribute ControlSet_TextBox.VB_VarHelpID = -1
Private WithEvents ControlSet_CommandButton As MSForms.CommandButton 'ここをコントロール部品に合わせてラベルやテキストボックスに変更
Attribute ControlSet_CommandButton.VB_VarHelpID = -1
Private ControlName As String  'コントロール名称
Private MyForm As Object


'各部品ごとに設定する
Public Sub SetControl_TextBox(NewControl As MSForms.TextBox, NewName As String, NewForm As Object)

    Set ControlSet_TextBox = NewControl
    
    ControlName = NewName
    
    Set MyForm = NewForm
    

End Sub
Public Sub SetControl_CommandButton(NewControl As MSForms.CommandButton, NewName As String, NewForm As Object)

    Set ControlSet_CommandButton = NewControl
    
    ControlName = NewName
    
    Set MyForm = NewForm


End Sub


Private Sub ControlSet_CommandButton_Click()
    
    
    'テキストボックスとコマンドボタンのインデックスを一致させておく
    MyForm.Controls(Replace(ControlName, "CommandButton", "TextBox")).Value = ""
    

End Sub

Private Sub ControlSet_TextBox_Change()

    If Len(MyForm.Controls(ControlName).Value) > 4 Then
    
        MsgBox "桁数の制限を超えています。"
        
        MyForm.Controls(ControlName).Value = Left(MyForm.Controls(ControlName).Value, 4)
        
    End If
    
End Sub
