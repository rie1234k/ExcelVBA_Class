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

Private WithEvents ControlSet_TextBox As MSForms.TextBox 'ここを部品に合わせて変更する
Attribute ControlSet_TextBox.VB_VarHelpID = -1
Private WithEvents ControlSet_CommandButton As MSForms.CommandButton 'ここをコントロール部品に合わせてラベルやテキストボックスに変更
Attribute ControlSet_CommandButton.VB_VarHelpID = -1


'各部品ごとに設定する
Public Property Set SetControl_TextBox(ByVal NewControl As MSForms.TextBox)

     Set ControlSet_TextBox = NewControl

End Property

Public Property Set SetControl_CommandButton(ByVal NewControl As MSForms.CommandButton)

    Set ControlSet_CommandButton = NewControl

End Property



Private Sub ControlSet_CommandButton_Click()
Dim ControlName As String
Dim MyForm As MSForms.UserForm

    ControlName = ControlSet_CommandButton.Name
    Set MyForm = ControlSet_CommandButton.Parent
    
    'テキストボックスとコマンドボタンのインデックスを一致させておく
    MyForm.Controls(Replace(ControlName, "CommandButton", "TextBox")).Value = ""
    

End Sub

Private Sub ControlSet_TextBox_Change()

    If Len(ControlSet_TextBox.Value) > 4 Then
    
        MsgBox "桁数の制限を超えています。"
        ControlSet_TextBox.Value = Left(ControlSet_TextBox.Value, 4)
        
    End If
       
        
    If Not IsNumeric(ControlSet_TextBox.Value) And ControlSet_TextBox.Value <> "" Then
            
        MsgBox "入力された値が数値ではありません。"
        ControlSet_TextBox.Value = ""
        
    End If
    
End Sub

