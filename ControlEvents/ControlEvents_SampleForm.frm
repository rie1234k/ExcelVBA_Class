VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ControlEvents_SampleForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3480
   OleObjectBlob   =   "ControlEvents_SampleForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ControlEvents_SampleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ControlArray() As ControlSet  'Form表示中の間、有効にする

Private Sub CB_Close_Click()
    
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()

Dim myControl As Object
    

   For Each myControl In Me.Controls
        
        'コントロールごとに設定
        Select Case True
        
            Case myControl.Name Like "TextBox*"
            
                
                '配列の設定
                If (Not ControlArray) = -1 Then
                        
                    ReDim ControlArray(0)
                    
                Else
                
                    ReDim Preserve ControlArray(UBound(ControlArray) + 1)
                
                End If
                
                'コントロールごとに、インスタンス化
                Set ControlArray(UBound(ControlArray)) = New ControlSet
                
                'コントロールタイプに合わせて、コントロールを設定
                ControlArray(UBound(ControlArray)).SetControl_TextBox myControl, myControl.Name, Me
                
                myControl.IMEMode = fmIMEModeOff
                
            
            Case myControl.Name Like "CommandButton*"
            
                '配列の設定
                If (Not ControlArray) = -1 Then
                        
                    ReDim ControlArray(0)
                    
                Else
                
                    ReDim Preserve ControlArray(UBound(ControlArray) + 1)
                
                End If
                
                'コントロールごとに、インスタンス化
                Set ControlArray(UBound(ControlArray)) = New ControlSet
                
                'コントロールタイプに合わせて、コントロールを設定
                ControlArray(UBound(ControlArray)).SetControl_CommandButton myControl, myControl.Name, Me

      
                        
        End Select
        
        
   Next myControl
   
    

End Sub
