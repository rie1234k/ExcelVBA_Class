VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private myform_ As Object


Public Property Get TargetForm() As Object

    Set TargetForm = myform_

End Property

Public Property Set TargetForm(ByVal NewForm As Object)

    Set myform_ = NewForm

End Property

'プログレスバーを描画する
Public Sub ProgressBarPaint(StartParcent As Long, EndParcent As Long, Unit As Long, Messege As String)
Dim myTimer As Double
Dim myInterval As Double
Dim Max As Long
Dim GaugeCtr As Long
Dim WidthParcent As Single

    myform_.Label1.Caption = Messege
     
    Max = myform_.Label1.Width

    
    myInterval = 50 / 10000


    For GaugeCtr = Int(StartParcent / Unit) To Application.WorksheetFunction.RoundUp(EndParcent / Unit, 0)
 
        myTimer = Timer()
        
        Do While Timer() - myTimer < myInterval
        
        Loop
        
        If GaugeCtr * 0.01 * Unit <= EndParcent * 0.01 Then
        
            WidthParcent = GaugeCtr * 0.01 * Unit
        
        Else
        
            WidthParcent = EndParcent * 0.01
        
        End If
        
    
        myform_.L_Gauge.Width = Max * WidthParcent

        DoEvents
        
        
    Next GaugeCtr

    
    myform_.Repaint
    
End Sub

'ユーザーフォーム（プログレスバー）を表示
Public Sub ShowForm()

    myform_.Show vbModeless


End Sub

'ユーザーフォーム（プログレスバー）を閉じる
Public Sub UnloadForm()

    Unload myform_

End Sub


'ユーザーフォーム（プログレスバー）のメッセージを取得
Public Property Get Message() As String

   Message = myform_.Label1.Caption
   
End Property

'ユーザーフォーム（プログレスバー）のメッセージを設定
Public Property Let Message(ByVal NewMessage As String)

    myform_.Label1.Caption = NewMessage
    myform_.Repaint


End Property

'ユーザーフォーム（プログレスバー)の表示を待つ
Public Sub FinalWait()

    
     Application.Wait [Now()] + 500 / 86400000
    

End Sub
