VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalendarForm 
   Caption         =   "日付選択"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4200
   OleObjectBlob   =   "CalendarForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
'フォームが表示されている間は有効
Private DayLabelArray(1 To 37) As CalendarEvents





Private Sub CB_Close_Click()

    MsgBox "キャンセルされました。" & vbCrLf & "終了します。"
    Unload Me
    
    End
    
End Sub

Private Sub CB_Month_Change()

    Call DateSet

End Sub

Private Sub CB_OK_Click()

    If IsDate(TB_Date.Value) Then
    
        'iDateをグローバル変数として、標準モジュールにPublicで宣言しておく
        iDate = TB_Date.Value
        
        Unload Me
        
        Else
        
        MsgBox "日付を選択してください。"
    
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

'------- コンボボックス設定 -------
    For i = -1 To 1
    
        CB_Year.AddItem Year(Date) + i
        
    Next i
    
    For i = 1 To 12
    
        CB_Month.AddItem i
        
    Next i
    
    CB_Year.Text = Year(Date)
    CB_Month.Text = Month(Date)
    
    TB_Date.Text = Date
    
'------- 日付ラベル設定 -------

    For i = 1 To 37
    
        'イベント用ラベルクラスを生成
        Set DayLabelArray(i) = New CalendarEvents
        'ラベルをセット
        DayLabelArray(i).SetLabel Controls("Label" & i), i
    
    Next i
    
    

End Sub



Private Sub DateSet()
Dim i As Long
Dim j As Long



    If CB_Year = "" Or CB_Month = "" Then Exit Sub
    
    
'------- ラベル初期化 -------
    
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
    
    
'------- 日付ラベルセット -------
Dim iYear As Long
Dim iMonth As Long
Dim FirstWeekDay As Long
Dim EndDay As Long
  
    iYear = CB_Year
    iMonth = CB_Month
    
    'その月の1日の曜日番号を取得（月曜始まり）
    FirstWeekDay = Weekday(DateSerial(iYear, iMonth, 1), vbMonday)
    
    'その月の最終日を取得（翌月の1日前）
    EndDay = Day(DateAdd("d", -1, DateAdd("m", 1, DateSerial(iYear, iMonth, 1))))
    
    
    For i = 1 To EndDay
    
        'ラベル番号に始まりの曜日を足して1を引くと始まりのラベルになる
        Controls("Label" & i + FirstWeekDay - 1).Caption = i
        Controls("Label" & i + FirstWeekDay - 1).Visible = True
        
        '当日なら色を付ける
        If DateSerial(iYear, iMonth, i) = Date Then
        
            Controls("Label" & i + FirstWeekDay - 1).BackColor = RGB(200, 200, 200)
            
        End If
        
        
        '祝日に色を付ける
        j = 1
        
        Do
        
            If DateSerial(iYear, iMonth, i) = ThisWorkbook.Sheets("休日リスト").Cells(j, 1).Value Then
            
                Controls("Label" & i + FirstWeekDay - 1).ForeColor = DateColor.Sunday
            
            End If
            
            j = j + 1
        
        Loop While ThisWorkbook.Sheets("休日リスト").Cells(j, 1) <> ""
        
    Next i
        
    TB_Date.SetFocus
    

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
    
        MsgBox "キャンセルボタンで終了してください。"
        
        Cancel = True
    
    End If
    


End Sub
