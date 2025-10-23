Attribute VB_Name = "サンプル2"
Option Explicit

Public Sub サンプル2a()

Dim iBar As ProgressBar
Dim i As Long

Dim iCount As Long

    
'------- 初期設定 -------

    Set iBar = New ProgressBar
    
    Set iBar.TargetForm = UF_ProgressBar
    
    iBar.ShowForm

    
'------- 処理 -------

    iCount = 100
     
    For i = 1 To iCount

        '処理
        Application.Wait [Now()] + 10 / 86400000
        
        
        iBar.ProgressBarPaint (i - 1) / (iCount) * 100, (i - 1 + 1) / (iCount) * 100, 1, i & "/" & iCount & " 件目を処理中です。"
        
    Next i
    
'------- 終了処理 -------

    iBar.FinalWait
    
    iBar.UnloadForm
    
End Sub

Public Sub サンプル2b()

Dim iBar As ProgressBar
Dim i As Long

Dim iCount As Long

'------- 初期設定 -------

    Set iBar = New ProgressBar
    
    Set iBar.TargetForm = UF_ProgressBar
    
    iBar.ShowForm
    
'------- 処理 -------

    iCount = 100
    
    For i = 1 To iCount
    
         '処理
        Application.Wait [Now()] + 10 / 86400000
        
        iBar.Message = i & "件目を処理中です。"
        
        If i Mod 2 = 0 Then    '表示する間隔をあけると表示が速くなる

            'メッセージをiBar.Messageとすると現在のメッセージを表示する（メッセージを変更しない）
            iBar.ProgressBarPaint i / (iCount + 1) * 100, (i + 1) / (iCount + 1) * 100, 1, iBar.Message
   
   
        End If
    
    Next i
    
        
'------- 終了処理 -------

    'まとめた単位によっては、100％まで到達しないことがあるので、最後に100％表示する
    iBar.ProgressBarPaint (i - 1) / (iCount + 1) * 100, i / (iCount + 1) * 100, 1, iBar.Message
    
    iBar.FinalWait
    
    iBar.UnloadForm
    
End Sub


Public Sub サンプル2c()

Dim iBar As ProgressBar
Dim i As Long

Dim iCount As Long

'------- 初期設定 -------

    Set iBar = New ProgressBar
    
    Set iBar.TargetForm = UF_ProgressBar
    
    iBar.ShowForm
    
'------- 処理 -------

    iCount = 100
    
    For i = 1 To iCount
    
         '処理
        Application.Wait [Now()] + 10 / 86400000
        
        iBar.Message = i & "件目を処理中です。"
        
        If i Mod 3 = 0 Then
            
            iBar.ProgressBarPaint i / (iCount + 1) * 100, (i + 1) / (iCount + 1) * 100, 1, iBar.Message
             
             
        End If
    
    Next i
    
        
'------- 終了処理 -------

    iBar.ProgressBarPaint (i - 1) / (iCount + 1) * 100, i / (iCount + 1) * 100, 1, iBar.Message
    
    iBar.FinalWait
    
    iBar.UnloadForm
    
End Sub


