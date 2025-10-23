Attribute VB_Name = "サンプル1"
Option Explicit


Public Sub サンプル1a()

Dim iBar As ProgressBar
    
'------- 初期設定 -------

    'インスタンス化
    Set iBar = New ProgressBar
    
    'ユーザーフォームを設定
    Set iBar.TargetForm = UF_ProgressBar
    
    iBar.ShowForm
               
    iBar.Message = "処理を開始します。"   'プログレスバーのメッセージを設定
   
'------- 処理 -------
    
    '処理１
    Application.Wait Now() + TimeValue("00:00:02")
    
    '引数は、開始パーセント(0～100)、終了パーセント(0～100)、増加単位(1～)、メッセージ
    iBar.ProgressBarPaint 0, 33, 1, "処理１を処理中です。"
    
    
    
    '処理２
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 33, 66, 1, "処理２を処理中です。"
    
    
    
    '処理３
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 66, 100, 1, "処理３を処理中です。"
    
'------- 終了処理 -------
 
    iBar.FinalWait  '少し待って最後の100%を表示させる（なくてもよい）
    iBar.UnloadForm
    
End Sub
Public Sub サンプル1b()

Dim iBar As ProgressBar
    
'------- 初期設定 -------

 
    Set iBar = New ProgressBar
    

    Set iBar.TargetForm = UF_ProgressBar
    
    iBar.ShowForm
    
    iBar.Message = "処理を開始します。"
   
'------- 処理 -------
    
    '処理１
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 0, 33, 2, "処理１を処理中です。"
    
    
    
    '処理２
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 33, 66, 2, "処理２を処理中です。"
    
    
    
    '処理３
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 66, 100, 2, "処理３を処理中です。"
    
'------- 終了処理 -------
 
    iBar.FinalWait
    iBar.UnloadForm
    
End Sub
Public Sub サンプル1c()

Dim iBar As ProgressBar
    
'------- 初期設定 -------

    Set iBar = New ProgressBar
    

    Set iBar.TargetForm = UF_ProgressBar
    
    iBar.ShowForm
    
     iBar.Message = "処理を開始します。"
   
'------- 処理 -------
    
    '処理１
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 0, 33, 3, "処理１を処理中です。"
    
    
    '処理２
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 33, 66, 3, "処理２を処理中です。"
    
    
    
    '処理３
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 66, 100, 3, "処理３を処理中です。"
    
'------- 終了処理 -------
 
    iBar.FinalWait
    iBar.UnloadForm
        
End Sub
