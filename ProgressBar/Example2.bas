Attribute VB_Name = "�T���v��2"
Option Explicit

Public Sub �T���v��2a()

Dim iBar As ProgressBar
Dim i As Long

Dim iCount As Long

    
'------- �����ݒ� -------

    Set iBar = New ProgressBar
    
    Set iBar.TargetForm = UF_ProgressBar
    
    iBar.ShowForm

    
'------- ���� -------

    iCount = 100
     
    For i = 1 To iCount

        '����
        Application.Wait [Now()] + 10 / 86400000
        
        
        iBar.ProgressBarPaint (i - 1) / (iCount) * 100, (i - 1 + 1) / (iCount) * 100, 1, i & "/" & iCount & " ���ڂ��������ł��B"
        
    Next i
    
'------- �I������ -------

    iBar.FinalWait
    
    iBar.UnloadForm
    
End Sub

Public Sub �T���v��2b()

Dim iBar As ProgressBar
Dim i As Long

Dim iCount As Long

'------- �����ݒ� -------

    Set iBar = New ProgressBar
    
    Set iBar.TargetForm = UF_ProgressBar
    
    iBar.ShowForm
    
'------- ���� -------

    iCount = 100
    
    For i = 1 To iCount
    
         '����
        Application.Wait [Now()] + 10 / 86400000
        
        iBar.Message = i & "���ڂ��������ł��B"
        
        If i Mod 2 = 0 Then    '�\������Ԋu��������ƕ\���������Ȃ�

            '���b�Z�[�W��iBar.Message�Ƃ���ƌ��݂̃��b�Z�[�W��\������i���b�Z�[�W��ύX���Ȃ��j
            iBar.ProgressBarPaint i / (iCount + 1) * 100, (i + 1) / (iCount + 1) * 100, 1, iBar.Message
   
   
        End If
    
    Next i
    
        
'------- �I������ -------

    '�܂Ƃ߂��P�ʂɂ���ẮA100���܂œ��B���Ȃ����Ƃ�����̂ŁA�Ō��100���\������
    iBar.ProgressBarPaint (i - 1) / (iCount + 1) * 100, i / (iCount + 1) * 100, 1, iBar.Message
    
    iBar.FinalWait
    
    iBar.UnloadForm
    
End Sub


Public Sub �T���v��2c()

Dim iBar As ProgressBar
Dim i As Long

Dim iCount As Long

'------- �����ݒ� -------

    Set iBar = New ProgressBar
    
    Set iBar.TargetForm = UF_ProgressBar
    
    iBar.ShowForm
    
'------- ���� -------

    iCount = 100
    
    For i = 1 To iCount
    
         '����
        Application.Wait [Now()] + 10 / 86400000
        
        iBar.Message = i & "���ڂ��������ł��B"
        
        If i Mod 3 = 0 Then
            
            iBar.ProgressBarPaint i / (iCount + 1) * 100, (i + 1) / (iCount + 1) * 100, 1, iBar.Message
             
             
        End If
    
    Next i
    
        
'------- �I������ -------

    iBar.ProgressBarPaint (i - 1) / (iCount + 1) * 100, i / (iCount + 1) * 100, 1, iBar.Message
    
    iBar.FinalWait
    
    iBar.UnloadForm
    
End Sub


