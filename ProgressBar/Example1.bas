Attribute VB_Name = "�T���v��1"
Option Explicit


Public Sub �T���v��1a()

Dim iBar As ProgressBar
    
'------- �����ݒ� -------

    '�C���X�^���X��
    Set iBar = New ProgressBar
    
    '���[�U�[�t�H�[����ݒ�
    Set iBar.TargetForm = UF_ProgressBar
    
    iBar.ShowForm
               
    iBar.Message = "�������J�n���܂��B"   '�v���O���X�o�[�̃��b�Z�[�W��ݒ�
   
'------- ���� -------
    
    '�����P
    Application.Wait Now() + TimeValue("00:00:02")
    
    '�����́A�J�n�p�[�Z���g(0�`100)�A�I���p�[�Z���g(0�`100)�A�����P��(1�`)�A���b�Z�[�W
    iBar.ProgressBarPaint 0, 33, 1, "�����P���������ł��B"
    
    
    
    '�����Q
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 33, 66, 1, "�����Q���������ł��B"
    
    
    
    '�����R
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 66, 100, 1, "�����R���������ł��B"
    
'------- �I������ -------
 
    iBar.FinalWait  '�����҂��čŌ��100%��\��������i�Ȃ��Ă��悢�j
    iBar.UnloadForm
    
End Sub
Public Sub �T���v��1b()

Dim iBar As ProgressBar
    
'------- �����ݒ� -------

 
    Set iBar = New ProgressBar
    

    Set iBar.TargetForm = UF_ProgressBar
    
    iBar.ShowForm
    
    iBar.Message = "�������J�n���܂��B"
   
'------- ���� -------
    
    '�����P
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 0, 33, 2, "�����P���������ł��B"
    
    
    
    '�����Q
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 33, 66, 2, "�����Q���������ł��B"
    
    
    
    '�����R
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 66, 100, 2, "�����R���������ł��B"
    
'------- �I������ -------
 
    iBar.FinalWait
    iBar.UnloadForm
    
End Sub
Public Sub �T���v��1c()

Dim iBar As ProgressBar
    
'------- �����ݒ� -------

    Set iBar = New ProgressBar
    

    Set iBar.TargetForm = UF_ProgressBar
    
    iBar.ShowForm
    
     iBar.Message = "�������J�n���܂��B"
   
'------- ���� -------
    
    '�����P
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 0, 33, 3, "�����P���������ł��B"
    
    
    '�����Q
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 33, 66, 3, "�����Q���������ł��B"
    
    
    
    '�����R
    Application.Wait Now() + TimeValue("00:00:02")
    
    iBar.ProgressBarPaint 66, 100, 3, "�����R���������ł��B"
    
'------- �I������ -------
 
    iBar.FinalWait
    iBar.UnloadForm
        
End Sub
