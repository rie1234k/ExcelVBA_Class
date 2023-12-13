VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BookCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private filepath_ As String
'�t�@�C���p�X���擾
Public Property Get FilePath() As String

    FilePath = filepath_

End Property
'�t�@�C���p�X��ݒ�
Public Property Let FilePath(ByVal vNewValue As String)

    filepath_ = vNewValue


End Property
'�t�@�C�����J����Ă��邩�m�F����i�t�@�C���T�[�o�[��ő��̃��[�U�[���J���Ă��邩�ǂ����m�F�j
Public Property Get Opened() As Boolean

    On Error Resume Next
    
    Open filepath_ For Append As #1
    Close #1
    
    If Err.Number > 0 Then
    
        Opened = True
    
    Else
    
        Opened = False
    
    End If
    

End Property
'�t�@�C�������݂��邩�m�F����
Public Property Get Exist() As Boolean
Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Exist = FSO.fileExists(filepath_)
    
    Set FSO = Nothing
     
     
End Property
'�t�@�C�������擾
Public Property Get FileName() As String
Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    FileName = FSO.GetFileName(filepath_)
    
    Set FSO = Nothing
    

End Property
'�t�H���_�����擾
Public Property Get ForderName() As String
Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ForderName = Replace(filepath_, "\" & FSO.GetFileName(filepath_), "")
    
    Set FSO = Nothing
    

End Property
'�������t�@�C�����J���Ă��邩�ǂ����m�F
Public Property Get OpenedMe() As Boolean
Dim iBook As Workbook

    OpenedMe = False
    
    For Each iBook In Workbooks
    
        If iBook.Name = FileName Then
    
            OpenedMe = True
            
        End If
        
    Next iBook


End Property
'�u�b�N���A�N�e�B�u�ɂ���
Public Sub BookActive()


    Workbooks(FileName).Activate


End Sub
'�u�b�N��ۑ������ɕ���
Public Sub BookCloseNoSave()

    Workbooks(FileName).Saved = True
    
    Workbooks(FileName).Close False
    

End Sub
'�t�@�C�����J��
Public Function BookOpen() As Workbook

    Set BookOpen = Workbooks.Open(filepath_)

End Function
'�p�X���[�h�t���t�@�C�����J��
Public Function BookOpenLocked(myPassword As String) As Workbook

    Set BookOpenLocked = Workbooks.Open(FileName:=filepath_, Password:=myPassword)

End Function
'�ǂݎ���p�ŊJ��
Public Function ReadOnlyOpen() As Workbook

    Set ReadOnlyOpen = Workbooks.Open(FileName:=filepath_, ReadOnly:=True)

End Function
'�p�X���[�h�t���t�@�C����ǂݎ���p�ŊJ��
Public Function ReadOnlyOpenLocked(myPassword As String) As Workbook

    Set ReadOnlyOpenLocked = Workbooks.Open(FileName:=filepath_, Password:=myPassword, ReadOnly:=True)

End Function
'�t�@�C���������N���X�V�����ɊJ��
Public Function NoUpdateLinksOpen() As Workbook

    Set NoUpdateLinksOpen = Workbooks.Open(FileName:=filepath_, UpdateLinks:=False)


End Function
'�t�@�C���������N���X�V�����ɓǂݎ���p�ŊJ��
Public Function NoUpdateLinksReadOnlyOpen() As Workbook

    Set NoUpdateLinksOpen = Workbooks.Open(FileName:=filepath_, UpdateLinks:=False, ReadOnly:=True)


End Function
'�u�b�N�I�u�W�F�N�g���擾
Public Property Get SetBook() As Workbook

Set SetBook = Workbooks(FileName)

End Property
