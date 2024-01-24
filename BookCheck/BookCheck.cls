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
'ファイルパスを取得
Public Property Get FilePath() As String

    FilePath = filepath_

End Property
'ファイルパスを設定
Public Property Let FilePath(ByVal vNewValue As String)

    filepath_ = vNewValue


End Property
'ファイルが開かれているか確認する（ファイルサーバー上で他のユーザーが開いているかどうか確認）
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
'ファイルが存在するか確認する
Public Property Get Exist() As Boolean
Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Exist = FSO.fileExists(filepath_)
    
    Set FSO = Nothing
     
     
End Property
'ファイル名を取得
Public Property Get FileName() As String
Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    FileName = FSO.GetFileName(filepath_)
    
    Set FSO = Nothing
    

End Property
'フォルダ名を取得
Public Property Get ForderName() As String
Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ForderName = Replace(filepath_, "\" & FSO.GetFileName(filepath_), "")
    
    Set FSO = Nothing
    

End Property
'自分がファイルを開いているかどうか確認
Public Property Get OpenedMe() As Boolean
Dim iBook As Workbook

    OpenedMe = False
    
    For Each iBook In Workbooks
    
        If iBook.Name = FileName Then
    
            OpenedMe = True
            
        End If
        
    Next iBook


End Property
'ブックをアクティブにする
Public Sub BookActive()


    Workbooks(FileName).Activate


End Sub
'ブックを保存せずに閉じる
Public Sub BookCloseNoSave()

    Workbooks(FileName).Saved = True
    
    Workbooks(FileName).Close False
    

End Sub
'ファイルを開く
Public Function BookOpen() As Workbook

    Set BookOpen = Workbooks.Open(filepath_)

End Function
'パスワード付きファイルを開く
Public Function BookOpenLocked(myPassword As String) As Workbook

    Set BookOpenLocked = Workbooks.Open(FileName:=filepath_, Password:=myPassword)

End Function
'読み取り専用で開く
Public Function ReadOnlyOpen() As Workbook

    Set ReadOnlyOpen = Workbooks.Open(FileName:=filepath_, ReadOnly:=True)

End Function
'パスワード付きファイルを読み取り専用で開く
Public Function ReadOnlyOpenLocked(myPassword As String) As Workbook

    Set ReadOnlyOpenLocked = Workbooks.Open(FileName:=filepath_, Password:=myPassword, ReadOnly:=True)

End Function
'ファイルをリンクを更新せずに開く
Public Function NoUpdateLinksOpen() As Workbook

    Set NoUpdateLinksOpen = Workbooks.Open(FileName:=filepath_, UpdateLinks:=False)


End Function
'ファイルをリンクを更新せずに読み取り専用で開く
Public Function NoUpdateLinksReadOnlyOpen() As Workbook

    Set NoUpdateLinksReadOnlyOpen = Workbooks.Open(FileName:=filepath_, UpdateLinks:=False, ReadOnly:=True)


End Function
'ブックオブジェクトを取得
Public Property Get SetBook() As Workbook

Set SetBook = Workbooks(FileName)

End Property

