VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Picker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private setrange_ As Range

Public Property Set SetRange(ByVal NewRange As Range)

    Set setrange_ = NewRange

End Property

Public Property Get SetRange() As Range

    Set SetRange = setrange_

End Property

'ダイアログボックスに最初に表示されるフォルダの設定用
Public Property Get ForderPath() As String
Dim FSO As Object

    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If FSO.FolderExists(setrange_.Value) Then
    
        ForderPath = FSO.GetParentFolderName(setrange_.Value)

    End If
    
    If FSO.fileExists(setrange_.Value) Then
    
        ForderPath = Replace(setrange_.Value, "\" & FSO.GetFileName(setrange_.Value), "")
 
    End If
    
    If ForderPath = "" Then ForderPath = ThisWorkbook.Path
    
    Set FSO = Nothing
    

End Property

Public Sub FolderPicker()
Dim SelectFolderPath As Variant

  With Application.FileDialog(msoFileDialogFolderPicker)
   
   .InitialFileName = ForderPath
   
   If .Show = 0 Then Exit Sub
   
   SelectFolderPath = .SelectedItems(1)
   
   
   
  End With
  
  setrange_.Value = SelectFolderPath
  
  
End Sub



Public Sub FilePicker()
Dim SelectFilePath As Variant

 With Application.FileDialog(msoFileDialogFilePicker)
   
   .InitialFileName = ForderPath
   
   If .Show = 0 Then Exit Sub
   
   SelectFilePath = .SelectedItems(1)
   
  End With
  
  setrange_.Value = SelectFilePath

End Sub

Public Sub HyperlinkAdd()
     
    If setrange_.Value <> "" Then
        
        setrange_.Worksheet.Hyperlinks.Add Anchor:=setrange_, Address:=setrange_.Value, TextToDisplay:=setrange_.Value
    
    End If
    
End Sub

