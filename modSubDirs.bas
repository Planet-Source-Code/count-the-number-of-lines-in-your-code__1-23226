Attribute VB_Name = "modSubDirs"
Option Explicit

Public Function GetAllDirsFrom(ByVal pstrDir As String, ByVal Extension As String, ByVal ListBox As ListBox)
    
Dim fso As FileSystemObject
Dim fldrMain As Folder
Dim fldrsSub As Folders
Dim fldr As Folder
    
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set fldrMain = fso.GetFolder(pstrDir & "\")
  
  If Right(fldrMain.Path, 1) = "\" Then
      AddAllFilesFrom Left(fldrMain.Path, Len(fldrMain.Path) - 1), Extension, ListBox
  Else
      AddAllFilesFrom fldrMain.Path, Extension, ListBox
  End If
  
  ' Recurse subdirectories
  Set fldrsSub = fldrMain.SubFolders
  For Each fldr In fldrsSub
      GetAllDirsFrom fldr.Path, Extension, ListBox
  Next
  
  ListBox.Refresh
  
End Function

Public Function AddAllFilesFrom(ByVal pstrDir As String, ByVal Extension As String, ByVal ListBox As ListBox)

Dim strfile

  strfile = pstrDir & "\" & Dir(pstrDir & "\*." & Extension)
  Do Until strfile = pstrDir & "\"
    ListBox.AddItem strfile
    strfile = pstrDir & "\" & Dir
  Loop
    
End Function
