Attribute VB_Name = "modVBLineCount"
Option Explicit

Global cCode As Long, cComments As Long
Global cBlank As Long, cTotal As Long

Global cForms As Long, cModules As Long
Global cClasses As Long

Public Function GetLineCount(ByVal File As String, ByRef CodeCount As Long, CommentCount As Long, BlankCount As Long)

Dim fName As String, fNum As Integer
Dim strData As String, aFound As Boolean

  CodeCount = 0
  CommentCount = 0
  BlankCount = 0
  
  fName = File
  fNum = FreeFile
  
  If fName = "" Then
    MsgBox "Invalid File Name!", vbCritical, "Error"
    Exit Function
  End If
  
  Open fName For Input As fNum
  
  aFound = False
  
  Do Until EOF(fNum)
  
    Line Input #fNum, strData
    
    If Left(strData, 9) = "Attribute" And aFound = False Then
      aFound = True
    ElseIf Left(strData, 9) <> "Attribute" And aFound = True Then
            
      Call StripBeginingSpaces(strData)
      If strData = "" Then
        BlankCount = BlankCount + 1
      ElseIf Left(strData, 1) = "'" Then
        CommentCount = CommentCount + 1
      Else
        CodeCount = CodeCount + 1
      End If
      
    End If
    
  Loop
  
  Close #fNum

End Function

Public Function StripBeginingSpaces(ByRef strData As String)

  Do Until Left(strData, 1) <> " "
    strData = Right(strData, Len(strData) - 1)
  Loop

End Function

Public Function GetFilePath(ByVal Data As String, ByVal FilePath As String)

Dim fPath As String, strData As String
Dim x As Integer, fName As String
Dim fDir As String, oas As New OpenSaveDialog

  strData = Data
  fPath = FilePath
  
  If InStr(1, strData, "\") = 0 And InStr(1, strData, ";") = 0 Then
    
    x = InStr(1, strData, "=")
    GetFilePath = fPath & Right(strData, Len(strData) - x)
    
  ElseIf InStr(1, strData, "\") = 0 And InStr(1, strData, ";") <> 0 Then
  
    x = InStr(1, strData, ";")
    GetFilePath = fPath & Right(strData, Len(strData) - (x + 1))
        
  ElseIf InStr(1, strData, "\") <> 0 Then
    
    x = InStrRev(strData, "\")
    fName = Right(strData, Len(strData) - x)
    x = InStr(fPath, "\")
    fDir = Left(fPath, x)
    'Now that we have the filename, find the file
    GetFilePath = FindFile(fDir, fName)
    
    'If we can't find the file on the current drive, ask the user to point it out to us
    If GetFilePath = vbNullString Then
      
      MsgBox "The system could not find the file " & fName & " on your " & fDir & " directory.  Please select the file from the following window.", vbExclamation, "Cannot find file"
      GetFilePath = oas.OpenDialogBox(frmMain, fCustom, "U:\VB", , "VB Forms (*.frm)" + Chr$(0) + "*.frm" + Chr$(0) + "VB Modules (*.bas)" + Chr$(0) + "*.bas" + Chr$(0) + "VB Class Modules (*.cls)" + Chr$(0) + "*.cls" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0))
    
    End If
    
  End If

End Function
