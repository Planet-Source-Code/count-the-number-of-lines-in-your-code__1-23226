VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OpenSaveDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type OpenFileName
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
  "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
  
Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
  Alias "GetSaveFileNameA" (pOpenfilename As OpenFileName) As Long
  
Public Enum FilterType
  fText = 0
  fDatabase = 1
  fExcel = 2
  fSystem = 3
  fCustom = 4
End Enum

Public Function OpenDialogBox(FormName As Form, Filter As FilterType, Optional InitialDir As String, Optional DialogTitle As String, Optional CustomFilter As String) As String
'
Dim ofn As OpenFileName, a
Dim strInitDir As String, dlgTitle As String
Dim strFilter
  '
  If Not IsEmpty(InitialDir) Then
    strInitDir = InitialDir
  Else
    strInitDir = CurDir
  End If
  '
  If Not IsEmpty(DialogTitle) Then
    dlgTitle = DialogTitle
  Else
    dlgTitle = "Open"
  End If
  '
  If IsEmpty(CustomFilter) And Filter = fCustom Then
    Filter = fText
  Else
    strFilter = CustomFilter
  End If
  '
  If Filter = fText Then
    strFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Log Files (*.log)" + Chr$(0) + "*.log" + Chr$(0) + "List Files (*.lst)" + Chr$(0) + "*.lst" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  ElseIf Filter = fDatabase Then
    strFilter = "MSAccess Database (*.mdb)" + Chr$(0) + "*.mdb" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  ElseIf Filter = fExcel Then
    strFilter = "Excel Files (*.xls)" + Chr$(0) + "*.xls" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  ElseIf Filter = fSystem Then
    strFilter = "System Files (*.dll)" + Chr$(0) + "*.dll" + Chr$(0) + "ActiveX Files (*.ocx)" + Chr$(0) + "*.ocx" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  End If
  '
  ofn.lStructSize = Len(ofn)
  ofn.hwndOwner = FormName.hWnd
  ofn.hInstance = App.hInstance
  'ofn.lpstrFilter = "Log Files (*.log)" + Chr$(0) + "*.log" + Chr$(0) + "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  ofn.lpstrFilter = strFilter
  ofn.lpstrFile = Space$(254)
  ofn.nMaxFile = 255
  ofn.lpstrFileTitle = Space$(254)
  ofn.nMaxFileTitle = 255
  ofn.lpstrInitialDir = strInitDir
  ofn.lpstrTitle = dlgTitle
  ofn.flags = 0
  '
  a = GetOpenFileName(ofn)
  '
  If (a) Then
    'Do the file open stuff here
    OpenDialogBox = Trim$(ofn.lpstrFile)
    '
  Else
    'Cancel was pressed
    OpenDialogBox = ""
    '
  End If
  '
  '
End Function

Public Function SaveDialogBox(FormName As Form, Filter As FilterType, Optional InitialDir As String, Optional DialogTitle As String, Optional CustomFilter As String) As String
'
Dim ofn As OpenFileName, a
Dim strInitDir As String, dlgTitle As String
Dim strFilter As String
  '
  If Not IsEmpty(InitialDir) Then
    strInitDir = InitialDir
  Else
    strInitDir = CurDir
  End If
  '
  If Not IsEmpty(DialogTitle) Then
    dlgTitle = DialogTitle
  Else
    dlgTitle = "Save As"
  End If
  '
  If IsEmpty(CustomFilter) And Filter = fCustom Then
    Filter = fText
  Else
    strFilter = CustomFilter
  End If
  '
  If Filter = fText Then
    strFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Log Files (*.log)" + Chr$(0) + "*.log" + Chr$(0) + "List Files (*.lst)" + Chr$(0) + "*.lst" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  ElseIf Filter = fDatabase Then
    strFilter = "MSAccess Database (*.mdb)" + Chr$(0) + "*.mdb" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  ElseIf Filter = fExcel Then
    strFilter = "Excel Files (*.xls)" + Chr$(0) + "*.xls" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
  End If
  '
  ofn.lStructSize = Len(ofn)
  ofn.hwndOwner = FormName.hWnd
  ofn.hInstance = App.hInstance
  'ofn.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Rich Text Files (*.rtf)" + Chr$(0) + "*.rtf" + Chr$(0)
  ofn.lpstrFilter = strFilter
  ofn.lpstrFile = Space$(254)
  ofn.nMaxFile = 255
  ofn.lpstrFileTitle = Space$(254)
  ofn.nMaxFileTitle = 255
  ofn.lpstrInitialDir = strInitDir
  ofn.lpstrTitle = dlgTitle
  ofn.flags = 0
  '
  a = GetSaveFileName(ofn)
  '
  If (a) Then
    'FileSave Stuff Here
    SaveDialogBox = Trim$(ofn.lpstrFile)
  Else
    'Cancel was pressed
    SaveDialogBox = False
  End If
  '
End Function

