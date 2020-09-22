VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Line Counter - Beta v1.1"
   ClientHeight    =   7485
   ClientLeft      =   150
   ClientTop       =   630
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFiles 
      Height          =   1425
      Left            =   360
      TabIndex        =   24
      Top             =   2280
      Width           =   5775
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset Counts"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Get Line Count"
      Default         =   -1  'True
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Frame frameResults 
      Caption         =   "Results:"
      Height          =   2535
      Left            =   720
      TabIndex        =   7
      Top             =   3840
      Width           =   5055
      Begin VB.Label lblClasses 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   840
         Width           =   1995
      End
      Begin VB.Label lblModules 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label lblForms 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2400
         TabIndex        =   21
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label Label9 
         Caption         =   "Number of Forms:"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Number of Modules:"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Number of Classes:"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   2040
         Width           =   1995
      End
      Begin VB.Label lblBlank 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label lblComments 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label lblCode 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label Label6 
         Caption         =   "Total Lines:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Blank Lines:"
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Lines of Comments:"
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Lines of Code:"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   6015
      Begin VB.CheckBox chkSub 
         Caption         =   "Search SubDirectories"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   285
         Left            =   5280
         TabIndex        =   2
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   1200
         Width           =   5055
      End
      Begin VB.ComboBox cmbMethod 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File or Directory Location:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1800
      End
      Begin VB.Label Label2 
         Caption         =   "Select the type of file(s) you want to get a line count for:"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuhlp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "A&bout"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim reg As New clsRegistry

Private Sub cmbMethod_Click()

  If cmbMethod.ListIndex = 0 Then
    chkSub.Enabled = True
  Else
    chkSub.Enabled = False
  End If

End Sub

Private Sub cmdBrowse_Click()

Dim oas As New OpenSaveDialog

  Select Case cmbMethod.ListIndex
    Case 0
      txtFile.Text = BrowseForFolder(Me.hwnd, "Select the foder you want to total all the lines in:")
    Case 1
      txtFile.Text = oas.OpenDialogBox(frmMain, fCustom, , , "VB Forms (*.frm)" + Chr$(0) + "*.frm" + Chr$(0) + "VB Modules (*.bas)" + Chr$(0) + "*.bas" + Chr$(0) + "VB Class Modules (*.cls)" + Chr$(0) + "*.cls" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0))
    Case 2
      txtFile.Text = oas.OpenDialogBox(frmMain, fCustom, , , "VB Projects (*.vbp)" + Chr$(0) + "*.vbp" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0))
    Case 3
      txtFile.Text = oas.OpenDialogBox(frmMain, fText)
    Case Else
      MsgBox "You must first select a method to browse by", vbCritical, "Duh"
  End Select
  
End Sub

Private Sub cmdExit_Click()
  
  Unload Me
  
End Sub

Private Sub cmdOK_Click()

Dim fName As String, fPath As String
Dim fFile As String, lCode As Long
Dim lComments As Long, lBlank As Long
Dim fNum As Integer, strData As String
Dim pName As String, pMajor As String
Dim pMinor As String, pRev As String
Dim x As Integer
Static FolderList As Collection

  'Make sure that the top half is filled in
  If cmbMethod.Text = "" Or txtFile.Text = "" Then
    MsgBox "You must select a method to search by and a file to search!", vbCritical, "Dumb Ass"
    Exit Sub
  End If
  
  Call ResetCounts(True)
  cmdOK.Enabled = False
  VB.Screen.MousePointer = 11
  
  'Reset counters
  cCode = 0
  cComments = 0
  cBlank = 0
  cTotal = 0
  cForms = 0
  cModules = 0
  cClasses = 0
  
  frameResults.Caption = "Results:"
  lstFiles.AddItem "All Files"
  
  If cmbMethod.ListIndex = 0 Then 'Directory
      
    If Left(txtFile.Text, 1) <> "\" Then
      fPath = txtFile.Text & "\"
    Else
      fPath = txtFile.Text
    End If
    
    If chkSub.Value = 1 Then
      Call GetAllDirsFrom(fPath, "frm", lstFiles)
      Call GetAllDirsFrom(fPath, "bas", lstFiles)
      Call GetAllDirsFrom(fPath, "cls", lstFiles)
    Else
    
      fName = Dir(fPath & "*.frm")
      Do Until fName = ""
        fFile = fPath & fName
        lstFiles.AddItem fFile
        fName = Dir
      Loop
      
      fName = Dir(fPath & "*.bas")
      Do Until fName = ""
        fFile = fPath & fName
        lstFiles.AddItem fFile
        fName = Dir
      Loop
      
      fName = Dir(fPath & "*.cls")
      Do Until fName = ""
        fFile = fPath & fName
        lstFiles.AddItem fFile
        fName = Dir
      Loop
      
    End If
    
    If lstFiles.ListCount = 1 Then
      MsgBox "There were no VB Forms, Modules or Class Modules found in this directory.  Please re-select and try again.", vbInformation, "No files found"
      Call ResetCounts(True)
      VB.Screen.MousePointer = 0
      cmdOK.Enabled = True
      Exit Sub
    End If
  
  ElseIf cmbMethod.ListIndex = 1 Then 'Single File
  
    If Right(txtFile.Text, 3) = "frm" Then
    ElseIf Right(txtFile.Text, 3) = "bas" Then
    ElseIf Right(txtFile.Text, 3) = "cls" Then
    Else
      'Not a valid file
      txtFile.Text = ""
      lstFiles.Clear
      MsgBox "The file you have selected is not a Visual Basic Form, Module, or Class Module.  Please re-select and try again.", vbCritical, "Bozo"
      VB.Screen.MousePointer = 0
      cmdOK.Enabled = True
      Exit Sub
    End If
    
    'Add file to listbox
    lstFiles.AddItem txtFile.Text
      
  ElseIf cmbMethod.ListIndex = 2 Then 'VB Project File
  
    fName = txtFile.Text
    If Right(fName, 3) <> "vbp" Then
      'Not a valid file
      txtFile.Text = ""
      lstFiles.Clear
      MsgBox "The file you have selected is not a Visual Basic Project file.  Please re-select and try again.", vbCritical, "Wake Up"
      VB.Screen.MousePointer = 0
      cmdOK.Enabled = True
      Exit Sub
    End If
    
    x = InStrRev(fName, "\")
    fPath = Left(fName, x)
    
    fNum = FreeFile
    
    Open fName For Input As fNum
    
      Do Until EOF(fNum)
        
        Line Input #fNum, strData
        
        If Left(strData, 5) = "Name=" Then
          pName = Mid(strData, 7)
          pName = Left(pName, Len(pName) - 1)
        ElseIf Left(strData, 9) = "MajorVer=" Then
          pMajor = Mid(strData, 10)
        ElseIf Left(strData, 9) = "MinorVer=" Then
          pMinor = Mid(strData, 10)
        ElseIf Left(strData, 12) = "RevisionVer=" Then
          pRev = Mid(strData, 13)
        ElseIf Left(strData, 5) = "Form=" Then
        
          fFile = GetFilePath(strData, fPath)
          lstFiles.AddItem fFile
            
        ElseIf Left(strData, 7) = "Module=" Then
        
          fFile = GetFilePath(strData, fPath)
          lstFiles.AddItem fFile
          
        ElseIf Left(strData, 6) = "Class=" Then
        
          fFile = GetFilePath(strData, fPath)
          lstFiles.AddItem fFile
          
        End If
        
      Loop
      
    Close #fNum
    
    frameResults.Caption = "Results: " & pName & " " & pMajor & "." & pMinor & "." & pRev
    
  ElseIf cmbMethod.ListIndex = 3 Then 'Other
  
    fNum = FreeFile
    Open txtFile.Text For Input As fNum
  
    Do Until EOF(fNum)
    
      Line Input #fNum, strData
      
      Call StripBeginingSpaces(strData)
      If strData = "" Then
        cBlank = cBlank + 1
      Else
        cCode = cCode + 1
      End If
    
    Loop
    
    Close #fNum
    
    lstFiles.Clear
    lblForms.Caption = cForms
    lblModules.Caption = cModules
    lblClasses.Caption = cClasses
    lblCode.Caption = cCode
    lblComments.Caption = cComments
    lblBlank.Caption = cBlank
    lblTotal.Caption = cCode + cComments + cBlank
    
    GoTo Other
    
  End If
    
  lstFiles.Refresh
  lstFiles.Selected(0) = True
  Call AnalyzeFile(lstFiles.Text)
  
Other:
  
  VB.Screen.MousePointer = 0
  cmdOK.Enabled = True
  
  reg.SaveSettingString Local_Machine, "Software\Rossi\VBLineCounter", "Method", cmbMethod.Text
  reg.SaveSettingString Local_Machine, "Software\Rossi\VBLineCounter", "File", txtFile.Text
  reg.SaveSettingLong Local_Machine, "Software\Rossi\VBLineCounter", "CheckSubs", chkSub.Value
  
End Sub

Private Sub cmdReset_Click()

  Call ResetCounts(True)

End Sub

Private Sub Form_Load()

  With cmbMethod
    .AddItem "Directory", 0
    .AddItem "Single VB Item (Form, Class, Module)", 1
    .AddItem "Visual Basic Project (.vbp)", 2
    .AddItem "Other", 3
  End With
  
  cmbMethod.Text = reg.GetSettingString(Local_Machine, "Software\Rossi\VBLineCounter", "Method", "Visual Basic Project (.vbp)")
  txtFile.Text = reg.GetSettingString(Local_Machine, "Software\Rossi\VBLineCounter", "File")
  chkSub.Value = reg.GetSettingLong(Local_Machine, "Software\Rossi\VBLineCounter", "CheckSubs")
  
End Sub

Public Sub ResetCounts(Optional ClearListBox As Boolean = False)
  
  If ClearListBox = True Then lstFiles.Clear
  
  'Reset counters
  cCode = 0
  cComments = 0
  cBlank = 0
  cTotal = 0
  cForms = 0
  cModules = 0
  cClasses = 0
  
  'Reset labels
  lblForms.Caption = 0
  lblModules.Caption = 0
  lblClasses.Caption = 0
  lblCode.Caption = 0
  lblComments.Caption = 0
  lblBlank.Caption = 0
  lblTotal.Caption = 0
  frameResults.Caption = "Results:"
  
End Sub

Private Sub lstFiles_DblClick()

  Call AnalyzeFile(lstFiles.Text)

End Sub

Public Function AnalyzeFile(FileName As String)

Dim fFile As String, lCode As Long
Dim lComments As Long, lBlank As Long
Dim x As Integer, lCount As Integer

  VB.Screen.MousePointer = 11
  
  fFile = FileName
  
  If fFile = "All Files" Then
    
    Call ResetCounts
    
    x = 1
    For x = 1 To lstFiles.ListCount - 1
      
      fFile = lstFiles.List(x)
      
      Call GetLineCount(fFile, lCode, lComments, lBlank)
      cCode = cCode + lCode
      cComments = cComments + lComments
      cBlank = cBlank + lBlank
      
      If Right(fFile, 3) = "frm" Then
        cForms = cForms + 1
      ElseIf Right(fFile, 3) = "bas" Then
        cModules = cModules + 1
      ElseIf Right(fFile, 3) = "cls" Then
        cClasses = cClasses + 1
      End If
      
      lblForms.Caption = cForms
      lblModules.Caption = cModules
      lblClasses.Caption = cClasses
      lblCode.Caption = cCode
      lblComments.Caption = cComments
      lblBlank.Caption = cBlank
      lblTotal.Caption = cCode + cComments + cBlank
      DoEvents
      
    Next x
    
  Else
  
    Call ResetCounts
  
    Call GetLineCount(fFile, lCode, lComments, lBlank)
    cCode = cCode + lCode
    cComments = cComments + lComments
    cBlank = cBlank + lBlank
    
    If Right(fFile, 3) = "frm" Then
      cForms = cForms + 1
    ElseIf Right(fFile, 3) = "bas" Then
      cModules = cModules + 1
    ElseIf Right(fFile, 3) = "cls" Then
      cClasses = cClasses + 1
    End If
    
    lblForms.Caption = cForms
    lblModules.Caption = cModules
    lblClasses.Caption = cClasses
    lblCode.Caption = cCode
    lblComments.Caption = cComments
    lblBlank.Caption = cBlank
    lblTotal.Caption = cCode + cComments + cBlank
    DoEvents
    
  End If
  
  VB.Screen.MousePointer = 0
  
End Function

Private Sub mnuAbout_Click()

  frmAbout.Show

End Sub

Private Sub mnuExit_Click()

  Unload Me

End Sub

