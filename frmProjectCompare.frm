VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmProjectCompare 
   Caption         =   "Project compare"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCompare 
      Caption         =   "Compare"
      Enabled         =   0   'False
      Height          =   495
      Left            =   90
      TabIndex        =   8
      Top             =   7380
      Width           =   1605
   End
   Begin RichTextLib.RichTextBox txtResults 
      Height          =   2505
      Left            =   60
      TabIndex        =   4
      Top             =   4830
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   4419
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmProjectCompare.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   5190
      Top             =   8160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSecondProject 
      Caption         =   "Second project..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   6210
      TabIndex        =   3
      Top             =   3840
      Width           =   1605
   End
   Begin VB.ListBox lstSecondProject 
      Height          =   3570
      Left            =   6210
      TabIndex        =   2
      Top             =   210
      Width           =   5595
   End
   Begin VB.CommandButton cmdFirstProject 
      Caption         =   "First project..."
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   3840
      Width           =   1605
   End
   Begin VB.ListBox lstFirstProject 
      Height          =   3570
      Left            =   0
      TabIndex        =   0
      Top             =   210
      Width           =   5895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Second project files:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   6210
      TabIndex        =   7
      Top             =   0
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "First Project files:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Results:"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   4620
      Width           =   570
   End
End
Attribute VB_Name = "frmProjectCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type


Private Type FileCompare
    FirstFile As String
    SecondFile As String
End Type

Private Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FileTime
    ftLastAccessTime As FileTime
    ftLastWriteTime As FileTime
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type

Const OFS_MAXPATHNAME = 128
Const OF_CREATE = &H1000
Const OF_READ = &H0
Const OF_WRITE = &H1
Private Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type


Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function CompareFileTime Lib "kernel32" (lpFileTime1 As FileTime, lpFileTime2 As FileTime) As Long

Private Sub cmdCompare_Click()

    Me.txtResults.Text = ""
    ShowResults
    
End Sub

Private Sub cmdFirstProject_Click()
    
    '*Purpose:Brows for a vb project or just one more free file
    '*        if vb project found than i load all its relevant code files
    
    Dim sFileName       As String
    Dim oDir            As Scripting.FileSystemObject
    
    
    'open dialog box
    dlg1.Filter = "VB Project (*.vbp)|*.vbp|All files (*.*)|*.*"
    dlg1.DefaultExt = ".vbp"
    dlg1.ShowOpen
    
    Set oDir = New Scripting.FileSystemObject
    
    'checks for valid file
    If Not dlg1.CancelError Then
        sFileName = dlg1.Filename
        
        'Clear listboxes
        Me.lstFirstProject.Clear
                
        'checks for file type
        If oDir.GetExtensionName(sFileName) <> "vbp" Then
            AddFileName Me.lstFirstProject, sFileName   'other file
        Else
            AddProject Me.lstFirstProject, sFileName 'vb project - add all relevant files
        End If
        
    End If
    
    Me.cmdSecondProject.Enabled = True
    
End Sub

Private Sub AddFileName(ByRef Lst As Object, ByVal sFileName As String)


    On Error GoTo Err_Proc

    '*Purpose:adds new file to the files list
    
    If LenB(sFileName) > 0 Then
        If Not FileInList(Lst, sFileName) Then
            Lst.AddItem sFileName
            Lst.Selected(Lst.NewIndex) = True
        End If
    End If

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "AddFileName", Err, Err_Handle_Mode
    Resume Exit_Proc


End Sub

Private Sub AddProject(ByRef Lst As Object, ByVal sFileName As String)


    On Error GoTo Err_Proc

    '*Purpose: adds the selected project (all its files) to the system manager
    
    Dim oDir        As Scripting.FileSystemObject
    Dim ff          As Long
    Dim i           As Long
    Dim sline       As String
    Dim sObjectName As String
    Dim sPath       As String
    
    Set oDir = New Scripting.FileSystemObject
    
    '*ensures backslash is exists
    sPath = oDir.GetParentFolderName(sFileName)
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    'open file port
    ff = FreeFile
    
    Open sFileName For Input As #ff
    
    'scan vb project file
    Do Until EOF(ff)
        Line Input #ff, sline 'read next line in the project file
        
        'check for the next object:
        If InStr(1, LCase$(sline), "form=") > 0 Then
            i = InStr(1, sline, "=") + 1
            sObjectName = Mid$(sline, i, Len(sline) - i + 1) 'find object name
            
            '*check that there is no (") in the object name
            If InStr(1, sObjectName, Chr$(34)) = 0 Then
                AddFileName Lst, sPath & sObjectName 'add file to list
            End If
            
        End If
        
        If InStr(1, LCase$(sline), "class=") > 0 Then
            i = InStr(1, sline, ";") + 2
            sObjectName = Mid$(sline, i, Len(sline) - i + 1)
            AddFileName Lst, sPath & sObjectName 'add file to list
        End If
        
        If InStr(1, LCase$(sline), "module=") > 0 Then
            i = InStr(1, sline, ";") + 2
            sObjectName = Mid$(sline, i, Len(sline) - i + 1)
            AddFileName Lst, sPath & sObjectName 'add file to list
        End If
        
        If InStr(1, LCase$(sline), "usercontrol=") > 0 Then
            i = InStr(1, sline, "=") + 1
            sObjectName = Mid$(sline, i, Len(sline) - i + 1)
            AddFileName Lst, sPath & sObjectName 'add file to list
        End If
        
    Loop
    
    'close project file port
    Close #ff
    
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "AddProject", Err, Err_Handle_Mode
    Resume Exit_Proc


End Sub

Private Function FileInList(ByRef Lst As Object, ByVal sFileName As String) As Boolean

    'check if the specified file is allready in the list
    
    On Error GoTo Err_Proc
    
    Dim i           As Long
    
    FileInList = False
    
    For i = 0 To Lst.ListCount - 1
        FileInList = (sFileName = Lst.List(i))
        If FileInList Then Exit For
    Next i
    
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "FileInList", Err, Err_Handle_Mode
Resume Exit_Proc


End Function

Private Sub cmdSecondProject_Click()
    '*Purpose:Brows for a vb project or just one more free file
    '*        if vb project found than i load all its relevant code files
    
    Dim sFileName       As String
    Dim oDir            As Scripting.FileSystemObject
    
    
    'open dialog box
    dlg1.Filter = "VB Project (*.vbp)|*.vbp|All files (*.*)|*.*"
    dlg1.DefaultExt = ".vbp"
    dlg1.ShowOpen
    
    Set oDir = New Scripting.FileSystemObject
    
    'checks for valid file
    If Not dlg1.CancelError Then
        sFileName = dlg1.Filename
        
        'Clear listboxes
        Me.lstSecondProject.Clear
                
        'checks for file type
        If oDir.GetExtensionName(sFileName) <> "vbp" Then
            AddFileName Me.lstSecondProject, sFileName   'other file
        Else
            AddProject Me.lstSecondProject, sFileName 'vb project - add all relevant files
        End If
        
    End If

    Me.cmdCompare.Enabled = True
    
End Sub

Private Function ShowResults(Optional ByRef ErrorMsg As String = "") As Boolean

    Dim objFile     As Scripting.FileSystemObject
    Dim iCheckList  As Long
    
    Set objFile = New Scripting.FileSystemObject

    ShowResults = False
    
    If Me.lstFirstProject.ListCount = 0 Or Me.lstSecondProject.ListCount = 0 Then
        ErrorMsg = "One project is missing, please select project and try again."
        Exit Function
    End If
    
    Dim i           As Long
    Dim objLst      As Object
    Dim objLst2     As Object
    Dim sTmp        As String
    
    If Me.lstFirstProject.ListCount > Me.lstSecondProject.ListCount Then
        iCheckList = 1
        Set objLst = Me.lstFirstProject
        Set objLst2 = Me.lstSecondProject
    Else
        iCheckList = 2
        Set objLst = Me.lstSecondProject
        Set objLst2 = Me.lstFirstProject
    End If
 
 
    Dim arrFiles()      As FileCompare
    Dim iCounter        As Long
        
    ' Build file compare collection
    ReDim arrFiles(0)
    iCounter = 0
    For i = 0 To objLst.ListCount - 1
        sTmp = FileExists(objLst2, objLst.List(i))
        If sTmp <> "" Then
            ReDim Preserve arrFiles(iCounter)
            
            If iCheckList = 1 Then
                arrFiles(iCounter).FirstFile = objLst.List(i)
                arrFiles(iCounter).SecondFile = sTmp
            Else
                arrFiles(iCounter).FirstFile = sTmp
                arrFiles(iCounter).SecondFile = objLst.List(i)
            End If
            
            iCounter = iCounter + 1
        End If
    Next i
    
    Dim hFile           As Long
    Dim hFile2          As Long
    Dim iRet            As Long
    Dim iSelCursor      As Long
    Dim FileInfo        As BY_HANDLE_FILE_INFORMATION
    Dim FileInfo2       As BY_HANDLE_FILE_INFORMATION
    Dim OF              As OFSTRUCT
    
    iSelCursor = 0
    For i = 0 To UBound(arrFiles)
    
        hFile = OpenFile(arrFiles(i).FirstFile, OF, OF_READ)
        hFile2 = OpenFile(arrFiles(i).SecondFile, OF, OF_READ)
        
        GetFileInformationByHandle hFile, FileInfo
        GetFileInformationByHandle hFile2, FileInfo2
        
        ' Compare file time
        iRet = CompareFileTime(FileInfo.ftLastWriteTime, FileInfo2.ftLastWriteTime)
        
        With Me.txtResults
            
            
            Select Case iRet
                Case -1 ' First file is less than the second
                    sTmp = "File: " & objFile.GetFilename(arrFiles(i).FirstFile) & "  ( Second file is newer )." & vbNewLine
                    .Text = .Text & sTmp
                    ApplyFontFormat iSelCursor, sTmp
                    iSelCursor = iSelCursor + Len(sTmp)
                Case 0 ' File times are equal
                
                Case 1 ' First file is greater
                    sTmp = "File: " & objFile.GetFilename(arrFiles(i).FirstFile) & "  ( First file is newer )." & vbNewLine
                    .Text = .Text & sTmp
                    ApplyFontFormat iSelCursor, sTmp
                    iSelCursor = iSelCursor + Len(sTmp)
            End Select
        End With
        
        CloseHandle hFile
        CloseHandle hFile2
        
    Next i
        
End Function

Private Function FileExists(Lst As Object, ByVal str As String) As String

    Dim i           As Long
    Dim objFile     As Scripting.FileSystemObject
    
    Set objFile = New Scripting.FileSystemObject
    For i = 0 To Lst.ListCount - 1
        If Trim(objFile.GetFilename(Lst.List(i))) = objFile.GetFilename(str) Then
            FileExists = Lst.List(i)
            Exit For
        End If
    Next i
    
    
End Function

Private Function ApplyFontFormat(SelStart As Long, ByRef str As String) As Boolean

    Dim i As Long, i2 As Long
    
    With Me.txtResults
        ' Set 'file' word
        .SelStart = SelStart
        .SelLength = 5
        .SelBold = True
        .SelColor = vbRed
        
        ' Set result desc
        i = InStr(1, str, "("): i2 = InStr(i, str, ")")
        .SelStart = SelStart + i
        .SelLength = i2 - i - 1
        .SelColor = vbBlue
        
        ' Set file name to bold
        .SelStart = SelStart + 5
        .SelLength = i - 5 - 1
        .SelBold = True
        
        
    End With
    
End Function
