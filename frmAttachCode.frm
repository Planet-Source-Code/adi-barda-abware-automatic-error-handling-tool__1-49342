VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAttachCode 
   Caption         =   "Attach Code"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   11880
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   3540
      Top             =   4620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrows 
      Caption         =   "Brows"
      Height          =   375
      Left            =   60
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Add vb project to the list"
      Top             =   4650
      Width           =   855
   End
   Begin VB.ListBox lstSelectedFiles 
      Height          =   4335
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   270
      Width           =   7365
   End
   Begin VB.ListBox lstFunctions 
      DragIcon        =   "frmAttachCode.frx":0000
      Height          =   4335
      Left            =   7590
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Drag one of the functions from this list to the code textbox in the previous window"
      Top             =   270
      Width           =   4125
   End
   Begin VB.CommandButton cmdSelectFiles 
      Caption         =   "+"
      Height          =   285
      Left            =   7020
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Select all"
      Top             =   4650
      Width           =   405
   End
   Begin VB.CommandButton cmdUnSelectFiles 
      Caption         =   "-"
      Height          =   285
      Left            =   6570
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Un select"
      Top             =   4650
      Width           =   405
   End
   Begin VB.CommandButton cmdSelectFunc 
      Caption         =   "+"
      Height          =   285
      Left            =   11310
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Select all"
      Top             =   4650
      Width           =   405
   End
   Begin VB.CommandButton cmdUnSelectFunc 
      Caption         =   "-"
      Height          =   285
      Left            =   10860
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Un select"
      Top             =   4650
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CTRL+B"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   9
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Selected files:"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   2
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "Selected functions:"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   7590
      TabIndex        =   7
      Top             =   60
      Width           =   1515
   End
End
Attribute VB_Name = "frmAttachCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_FilesCounter          As Long
Private m_bAvoidClick           As Boolean
Private m_AControlsPrefix()     As String


Private Sub cmdBrows_Click()

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
        
        '*Init functions array:
        ReDim g_AFunctions(3, 0)
        g_AFunctions(0, 0) = -1
        
        'Clear listboxes
        Me.lstFunctions.Clear
        Me.lstSelectedFiles.Clear
                
        'checks for file type
        If oDir.GetExtensionName(sFileName) <> "vbp" Then
            AddFileName sFileName 'other file
        Else
            AddProject sFileName 'vb project - add all relevant files
        End If
        
    End If
    
    '*Parse the selected files and make temporary new files on the fly
    ProcessFiles False 'dont use the previuse definition
    If Me.lstSelectedFiles.ListCount > 0 Then
        Me.lstSelectedFiles.ListIndex = 0 'focus on the first file
        lstSelectedFiles_Click 'force showing the first file's functions
    End If
    
End Sub

Private Sub AddFileName(ByVal sFileName As String)


    On Error GoTo Err_Proc

    '*Purpose:adds new file to the files list
    
    If LenB(sFileName) > 0 Then
        If Not FileInList(sFileName) Then
            lstSelectedFiles.AddItem sFileName
            lstSelectedFiles.Selected(lstSelectedFiles.NewIndex) = True
        End If
    End If

Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "AddFileName", Err, Err_Handle_Mode
    Resume Exit_Proc


End Sub

Private Function FileInList(ByVal sFileName As String) As Boolean

    'check if the specified file is allready in the list
    
    On Error GoTo Err_Proc
    
    Dim i           As Long
    
    FileInList = False
    
    For i = 0 To Me.lstSelectedFiles.ListCount - 1
        FileInList = (sFileName = Me.lstSelectedFiles.List(i))
        If FileInList Then Exit For
    Next i
    
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "FileInList", Err, Err_Handle_Mode
Resume Exit_Proc


End Function




Private Sub ProcessFiles(Optional ByVal UseDefinitions As Boolean = False)


    '*Purpose: parse all the selected files in the files list and generate
    '          err handling code for all the selected functions
    
    On Error GoTo Err_Proc
    
    Dim i           As Long
    
    If CheckValidation() Then
        
         'scan the files list
         For i = 0 To Me.lstSelectedFiles.ListCount - 1
             If Me.lstSelectedFiles.Selected(i) Then
                 If LenB(Me.lstSelectedFiles.List(i)) > 0 Then
                    'add err handling to the destination temp file
                     AddErrHandling Me.lstSelectedFiles.List(i), i, UseDefinitions
                 End If
             End If
             
        Next i
        'MsgBox "File definition completed successfully"
    Else
        MsgBox "Cannot commit because one of the parameters is wrong"
        
   End If
   
   
Exit_Proc:
Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "cmdAdd_Click", Err, Err_Handle_Mode
Resume Exit_Proc

End Sub


Private Sub cmdSelectFiles_Click()

    'select all files
    
    Dim i           As Long
    
    For i = 0 To Me.lstSelectedFiles.ListCount - 1
        Me.lstSelectedFiles.ListIndex = i
        Me.lstSelectedFiles.Selected(i) = True
    Next i

End Sub

Private Sub cmdSelectFunc_Click()

    'select all functions
    
    Dim i           As Long
    
    For i = 0 To Me.lstFunctions.ListCount - 1
        Me.lstFunctions.ListIndex = i
        Me.lstFunctions.Selected(i) = True
    Next i

End Sub



'* Function: AddErrHandling
'* Purpose: Add error handling to a certain file

Private Function AddErrHandling(ByVal sFilePath As String, _
                                ByVal FileNum As Long, _
                                Optional ByVal UseDefinition As Boolean = False) As Boolean

    'Purpose: Add error handling to the temporary file
    '         if UseDefinition = true,

    On Error GoTo Err_Proc
    
    Const PROCESS_REMARK = "'"
    
    Dim ff          As Long 'source file
    Dim s           As String
    Dim sline       As String
    Dim sDest       As String
    Dim sModuleName As String 'current module name
    Dim sProcName   As String 'current procedure name
    Dim ProcIndex   As Long 'function index in array
    Dim i           As Long
    
    Dim bStartSub     As Boolean 'recognize function
    Dim bStartFunc    As Boolean 'recognize function
    Dim bEndSub       As Boolean 'recognize end of sub or function
    Dim bFoundModuleName As Boolean 'flag-found thew module name
    
    Dim iTopIndex     As Long 'optimization flag
    Dim oDir          As Scripting.FileSystemObject
    Dim sDesc         As String 'temp variable to store function description
    Dim iDesc         As Long 'function description counter
    Dim iFunc         As Long 'function index
    Dim sFunc         As String
    
    
    'Init functions array
    iFunc = 0
    ReDim m_Functions(0)
    
    'Init interface description array
    ReDim g_InterfaceDesc(0)
    sDesc = ""
    iDesc = 1
    
    Set oDir = New Scripting.FileSystemObject
    
    'gets the array size-number of functions in the system
    iTopIndex = UBound(g_AFunctions, 2)
    If g_AFunctions(0, 0) <> -1 Then iTopIndex = iTopIndex + 1 'case its not the first time
    
    'init vars
    sModuleName = ""
    sProcName = ""
    
    'open source file
    ff = FreeFile
    Open sFilePath For Input As #ff
    
    'init algorithm flags
    s = ""
    bStartSub = False
    bEndSub = False
    bStartFunc = False
    bFoundModuleName = False
    
    'main scanning loop
    Do Until EOF(ff)
    
        'read the current line from the file
        Line Input #ff, sline
        
        'init dest line
        sDest = ""
        
        '*Check for the module name
        If Not bFoundModuleName Then
            sModuleName = GetModuleName(sline)
            bFoundModuleName = (LenB(sModuleName) <> 0)
        End If
        
        '* check if its a begining of a sub or function
        '* Check subs:
        If (Not bStartSub) Then
            If LCase(Left$(sline, 11)) = "public sub " Then
                sProcName = GetProcName(sline, 12)
                bStartSub = ((FunctionSelected(FileNum, sProcName, UseDefinition)))
            ElseIf LCase(Left$(sline, 4)) = "sub " Then
                sProcName = GetProcName(sline, 5)
                bStartSub = ((FunctionSelected(FileNum, sProcName, UseDefinition)))
            ElseIf LCase(Left$(sline, 12)) = "private sub " Then
                sProcName = GetProcName(sline, 13)
                bStartSub = ((FunctionSelected(FileNum, sProcName, UseDefinition)))
            End If
        Else
            If LCase(Left$(sline, 7)) = "end sub" Then
                bEndSub = True
            End If
        End If
        
        '* Check functions:
        If (Not bStartFunc) Then
            If LCase(Left$(sline, 16)) = "public function " Then
                sProcName = GetProcName(sline, 17)
                bStartFunc = ((FunctionSelected(FileNum, sProcName, UseDefinition)))
            ElseIf LCase(Left$(sline, 9)) = "function " Then
                sProcName = GetProcName(sline, 10)
                bStartFunc = ((FunctionSelected(FileNum, sProcName, UseDefinition)))
            ElseIf LCase(Left$(sline, 17)) = "private function " Then
                sProcName = GetProcName(sline, 18)
                bStartFunc = ((FunctionSelected(FileNum, sProcName, UseDefinition)))
            End If
        Else
            If LCase(Left$(sline, 12)) = "end function" Then
                bEndSub = True
            End If
        End If
        
        
        'START OF SOME FUNCTION
        If ((bStartSub) Or (bStartFunc)) Then
            '* Add on error goto...
            sDest = sDest & sline
            sFunc = sFunc & sline & vbNewLine 'add function line
            
        End If
        
        
        '*END OF SOME FUNCTION
        If (bEndSub) Then
                        
            '*Update functions array:
            If Not UseDefinition Then
                ReDim Preserve g_AFunctions(3, iTopIndex) 'allocates new memory unit
                g_AFunctions(C_MODULE_NAME, iTopIndex) = FileNum  'file num in files lst
                g_AFunctions(C_PROC_NAME, iTopIndex) = sProcName  'function name
                g_AFunctions(C_SELECTED, iTopIndex) = 1  'put err handling by default
                g_AFunctions(C_CODE, iTopIndex) = sFunc  'function's code
                
                iTopIndex = iTopIndex + 1
            End If
            
            
            '*Clear variables:
            bStartSub = False
            bEndSub = False
            bStartFunc = False
            sProcName = ""
            sDesc = ""
            sFunc = ""
            
        End If
        
    Loop
    
    'close file ports
    Close #ff
    
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "AddErrHandling", Err, Err_Handle_Mode
    'Resume
Resume Exit_Proc


End Function

Private Sub AddProject(ByVal sFileName As String)


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
                AddFileName sPath & sObjectName 'add file to list
            End If
            
        End If
        
        If InStr(1, LCase$(sline), "class=") > 0 Then
            i = InStr(1, sline, ";") + 2
            sObjectName = Mid$(sline, i, Len(sline) - i + 1)
            AddFileName sPath & sObjectName 'add file to list
        End If
        
        If InStr(1, LCase$(sline), "module=") > 0 Then
            i = InStr(1, sline, ";") + 2
            sObjectName = Mid$(sline, i, Len(sline) - i + 1)
            AddFileName sPath & sObjectName 'add file to list
        End If
        
        If InStr(1, LCase$(sline), "usercontrol=") > 0 Then
            i = InStr(1, sline, "=") + 1
            sObjectName = Mid$(sline, i, Len(sline) - i + 1)
            AddFileName sPath & sObjectName 'add file to list
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

Private Sub cmdUnSelectFiles_Click()

    'unselect all the files
    
    Dim i           As Long
    
    For i = 0 To Me.lstSelectedFiles.ListCount - 1
        Me.lstSelectedFiles.ListIndex = i
        Me.lstSelectedFiles.Selected(i) = False
    Next i
    
End Sub

Private Sub cmdUnSelectFunc_Click()

    'unselect all the functions
    
    Dim i           As Long
    
    For i = 0 To Me.lstFunctions.ListCount - 1
        Me.lstFunctions.ListIndex = i
        Me.lstFunctions.Selected(i) = False
    Next i

End Sub

Private Sub cmdView_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        Case vbKeyB And Shift = 2
            Me.cmdBrows.Value = True
            
    End Select
    
End Sub

Private Sub Form_Load()
    
    Me.Width = 12000
    Me.Height = 5650
    
End Sub

Private Sub lstFunctions_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lstFunctions.Drag 1
    
End Sub

Private Sub lstSelectedFiles_Click()

    'show all the functions in the module
    m_bAvoidClick = True
    ShowFunctions Me.lstSelectedFiles.ListIndex
    m_bAvoidClick = False
    
End Sub

Private Function CheckValidation() As Boolean

    '*Purpose: check that all the nesesary fields has data
    
    On Error GoTo Err_Proc


    Dim obj     As Control
    
    CheckValidation = False
    
    
    For Each obj In frmAttachCode
        If TypeOf obj Is TextBox Then
            If obj.Name <> "txtSource" And obj.Name <> "txtDest" Then
                If Trim(obj.Text) = "" Then
                    Exit Function
                End If
            End If
            
        End If
        
        
    Next obj
    
    CheckValidation = True
    
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "CheckValidation", Err, Err_Handle_Mode
Resume Exit_Proc


End Function

Private Function FunctionSelected(ByVal ModuleIndex As Long, ByVal sProcName As String, ByVal UseDefinition As Boolean, _
                                  Optional ByRef ProcIndex As Long) As Boolean


    On Error GoTo Err_Proc

    '*Purpose: checks if the function was selected and not ignored
    '*         function is ignored when user unmark its checkbox
    
    Dim i           As Long
    
    FunctionSelected = True
    ProcIndex = -1
    
    If Not UseDefinition Then Exit Function
    
    'scan the functions array
    For i = 0 To UBound(g_AFunctions, 2)
        If (g_AFunctions(1, i) = sProcName) And (g_AFunctions(0, i) = ModuleIndex) Then
            FunctionSelected = (g_AFunctions(C_SELECTED, i) = 1)
            ProcIndex = i
            Exit For
        End If
    Next i
    
    
Exit_Proc:
    Exit Function


Err_Proc:
    Err_Handler " frmMain ", "FunctionSelected", Err, Err_Handle_Mode
    Resume Exit_Proc


End Function


Private Sub ShowFunctions(ByVal FileIndex As Long)


    On Error GoTo Err_Proc

    '*Purpose: show all the functions in the selected module
    
    Dim i               As Long
    Dim s               As String
    Dim sFuncName       As String
    Dim bFirstElement   As Boolean
    Dim bNoMore         As Boolean
    Dim iTopIndx        As Long
    
    bFirstElement = False
    bNoMore = False
    Me.lstFunctions.Clear
    
    i = 0
    iTopIndx = UBound(g_AFunctions, 2)
    
    'scan the functions array
    Do
        If g_AFunctions(0, i) = FileIndex Then
            If (Not bFirstElement) Then bFirstElement = (Not bFirstElement)
            sFuncName = g_AFunctions(1, i)
            Me.lstFunctions.AddItem sFuncName
            Me.lstFunctions.ItemData(Me.lstFunctions.NewIndex) = i
            Me.lstFunctions.Selected(Me.lstFunctions.NewIndex) = (g_AFunctions(2, i) = 1)
        Else
            If bFirstElement Then 'no more relevant functions
                bNoMore = True
            End If
            
        End If
        
        i = i + 1
        bNoMore = (i > iTopIndx)
        
    Loop Until bNoMore 'no more relevant functions
    
    
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "ShowFunctions", Err, Err_Handle_Mode
    Resume Exit_Proc


End Sub

