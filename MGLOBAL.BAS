Attribute VB_Name = "MGlobal"
Option Explicit

Public Const C_MODULE_NAME = 0
Public Const C_PROC_NAME = 1
Public Const C_SELECTED = 2
Public Const C_CODE = 3

Public cn                 As ADODB.Connection

Public Err_Handle_Mode    As Boolean
Public g_InterfaceDesc()  As String
Public g_AFunctions()     As Variant

Public Function InitConnection() As Boolean

    '*Purpose: init connection to database
    
    Dim s           As String
    
    On Error GoTo Err_Proc
    
    InitConnection = False
    s = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\Code\Code.abd" & ";Persist Security Info=False"
    Set cn = New ADODB.Connection
    cn.ConnectionString = s
    cn.Open
    InitConnection = True
    Exit Function
    
Err_Proc:
    'connection faild
    
End Function

Public Sub CloseConnection()

    '*Purpose: close connection to database
    
    If Not (cn Is Nothing) Then
        If cn.State <> adStateClosed Then cn.Close
        Set cn = Nothing
    End If
    
End Sub



Public Sub Err_Handler(ByVal Module As String, ByVal Proc As String, Err As ErrObject, Err_Handle_Mode As Boolean)


    On Error GoTo Err_Proc

    'Centeral error handling procedure
    
    If Err_Handle_Mode Then
        MsgBox "çìä ùâéàä áîåãåì:" & Module & "    áôåð÷öéä:" & Proc & vbNewLine & _
        "úàåø äùâéàä: " & Err.Description
        
    End If
    
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler "MGlobal", "Err_Handler", Err, Err_Handle_Mode
    Resume Exit_Proc


End Sub

Public Function GetModuleName(ByVal sline As String) As String

    'Purpose: parse the module name from the initializing line
    
    On Error GoTo Err_Proc


    Dim iStart          As Long
    Dim iEnd            As Long
    Dim sEndChar        As String
    
    GetModuleName = ""
    
    
    If InStr(1, sline, "Attribute VB_Name = ") <> 0 Then
        iStart = InStr(1, sline, "Attribute VB_Name") + 21
        sEndChar = Chr$(34)
    ElseIf InStr(1, sline, "Begin VB.Form") <> 0 Then
        iStart = InStr(1, sline, "Begin VB.Form") + 13
        sEndChar = " "
    End If
    
    
    If iStart > 0 Then
        iEnd = InStr(iStart, sline, sEndChar)
        GetModuleName = Mid$(sline, iStart, iEnd)
    End If
    
    If Right$(GetModuleName, 1) = Chr$(34) Then
        GetModuleName = Left$(GetModuleName, Len(GetModuleName) - 1)
    End If
    
    
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "GetModuleName", Err, Err_Handle_Mode
Resume Exit_Proc


End Function

Public Function GetProcName(ByVal sline As String, ByVal StartPoint As Long) As String

    'Purpose: parse the procedure name from the initializing line
    
    On Error GoTo Err_Proc

    Dim iStartBr    As Long
    
    StartPoint = StartPoint - 1
    
    Do Until Len(GetProcName) > 0
        iStartBr = InStrRev(sline, " ", StartPoint, vbTextCompare)
        
        If StartPoint - iStartBr > 0 Then
            GetProcName = Trim$(Mid$(sline, iStartBr + 1, (StartPoint - iStartBr)))
        End If
        
        If LenB(GetProcName) = 0 Then
            StartPoint = iStartBr - 1
        End If
    Loop
    
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "GetProcName", Err, Err_Handle_Mode
Resume Exit_Proc


End Function


Public Function MBX(ByVal msg As String, _
                    Optional MsgStyle As VbMsgBoxStyle = vbOKOnly Or vbInformation Or vbMsgBoxRight Or vbMsgBoxRtlReading) As VbMsgBoxResult

    MBX = MsgBox(msg, MsgStyle, "Code manager")
    
End Function

Public Function FormExists(ByVal sFormName As String) As Boolean

    FormExists = (FormCount(sFormName) > 0)

End Function

Private Function FormCount(ByVal frmName As String) As Long

' Return the number of instances of a form
' that are currently loaded

    Dim frm As Form
    For Each frm In Forms
        If StrComp(frm.Name, frmName, vbTextCompare) = 0 Then
            FormCount = FormCount + 1
        End If
    Next frm
    
End Function


Public Function GetDestFileName(ByVal sFilePath As String) As String

    'Purpose: return the temporary file name wich is going to be worked on
    
    On Error GoTo Err_Proc

    Dim oDir        As Scripting.FileSystemObject
    Dim sFileName   As String
    
    Set oDir = New Scripting.FileSystemObject
    
    sFileName = oDir.GetFilename(sFilePath)
    
    GetDestFileName = sFileName & ".tmp"
    
Exit_Proc:
Exit Function


Err_Proc:
    Err_Handler " frmMain ", "GetDestFileName", Err, Err_Handle_Mode
Resume Exit_Proc


End Function



