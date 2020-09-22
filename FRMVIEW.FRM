VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmView 
   Caption         =   "View File"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox txtSource 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   210
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   13573
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmView.frx":0000
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
   Begin RichTextLib.RichTextBox txtDest 
      Height          =   7695
      Left            =   5610
      TabIndex        =   1
      Top             =   210
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   13573
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmView.frx":00B5
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
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Press CTRL+P to print the file"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   30
      TabIndex        =   4
      Top             =   7950
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Source file:"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   945
   End
   Begin VB.Label Label1 
      Caption         =   "Dest file:"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   2
      Top             =   0
      Width           =   945
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_objEditor As CEditor 'main editor object
Private m_objEditor2 As CEditor 'main editor object


Public Sub ShowEX(ByVal FilePath As String, Optional ShowInterface As Boolean = False)


    On Error GoTo Err_Proc

    Me.txtSource.Text = ""
    Me.txtDest.Text = ""

    UpdateTextView FilePath, ShowInterface
    If ShowInterface Then
        With Me.txtDest
            .Width = .Left + .Width
            .Left = Me.txtSource.Left
            .ZOrder 0
            Label1(0).Visible = False
            Label1(1).Left = Label1(0).Left
        End With
    End If
    
    Me.Show
    m_objEditor.PaintText
    m_objEditor2.PaintText
    
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler " frmView ", "ShowEX", Err, Err_Handle_Mode
    Resume Exit_Proc


End Sub

Private Sub UpdateTextView(ByVal sFilePath As String, Optional ByVal ShowInterface As Boolean = False)


    On Error GoTo Err_Proc

    Dim ff          As Long
    Dim s           As String
    Dim sExt        As String
    Dim sView       As String
    
    
    sExt = IIf(ShowInterface, ".desc", "")
    
    ff = FreeFile
    Open sFilePath For Input As #ff
    sView = ""
    Do Until EOF(ff)
        Line Input #ff, s
        sView = sView & s & vbNewLine
    Loop
    Me.txtSource.Text = sView
    Close #ff
    
    
    '* Show dest file:
    ff = FreeFile
    sFilePath = GetDestFileName(sFilePath)
    Open App.path & "\DestTmp\" & sFilePath & sExt For Input As #ff
    
    sView = ""
    Do Until EOF(ff)
        Line Input #ff, s
        sView = sView & s & vbNewLine
    Loop
    Me.txtDest.Text = sView
    Close #ff
    

Exit_Proc:
Exit Sub


Err_Proc:
    Err_Handler " frmMain ", "UpdateTextView", Err, Err_Handle_Mode
Resume Exit_Proc


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)


    On Error GoTo Err_Proc

    Select Case KeyCode
        Case vbKeyP And Shift = 2
            Printer.Print Me.txtDest.Text
            Printer.EndDoc
    End Select
    
Exit_Proc:
    Exit Sub


Err_Proc:
    Err_Handler " frmView ", "Form_KeyDown", Err, Err_Handle_Mode
    Resume Exit_Proc


End Sub

Private Sub Form_Load()

    Set m_objEditor = New CEditor
    Set m_objEditor2 = New CEditor
    
    'set editor objects
    m_objEditor.SetEditorObjects Me.txtSource, Nothing, Nothing, Nothing, Nothing, Me.txtDest
    InitWords m_objEditor
    
    m_objEditor2.SetEditorObjects Me.txtDest, Nothing, Nothing, Nothing, Nothing, Me.txtSource
    InitWords m_objEditor2

End Sub

Private Sub InitWords(ByRef obj As CEditor)

    'hard code init the basic vb script words -
    'you can init any words you want with any colors you like
    obj.AddEditorWord "Dim", vbRed
    obj.AddEditorWord "Select ", vbBlue
    obj.AddEditorWord "Until", vbBlue
    obj.AddEditorWord "Set", vbBlue
    obj.AddEditorWord "Private Function", vbBlue
    obj.AddEditorWord "Function", vbBlue
    obj.AddEditorWord "Public", vbBlue
    obj.AddEditorWord "Sub", vbBlue
    obj.AddEditorWord "Private", vbBlue
    obj.AddEditorWord "Case", vbBlue
    obj.AddEditorWord "End", vbBlue
    obj.AddEditorWord "For", vbBlue
    obj.AddEditorWord "Next", vbBlue
    obj.AddEditorWord "Do", vbBlue
    obj.AddEditorWord "Loop", vbBlue
    obj.AddEditorWord "If", vbBlue
    obj.AddEditorWord "Select", vbBlue
    obj.AddEditorWord "Then", vbBlue
    obj.AddEditorWord "Else", vbBlue

    obj.AddEditorWord "ElseIf", vbBlue
    obj.AddEditorWord "Open", vbBlue
    obj.AddEditorWord "Exit", vbBlue
    obj.AddEditorWord "On", vbBlue
    obj.AddEditorWord "Resume", vbBlue
    obj.AddEditorWord "New", vbBlue
    obj.AddEditorWord "Close", vbBlue
    obj.AddEditorWord "Print", vbBlue
    obj.AddEditorWord "Redim", vbBlue
    obj.AddEditorWord "Preserve", vbBlue
    obj.AddEditorWord "Error", vbBlue
    obj.AddEditorWord "Goto", vbBlue
    
End Sub
