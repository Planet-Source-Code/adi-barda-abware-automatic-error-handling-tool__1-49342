VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCodeStorage 
   Caption         =   "Code Storage"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   555
      Left            =   10500
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6780
      Width           =   1335
   End
   Begin VB.PictureBox pic2 
      Height          =   6225
      Left            =   4530
      ScaleHeight     =   6165
      ScaleWidth      =   7245
      TabIndex        =   11
      Top             =   450
      Width           =   7305
      Begin VB.Frame Frame1 
         Caption         =   "Properties:"
         ForeColor       =   &H00FF0000&
         Height          =   1815
         Left            =   30
         TabIndex        =   12
         Top             =   4320
         Width           =   7155
         Begin VB.TextBox txtDesc 
            Height          =   825
            Left            =   150
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   900
            Width           =   6645
         End
         Begin VB.TextBox txtCreationDate 
            BackColor       =   &H00C0C0C0&
            Height          =   345
            Left            =   1230
            TabIndex        =   13
            Top             =   300
            Width           =   1995
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Date Created:"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   300
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Description:"
            ForeColor       =   &H00404040&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   630
            Width           =   840
         End
      End
      Begin RichTextLib.RichTextBox txtCode 
         Height          =   4275
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   7541
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmCodeStorage.frx":0000
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
   End
   Begin VB.PictureBox pic1 
      BackColor       =   &H00404040&
      Height          =   6225
      Left            =   60
      ScaleHeight     =   6165
      ScaleWidth      =   4365
      TabIndex        =   9
      Top             =   450
      Width           =   4425
      Begin VBTools.abTreeView tvCategories 
         Height          =   6105
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   4305
         _ExtentX        =   7594
         _ExtentY        =   10769
         ID_Field        =   ""
         Father_Field    =   ""
         Name_Field      =   ""
         Table_Name      =   ""
         DataSourceType  =   0
      End
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy Code"
      Height          =   555
      Left            =   4290
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6780
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveItem 
      Caption         =   "Save Item"
      Height          =   555
      Left            =   2910
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6780
      Width           =   1335
   End
   Begin VB.CommandButton cmdAttachCode 
      Caption         =   "Attach Code"
      Height          =   555
      Left            =   1500
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6780
      Width           =   1335
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "Add Item"
      Height          =   555
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6780
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " ESC"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   10950
      TabIndex        =   20
      Top             =   7350
      Width           =   360
   End
   Begin VB.Label lblCode 
      Caption         =   "Code:"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   4530
      TabIndex        =   18
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CTRL+P"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   7
      Left            =   4620
      TabIndex        =   8
      Top             =   7350
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CTRL+S"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   6
      Left            =   3240
      TabIndex        =   6
      Top             =   7350
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CTRL+T"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   5
      Left            =   1830
      TabIndex        =   4
      Top             =   7350
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CTRL+I"
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   4
      Left            =   420
      TabIndex        =   3
      Top             =   7350
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Storage:"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   210
      Width           =   735
   End
End
Attribute VB_Name = "frmCodeStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAX_HEIGHT = 6225 'treeview max height

Private m_LastContainerSize     As Long

Private m_Resize                As CResize
Private m_Dist                  As Long
Private m_DescDist              As Long


Private Sub cmdAddItem_Click()

    Dim s           As String
    
    s = InputBox("ä÷ìã ùí ôøéè ùáøöåðê ìäåñéó", "ôøéè çãù", "ôøéè çãù")
    If LenB(Trim(s)) > 0 Then
        tvCategories.Add_Branch s
    End If
    
End Sub

Private Sub cmdAttachCode_Click()
    frmAttachCode.Show
    
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText Me.txtCode.Text
    
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSaveItem_Click()
    SaveItem
    
End Sub

Private Sub Form_Activate()

    With Me.tvCategories
        .SetFocus
        
    End With
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        ' Attach code to the current item
        Case vbKeyT And Shift = 2
            Me.cmdAttachCode.Value = Me.cmdAttachCode.Enabled
        ' Add new item to the current node
        Case vbKeyI And Shift = 2
            Me.cmdAddItem.Value = True
        ' Save current item
        Case vbKeyS And Shift = 2
            Me.cmdSaveItem.Value = True
        ' Copy the current item
        Case vbKeyP And Shift = 2
            Me.cmdCopy.Value = True
        ' Get out
        Case vbKeyEscape
            Unload Me
            
    End Select
    
    
End Sub

Private Sub Form_Load()

    m_LastContainerSize = Me.pic2.Width
    Me.Width = Screen.Width
    Set m_Resize = New CResize 'resize component
    
    
    m_Resize.CHWND = Me.hwnd
    m_Resize.MakeControlResize Me.pic1
    
    With tvCategories
        .Use_Connection cn
        .Table_Name = "tblStorageDef"
        .ID_Field = "ID"
        .Father_Field = "Father"
        .Name_Field = "Name"
        .Build_Tree
    End With
    
    pic1_Resize
    m_Dist = pic2.Left - (pic1.Left + pic1.Width)
    m_DescDist = pic2.Width - (Me.txtDesc.Left + Me.txtDesc.Width)
        
End Sub

Private Sub pic1_Resize()

    Dim r           As Single
    
    
    If pic1.Height > MAX_HEIGHT Then
        pic1.Height = MAX_HEIGHT
    End If
    
    With Me.tvCategories
        .Width = pic1.Width - Screen.TwipsPerPixelX * 6
        .Height = pic1.Height - Screen.TwipsPerPixelY * 6
        .Top = 0
        .Left = 0
    End With
    
    'place pic2 container
    pic2.Left = pic1.Left + pic1.Width + m_Dist
    Me.lblCode.Left = Me.pic2.Left
    Me.pic2.Width = Me.Width - Me.pic2.Left - 5 * Screen.TwipsPerPixelX
    r = Me.pic2.Width / m_LastContainerSize
    m_LastContainerSize = Me.pic2.Width
    'Apply inner controls size
    Me.txtCode.Width = Me.txtCode.Width * r
    Me.txtDesc.Width = Me.txtDesc.Width * r
    Me.Frame1.Width = Me.Frame1.Width * r
        
End Sub

Private Sub tvCategories_NodeSelected(ByVal oNode As MSComctlLib.Node, ByVal ItemCode As String)
    
    Me.cmdAttachCode.Enabled = (ItemCode <> "OOT")
    LoadItem
    
End Sub

Private Sub txtCode_DragDrop(source As Control, X As Single, Y As Single)

    '*Purpose: drop function's code into text box
    
    If LenB(Trim(Me.txtCode.Text)) > 0 Then
        If MsgBox("äàí àúä áèåç ùáøöåðê ìäçìéó àú ä÷åã äðåëçé á÷åã äçãù ?", vbYesNo Or vbMsgBoxRight Or vbMsgBoxRtlReading Or vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    
    Me.txtCode.Text = g_AFunctions(C_CODE, source.ItemData(source.ListIndex))
    SaveItem
    Me.txtCode.BackColor = vbWhite
    
End Sub

Private Sub txtCode_DragOver(source As Control, X As Single, Y As Single, State As Integer)
    Me.txtCode.BackColor = vbGreen
    
End Sub

Private Sub SaveItem()

    '*Purpose: save the current item to the code folder
    
    Dim sFile       As String
    Dim s           As String
    Dim ff          As Long
    
    ff = FreeFile
    
    sFile = "Proc" & Me.tvCategories.CurrentKey & ".abd"
    Open App.path & "\Code\" & sFile For Output As #ff
    Print #ff, Me.txtCode.Text
    Close #ff
        
        
    'Save description
    s = "UPDATE tblStorageDef SET Comments='" & Me.txtDesc.Text & " '" & _
        ", UpdateDate=#" & Format(Date, "yyyy/mm/dd") & "#" & _
        " WHERE ID=" & Me.tvCategories.CurrentKey
        
    cn.Execute s
    
End Sub

Private Sub LoadItem()

    '*Purpose: load current item from code folder
    
    Dim sFile       As String
    Dim s           As String
    Dim ff          As Long
    Dim oDir        As Scripting.FileSystemObject
    Dim rs          As ADODB.Recordset
    
    
    ff = FreeFile
    Set oDir = New Scripting.FileSystemObject
    
    sFile = "Proc" & Me.tvCategories.CurrentKey & ".abd"
    Me.txtCode.Text = ""
    
    If Me.tvCategories.CurrentKey = "OOT" Then Exit Sub
    
    If oDir.FileExists(App.path & "\Code\" & sFile) Then
        Open App.path & "\Code\" & sFile For Input As #ff
        sFile = ""
        Do Until EOF(ff)
            Line Input #ff, s
            sFile = sFile & s & vbNewLine
        Loop
        Close #ff
        Me.txtCode.Text = sFile
    End If
    
    Set rs = New ADODB.Recordset
    s = "SELECT UpdateDate,Comments FROM tblStorageDef WHERE ID=" & Me.tvCategories.CurrentKey
    rs.Open s, cn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        Me.txtDesc.Text = rs.Fields("Comments").Value & ""
    End If
    rs.Close
    Set rs = Nothing
        
End Sub
