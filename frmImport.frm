VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restore"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImport 
      Caption         =   "Restore database"
      Height          =   555
      Left            =   1050
      TabIndex        =   0
      Top             =   1110
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   2940
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   3060
      Left            =   510
      Picture         =   "frmImport.frx":0000
      Stretch         =   -1  'True
      Top             =   90
      Width           =   3180
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_ZIP           As CZip


Private Sub cmdImport_Click()
    
    
    Dim sFile           As String
    
    On Error GoTo Err_Proc
    
    dlg1.CancelError = True
    dlg1.ShowOpen
    sFile = dlg1.Filename
    
    If Dir$(sFile) = "" Then
        Exit Sub
    End If
    Me.cmdImport.Enabled = False
    
    'First close the database connection
    CloseConnection
    Unload frmCodeStorage 'if its loaded - unload it
        
    m_ZIP.ExtractZipFile sFile, App.path & "\Code"   'extract
    
    'Re open the database connection
    InitConnection
    
Err_Proc:
    
End Sub

Private Sub Form_Load()

    Set m_ZIP = New CZip
    
End Sub

Private Sub m_ZIP_OnUnzipComplete(ByVal Successful As Boolean)

    If Successful Then
        MBX "Import data completed successfully"
        Me.cmdImport.Enabled = True
    End If

End Sub

