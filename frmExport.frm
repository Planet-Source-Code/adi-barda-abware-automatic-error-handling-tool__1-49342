VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup database"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   3600
      Top             =   1380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Backup database"
      Height          =   555
      Left            =   1050
      TabIndex        =   0
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   3060
      Left            =   510
      Picture         =   "frmExport.frx":0000
      Stretch         =   -1  'True
      Top             =   60
      Width           =   3180
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExport_Click()

    Dim sPath           As String
    Dim objZip          As CZip
    
    On Error GoTo Err_Proc
    
    CloseConnection 'close database
    dlg1.CancelError = True
    dlg1.Flags = cdlOFNPathMustExist
    dlg1.Filename = "backup"
    dlg1.ShowSave
    sPath = dlg1.Filename
    
    Set objZip = New CZip
    With objZip
        .MakeZipFile App.path & "\Code", sPath & ".abd"
    End With
    MBX "Backup data completed"
        
    InitConnection 'reopen database
    Exit Sub
    
Err_Proc:
    InitConnection 'reopen database
      
End Sub

