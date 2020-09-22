VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "VBTools"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuErrHandler 
         Caption         =   "Error handling tool..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MDIForm_Activate()
    frmErrHandler.Show
End Sub

Private Sub mnuErrHandler_Click()
    frmErrHandler.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
    
End Sub

