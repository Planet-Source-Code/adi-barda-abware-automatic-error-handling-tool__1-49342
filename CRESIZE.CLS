VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Resize class by adi barda

Option Explicit

'windows constants
Private Const SWP_DRAWFRAME = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or _
                         SWP_NOMOVE Or SWP_DRAWFRAME
                         
                         
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000

Private Declare Function GetWindowLong Lib _
    "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib _
    "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
     ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
     ByVal cy As Long, ByVal wFlags As Long) As Long


Private Const INIT_FRAME_STYLE = 1409351872 'by default no resizing

Private m_HWND                  As Long 'containers hwnd
Private m_Ctrl                  As Object


Public Property Let CHWND(ByVal newval As Long)
    m_HWND = newval
End Property

Private Sub SetControlStyle(ByVal style As Long, ctl As Control)
   
    'Activate / Deactivate the control's resizing
    If style Then
      Call SetWindowLong(ctl.hwnd, GWL_STYLE, style)
      Call SetWindowPos(ctl.hwnd, m_HWND, 0, 0, 0, 0, SWP_FLAGS)
    End If

End Sub


Public Sub MakeControlResize(ctl As Control)

    Dim iStyle As Long
    
    'restore control's frame
    If Not (m_Ctrl Is Nothing) Then SetControlStyle INIT_FRAME_STYLE, m_Ctrl
    
    Set m_Ctrl = ctl 'set new control to be resized
       
    'get the current style attributes for the textbox
    iStyle = GetWindowLong(m_Ctrl.hwnd, GWL_STYLE)
 
    'modify the style to show the sizing frame
    iStyle = iStyle Or WS_THICKFRAME

    SetControlStyle iStyle, m_Ctrl
    
End Sub

Public Sub CancelResize()

    'restore control's frame
    If Not (m_Ctrl Is Nothing) Then SetControlStyle INIT_FRAME_STYLE, m_Ctrl

End Sub


