VERSION 5.00
Begin VB.UserControl abTextBox 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1110
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   177
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   375
   ScaleWidth      =   1110
   Begin VB.TextBox txtMain 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   1065
   End
End
Attribute VB_Name = "abTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Textbox fillters struct
Public Enum FillterType
    FillterOff = 0
    LettersOnly = 1
    IntegerNumber = 2
    FloatNumber = 3
End Enum

'Possible Numbers struct
Public Enum NumberType
    Free = 0
    Positive = 1
    Negative = 2
End Enum

''Set focus options
Public Enum MoveActions
    NoAction = 0
    MoveForward = 1
End Enum


'Default Property Values:
Const m_def_EnterKey = MoveActions.MoveForward
Const m_def_BackColor = vbWhite
Const m_def_ForeColor = 0
Const m_def_Enabled = True
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 1
Const m_def_Alignment = 1
Const m_def_Locked = 0
Const m_def_MaxLength = 0
Const m_def_MultiLine = 0
Const m_def_PasswordChar = ""
Const m_def_RightToLeft = True
Const m_def_ScrollBars = 0
Const m_def_SelLength = 0
Const m_def_SelStart = 0
Const m_def_SelText = ""
Const m_def_Text = ""
Const m_def_ToolTipText = ""

'Property Variables:
Private m_KeyEnter                  As MoveActions
Dim m_BackColor                     As OLE_COLOR
Dim m_ForeColor                     As OLE_COLOR
Dim m_Enabled                       As Boolean
Dim m_Font                          As Font
Dim m_BackStyle                     As Integer
Dim m_BorderStyle                   As Integer
Dim m_Alignment                     As AlignmentConstants
Dim m_Locked                        As Boolean
Dim m_MaxLength                     As Long
Dim m_MultiLine                     As Boolean
Dim m_PasswordChar                  As String
Dim m_RightToLeft                   As Boolean
Dim m_ScrollBars                    As Integer
Dim m_SelLength                     As Long
Dim m_SelStart                      As Long
Dim m_SelText                       As String
Dim m_Text                          As String
Dim m_ToolTipText                   As String
Dim m_FillterType                   As FillterType
Dim m_AutoAlign                     As Boolean
Dim m_FormatString                  As String
Dim m_ExtraTag                      As String

'Event Declarations:
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Change()
Event Validate(Cancel As Boolean)
Event WriteProperties(PropBag As PropertyBag)

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns / Sets the controls back color"
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    txtMain.BackColor = m_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns / Sets  the controls fore color"
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    txtMain.ForeColor = m_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns / Sets whether the control response to general events or methods"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    txtMain.Enabled = m_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns / Sets the controls default font"
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    
    With txtMain
        .Font = m_Font
        .FontBold = m_Font.Bold
        .FontSize = m_Font.Size
        
    End With
    
    PropertyChanged "Font"
End Property

Public Property Get BackStyle() As Integer
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns / Sets the controls border style - 2d or 3d"
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    txtMain.BorderStyle = m_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Sub Refresh()
     
End Sub

Public Property Get Alignment() As AlignmentConstants
Attribute Alignment.VB_Description = "Returns/Sets the abTextbox alignment"
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentConstants)
    m_Alignment = New_Alignment
    txtMain.Alignment = m_Alignment
    PropertyChanged "Alignment"
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Returns / Sets whether pressing on a key will change the controls content"
    Locked = m_Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    m_Locked = New_Locked
    txtMain.Locked = m_Locked
    PropertyChanged "Locked"
End Property

Public Property Get MaxLength() As Long
    MaxLength = m_MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    m_MaxLength = New_MaxLength
    txtMain.MaxLength = m_MaxLength
    PropertyChanged "MaxLength"
End Property

Public Property Get PasswordChar() As String
    PasswordChar = m_PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    m_PasswordChar = New_PasswordChar
    txtMain.PasswordChar = m_PasswordChar
    PropertyChanged "PasswordChar"
End Property

Public Property Get RightToLeft() As Boolean
    RightToLeft = m_RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
    m_RightToLeft = New_RightToLeft
    txtMain.RightToLeft = m_RightToLeft
    PropertyChanged "RightToLeft"
End Property

Public Property Get SelLength() As Long
    SelLength = m_SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    m_SelLength = New_SelLength
    txtMain.SelLength = m_SelLength
    PropertyChanged "SelLength"
End Property

Public Property Get SelStart() As Long
    SelStart = m_SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    m_SelStart = New_SelStart
    txtMain.SelStart = m_SelStart
    PropertyChanged "SelStart"
End Property

Public Property Get SelText() As String
    SelText = m_SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    m_SelText = New_SelText
    txtMain.SelText = m_SelText
    PropertyChanged "SelText"
End Property

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    txtMain.Text = m_Text
    PropertyChanged "Text"
End Property

Public Property Get ToolTipText() As String
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    txtMain.ToolTipText = m_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Private Sub txtMain_Change()

    Me.Text = txtMain.Text
    
    RaiseEvent Change

End Sub

Private Sub txtMain_GotFocus()

    If AutoAlignment = True Then
        If TypeFillter = IntegerNumber Or TypeFillter = FloatNumber Then
            txtMain.Alignment = AlignmentConstants.vbLeftJustify
        End If
    End If
    txtMain.SelStart = 0
    txtMain.SelLength = Len(txtMain.Text)
    
End Sub

Private Sub txtMain_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            If Me.KeyEnter = MoveForward Then
                SendKeys ("{TAB}")
                
            End If
            
    End Select
    
End Sub

Private Sub txtMain_KeyPress(KeyAscii As Integer)

    Dim s               As String
    Dim i               As Long
    
    i = KeyAscii
    s = Chr$(KeyAscii)
    
    If TypeFillter = FillterOff Then Exit Sub
    If TypeFillter = FloatNumber Then
        i = IIf(s Like "[0-9]" Or s = "." Or s = "-" Or KeyAscii = vbKeyBack, KeyAscii, 0)
    End If
    If TypeFillter = IntegerNumber Then
        i = IIf(s Like "[0-9]" Or s = "-" Or KeyAscii = vbKeyBack, KeyAscii, 0)
    End If
    If TypeFillter = LettersOnly Then
        i = IIf(s Like "[a-z]" Or s Like "[A-Z]" Or s Like "[à-ú]" Or KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace, KeyAscii, 0)
    End If
    KeyAscii = i
    
End Sub

Private Sub txtMain_LostFocus()
    txtMain.Alignment = 1
    If Trim$(FormatString) <> "" Then
        txtMain.Text = Format(txtMain.Text, FormatString)
    End If
    
End Sub

Private Sub UserControl_GotFocus()
    txtMain.SetFocus
    
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_KeyEnter = m_def_EnterKey
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_Alignment = m_def_Alignment
    m_Locked = m_def_Locked
    m_MaxLength = m_def_MaxLength
    m_MultiLine = m_def_MultiLine
    m_PasswordChar = m_def_PasswordChar
    m_RightToLeft = m_def_RightToLeft
    m_ScrollBars = m_def_ScrollBars
    m_SelLength = m_def_SelLength
    m_SelStart = m_def_SelStart
    m_SelText = m_def_SelText
    m_Text = m_def_Text
    m_ToolTipText = m_def_ToolTipText
    m_ExtraTag = ""
    AutoAlignment = False
    FormatString = ""
    
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
    Locked = PropBag.ReadProperty("Locked", m_def_Locked)
    MaxLength = PropBag.ReadProperty("MaxLength", m_def_MaxLength)
    PasswordChar = PropBag.ReadProperty("PasswordChar", m_def_PasswordChar)
    RightToLeft = PropBag.ReadProperty("RightToLeft", m_def_RightToLeft)
   ' ScrollBars = PropBag.ReadProperty("ScrollBars", m_def_ScrollBars)
    SelLength = PropBag.ReadProperty("SelLength", m_def_SelLength)
    SelStart = PropBag.ReadProperty("SelStart", m_def_SelStart)
    SelText = PropBag.ReadProperty("SelText", m_def_SelText)
    Text = PropBag.ReadProperty("Text", m_def_Text)
    ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
    TypeFillter = PropBag.ReadProperty("TypeFillter", 0)
    AutoAlignment = PropBag.ReadProperty("AutoAlignment", False)
    FormatString = PropBag.ReadProperty("FormatString", "")
    Me.KeyEnter = PropBag.ReadProperty("KeyEnter", MoveActions.MoveForward)
    m_ExtraTag = PropBag.ReadProperty("ExtraTag", "")
    
End Sub

Private Sub UserControl_Resize()
txtMain.Left = 0
txtMain.Top = 0
txtMain.Width = Width
txtMain.Height = Height

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
    Call PropBag.WriteProperty("Locked", m_Locked, m_def_Locked)
    Call PropBag.WriteProperty("MaxLength", m_MaxLength, m_def_MaxLength)
    Call PropBag.WriteProperty("MultiLine", m_MultiLine, m_def_MultiLine)
    Call PropBag.WriteProperty("PasswordChar", m_PasswordChar, m_def_PasswordChar)
    Call PropBag.WriteProperty("RightToLeft", m_RightToLeft, m_def_RightToLeft)
    Call PropBag.WriteProperty("ScrollBars", m_ScrollBars, m_def_ScrollBars)
    Call PropBag.WriteProperty("SelLength", m_SelLength, m_def_SelLength)
    Call PropBag.WriteProperty("SelStart", m_SelStart, m_def_SelStart)
    Call PropBag.WriteProperty("SelText", m_SelText, m_def_SelText)
    Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
    Call PropBag.WriteProperty("TypeFillter", m_FillterType, 0)
    Call PropBag.WriteProperty("AutoAlignment", AutoAlignment, False)
    Call PropBag.WriteProperty("FormatString", FormatString, "")
    Call PropBag.WriteProperty("KeyEnter", m_KeyEnter, MoveActions.MoveForward)
    Call PropBag.WriteProperty("ExtraTag", m_ExtraTag, "")
    
End Sub

Public Property Get TypeFillter() As FillterType
Attribute TypeFillter.VB_Description = "Returns / Sets the specified keyboard fillter - usefull for rule validation"
    TypeFillter = m_FillterType

End Property

Public Property Let TypeFillter(ByVal vNewValue As FillterType)
    m_FillterType = vNewValue
    PropertyChanged "TypeFillter"

End Property

Public Property Get AutoAlignment() As Boolean
Attribute AutoAlignment.VB_Description = "Returns / Sets whether alignment to the left will automaticaly set on got focus - only in numbers fillter"
    AutoAlignment = m_AutoAlign
    
End Property

Public Property Let AutoAlignment(ByVal vNewValue As Boolean)
    m_AutoAlign = vNewValue
    PropertyChanged "AutoAlignment"
    
End Property


Public Property Get FormatString() As String
Attribute FormatString.VB_Description = "Returns / Sets the text format wich will be set after lost focus event"
    FormatString = m_FormatString
    
End Property

Public Property Let FormatString(ByVal vNewValue As String)
    m_FormatString = vNewValue
    PropertyChanged "FormatString"
    
End Property


Public Property Get KeyEnter() As MoveActions
Attribute KeyEnter.VB_Description = "Returns / Sets whether hitting the enter key will move the focus to the next indexed control"

    KeyEnter = m_KeyEnter
    
End Property

Public Property Let KeyEnter(ByVal vNewValue As MoveActions)
    
    m_KeyEnter = vNewValue
    PropertyChanged "KeyEnter"
    
End Property

Public Property Get ExtraTag() As String
    ExtraTag = m_ExtraTag
End Property

Public Property Let ExtraTag(ByVal vNewValue As String)
    m_ExtraTag = vNewValue
End Property
