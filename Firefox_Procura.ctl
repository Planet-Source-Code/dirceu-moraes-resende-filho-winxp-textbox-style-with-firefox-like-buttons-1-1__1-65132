VERSION 5.00
Begin VB.UserControl SearchBox 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   ScaleHeight     =   330
   ScaleWidth      =   6840
   ToolboxBitmap   =   "Firefox_Procura.ctx":0000
   Begin VB.Timer tmrMouseLeave 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3510
      Top             =   1020
   End
   Begin VB.PictureBox picFundo 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   480
      ScaleHeight     =   255
      ScaleWidth      =   5985
      TabIndex        =   2
      Top             =   40
      Width           =   5985
      Begin VB.TextBox txtProcura 
         BorderStyle     =   0  'None
         Height          =   290
         Left            =   30
         ScrollBars      =   1  'Horizontal
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   0
         Width           =   5985
      End
   End
   Begin VB.PictureBox imgMouse 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   4440
      Picture         =   "Firefox_Procura.ctx":0312
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   870
      Width           =   480
   End
   Begin VB.Label lblArrow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   0
      ToolTipText     =   "Clique aqui para opções"
      Top             =   110
      Width           =   180
   End
   Begin VB.Image imgIcone 
      Height          =   240
      Left            =   60
      Picture         =   "Firefox_Procura.ctx":0464
      Stretch         =   -1  'True
      ToolTipText     =   "Clique aqui para opções"
      Top             =   40
      Width           =   240
   End
   Begin VB.Image picFechar 
      Height          =   240
      Left            =   6540
      Picture         =   "Firefox_Procura.ctx":12A6
      Stretch         =   -1  'True
      ToolTipText     =   "Limpar texto"
      Top             =   50
      Width           =   240
   End
   Begin VB.Shape FundoTexto 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00B99D7F&
      FillColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   6825
   End
End
Attribute VB_Name = "SearchBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'isMouseOver
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private OverAnt As Boolean

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private m_OnFocusAutoSelect As Boolean
Private m_AutoTab As Boolean
Private m_EnterGotoNextTab As Boolean
Private m_BorderStyle As BorderStyleConstants
Private m_OnFocusBorderStyle As BorderStyleConstants
Private m_ButtonClear As Boolean
Private m_ButtonIcon As Boolean
Private m_ColorBorder As OLE_COLOR
Private m_OnFocusColorTextBorder As OLE_COLOR
Private m_ColorText As OLE_COLOR
Private m_CharSet As SearchCharCase
Private m_Font As New StdFont
Private m_InputType As SearchInputType
Private m_InputTypeCustomize As String
Private m_OnFocusFont As New StdFont
Private m_OnFocusColorText As OLE_COLOR
Private varxUltUp As Boolean

Enum SearchJustify
    SrchLeft = 0
    SrchRight = 1
    SrchCenter = 2
End Enum

Enum SearchCharCase
    SrchNone = 0
    SrchLowerCase = 1
    SrchUpperCase = 2
    SrchProperCase = 3
End Enum

Enum SearchInputType
    SrchNone = 0
    SrchAlphabetic = 1
    SrchNumeric = 2
    SrchAlphaNumeric = 3
    SrchCustomize = 4
End Enum

'Eventos
Event ArrowClick(Button As Integer)
Event ClearClick(Button As Integer)
Event Change()
Event Click()
Event DblClick()
Event IconClick(Button As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event MouseEnter()
Event MouseLeave()
Event MouseUp(Button As Integer)

'Eventos de funções
Private Sub imgIcone_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If picFundo.Enabled Then RaiseEvent IconClick(Button)
End Sub
Private Sub lblArrow_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If picFundo.Enabled Then RaiseEvent ArrowClick(Button)
End Sub
Private Sub picFechar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If picFundo.Enabled Then
        If (Button = 1) Then
            txtProcura.Text = vbNullString
            picFechar.Visible = False
            txtProcura.SetFocus
            varxUltUp = False
        End If
        RaiseEvent ClearClick(Button)
    End If
End Sub

'Eventos de texto
Private Sub txtProcura_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
    If m_EnterGotoNextTab Then
        If (KeyCode = vbKeyReturn) Then SendKeys "{TAB}"
    End If
End Sub
Private Sub txtProcura_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub txtProcura_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    Select Case m_InputType
        Case 0:
            KeyAscii = ModificaAscii(KeyAscii)
        Case 1:
            Select Case KeyAscii
                Case 65 To 90, 97 To 122, 44, 46, 8, 13, 32: KeyAscii = ModificaAscii(KeyAscii)
                Case Else: KeyAscii = 0
            End Select
        Case 2:
            Select Case KeyAscii
                Case 48 To 57, 46, 8, 13, 44: Exit Sub
                Case Else: KeyAscii = 0
            End Select
        Case 3:
            Select Case KeyAscii
                Case 48 To 57, 65 To 90, 97 To 122, 44, 46, 8, 13, 32: KeyAscii = ModificaAscii(KeyAscii)
                Case Else: KeyAscii = 0
            End Select
        Case 4:
            If (KeyAscii <> 8) And (KeyAscii <> 13) And (InStr(m_InputTypeCustomize, Chr(KeyAscii)) = 0) Then KeyAscii = 0
            Exit Sub
    End Select
End Sub
Private Sub txtProcura_Change()
    If picFundo.Enabled Then
        If m_ButtonClear Then
            picFechar.Visible = (txtProcura.Text <> vbNullString)
        End If
        If m_AutoTab Then If (Len(txtProcura.Text) = txtProcura.MaxLength) And (txtProcura.MaxLength <> 0) Then SendKeys "{TAB}"
        RaiseEvent Change
        UserControl_Resize
        If (Len(txtProcura.Text) = 0) Then varxUltUp = False
    End If
End Sub
Private Sub txtProcura_Click()
    RaiseEvent Click
End Sub
Private Sub txtProcura_DblClick()
    RaiseEvent DblClick
End Sub
Private Sub txtProcura_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseOverAct "Enter"
End Sub


'Propriedades
Public Property Get Alignment() As SearchJustify
    Alignment = txtProcura.Alignment
End Property
Public Property Let Alignment(ByVal New_Alignment As SearchJustify)
    If New_Alignment > 2 Then New_Alignment = 0
    txtProcura.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

Public Property Get AutoTab() As Boolean
    AutoTab = m_AutoTab
End Property
Public Property Let AutoTab(ByVal New_AutoTab As Boolean)
    m_AutoTab = New_AutoTab
    PropertyChanged "AutoTab"
End Property

Public Property Get BorderStyle() As BorderStyleConstants
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    m_BorderStyle = New_BorderStyle
    FundoTexto.BorderStyle = m_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get ButtonClear() As Boolean
    ButtonClear = m_ButtonClear
End Property
Public Property Let ButtonClear(ByVal New_ButtonClear As Boolean)
    m_ButtonClear = New_ButtonClear
    picFechar.Visible = m_ButtonClear
    PropertyChanged "ButtonClear"
    UserControl_Resize
End Property

Public Property Get ButtonIcon() As Boolean
    ButtonIcon = m_ButtonIcon
End Property
Public Property Let ButtonIcon(ByVal New_ButtonIcon As Boolean)
    m_ButtonIcon = New_ButtonIcon
    imgIcone.Visible = m_ButtonIcon
    lblArrow.Visible = m_ButtonIcon
    PropertyChanged "ButtonIcon"
    UserControl_Resize
End Property

Public Property Get CharSet() As SearchCharCase
    CharSet = m_CharSet
End Property
Public Property Let CharSet(ByVal New_CharSet As SearchCharCase)
    m_CharSet = New_CharSet
    AtualizaCharSets
    PropertyChanged "CharSet"
End Property

Public Property Get ColorArrow() As OLE_COLOR
    ColorArrow = lblArrow.ForeColor
End Property
Public Property Let ColorArrow(ByVal New_ColorArrow As OLE_COLOR)
    lblArrow.ForeColor = New_ColorArrow
    PropertyChanged "ColorArrow"
End Property

Public Property Get ColorText() As OLE_COLOR
    ColorText = m_ColorText
End Property
Public Property Let ColorText(ByVal New_ColorText As OLE_COLOR)
    m_ColorText = New_ColorText
    txtProcura.ForeColor = m_ColorText
    PropertyChanged "ColorText"
End Property

Public Property Get ColorBack() As OLE_COLOR
    ColorBack = txtProcura.BackColor
End Property
Public Property Let ColorBack(ByVal New_ColorBack As OLE_COLOR)
    txtProcura.BackColor = New_ColorBack
    FundoTexto.BackColor = New_ColorBack
    PropertyChanged "ColorBack"
End Property

Public Property Get ColorBorder() As OLE_COLOR
    ColorBorder = m_ColorBorder
End Property
Public Property Let ColorBorder(ByVal New_ColorBorder As OLE_COLOR)
    m_ColorBorder = New_ColorBorder
    FundoTexto.BorderColor = m_ColorBorder
    PropertyChanged "ColorBorder"
End Property

Public Property Get Enabled() As Boolean
    Enabled = picFundo.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    picFundo.Enabled = New_Enabled
    imgIcone.Enabled = New_Enabled
    lblArrow.Enabled = New_Enabled
    picFechar.Enabled = New_Enabled
    MudaMouse picFechar: MudaMouse imgIcone: MudaMouse lblArrow
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
    Set Font = m_Font
End Property
Public Property Set Font(ByRef New_Font As Font)
    Set m_Font = New_Font
    Set txtProcura.Font = m_Font
    PropertyChanged "Font"
End Property

Public Property Get EnterGotoNextTab() As Boolean
    EnterGotoNextTab = m_EnterGotoNextTab
End Property
Public Property Let EnterGotoNextTab(ByVal New_EnterGotoNextTab As Boolean)
    m_EnterGotoNextTab = New_EnterGotoNextTab
    PropertyChanged "EnterGotoNextTab"
End Property

Public Property Get Icon() As StdPicture
    Set Icon = imgIcone.Picture
End Property
Public Property Set Icon(ByVal New_Icon As StdPicture)
    On Error Resume Next
    If (New_Icon Is Nothing) Then Exit Property
    Set imgIcone.Picture = New_Icon
    PropertyChanged "Icon"
End Property

Public Property Get InputType() As SearchInputType
    InputType = m_InputType
End Property
Public Property Let InputType(ByVal New_InputType As SearchInputType)
    m_InputType = New_InputType
    PropertyChanged "InputType"
End Property

Public Property Get InputTypeCustomize() As String
    InputTypeCustomize = m_InputTypeCustomize
End Property
Public Property Let InputTypeCustomize(ByVal New_InputTypeCustomize As String)
    m_InputTypeCustomize = New_InputTypeCustomize
    PropertyChanged "InputTypeCustomize"
End Property

Public Property Get Locked() As Boolean
    Locked = txtProcura.Locked
End Property
Public Property Let Locked(ByVal New_Locked As Boolean)
    txtProcura.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

Public Property Get MaxLength() As Long
    MaxLength = txtProcura.MaxLength
End Property
Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtProcura.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

Public Property Get OnFocusAutoSelect() As Boolean
    OnFocusAutoSelect = m_OnFocusAutoSelect
End Property
Public Property Let OnFocusAutoSelect(ByVal New_OnFocusAutoSelect As Boolean)
    m_OnFocusAutoSelect = New_OnFocusAutoSelect
    PropertyChanged "OnFocusAutoSelect"
End Property

Public Property Get OnFocusColorBorder() As OLE_COLOR
    OnFocusColorBorder = m_OnFocusColorTextBorder
End Property
Public Property Let OnFocusColorBorder(ByVal New_OnFocusColorBorder As OLE_COLOR)
    m_OnFocusColorTextBorder = New_OnFocusColorBorder
    PropertyChanged "OnFocusColorBorder"
End Property

Public Property Get OnFocusColorText() As OLE_COLOR
    OnFocusColorText = m_OnFocusColorText
End Property
Public Property Let OnFocusColorText(ByVal New_OnFocusColor As OLE_COLOR)
    m_OnFocusColorText = New_OnFocusColor
    PropertyChanged "OnFocusColorText"
End Property

Public Property Get OnFocusFont() As Font
    Set OnFocusFont = m_OnFocusFont
End Property
Public Property Set OnFocusFont(ByRef New_OnFocusFont As Font)
    Set m_OnFocusFont = New_OnFocusFont
    PropertyChanged "OnFocusFont"
End Property

Public Property Get OnFocusBorderStyle() As BorderStyleConstants
    OnFocusBorderStyle = m_OnFocusBorderStyle
End Property
Public Property Let OnFocusBorderStyle(ByVal New_OnFocusBorderStyle As BorderStyleConstants)
    m_OnFocusBorderStyle = New_OnFocusBorderStyle
    PropertyChanged "OnFocusBorderStyle"
End Property

Public Property Get PasswordChar() As String
    PasswordChar = txtProcura.PasswordChar
End Property
Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    txtProcura.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

Public Property Get SelStart() As Long
    SelStart = txtProcura.SelStart
End Property
Public Property Let SelStart(ByVal New_SelStart As Long)
    txtProcura.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

Public Property Get SelLength() As Long
    SelLength = txtProcura.SelLength
End Property
Public Property Let SelLength(ByVal New_SelLength As Long)
    txtProcura.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

Public Property Get Text() As String
    Text = txtProcura.Text
End Property
Public Property Let Text(ByVal New_Text As String)
    On Error Resume Next
    txtProcura.Text = New_Text
    PropertyChanged "Text"
End Property

Public Property Get ToolTipIcon() As String
    ToolTipIcon = imgIcone.ToolTipText
End Property
Public Property Let ToolTipIcon(ByVal New_ToolTipIcon As String)
    imgIcone.ToolTipText = New_ToolTipIcon
    PropertyChanged "ToolTipIcon"
End Property

Public Property Get ToolTipArrow() As String
    ToolTipArrow = lblArrow.ToolTipText
End Property
Public Property Let ToolTipArrow(ByVal New_ToolTipArrow As String)
    lblArrow.ToolTipText = New_ToolTipArrow
    PropertyChanged "ToolTipArrow"
End Property

Public Property Get ToolTipFechar() As String
    ToolTipFechar = picFechar.ToolTipText
End Property
Public Property Let ToolTipFechar(ByVal New_ToolTipFechar As String)
    picFechar.ToolTipText = New_ToolTipFechar
    PropertyChanged "ToolTipFechar"
End Property

Public Property Get RoundCorners() As Boolean
    RoundCorners = (FundoTexto.Shape = 4)
End Property
Public Property Let RoundCorners(ByVal New_RoundCorner As Boolean)
    FundoTexto.Shape = IIf(New_RoundCorner, 4, 0)
    PropertyChanged "RoundCorners"
End Property

'Outros eventos do controle
Private Sub UserControl_Initialize()
    txtProcura_Change
End Sub
Private Sub UserControl_InitProperties()
    m_OnFocusAutoSelect = False
    m_AutoTab = False
    m_EnterGotoNextTab = True
    m_InputType = 0
    m_ButtonClear = True
    m_ButtonIcon = True
    m_ColorBorder = &HB99D7F
    m_OnFocusColorTextBorder = &H96E7&
    m_BorderStyle = 1
    m_OnFocusBorderStyle = 1
    
    m_Font.Name = "MS Sans Serif"
    m_Font.Size = "8"
    m_OnFocusFont.Name = "MS Sans Serif"
    m_OnFocusFont.Size = "8"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    txtProcura.Alignment = PropBag.ReadProperty("Alignment", 0)
    m_OnFocusAutoSelect = PropBag.ReadProperty("OnFocusAutoSelect", False)
    m_AutoTab = PropBag.ReadProperty("AutoTab", False)
    If m_OnFocusAutoSelect Then txtProcura.SelStart = 0: txtProcura.SelLength = Len(txtProcura.Text)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", 1): FundoTexto.BorderStyle = m_BorderStyle
    m_ButtonClear = PropBag.ReadProperty("ButtonClear", True)
    m_ButtonIcon = PropBag.ReadProperty("ButtonIcon", True)
    m_CharSet = PropBag.ReadProperty("CharSet", 0)
    lblArrow.ForeColor = PropBag.ReadProperty("ColorArrow", &H80000008)
    txtProcura.BackColor = PropBag.ReadProperty("ColorBack", &HFFFFFF)
    FundoTexto.BackColor = PropBag.ReadProperty("ColorBack", &HFFFFFF)
    m_ColorBorder = PropBag.ReadProperty("ColorBorder", &HB99D7F): FundoTexto.BorderColor = m_ColorBorder
    m_ColorText = PropBag.ReadProperty("ColorText", &H80000008): txtProcura.ForeColor = m_ColorText
    picFundo.Enabled = PropBag.ReadProperty("Enabled", True)
    imgIcone.Enabled = Enabled: lblArrow.Enabled = Enabled: picFechar.Enabled = Enabled
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font): Set txtProcura.Font = m_Font
    m_EnterGotoNextTab = PropBag.ReadProperty("EnterGotoNextTab", True)
    Set imgIcone.Picture = PropBag.ReadProperty("Icon", imgIcone.Picture)
    m_InputType = PropBag.ReadProperty("InputType", 0)
    m_InputTypeCustomize = PropBag.ReadProperty("InputTypeCustomize", vbNullString)
    txtProcura.Locked = PropBag.ReadProperty("Locked", False)
    txtProcura.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    m_OnFocusBorderStyle = PropBag.ReadProperty("OnFocusBorderStyle", 1)
    m_OnFocusColorText = PropBag.ReadProperty("OnFocusColorText", &H80000008)
    m_OnFocusColorTextBorder = PropBag.ReadProperty("OnFocusColorBorder", &H96E7&)
    Set m_OnFocusFont = PropBag.ReadProperty("OnFocusFont")
    txtProcura.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    FundoTexto.Shape = IIf(PropBag.ReadProperty("RoundCorners", 4), 4, 0)
    txtProcura.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtProcura.SelText = PropBag.ReadProperty("SelText", "")
    txtProcura.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtProcura.Text = PropBag.ReadProperty("Text", "Text1")
    imgIcone.ToolTipText = PropBag.ReadProperty("ToolTipIcon", "")
    lblArrow.ToolTipText = PropBag.ReadProperty("ToolTipArrow", "Clique aqui para mais opções")
    picFechar.ToolTipText = PropBag.ReadProperty("ToolTipFechar", "Limpar texto")
    MudaMouse picFechar: MudaMouse imgIcone: MudaMouse lblArrow
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Alignment", txtProcura.Alignment, 0)
    Call PropBag.WriteProperty("OnFocusAutoSelect", m_OnFocusAutoSelect, False)
    Call PropBag.WriteProperty("AutoTab", m_AutoTab, False)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, 1)
    Call PropBag.WriteProperty("OnFocusBorderStyle", m_OnFocusBorderStyle, 1)
    Call PropBag.WriteProperty("ButtonClear", m_ButtonClear, True)
    Call PropBag.WriteProperty("ButtonIcon", m_ButtonIcon, True)
    Call PropBag.WriteProperty("CharSet", m_CharSet, 0)
    Call PropBag.WriteProperty("ColorArrow", lblArrow.ForeColor, &H80000008)
    Call PropBag.WriteProperty("ColorBack", txtProcura.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("ColorText", m_ColorText, &H80000008)
    Call PropBag.WriteProperty("ColorBorder", m_ColorBorder, &HB99D7)
    Call PropBag.WriteProperty("OnFocusColorBorder", m_OnFocusColorTextBorder, &H96E7&)
    Call PropBag.WriteProperty("Enabled", picFundo.Enabled, True)
    Call PropBag.WriteProperty("Font", m_Font)
    Call PropBag.WriteProperty("EnterGotoNextTab", m_EnterGotoNextTab, True)
    Call PropBag.WriteProperty("Icon", imgIcone.Picture)
    Call PropBag.WriteProperty("InputType", m_InputType, 0)
    Call PropBag.WriteProperty("InputTypeCustomize", m_InputTypeCustomize, vbNullString)
    Call PropBag.WriteProperty("Locked", txtProcura.Locked, False)
    Call PropBag.WriteProperty("MaxLength", txtProcura.MaxLength, 0)
    Call PropBag.WriteProperty("OnFocusColorText", m_OnFocusColorText, &H80000008)
    Call PropBag.WriteProperty("OnFocusFont", m_OnFocusFont)
    Call PropBag.WriteProperty("PasswordChar", txtProcura.PasswordChar, "")
    Call PropBag.WriteProperty("RoundCorners", FundoTexto.Shape, 4)
    Call PropBag.WriteProperty("SelStart", txtProcura.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtProcura.SelText, "")
    Call PropBag.WriteProperty("SelLength", txtProcura.SelLength, 0)
    Call PropBag.WriteProperty("Text", txtProcura.Text, "Text1")
    Call PropBag.WriteProperty("ToolTipIcon", imgIcone.ToolTipText, "")
    Call PropBag.WriteProperty("ToolTipArrow", lblArrow.ToolTipText, "Clique aqui para mais opções")
    Call PropBag.WriteProperty("ToolTipFechar", picFechar.ToolTipText, "Limpar texto")
End Sub


'Interface
Private Sub UserControl_EnterFocus()
    If picFundo.Enabled Then
        txtProcura.ForeColor = m_OnFocusColorText
        Set txtProcura.Font = m_OnFocusFont
        FundoTexto.BorderStyle = m_OnFocusBorderStyle
        FundoTexto.BorderColor = m_OnFocusColorTextBorder
        If m_OnFocusAutoSelect Then
            txtProcura.SelStart = 0
            txtProcura.SelLength = Len(txtProcura.Text)
        End If
        txtProcura.SetFocus
    End If
End Sub
Private Sub txtProcura_LostFocus()
    txtProcura.ForeColor = m_ColorText
    Set txtProcura.Font = m_Font
    txtProcura.ForeColor = m_ColorText
    FundoTexto.BorderColor = m_ColorBorder
    FundoTexto.BorderStyle = m_BorderStyle
    AtualizaCharSets
End Sub


Private Sub UserControl_Resize()
    On Error Resume Next
    UserControl.Height = FundoTexto.Height
    FundoTexto.Width = UserControl.Width
    txtProcura.Move 0, 0, picFundo.Width, picFundo.Height
    picFechar.Move FundoTexto.Width - 280
    
    imgIcone.Visible = m_ButtonIcon
    lblArrow.Visible = m_ButtonIcon
    If m_ButtonIcon Then
        If m_ButtonClear Then
            picFechar.Visible = (txtProcura.Text <> vbNullString)
            picFundo.Move 450, 60, (FundoTexto.Width - 760)
        Else
            picFechar.Visible = False
            picFundo.Move 450, 60, (FundoTexto.Width - 840) + 340
        End If
    Else
        If m_ButtonClear Then
            picFechar.Visible = (txtProcura.Text <> vbNullString)
            picFundo.Move 450 - 400, 60, (FundoTexto.Width - 740) + 380
        Else
            picFechar.Visible = False
            picFundo.Move 450 - 400, 60, (FundoTexto.Width - 840) + 740
        End If
    End If
End Sub


'Funções extras
Private Sub MudaMouse(ByVal Ctl As Control)
    On Error Resume Next
    If picFundo.Enabled Then
        Ctl.MousePointer = vbCustom
        Set Ctl.MouseIcon = imgMouse.Picture
    Else
        Ctl.MousePointer = vbDefault
        Set Ctl.MouseIcon = Nothing
    End If
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseOverAct "Enter"
End Sub
Private Sub imgIcone_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseOverAct "Enter"
End Sub
Private Sub lblArrow_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseOverAct "Enter"
End Sub
Private Sub picFechar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    MouseOverAct "Enter"
End Sub
Private Sub tmrMouseLeave_Timer()
    MouseOverAct "Leave"
End Sub
Private Sub MouseOverAct(ByVal Acao As String)
    If (Acao = "Enter") Then
        If Not OverAnt Then
            tmrMouseLeave.Enabled = True
            OverAnt = True
            RaiseEvent MouseEnter
        End If
    Else
        If Not isMouseOver Then
            If OverAnt Then
                tmrMouseLeave.Enabled = False
                OverAnt = False
                RaiseEvent MouseLeave
            End If
        End If
    End If
End Sub
Private Function isMouseOver() As Boolean
    Dim pointCurPos As POINTAPI
    Dim rectCtl As RECT
    GetCursorPos pointCurPos
    GetWindowRect UserControl.hWnd, rectCtl
    With pointCurPos
        isMouseOver = (.x >= rectCtl.Left) And (.x <= rectCtl.Right) And (.y >= rectCtl.Top) And (.y <= rectCtl.Bottom)
    End With
End Function
Private Function ModificaAscii(ByVal KeyAscii As Integer)
    Select Case m_CharSet
        Case 0:
            ModificaAscii = KeyAscii: Exit Function
        Case 1:
            If (KeyAscii > 64 And KeyAscii < 91) Then ModificaAscii = KeyAscii + 32: Exit Function
        Case 2:
            If (KeyAscii > 96 And KeyAscii < 123) Then ModificaAscii = KeyAscii - 32: Exit Function
        Case 3:
            If ((KeyAscii > 64 And KeyAscii < 91) Or (KeyAscii > 96 And KeyAscii < 123)) Or (KeyAscii = 32) Then
                If (KeyAscii = 32) Then
                    varxUltUp = False
                Else
                    If varxUltUp Then
                        If (KeyAscii > 64 And KeyAscii < 91) Then ModificaAscii = KeyAscii + 32: Exit Function
                    Else
                        varxUltUp = True
                        If (KeyAscii > 96 And KeyAscii < 123) Then ModificaAscii = KeyAscii - 32: Exit Function
                    End If
                End If
            End If
    End Select
    ModificaAscii = KeyAscii
End Function
Private Sub AtualizaCharSets()
    Select Case m_CharSet
        Case 1:
            txtProcura.Text = LCase(txtProcura.Text)
        Case 2:
            txtProcura.Text = UCase(txtProcura.Text)
        Case 3:
            Dim x As Long, xSpaces As Boolean, xUltUp As Boolean
            With txtProcura
                For x = 1 To Len(.Text)
                    If (Mid(.Text, x, 1) = Chr(32)) Then
                        xSpaces = True
                        xUltUp = False
                    Else
                        If xUltUp Then
                            .Text = Left(.Text, x - 1) & LCase(Mid(.Text, x, 1)) & Mid(.Text, x + 1)
                        Else
                            .Text = Left(.Text, x - 1) & UCase(Mid(.Text, x, 1)) & Mid(.Text, x + 1)
                            xUltUp = True
                        End If
                    End If
                Next
            End With
    End Select
End Sub


'Créditos
Public Sub About()
    MsgBox "Textbox estilo Windows XP + Firefox buttons by Body_of_Rays" & vbCrLf & vbCrLf & "E-mail: body_of_rays@yahoo.com.br", vbInformation + vbOKOnly
End Sub


Private Sub UserControl_Terminate()
    tmrMouseLeave.Enabled = False
End Sub
