VERSION 5.00
Begin VB.UserControl GraFxButton 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3345
   DefaultCancel   =   -1  'True
   ScaleHeight     =   119
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   223
   ToolboxBitmap   =   "GraFxButton.ctx":0000
   Begin VB.PictureBox pictTempHighlight 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox pictTempDestination 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   1800
      ScaleHeight     =   600
      ScaleWidth      =   1440
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox imgPicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   240
      ScaleHeight     =   600
      ScaleWidth      =   1440
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "GraFxButton.ctx":0312
      Top             =   1320
      Width           =   1260
   End
End
Attribute VB_Name = "GraFxButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const BLACKNESS = &H42
Private Const NOTSRCCOPY = &H330008
Private Const NOTSRCERASE = &H1100A6
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020
Private Const SRCERASE = &H440328
Private Const SRCINVERT = &H660046
Private Const SRCPAINT = &HEE0086
Private Const WHITENESS = &HFF006

Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2

Private Const BDR_RAISED = &H5
Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC

Private Const BF_ADJUST = &H2000
Private Const BF_FLAT = &H4000
Private Const BF_MONO = &H8000
Private Const BF_SOFT = &H1000
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)

Private Const DT_CENTER = &H1
Private Const DT_RTLREADING = &H20000
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4

Private Const DST_COMPLEX = &H0
Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4

Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10                   '/* Gray string appearance */
Private Const DSS_DISABLED = &H20
Private Const DSS_RIGHT = &H8000

'//---------------------------------------------------------------------------------------
'// Windows API types
'//---------------------------------------------------------------------------------------
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'//---------------------------------------------------------------------------------------
'// Windows API declarations
'//---------------------------------------------------------------------------------------
Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function DrawStateText Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As String, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Enum StateConstants
    btDown = 0
    btUp = 1
    btOver = 2
    btDisabled = 3
    btFocus = 4
End Enum

Private Enum RasterOperationConstants
    roNotSrcCopy = NOTSRCCOPY
    roNotSrcErase = NOTSRCERASE
    roSrcAnd = SRCAND
    roSrcCopy = SRCCOPY
    roSrcErase = SRCERASE
    roSrcInvert = SRCINVERT
    roSrcPaint = SRCPAINT
End Enum

Private Const clTop As Long = 6
Private Const clLeft As Long = 6
Private Const clFocusOffset As Long = 4
Private Const clDownOffset As Long = 1

Private tPrevEvent As String
Private lState As StateConstants
Private bLeftFocus As Boolean
Private bHasFocus As Boolean

Private Enum AppearanceConstants
    Flat = 0
    [3D] = 1
    Skin = 2
End Enum

Private Enum StyleConstants
    Default = 0
    ButtonGroup = 1
End Enum

Public Enum ValueConstants
    Down = 0
    Up = 1
End Enum

Private Const m_def_Appearance = Skin
Private Const m_def_BackColor = vbButtonFace
Private Const m_def_Caption = "GraFxButton"
Private Const m_def_CaptionOffsetX = 0
Private Const m_def_CaptionOffsetY = 0
Private Const m_def_Enabled = True
Private Const m_def_ForeColor = vbButtonText
Private Const m_def_HighlightColor = vbButtonText
Private Const m_def_HighlightPicture = False
Private Const m_def_MousePointer = vbDefault
Private Const m_def_PictureOffsetX = 0
Private Const m_def_PictureOffsetY = 0
'Private Const m_def_RightToLeft = False
Private Const m_def_Style = 0
'Private Const m_def_ToolTipText = ""
Private Const m_def_TransparentColor = vbMagenta
Private Const m_def_Value = Up
'Private Const m_def_WhatsThisHelpID = 0

Private m_Appearance As AppearanceConstants
Private m_BackColor As OLE_COLOR
Private m_Caption As String
Private m_CaptionOffsetX As Long
Private m_CaptionOffsetY As Long
Private m_Enabled As Boolean
Private m_ForeColor As OLE_COLOR
Private m_Font As Font
Private m_HighlightColor As OLE_COLOR
Private m_HighlightPicture As Boolean
Private m_MouseIcon As Picture
Private m_MousePointer As MousePointerConstants
Private m_Picture As Picture
Private m_PictureDisabled As Picture
Private m_PictureDown As Picture
Private m_PictureFocus As Picture
Private m_PictureOffsetX As Long
Private m_PictureOffsetY As Long
Private m_PictureOver As Picture
Private m_RightToLeft As Boolean
Private m_SkinDisabled As Picture
Private m_SkinDown As Picture
Private m_SkinFocus As Picture
Private m_SkinOver As Picture
Private m_SkinUp As Picture
Private m_Style As StyleConstants
Private m_ToolTipText As String
Private m_TransparentColor As OLE_COLOR
Private m_Value As ValueConstants
Private m_WhatsThisHelpID As Long

Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over the control."
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while the control has the focus."
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while the control has the focus."
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while the control has the focus."
Public Event MouseEnter()
Attribute MouseEnter.VB_Description = "Occurs when the user moves the mouse over the control after MouseExit event."
Public Event MouseExit()
Attribute MouseExit.VB_Description = "Occurs when the user moves the mouse out of the control after MouseEnter event."
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while the control has the focus."
Public Event Resize()
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of the control changes."

Private Property Get Appearance() As AppearanceConstants
    Appearance = m_Appearance
End Property

Private Property Let Appearance(ByVal NewValue As AppearanceConstants)
    m_Appearance = NewValue
        
    Call DrawButton(lState)
    
    PropertyChanged "Appearance"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in the control."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    m_BackColor = NewValue
    UserControl.BackColor = NewValue
    imgPicture.BackColor = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "BackColor"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in the control."
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal NewValue As String)
    Dim lPlace As Long
    
    m_Caption = NewValue
    
    lPlace = 0
    lPlace = InStr(lPlace + 1, NewValue, "&", vbTextCompare)
    Do While lPlace <> 0
        If Mid$(NewValue, lPlace + 1, 1) <> "&" Then
            UserControl.AccessKeys = Mid$(NewValue, lPlace + 1, 1)
            Exit Do
        Else
            lPlace = lPlace + 1
        End If
    
        lPlace = InStr(lPlace + 1, NewValue, "&", vbTextCompare)
    Loop
    
    Call DrawButton(lState)
    
    PropertyChanged "Caption"
End Property

Public Property Get CaptionOffsetX() As Long
Attribute CaptionOffsetX.VB_Description = "Returns/sets the horizontal offset for displaying the caption."
    CaptionOffsetX = m_CaptionOffsetX
End Property

Public Property Let CaptionOffsetX(ByVal NewValue As Long)
    m_CaptionOffsetX = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "CaptionOffsetX"
End Property

Public Property Get CaptionOffsetY() As Long
Attribute CaptionOffsetY.VB_Description = "Returns/sets the vertical offset for displaying the caption."
    CaptionOffsetY = m_CaptionOffsetY
End Property

Public Property Let CaptionOffsetY(ByVal NewValue As Long)
    m_CaptionOffsetY = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "CaptionOffsetY"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    m_Enabled = NewValue
    UserControl.Enabled = NewValue
    imgPicture.Enabled = NewValue
    
    If m_Enabled Then
        lState = btUp
    End If
    Call DrawButton(lState)
    
    PropertyChanged "Enabled"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in the control."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    m_ForeColor = NewValue
    UserControl.ForeColor = NewValue
    imgPicture.ForeColor = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "ForeColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns/sets a Font object used to display text in the control."
    Set Font = m_Font
End Property

Public Property Set Font(ByVal NewValue As Font)
    Set m_Font = NewValue
    Set UserControl.Font = NewValue
    Set imgPicture.Font = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "Font"
End Property

Public Property Get HighlightColor() As OLE_COLOR
Attribute HighlightColor.VB_Description = "Returns/sets the highlight color used to display text and graphics when the mouse is over the control."
    HighlightColor = m_HighlightColor
End Property

Public Property Let HighlightColor(ByVal NewValue As OLE_COLOR)
    m_HighlightColor = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "HighlightColor"
End Property

Public Property Get HighlightPicture() As Boolean
Attribute HighlightPicture.VB_Description = "Returns/sets whether or not to highlight the object's picture with the HighlightColor."
    HighlightPicture = m_HighlightPicture
End Property

Public Property Let HighlightPicture(ByVal NewValue As Boolean)
    m_HighlightPicture = NewValue
    
    Call DrawButton(btDisabled)
    
    PropertyChanged "HighlightPicture"
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Returns/sets a custom mouse icon."
    Set MouseIcon = m_MouseIcon
End Property

Public Property Set MouseIcon(ByVal NewValue As Picture)
    Set m_MouseIcon = NewValue
    Set UserControl.MouseIcon = NewValue
    Set imgPicture.MouseIcon = NewValue
    
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of the control."
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal NewValue As MousePointerConstants)
    m_MousePointer = NewValue
    UserControl.MousePointer = NewValue
    imgPicture.MousePointer = NewValue
    
    PropertyChanged "MousePointer"
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in the control."
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal NewValue As Picture)
    Set m_Picture = NewValue
    Set imgPicture.Picture = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "Picture"
End Property
'
'Public Property Get PictureDisabled() As Picture
'    Set PictureDisabled = m_PictureDisabled
'End Property
'
'Public Property Set PictureDisabled(ByVal NewValue As Picture)
'    Set m_PictureDisabled = NewValue
'    PropertyChanged "PictureDisabled"
'End Property
'
'Public Property Get PictureDown() As Picture
'    Set PictureDown = m_PictureDown
'End Property
'
'Public Property Set PictureDown(ByVal NewValue As Picture)
'    Set m_PictureDown = NewValue
'    PropertyChanged "PictureDown"
'End Property
'
'Public Property Get PictureFocus() As Picture
'    Set PictureFocus = m_PictureFocus
'End Property
'
'Public Property Set PictureFocus(ByVal New_PictureFocus As Picture)
'    Set m_PictureFocus = New_PictureFocus
'    PropertyChanged "PictureFocus"
'End Property

Public Property Get PictureOffsetX() As Long
Attribute PictureOffsetX.VB_Description = "Returns/sets the horizontal offset for displaying the picture."
    PictureOffsetX = m_PictureOffsetX
End Property

Public Property Let PictureOffsetX(ByVal NewValue As Long)
    m_PictureOffsetX = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "PictureOffsetX"
End Property

Public Property Get PictureOffsetY() As Long
Attribute PictureOffsetY.VB_Description = "Returns/sets the vertical offset for displaying the picture."
    PictureOffsetY = m_PictureOffsetY
End Property

Public Property Let PictureOffsetY(ByVal NewValue As Long)
    m_PictureOffsetY = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "PictureOffsetY"
End Property
'
'Public Property Get PictureOver() As Picture
'    Set PictureOver = m_PictureOver
'End Property
'
'Public Property Set PictureOver(ByVal New_PictureOver As Picture)
'    Set m_PictureOver = New_PictureOver
'    PropertyChanged "PictureOver"
'End Property
'
'Public Property Get RightToLeft() As Boolean
'    RightToLeft = m_RightToLeft
'End Property
'
'Public Property Let RightToLeft(ByVal NewValue As Boolean)
'    m_RightToLeft = NewValue
'    UserControl.RightToLeft = NewValue
'    imgPicture.RightToLeft = NewValue
'
'    Call DrawButton(lState)
'
'    PropertyChanged "RightToLeft"
'End Property

Public Property Get SkinDisabled() As Picture
Attribute SkinDisabled.VB_Description = "Returns/sets a graphic to be displayed for the control when it is disabled."
    Set SkinDisabled = m_SkinDisabled
End Property

Public Property Set SkinDisabled(ByVal NewValue As Picture)
    Set m_SkinDisabled = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "SkinDisabled"
End Property

Public Property Get SkinDown() As Picture
Attribute SkinDown.VB_Description = "Returns/sets a graphic to be displayed for the control the mouse has been pressed over it."
    Set SkinDown = m_SkinDown
End Property

Public Property Set SkinDown(ByVal NewValue As Picture)
    Set m_SkinDown = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "SkinDown"
End Property

Public Property Get SkinFocus() As Picture
Attribute SkinFocus.VB_Description = "Returns/sets a graphic to be displayed for the control when it default."
    Set SkinFocus = m_SkinFocus
End Property

Public Property Set SkinFocus(ByVal NewValue As Picture)
    Set m_SkinFocus = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "SkinFocus"
End Property

Public Property Get SkinOver() As Picture
Attribute SkinOver.VB_Description = "Returns/sets a graphic to be displayed for the control when the mouse is over it."
    Set SkinOver = m_SkinOver
End Property

Public Property Set SkinOver(ByVal NewValue As Picture)
    Set m_SkinOver = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "SkinOver"
End Property

Public Property Get SkinUp() As Picture
Attribute SkinUp.VB_Description = "Returns/sets a graphic to be displayed for the control."
    Set SkinUp = m_SkinUp
End Property

Public Property Set SkinUp(ByVal NewValue As Picture)
    Set m_SkinUp = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "SkinUp"
End Property

Private Property Get Style() As StyleConstants
    Style = m_Style
End Property

Private Property Let Style(ByVal NewValue As StyleConstants)
    m_Style = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "Style"
End Property
'
'Public Property Get ToolTipText() As String
'    ToolTipText = m_ToolTipText
'End Property
'
'Public Property Let ToolTipText(ByVal NewValue As String)
'    m_ToolTipText = NewValue
'    imgPicture.ToolTipText = NewValue
'
'    PropertyChanged "ToolTipText"
'End Property

Public Property Get TransparentColor() As OLE_COLOR
Attribute TransparentColor.VB_Description = "Returns/sets the color of the Picture property to make transparent."
    TransparentColor = m_TransparentColor
End Property

Public Property Let TransparentColor(ByVal NewValue As OLE_COLOR)
    m_TransparentColor = NewValue
    UserControl.MaskColor = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "TransparentColor"
End Property

Public Property Get Value() As ValueConstants
Attribute Value.VB_Description = "Returns/sets a default state for the control."
    Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As ValueConstants)
    m_Value = NewValue
    
    Call DrawButton(m_Value)
    
    PropertyChanged "Value"
End Property
'
'Public Property Get WhatsThisHelpID() As Long
'    WhatsThisHelpID = m_WhatsThisHelpID
'End Property
'
'Public Property Let WhatsThisHelpID(ByVal NewValue As Long)
'    m_WhatsThisHelpID = NewValue
'    imgPicture.WhatsThisHelpID = NewValue
'
'    PropertyChanged "WhatsThisHelpID"
'End Property

Private Sub imgPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, imgPicture.Left + (X \ Screen.TwipsPerPixelX), imgPicture.Top + (Y \ Screen.TwipsPerPixelY))
End Sub

Private Sub imgPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, imgPicture.Left + (X \ Screen.TwipsPerPixelX), imgPicture.Top + (Y \ Screen.TwipsPerPixelY))
End Sub

Private Sub imgPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, imgPicture.Left + (X \ Screen.TwipsPerPixelX), imgPicture.Top + (Y \ Screen.TwipsPerPixelY))
End Sub
Private Sub Timer1_Timer()
    Dim pnt As POINTAPI
    
    GetCursorPos pnt
    ScreenToClient UserControl.hWnd, pnt
    
    If pnt.X < UserControl.ScaleLeft Or _
            pnt.Y < UserControl.ScaleTop Or _
            pnt.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
            pnt.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
        Timer1.Enabled = False
    
        Call RaiseEventEx("MouseExit")
        
        If lState <> btUp Then
            Call DrawButton(btUp)
        End If
        bLeftFocus = True
    Else
        If bLeftFocus Then
            Call DrawButton(btDown)
        End If
    End If
End Sub
Private Sub UserControl_InitProperties()
    Appearance = m_def_Appearance
    BackColor = m_def_BackColor
    Caption = m_def_Caption
    CaptionOffsetX = m_def_CaptionOffsetX
    CaptionOffsetY = m_def_CaptionOffsetY
    Enabled = m_def_Enabled
    ForeColor = m_def_ForeColor
    Set Font = Ambient.Font
    HighlightColor = m_def_HighlightColor
    HighlightPicture = m_def_HighlightPicture
    Set MouseIcon = LoadPicture("")
    MousePointer = m_def_MousePointer
    PictureOffsetX = m_def_PictureOffsetX
    PictureOffsetY = m_def_PictureOffsetY
    Set SkinDisabled = LoadPicture("")
    Set SkinDown = LoadPicture("")
    Set SkinFocus = LoadPicture("")
    Set SkinOver = LoadPicture("")
    'Set SkinUp = LoadPicture("")
    Set SkinUp = Image1.Picture
    Style = m_def_Style
    TransparentColor = m_def_TransparentColor
    Value = m_def_Value
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    CaptionOffsetX = PropBag.ReadProperty("CaptionOffsetX", m_def_CaptionOffsetX)
    CaptionOffsetY = PropBag.ReadProperty("CaptionOffsetY", m_def_CaptionOffsetY)
    Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    HighlightColor = PropBag.ReadProperty("HighlightColor", m_def_HighlightColor)
    HighlightPicture = PropBag.ReadProperty("HighlightPicture", m_def_HighlightPicture)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    MousePointer = PropBag.ReadProperty("MousePointer", m_def_MousePointer)
    PictureOffsetX = PropBag.ReadProperty("PictureOffsetX", m_def_PictureOffsetX)
    PictureOffsetY = PropBag.ReadProperty("PictureOffsetY", m_def_PictureOffsetY)
    Set SkinDisabled = PropBag.ReadProperty("SkinDisabled", Nothing)
    Set SkinDown = PropBag.ReadProperty("SkinDown", Nothing)
    Set SkinFocus = PropBag.ReadProperty("SkinFocus", Nothing)
    Set SkinOver = PropBag.ReadProperty("SkinOver", Nothing)
    Set SkinUp = PropBag.ReadProperty("SkinUp", Nothing)
    Style = PropBag.ReadProperty("Style", m_def_Style)
    TransparentColor = PropBag.ReadProperty("TransparentColor", m_def_TransparentColor)
    Value = PropBag.ReadProperty("Value", m_def_Value)
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("CaptionOffsetX", m_CaptionOffsetX, m_def_CaptionOffsetX)
    Call PropBag.WriteProperty("CaptionOffsetY", m_CaptionOffsetY, m_def_CaptionOffsetY)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("HighlightColor", m_HighlightColor, m_def_HighlightColor)
    Call PropBag.WriteProperty("HighlightPicture", m_HighlightPicture, m_def_HighlightPicture)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("PictureDisabled", m_PictureDisabled, Nothing)
    Call PropBag.WriteProperty("PictureDown", m_PictureDown, Nothing)
    Call PropBag.WriteProperty("PictureFocus", m_PictureFocus, Nothing)
    Call PropBag.WriteProperty("PictureOffsetX", m_PictureOffsetX, m_def_PictureOffsetX)
    Call PropBag.WriteProperty("PictureOffsetY", m_PictureOffsetY, m_def_PictureOffsetY)
    Call PropBag.WriteProperty("PictureOver", m_PictureOver, Nothing)
'    Call PropBag.WriteProperty("RightToLeft", m_RightToLeft, m_def_RightToLeft)
    Call PropBag.WriteProperty("TransparentColor", m_TransparentColor, m_def_TransparentColor)
    Call PropBag.WriteProperty("MouseIcon", m_MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", m_MousePointer, m_def_MousePointer)
    Call PropBag.WriteProperty("SkinDisabled", m_SkinDisabled, Nothing)
    Call PropBag.WriteProperty("SkinDown", m_SkinDown, Nothing)
    Call PropBag.WriteProperty("SkinFocus", m_SkinFocus, Nothing)
    Call PropBag.WriteProperty("SkinOver", m_SkinOver, Nothing)
    Call PropBag.WriteProperty("SkinUp", m_SkinUp, Nothing)
    Call PropBag.WriteProperty("Style", m_Style, m_def_Style)
'    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
    Call PropBag.WriteProperty("TransparentColor", m_TransparentColor, m_def_TransparentColor)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
'    Call PropBag.WriteProperty("WhatsThisHelpID", m_WhatsThisHelpID, m_def_WhatsThisHelpID)
End Sub

Private Sub UserControl_Click()
    Call RaiseEventEx("Click")
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RaiseEventEx("KeyDown", KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    Call RaiseEventEx("KeyPress", KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    Call RaiseEventEx("KeyUp", KeyCode, Shift)
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    Call RaiseEventEx("Click")
End Sub
Public Function About() As String
         Inf.Show
End Function

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "DisplayAsDefault" Then
        If UserControl.Ambient.DisplayAsDefault Then
            bHasFocus = True
        Else
            bHasFocus = False
        End If
        Call DrawButton(lState)
    End If
End Sub

Private Sub UserControl_Initialize()
    UserControl.Width = 1200
    UserControl.Height = 360
End Sub

Private Sub UserControl_GotFocus()
    bHasFocus = True
    Call DrawButton(lState)
End Sub

Private Sub UserControl_LostFocus()
    bHasFocus = False
    Call DrawButton(lState)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bLeftFocus = False
    
    If Button = vbLeftButton Then
        If lState = btDown Then
            m_Value = Up
        Else
            m_Value = Down
        End If
        
        Call DrawButton(btDown)
    End If
    
    Call RaiseEventEx("MouseDown", Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bLeftFocus = False
    
    If UserControl.Ambient.UserMode = True And Not Timer1.Enabled Then
        Timer1.Enabled = True
    
    ElseIf Button = 0 Then
        If lState <> btOver Then
            Call DrawButton(btOver)
        End If

    ElseIf Button = vbLeftButton Then
        If lState <> btDown Then
            Call DrawButton(btDown)
        End If
    End If

    If X >= 0 And Y >= 0 And _
                X <= UserControl.ScaleWidth And Y <= UserControl.ScaleHeight Then
        Call RaiseEventEx("MouseEnter")
        Call RaiseEventEx("MouseMove", Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bLeftFocus = False
    
    If Button = vbLeftButton Then
        Call DrawButton(btUp)
    End If

    Call RaiseEventEx("MouseUp", Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
End Sub

Private Sub UserControl_Resize()
    Call DrawButton(btUp)
    Call RaiseEventEx("Resize")
End Sub
Private Sub TransparentBlt_New2(ByVal hdc As Long, ByVal Source As PictureBox, ByRef DestPoint As POINTAPI, ByRef SrcPoint As POINTAPI, ByVal Width As Long, ByVal Height As Long, Optional ByVal TransparentColor As OLE_COLOR = -1, Optional ByVal Clear As Boolean = False, Optional ByVal Resize As Boolean = False, Optional ByVal Refresh As Boolean = False)
    Dim MonoMaskDC As Long
    Dim hMonoMask As Long
    Dim MonoInvDC As Long
    Dim hMonoInv As Long
    Dim ResultDstDC As Long
    Dim hResultDst As Long
    Dim ResultSrcDC As Long
    Dim hResultSrc As Long
    Dim hPrevMask As Long
    Dim hPrevInv As Long
    Dim hPrevSrc As Long
    Dim hPrevDst As Long
    Dim OldBC As Long
    
    If TransparentColor = -1 Then
        TransparentColor = GetPixel(Source.hdc, 1, 1)
    End If
    
    MonoMaskDC = CreateCompatibleDC(hdc)
    MonoInvDC = CreateCompatibleDC(hdc)
    hMonoMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    hMonoInv = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
    hPrevInv = SelectObject(MonoInvDC, hMonoInv)
    
    ResultDstDC = CreateCompatibleDC(hdc)
    ResultSrcDC = CreateCompatibleDC(hdc)
    hResultDst = CreateCompatibleBitmap(hdc, Width, Height)
    hResultSrc = CreateCompatibleBitmap(hdc, Width, Height)
    hPrevDst = SelectObject(ResultDstDC, hResultDst)
    hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)
    
    OldBC = SetBkColor(Source.hdc, TransparentColor)
    Call BitBlt(MonoMaskDC, 0, 0, Width, Height, Source.hdc, SrcPoint.X, SrcPoint.Y, SRCCOPY)
    TransparentColor = SetBkColor(Source.hdc, OldBC)
    
    Call BitBlt(MonoInvDC, 0, 0, Width, Height, MonoMaskDC, 0, 0, NOTSRCCOPY)
    
    Call BitBlt(ResultDstDC, 0, 0, Width, Height, hdc, DestPoint.X, DestPoint.Y, SRCCOPY)
    
    Call BitBlt(ResultDstDC, 0, 0, Width, Height, MonoMaskDC, 0, 0, SRCAND)
    
    Call BitBlt(ResultSrcDC, 0, 0, Width, Height, Source.hdc, SrcPoint.X, SrcPoint.Y, SRCCOPY)
    
    Call BitBlt(ResultSrcDC, 0, 0, Width, Height, MonoInvDC, 0, 0, SRCAND)
    
    Call BitBlt(ResultDstDC, 0, 0, Width, Height, ResultSrcDC, 0, 0, SRCINVERT)
    
    Call BitBlt(hdc, DestPoint.X, DestPoint.Y, Width, Height, ResultDstDC, 0, 0, SRCCOPY)
    
    hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
    DeleteObject hMonoMask
    
    hMonoInv = SelectObject(MonoInvDC, hPrevInv)
    DeleteObject hMonoInv
    
    hResultDst = SelectObject(ResultDstDC, hPrevDst)
    DeleteObject hResultDst
    
    hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
    DeleteObject hResultSrc
    
    DeleteDC MonoMaskDC
    DeleteDC MonoInvDC
    DeleteDC ResultDstDC
    DeleteDC ResultSrcDC
End Sub

Private Function BitBltEx(ByVal Source As Object, ByVal Destination As Object, ByVal Operation As RasterOperationConstants, Optional ByVal xDest As Long = 0, Optional ByVal yDest As Long = 0, Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional ByVal Refresh As Boolean = False) As Boolean
    Dim lReturn As Long
    
    If Width = -1 Then
        Width = Source.Width \ Screen.TwipsPerPixelX
    End If
    If Height = -1 Then
        Height = Source.Height \ Screen.TwipsPerPixelX
    End If
    
    lReturn = BitBlt(Destination.hdc, xDest, yDest, Width, Height, Source.hdc, xSrc, ySrc, Operation)
    
    If Refresh Then
        Destination.Refresh
    End If
    
    If lReturn = 0 Then
        BitBltEx = False
    Else
        BitBltEx = True
    End If
End Function

Private Function MaskBltEx(ByVal Source As Object, ByVal Destination As Object, Optional ByVal MaskColor As OLE_COLOR = -1, Optional ByVal xDest As Long = 0, Optional ByVal yDest As Long = 0, Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional ByVal Refresh As Boolean = False) As Boolean
    Dim MonoMaskDC As Long
    Dim hMonoMask As Long
    Dim MonoInvDC As Long
    Dim hMonoInv As Long
    Dim ResultDstDC As Long
    Dim hResultDst As Long
    Dim ResultSrcDC As Long
    Dim hResultSrc As Long
    Dim hPrevMask As Long
    Dim hPrevInv As Long
    Dim hPrevSrc As Long
    Dim hPrevDst As Long
    Dim OldBC As Long
    Dim lReturn As Long
    
    If Width = -1 Then
        Width = Source.Width \ Screen.TwipsPerPixelX
    End If
    If Height = -1 Then
        Height = Source.Height \ Screen.TwipsPerPixelX
    End If
    
    If MaskColor = -1 Then
        MaskColor = GetPixel(Source.hdc, 0, 0)
    End If
    
    MonoMaskDC = CreateCompatibleDC(Destination.hdc)
    hMonoMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
    
    OldBC = SetBkColor(Source.hdc, MaskColor)
    lReturn = BitBlt(MonoMaskDC, 0, 0, Width, Height, Source.hdc, xSrc, ySrc, SRCCOPY)
    If lReturn <> 0 Then
        MaskColor = SetBkColor(Source.hdc, OldBC)
        
        lReturn = BitBlt(Destination.hdc, xDest, yDest, Width, Height, MonoMaskDC, 0, 0, SRCCOPY)
    End If
    
    hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
    DeleteObject hMonoMask
    DeleteDC MonoMaskDC

    If Refresh Then
         Destination.Refresh
    End If
    
    If lReturn = 0 Then
        MaskBltEx = False
    Else
        MaskBltEx = True
    End If
End Function

Private Function TransparentBltEx(ByVal Source As Object, ByVal Destination, Optional ByVal TransparentColor As OLE_COLOR = -1, Optional ByVal xDest As Long = 0, Optional ByVal yDest As Long = 0, Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional ByVal Refresh As Boolean = False) As Boolean
    Dim MonoMaskDC As Long
    Dim hMonoMask As Long
    Dim MonoInvDC As Long
    Dim hMonoInv As Long
    Dim ResultDstDC As Long
    Dim hResultDst As Long
    Dim ResultSrcDC As Long
    Dim hResultSrc As Long
    Dim hPrevMask As Long
    Dim hPrevInv As Long
    Dim hPrevSrc As Long
    Dim hPrevDst As Long
    Dim OldBC As Long
    Dim lReturn As Long
    
    If Width = -1 Then
        Width = Source.Width \ Screen.TwipsPerPixelX
    End If
    If Height = -1 Then
        Height = Source.Height \ Screen.TwipsPerPixelX
    End If
    
    If TransparentColor = -1 Then
        TransparentColor = GetPixel(Source.hdc, 0, 0)
    End If
    
    MonoMaskDC = CreateCompatibleDC(Destination.hdc)
    MonoInvDC = CreateCompatibleDC(Destination.hdc)
    hMonoMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    hMonoInv = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
    hPrevInv = SelectObject(MonoInvDC, hMonoInv)
    
    ResultDstDC = CreateCompatibleDC(Destination.hdc)
    ResultSrcDC = CreateCompatibleDC(Destination.hdc)
    hResultDst = CreateCompatibleBitmap(Destination.hdc, Width, Height)
    hResultSrc = CreateCompatibleBitmap(Destination.hdc, Width, Height)
    hPrevDst = SelectObject(ResultDstDC, hResultDst)
    hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)
    
    OldBC = SetBkColor(Source.hdc, TransparentColor)
    lReturn = BitBlt(MonoMaskDC, 0, 0, Width, Height, Source.hdc, xSrc, ySrc, SRCCOPY)
    If lReturn <> 0 Then
        TransparentColor = SetBkColor(Source.hdc, OldBC)
        
        lReturn = BitBlt(MonoInvDC, 0, 0, Width, Height, MonoMaskDC, 0, 0, NOTSRCCOPY)
        If lReturn <> 0 Then
                lReturn = BitBlt(ResultDstDC, 0, 0, Width, Height, Destination.hdc, xDest, yDest, SRCCOPY)
            If lReturn <> 0 Then
                lReturn = BitBlt(ResultDstDC, 0, 0, Width, Height, MonoMaskDC, 0, 0, SRCAND)
                If lReturn <> 0 Then
                    lReturn = BitBlt(ResultSrcDC, 0, 0, Width, Height, Source.hdc, xSrc, ySrc, SRCCOPY)
                    If lReturn <> 0 Then
                        lReturn = BitBlt(ResultSrcDC, 0, 0, Width, Height, MonoInvDC, 0, 0, SRCAND)
                        If lReturn <> 0 Then
                            lReturn = BitBlt(ResultDstDC, 0, 0, Width, Height, ResultSrcDC, 0, 0, SRCINVERT)
                            If lReturn <> 0 Then
                                lReturn = BitBlt(Destination.hdc, xDest, yDest, Width, Height, ResultDstDC, 0, 0, SRCCOPY)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
    DeleteObject hMonoMask
    
    hMonoInv = SelectObject(MonoInvDC, hPrevInv)
    DeleteObject hMonoInv
    
    hResultDst = SelectObject(ResultDstDC, hPrevDst)
    DeleteObject hResultDst
    
    hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
    DeleteObject hResultSrc
    
    DeleteDC MonoMaskDC
    DeleteDC MonoInvDC
    DeleteDC ResultDstDC
    DeleteDC ResultSrcDC

    If Refresh Then
        Destination.Refresh
    End If
    
    If lReturn = 0 Then
        TransparentBltEx = False
    Else
        TransparentBltEx = True
    End If
End Function

Private Function HighlightBltEx(ByVal Source As Object, ByVal Destination, ByVal TempDestination As Object, ByVal Highlight As Object, ByVal HighlightColor As OLE_COLOR, Optional ByVal xDest As Long = 0, Optional ByVal yDest As Long = 0, Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional ByVal Refresh As Boolean = False) As Boolean
    Highlight.BackColor = HighlightColor
    
    Call MaskBltEx(Source, TempDestination, -1, 0, 0, xSrc, ySrc, Width, Height)
    Call BitBltEx(TempDestination, Highlight, roSrcInvert, 0, 0, 0, 0, Width, Height)
    Call TransparentBltEx(Highlight, Destination, -1, xDest, yDest, 0, 0, Width, Height, Refresh)
End Function

Private Function RaiseEventEx(ByVal Name As String, ParamArray Params() As Variant)
        
    Select Case Name
        Case "Click"
            RaiseEvent Click
        
        Case "KeyDown"
            RaiseEvent KeyDown(CInt(Params(0)), CInt(Params(1)))
        
        Case "KeyPress"
            RaiseEvent KeyPress(CInt(Params(0)))
        
        Case "KeyUp"
            RaiseEvent KeyUp(CInt(Params(0)), CInt(Params(1)))
        
        Case "MouseDown"
            RaiseEvent MouseDown(CInt(Params(0)), CInt(Params(1)), CSng(Params(2)), CSng(Params(3)))
        
        Case "MouseMove"
            RaiseEvent MouseMove(CInt(Params(0)), CInt(Params(1)), CSng(Params(2)), CSng(Params(3)))
        
        Case "MouseUp"
            RaiseEvent MouseUp(CInt(Params(0)), CInt(Params(1)), CSng(Params(2)), CSng(Params(3)))
        
        Case "MouseExit"
            If tPrevEvent <> "MouseExit" Then
                RaiseEvent MouseExit
            End If
    
            tPrevEvent = Name
        
        Case "MouseEnter"
            If tPrevEvent <> "MouseEnter" Then
                RaiseEvent MouseEnter
            End If
    
            tPrevEvent = Name
        
        Case "Resize"
            RaiseEvent Resize
    End Select
End Function

Private Sub DrawButton(ByVal State As StateConstants)
    Dim bFocus As Boolean
    Dim bUserMode As Boolean
    
    bFocus = bHasFocus
    bUserMode = False
    Set UserControl.Picture = Nothing
    Set UserControl.MaskPicture = Nothing
    
    UserControl.Cls
    
    On Local Error Resume Next
    bUserMode = UserControl.Ambient.UserMode
    On Local Error GoTo 0
    
    If m_Style = ButtonGroup Then
        If m_Value = Down Then
            State = btDown
        Else
            If State <> btOver Then
                State = btUp
            End If
        End If
    End If
    
    If m_Appearance = Skin And Not (m_SkinUp Is Nothing) Then
        Call DrawSkin(State, bFocus And bUserMode)
    Else
        Call DrawStandard(State, bFocus And bUserMode)
    End If
    
    Call DrawPicture(State)
    Call DrawCaption(State)
End Sub

Private Sub DrawStandard(ByVal State As StateConstants, ByVal WithFocus As Boolean)
    Dim rct As RECT
    Dim lPrevColor As OLE_COLOR
    
    UserControl.BackStyle = 1
    
    With rct
        .Left = 0
        .Top = 0
        .Bottom = UserControl.ScaleHeight
        .Right = UserControl.ScaleWidth
    End With
    
    Select Case State
        Case btUp
            If m_Appearance = [3D] Then
                If WithFocus Then
                    Call DrawEdge(UserControl.hdc, rct, BDR_OUTER, BF_RECT Or BF_ADJUST Or BF_MONO)
                    Call DrawEdge(UserControl.hdc, rct, EDGE_RAISED, BF_RECT)
                Else
                    Call DrawEdge(UserControl.hdc, rct, EDGE_RAISED, BF_RECT)
                End If
            Else
                WithFocus = False
            End If
        
        Case btOver
            If WithFocus Then
                Call DrawEdge(UserControl.hdc, rct, BDR_OUTER, BF_RECT Or BF_ADJUST Or BF_MONO)
                Call DrawEdge(UserControl.hdc, rct, EDGE_RAISED, BF_RECT)
            Else
                Call DrawEdge(UserControl.hdc, rct, EDGE_RAISED, BF_RECT)
            End If
        
        Case btDown
            If WithFocus Then
                Call DrawEdge(UserControl.hdc, rct, BDR_OUTER, BF_RECT Or BF_ADJUST Or BF_MONO)
                Call DrawEdge(UserControl.hdc, rct, BDR_SUNKENOUTER, BF_RECT Or BF_FLAT)
            Else
                Call DrawEdge(UserControl.hdc, rct, EDGE_SUNKEN, BF_RECT)
            End If
    End Select

    If WithFocus Then
        With rct
            .Left = clFocusOffset
            .Top = clFocusOffset
            .Bottom = UserControl.ScaleHeight - clFocusOffset
            .Right = UserControl.ScaleWidth - clFocusOffset
        End With
        
        lPrevColor = UserControl.ForeColor
        UserControl.ForeColor = vbBlack
        Call DrawFocusRect(UserControl.hdc, rct)
        UserControl.ForeColor = lPrevColor
    End If
    
    lState = State
End Sub

Private Sub DrawSkin(ByVal State As StateConstants, ByVal WithFocus As Boolean)
    
    lState = State
    
    If Not m_Enabled Then
        State = btDisabled
        lState = State
    ElseIf WithFocus And State = btUp Then
        State = btFocus
    End If

    UserControl.BackStyle = 0
    Set UserControl.Picture = m_SkinUp
    
    Select Case State
        Case btDisabled
            If Not (m_SkinDisabled Is Nothing) Then
                Set UserControl.Picture = m_SkinDisabled
            End If
        
        Case btDown
            If Not (m_SkinDown Is Nothing) Then
                Set UserControl.Picture = m_SkinDown
            End If
        
        Case btUp
            Set UserControl.Picture = m_SkinUp
        
        Case btOver
            If Not (m_SkinOver Is Nothing) Then
                Set UserControl.Picture = m_SkinOver
            End If
        
        Case btFocus
            If Not (m_SkinFocus Is Nothing) Then
                Set UserControl.Picture = m_SkinFocus
            End If
    End Select
    
    If UserControl.Picture <> 0 Then
        Set UserControl.MaskPicture = UserControl.Picture
    
        UserControl.Width = UserControl.Picture.Width / 1.76
        UserControl.Height = UserControl.Picture.Height / 1.76
    End If
End Sub

Private Sub DrawCaption(ByVal State As StateConstants)
    Dim lFormat As Long
    Dim lLeft As Long
    Dim lTop As Long
    
    UserControl.ForeColor = m_ForeColor
    
    Select Case State
        Case btOver
            UserControl.ForeColor = m_HighlightColor
        
        Case btDown
            If tPrevEvent <> "MouseExit" Then
                UserControl.ForeColor = m_HighlightColor
            End If
    End Select
    
    If State = btDown And Not (m_Picture Is Nothing) Then
        lLeft = -1
    Else
        lLeft = 0
    End If
    lTop = -1
    
    If imgPicture.Picture <> 0 Then
        lLeft = lLeft + imgPicture.Left + imgPicture.Width
        lLeft = (((UserControl.ScaleWidth + lLeft) \ 2) - (UserControl.TextWidth(m_Caption) \ 2))
    Else
        lLeft = lLeft + ((UserControl.ScaleWidth \ 2) - (UserControl.TextWidth(m_Caption) \ 2))
    End If
    
    lTop = lTop + ((UserControl.ScaleHeight \ 2) - (UserControl.TextHeight(m_Caption) \ 2))
    
    If State = btDown Then
        lLeft = lLeft + clDownOffset
        lTop = lTop + clDownOffset
    End If
    
    lFormat = DST_PREFIXTEXT Or DSS_NORMAL
    If Not m_Enabled Then
        lFormat = lFormat Or DSS_DISABLED
    End If
    If m_RightToLeft Then
        lFormat = lFormat Or DSS_RIGHT
    End If
    
    Call DrawStateText(UserControl.hdc, 0, 0, m_Caption, Len(m_Caption), lLeft + m_CaptionOffsetX, lTop + m_CaptionOffsetY + clDownOffset, 0, 0, lFormat)
End Sub

Private Sub DrawPicture(ByVal State As StateConstants)
    Dim lLeft As Long
    Dim lTop As Long
    Dim ptDest As POINTAPI
    Dim ptSrc As POINTAPI
    
    Set imgPicture.Picture = m_Picture
    
    Select Case State
        Case btDisabled
            If Not (m_PictureDisabled Is Nothing) Then
                Set imgPicture.Picture = m_PictureDisabled
            End If
        
        Case btDown
            If Not (m_PictureDown Is Nothing) Then
                Set imgPicture.Picture = m_PictureDown
            End If
        
        Case btUp
            Set imgPicture.Picture = m_Picture
        
        Case btOver
            If Not (m_PictureOver Is Nothing) Then
                Set imgPicture.Picture = m_PictureOver
            End If

        Case btFocus
            If Not (m_PictureFocus Is Nothing) Then
                Set imgPicture.Picture = m_PictureFocus
            End If
    End Select
    
    With imgPicture
        If .Picture <> 0 Then
            If m_Appearance = Skin Then
                lLeft = 0
                lTop = (UserControl.ScaleHeight \ 2) - (.Height \ 2)
                If lTop < 0 Then
                    lTop = 0
                End If
            Else
                lLeft = clLeft
                lTop = (UserControl.ScaleHeight \ 2) - (.Height \ 2)
                If lTop < clTop Then
                    lTop = clTop
                End If
            End If
            
            If State = btDown Then
                lLeft = lLeft + clDownOffset
                lTop = lTop + clDownOffset
            End If
        
            lLeft = lLeft + m_PictureOffsetX
            lTop = lTop + m_PictureOffsetY
            
            If .Left <> lLeft Then
                .Left = lLeft
            End If
            If .Top <> lTop Then
                .Top = lTop
            End If
        
            ptDest.X = .Left
            ptDest.Y = .Top
            ptSrc.X = 0
            ptSrc.Y = 0
            
            If (State = btDown Or State = btOver Or (Not m_Enabled And State = btUp)) And m_HighlightPicture = True Then
                If m_Enabled Then
                    Call HighlightBltEx(imgPicture, UserControl, pictTempDestination, pictTempHighlight, m_HighlightColor, .Left, .Top, 0, 0, .Width, .Height)
                Else
                    Call HighlightBltEx(imgPicture, UserControl, pictTempDestination, pictTempHighlight, vbGrayText, .Left, .Top, 0, 0, .Width, .Height)
                End If
            Else
                Call TransparentBlt_New2(UserControl.hdc, imgPicture, ptDest, ptSrc, imgPicture.Width, imgPicture.Height, m_TransparentColor)
            End If
        End If
    End With
End Sub

