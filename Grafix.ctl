VERSION 5.00
Begin VB.UserControl GraFxConsole 
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   4755
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   4755
   ScaleWidth      =   4980
   ToolboxBitmap   =   "Grafix.ctx":0000
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   2880
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3240
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   2880
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   12
      Top             =   4080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3120
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   2520
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2880
      ScaleHeight     =   345
      ScaleWidth      =   105
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2400
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3120
      ScaleHeight     =   105
      ScaleWidth      =   345
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   2400
      ScaleHeight     =   105
      ScaleWidth      =   345
      TabIndex        =   6
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3120
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2880
      ScaleHeight     =   345
      ScaleWidth      =   105
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2400
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1710
      Left            =   0
      Picture         =   "Grafix.ctx":0312
      ScaleHeight     =   1680
      ScaleWidth      =   2100
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1710
      Left            =   0
      ScaleHeight     =   1680
      ScaleWidth      =   2100
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Image btn4 
      Height          =   135
      Left            =   1680
      Top             =   1000
      Width           =   135
   End
   Begin VB.Image Image14 
      Height          =   135
      Left            =   4200
      Top             =   4320
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image13 
      Height          =   150
      Left            =   2520
      Picture         =   "Grafix.ctx":BB14
      Top             =   4560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image Image12 
      Height          =   195
      Left            =   30
      Stretch         =   -1  'True
      Top             =   75
      Width           =   195
   End
   Begin VB.Image Image11 
      Height          =   135
      Left            =   4560
      Top             =   4080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image10 
      Height          =   135
      Left            =   4200
      Top             =   4080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image btn1 
      Height          =   150
      Left            =   1320
      ToolTipText     =   "Minimize"
      Top             =   1680
      Width           =   150
   End
   Begin VB.Image Image9 
      Appearance      =   0  'Flat
      Height          =   135
      Left            =   3840
      Top             =   4080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image8 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4440
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image7 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4200
      Top             =   3600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image6 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3720
      Top             =   3600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image5 
      Appearance      =   0  'Flat
      Height          =   135
      Left            =   4440
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      Height          =   135
      Left            =   3720
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4440
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4200
      Top             =   2880
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3720
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image btn3 
      Height          =   150
      Left            =   2040
      ToolTipText     =   "Close"
      Top             =   1680
      Width           =   150
   End
   Begin VB.Image btn2 
      Height          =   150
      Left            =   1680
      ToolTipText     =   "Restore/Maximize"
      Top             =   1680
      Width           =   150
   End
   Begin VB.Image R 
      Appearance      =   0  'Flat
      Height          =   1260
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   600
      Width           =   90
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   3000
      Top             =   600
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GraFx Console 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000010&
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   75
      Width           =   1500
   End
   Begin VB.Image br 
      Height          =   330
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   90
   End
   Begin VB.Image b 
      Height          =   90
      Left            =   120
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Image tr 
      Height          =   540
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   0
      Width           =   90
   End
   Begin VB.Image t 
      Height          =   285
      Left            =   120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3000
   End
   Begin VB.Image tl 
      Height          =   615
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   90
   End
   Begin VB.Image l 
      Height          =   1200
      Left            =   0
      Stretch         =   -1  'True
      Top             =   720
      Width           =   90
   End
   Begin VB.Image bl 
      Height          =   315
      Left            =   0
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   90
   End
End
Attribute VB_Name = "GraFxConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    Option Explicit
    
    Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
   Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
   Private Declare Function GetActiveWindow Lib "user32" () As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
    Private Declare Sub ReleaseCapture Lib "user32" ()

'Consts:
    Private Const WM_RBUTTONUP = &H205
    Private Const WM_MOUSEMOVE = &H200
    Private Const WM_LBUTTONDOWN = &H201
    Private Const WM_LBUTTONUP = &H202
    Private Const SRCCOPY = &HCC0020
    Private Const SWP_NOMOVE = &H2
    Private Const SWP_NOSIZE = &H1
    Private Const HWND_TOPMOST = -1
    Private Const HWND_NOTOPMOST = -2
    Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Private Const NIM_ADD = &H0
    Private Const NIM_DELETE = &H2
    Private Const NIF_ICON = &H2
    Private Const NIF_MESSAGE = &H1
    Private Const NIM_MODIFY = &H1
    Private Const NIF_TIP = &H4
    Private Const MAX_TOOLTIP As Integer = 64
    Private nfIconData As NOTIFYICONDATA
    Private Type NOTIFYICONDATA
    cbSize           As Long
    hWnd             As Long
    uId              As Long
    uFlags           As Long
    ucallbackMessage As Long
    hIcon            As Long
    szTip            As String * MAX_TOOLTIP
    End Type
    
'Default Property Values:
Const m_def_Movable = 1
Const m_def_ResizeAtRuntime = 0
'Const m_def_Icon = ""
Const m_def_ImageFile = ""
Const m_def_Btn_Max = 1
Const m_def_Btn_Min = 1
Const m_def_Btn_Close = 1
Const m_def_BackStyle = 0
Const m_def_BackColor = 0
'Const m_def_AutoLoad = 1
Const m_def_TitleColor = 0
Const m_def_MenuColor = 0
Const m_def_ShowMenu = False
Const m_def_HandleWidth = 200
Const m_def_MousePointer = 0
'Const m_def_OLEDragMode = 0
'Const m_def_OLEDropMode = 0
Const m_def_Caption = "GraFx Console 3"
'Property Variables:
Dim m_Icon As Picture
Dim m_Movable As Boolean
Dim m_ResizeAtRuntime As Boolean
'Dim m_Icon As String
Dim m_ImageFile As String
Dim m_Btn_Max As Boolean
Dim m_Btn_Min As Boolean
Dim m_Btn_Close As Boolean
Dim m_BackStyle As Integer
Dim m_BackColor As OLE_COLOR
'Dim m_AutoLoad As Boolean
'Dim m_INIFilename As String
Dim m_TitleColor As OLE_COLOR
Dim m_MenuColor As OLE_COLOR
Dim m_TitleFont As Font
Dim m_ShowMenu As Boolean

Dim m_HandleWidth As Integer
Dim m_MousePointer As Integer
'Dim m_OLEDragMode As Integer
'Dim m_OLEDropMode As Integer
Dim m_Caption As String
'Event Declarations:
Event RestoreClick() 'MappingInfo=btn4,btn4,-1,Click
Attribute RestoreClick.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event TitlebarDblClick() 'MappingInfo=t,t,-1,DblClick
Attribute TitlebarDblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event TitleBarMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=t,t,-1,MouseUp
Attribute TitleBarMouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event TitleBarMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=t,t,-1,MouseDown
Attribute TitleBarMouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event TitlebarClick() 'MappingInfo=t,t,-1,Click
Attribute TitlebarClick.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MinClick() 'MappingInfo=btn1,btn1,-1,Click
Attribute MinClick.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event CloseClick() 'MappingInfo=btn3,btn3,-1,Click
Attribute CloseClick.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MaxClick() 'MappingInfo=btn2,btn2,-1,Click
Attribute MaxClick.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event Resize()
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Event OLECompleteDrag(Effect As Long)
'Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
'Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
'Event OLESetData(Data As DataObject, DataFormat As Integer)
'Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

Private Sub br_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ResizeAtRuntime = False Then
'dont even bother
Else
    Dim nParam As Long
    
    With br
        '  You can change these coordinates
        'If (X > 0 And X < 100) Then
        '    nParam = HTLEFT
        'ElseIf (X > UserControl.Width - 100 And X < UserControl.Width) Then
            nParam = HTRIGHT
        'End If
        If nParam Then
            Call ReleaseCapture
            Call SendMessage(UserControl.hWnd, WM_NCLBUTTONDOWN, nParam, 0)
            Call SendMessage(UserControl.hWnd, WM_NCLBUTTONDOWN, 17, 0)
            UserControl.Refresh
            ManualRefresh
        End If
    End With
End If
End Sub

Private Sub br_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ResizeAtRuntime = False Then
'dont bother
Else
    Dim NewPointer As MousePointerConstants
    
    If (X > 0 And X < 100) Then
        NewPointer = vbSizeNWSE
    ElseIf (X > br.Width - 100 And X < br.Width) Then
        NewPointer = vbSizeNWSE
    Else
        NewPointer = vbDefault
    End If
    
    If NewPointer <> br.MousePointer Then
        br.MousePointer = NewPointer
    End If
End If
End Sub

Private Sub btn4_Click()
RaiseEvent RestoreClick
Btn4c
End Sub


Private Sub Image12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Movable = True Then
On Error Resume Next
MoveForm Parent
End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Movable = True Then
On Error Resume Next
MoveForm Parent
End If
End Sub


Private Sub tl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Movable = True Then
On Error Resume Next
MoveForm Parent
End If
End Sub

Private Sub tr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Movable = True Then
On Error Resume Next
MoveForm Parent
End If
End Sub

Private Sub UserControl_Hide()
t.Visible = False
tl.Visible = False
tr.Visible = False
l.Visible = False
R.Visible = False
br.Visible = False
bl.Visible = False
b.Visible = False
Label1.Visible = False
End Sub

Private Sub UserControl_Initialize()
On Error GoTo LoadingError
'start up code
Exit Sub
LoadingError:
ErrorCall "LoadingError"
End Sub

Private Sub UserControl_Resize()
On Error GoTo ResizeError
'sets all variables
'-fonts
Label1.Caption = Caption
Label1.ForeColor = TitleColor
Label1.Font.Name = TitleFont.Name
Label1.Font.Size = TitleFont.Size
Label1.Font.Bold = TitleFont.Bold
Label1.Font.Italic = TitleFont.Italic
Label1.Font.Strikethrough = TitleFont.Strikethrough
Label1.Font.Underline = TitleFont.Underline
'-menu
If ShowMenu = True Then Shape1.Visible = True
If ShowMenu = False Then Shape1.Visible = False
Shape1.FillColor = MenuColor
'-graphics
ChangeImage
'backstyle
If BackStyle = 0 Then
UserControl.BackStyle = 0
Else
    If BackStyle = 1 Then
    UserControl.BackStyle = 1
    Else
    BackStyle = 0
    End If
End If
'backcolor
UserControl.BackColor = BackColor
tl.Left = 0
tl.Top = 0
bl.Top = UserControl.Height - bl.Height
bl.Left = 0
l.Left = 0
l.Top = tl.Height
l.Height = UserControl.Height - tl.Height - bl.Height
tr.Top = 0
tr.Left = UserControl.Width - tr.Width
t.Top = 0
t.Left = tl.Width
t.Width = UserControl.Width - tl.Width - tr.Width
br.Top = UserControl.Height - br.Height
br.Left = UserControl.Width - br.Width
b.Left = bl.Width
b.Top = UserControl.Height - b.Height
b.Width = UserControl.Width - bl.Width - br.Width
R.Left = UserControl.Width - R.Width
R.Top = tr.Height
R.Height = UserControl.Height - tr.Height - br.Height
Shape1.Left = R.Left - Shape1.Width
Shape1.Top = R.Top
Shape1.Height = l.Height
btn1.Top = 45
btn3.Top = 45
btn1.Left = UserControl.Width - 900
btn2.Left = UserControl.Width - 600
btn3.Left = UserControl.Width - 300
btn4.Left = UserControl.Width - 600
'hnd.Left = UserControl.Width - hnd.Width
'hnd.Top = UserControl.Height - hnd.Height
Exit Sub
ResizeError:
ManualRefresh
ErrorCall "ResizeError"
Shape1.Left = R.Left - Shape1.Width
Shape1.Top = R.Top
Shape1.Height = l.Height
End Sub

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    m_MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
End Property
'
'Public Sub OLEDrag()
'
'End Sub
'
'Public Property Get OLEDragMode() As Integer
'    OLEDragMode = m_OLEDragMode
'End Property
'
'Public Property Let OLEDragMode(ByVal New_OLEDragMode As Integer)
'    m_OLEDragMode = New_OLEDragMode
'    PropertyChanged "OLEDragMode"
'End Property
'
'Public Property Get OLEDropMode() As Integer
'    OLEDropMode = m_OLEDropMode
'End Property
'
'Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
'    m_OLEDropMode = New_OLEDropMode
'    PropertyChanged "OLEDropMode"
'End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets the Controls Title\r\n"
    Caption = m_Caption
    Label1.Caption = Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    Label1.Caption = Caption
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_MousePointer = m_def_MousePointer
'    m_OLEDragMode = m_def_OLEDragMode
'    m_OLEDropMode = m_def_OLEDropMode
    m_Caption = m_def_Caption
    m_HandleWidth = m_def_HandleWidth
    m_ShowMenu = m_def_ShowMenu
    Set m_TitleFont = Ambient.Font
    m_TitleColor = m_def_TitleColor
    m_MenuColor = m_def_MenuColor
'    m_INIFilename = m_def_INIFilename
'    m_AutoLoad = m_def_AutoLoad
    m_BackStyle = m_def_BackStyle
    m_BackColor = m_def_BackColor
    m_Btn_Max = m_def_Btn_Max
    m_Btn_Min = m_def_Btn_Min
    m_Btn_Close = m_def_Btn_Close
    m_ImageFile = m_def_ImageFile
'    m_Icon = m_def_Icon
    m_ResizeAtRuntime = m_def_ResizeAtRuntime
    m_Movable = m_def_Movable
    Set m_Icon = LoadPicture("")
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_MousePointer = PropBag.ReadProperty("MousePointer", m_def_MousePointer)
'    m_OLEDragMode = PropBag.ReadProperty("OLEDragMode", m_def_OLEDragMode)
'    m_OLEDropMode = PropBag.ReadProperty("OLEDropMode", m_def_OLEDropMode)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_HandleWidth = PropBag.ReadProperty("HandleWidth", m_def_HandleWidth)
    m_ShowMenu = PropBag.ReadProperty("ShowMenu", m_def_ShowMenu)
    Set m_TitleFont = PropBag.ReadProperty("TitleFont", Ambient.Font)
    m_TitleColor = PropBag.ReadProperty("TitleColor", m_def_TitleColor)
    m_MenuColor = PropBag.ReadProperty("MenuColor", m_def_MenuColor)
'    m_INIFilename = PropBag.ReadProperty("INIFilename", m_def_INIFilename)
'    m_AutoLoad = PropBag.ReadProperty("AutoLoad", m_def_AutoLoad)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_Btn_Max = PropBag.ReadProperty("Btn_Max", m_def_Btn_Max)
    m_Btn_Min = PropBag.ReadProperty("Btn_Min", m_def_Btn_Min)
    m_Btn_Close = PropBag.ReadProperty("Btn_Close", m_def_Btn_Close)
    m_ImageFile = PropBag.ReadProperty("ImageFile", m_def_ImageFile)
'    m_Icon = PropBag.ReadProperty("Icon", m_def_Icon)
    m_ResizeAtRuntime = PropBag.ReadProperty("ResizeAtRuntime", m_def_ResizeAtRuntime)
    m_Movable = PropBag.ReadProperty("Movable", m_def_Movable)
    Set m_Icon = PropBag.ReadProperty("Icon", Nothing)
End Sub

Private Sub UserControl_Show()
On Error GoTo ShowError
t.Visible = True
tl.Visible = True
tr.Visible = True
l.Visible = True
R.Visible = True
br.Visible = True
bl.Visible = True
b.Visible = True
Label1.Visible = True
ChangeImage
Label1.Caption = Caption
Label1.ForeColor = TitleColor
Label1.Font.Name = TitleFont.Name
Label1.Font.Size = TitleFont.Size
Label1.Font.Bold = TitleFont.Bold
Label1.Font.Italic = TitleFont.Italic
Label1.Font.Strikethrough = TitleFont.Strikethrough
Label1.Font.Underline = TitleFont.Underline
'-menu
If ShowMenu = True Then Shape1.Visible = True
If ShowMenu = False Then Shape1.Visible = False
Shape1.FillColor = MenuColor
'-graphics
ChangeImage
ChangeIcon
'backstyle
If BackStyle = 0 Then
UserControl.BackStyle = 0
Else
    If BackStyle = 1 Then
    UserControl.BackStyle = 1
    Else
    BackStyle = 0
    End If
End If
'buttons
If Btn_Close = True Then
btn3.Visible = True
Else
btn3.Visible = False
End If

If Btn_Max = True Then
btn2.Visible = True
btn4.Visible = True
Else
btn2.Visible = False
btn4.Visible = False
End If
btn4.Top = -300
btn2.Top = 45
If Btn_Min = True Then
btn1.Visible = True
Else
btn1.Visible = False
End If

Exit Sub
ShowError:
ErrorCall "ShowControl"
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("MousePointer", m_MousePointer, m_def_MousePointer)
'    Call PropBag.WriteProperty("OLEDragMode", m_OLEDragMode, m_def_OLEDragMode)
'    Call PropBag.WriteProperty("OLEDropMode", m_OLEDropMode, m_def_OLEDropMode)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("HandleWidth", m_HandleWidth, m_def_HandleWidth)
    Call PropBag.WriteProperty("ShowMenu", m_ShowMenu, m_def_ShowMenu)
    Call PropBag.WriteProperty("TitleFont", m_TitleFont, Ambient.Font)
    Call PropBag.WriteProperty("TitleColor", m_TitleColor, m_def_TitleColor)
    Call PropBag.WriteProperty("MenuColor", m_MenuColor, m_def_MenuColor)
'    Call PropBag.WriteProperty("INIFilename", m_INIFilename, m_def_INIFilename)
'    Call PropBag.WriteProperty("AutoLoad", m_AutoLoad, m_def_AutoLoad)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Btn_Max", m_Btn_Max, m_def_Btn_Max)
    Call PropBag.WriteProperty("Btn_Min", m_Btn_Min, m_def_Btn_Min)
    Call PropBag.WriteProperty("Btn_Close", m_Btn_Close, m_def_Btn_Close)
    Call PropBag.WriteProperty("ImageFile", m_ImageFile, m_def_ImageFile)
'    Call PropBag.WriteProperty("Icon", m_Icon, m_def_Icon)
    Call PropBag.WriteProperty("ResizeAtRuntime", m_ResizeAtRuntime, m_def_ResizeAtRuntime)
    Call PropBag.WriteProperty("Movable", m_Movable, m_def_Movable)
    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
End Sub

Public Property Get HandleWidth() As Integer
Attribute HandleWidth.VB_Description = "Sets the Blue Handle Width"
    HandleWidth = m_HandleWidth
    Shape1.Width = HandleWidth
    Shape1.Left = R.Left - Shape1.Width
    End Property

Public Property Let HandleWidth(ByVal New_HandleWidth As Integer)
    m_HandleWidth = New_HandleWidth
    PropertyChanged "HandleWidth"
    Shape1.Width = HandleWidth
    Shape1.Left = R.Left - Shape1.Width
End Property
Private Sub REload()
On Error GoTo RELoadError:
'refreshs images from pictures
tl.Picture = Image1.Picture
t.Picture = Image2.Picture
tr.Picture = Image3.Picture
l.Picture = Image4.Picture
R.Picture = Image5.Picture
bl.Picture = Image6.Picture
b.Picture = Image7.Picture
br.Picture = Image8.Picture
btn1.Picture = Image9.Picture
btn2.Picture = Image10.Picture
btn3.Picture = Image11.Picture
btn4.Picture = Image14.Picture
tl.Stretch = False
tl.Stretch = True
t.Stretch = False
t.Stretch = True
tr.Stretch = False
tr.Stretch = True
R.Stretch = False
R.Stretch = True
l.Stretch = False
l.Stretch = True
bl.Stretch = False
bl.Stretch = True
b.Stretch = False
b.Stretch = True
br.Stretch = False
br.Stretch = True
btn1.Stretch = False
btn1.Stretch = True
btn2.Stretch = False
btn2.Stretch = True
btn3.Stretch = False
btn3.Stretch = True
btn4.Stretch = False
btn4.Stretch = True
ManualRefresh

Exit Sub
RELoadError:
ErrorCall "ReLoad"
End Sub
Private Sub ChangeImage()
On Error GoTo Error
If ImageFile = "" Then
LoadDefaultSkin
Else
    If ImageFile = "0" Then
    LoadDefaultSkin
    Else
    LoadPictures
    Exit Sub
    End If
End If
Refresh
Exit Sub
Error:
ImageFile = ""
ChangeImage
ErrorCall "ChangeImage"
Resume
End Sub

Public Property Get ShowMenu() As Boolean
Attribute ShowMenu.VB_Description = "Sets/returns whether the Menu is used"
    ShowMenu = m_ShowMenu
    If ShowMenu = False Then Shape1.Visible = False
    If ShowMenu = True Then Shape1.Visible = True
End Property

Public Property Let ShowMenu(ByVal New_ShowMenu As Boolean)
    m_ShowMenu = New_ShowMenu
    PropertyChanged "ShowMenu"
    If ShowMenu = False Then Shape1.Visible = False
    If ShowMenu = True Then Shape1.Visible = True
End Property

Public Property Get TitleFont() As Font
Attribute TitleFont.VB_Description = "Set Title font"
    Set TitleFont = m_TitleFont
    On Error Resume Next
    Label1.Font.Size = TitleFont.Size
    Label1.Font.Name = TitleFont.Name
    
    If TitleFont.Italic = True Then
    Label1.Font.Italic = True
    Else
    Label1.Font.Italic = False
    End If
    
    If TitleFont.Bold = True Then
    Label1.Font.Bold = True
    Else
    Label1.Font.Bold = False
    End If
    
    If TitleFont.Underline = True Then
    Label1.Font.Underline = True
    Else
    Label1.Font.Underline = False
    End If
End Property

Public Property Set TitleFont(ByVal New_TitleFont As Font)
    Set m_TitleFont = New_TitleFont
    PropertyChanged "TitleFont"
    On Error Resume Next
    Label1.Font.Size = TitleFont.Size
    Label1.Font.Name = TitleFont.Name
    
    If TitleFont.Italic = True Then
    Label1.Font.Italic = True
    Else
    Label1.Font.Italic = False
    End If
    
    If TitleFont.Bold = True Then
    Label1.Font.Bold = True
    Else
    Label1.Font.Bold = False
    End If
    
    If TitleFont.Underline = True Then
    Label1.Font.Underline = True
    Else
    Label1.Font.Underline = False
    End If
End Property

Public Property Get TitleColor() As OLE_COLOR
Attribute TitleColor.VB_Description = "Color of title text"
    TitleColor = m_TitleColor
    Label1.ForeColor = TitleColor
End Property

Public Property Let TitleColor(ByVal New_TitleColor As OLE_COLOR)
    m_TitleColor = New_TitleColor
    PropertyChanged "TitleColor"
    Label1.ForeColor = TitleColor
End Property

Public Property Get MenuColor() As OLE_COLOR
Attribute MenuColor.VB_Description = "Sets/Returns the Menu's Color"
    MenuColor = m_MenuColor
End Property

Public Property Let MenuColor(ByVal New_MenuColor As OLE_COLOR)
    m_MenuColor = New_MenuColor
    PropertyChanged "MenuColor"
    Shape1.FillColor = MenuColor
End Property

'Public Property Get INIFilename() As String
'    INIFilename = m_INIFilename
'End Property
'
'Public Property Let INIFilename(ByVal New_INIFilename As String)
'    m_INIFilename = New_INIFilename
'    PropertyChanged "INIFilename"
'    If FileReal(INIFilename) = True Then
'        If AutoLoad = True Then LoadSettingsINI
'    End If
'End Property
Private Sub ErrorCall(Area As String)
On Error Resume Next
Kill "C:\Windows\Profiles\samt\desktop\GrafxErrorLog.log"
Ext.WritePrivateProfile (Area & Date & " @ " & Time), "ErrorCode", Err, "GrafxErrorLog.log"
Ext.WritePrivateProfile (Area & Date & " @ " & Time), "ErrorMsg", Error(Err), "GrafxErrorLog.log"
End Sub
'
'Public Property Get AutoLoad() As Boolean
'    AutoLoad = m_AutoLoad
'End Property
'
'Public Property Let AutoLoad(ByVal New_AutoLoad As Boolean)
'    m_AutoLoad = New_AutoLoad
'    PropertyChanged "AutoLoad"
'End Property

Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
    If BackStyle = 0 Then
    UserControl.BackStyle = 0
    Else
    If BackStyle = 1 Then
    UserControl.BackStyle = 1
    Else
    BackStyle = 0
    End If
    End If
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
    If BackStyle = 0 Then
    UserControl.BackStyle = 0
    Else
        If BackStyle = 1 Then
        UserControl.BackStyle = 1
        Else
        BackStyle = 0
        End If
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
    UserControl.BackColor = BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = BackColor
End Property

Public Property Get Btn_Max() As Boolean
Attribute Btn_Max.VB_Description = "Show Maximize button"
    Btn_Max = m_Btn_Max
    If Btn_Max = True Then
    btn2.Visible = True
    btn4.Visible = True
    Else
    btn2.Visible = False
    btn4.Visible = False
    End If
End Property

Public Property Let Btn_Max(ByVal New_Btn_Max As Boolean)
    m_Btn_Max = New_Btn_Max
    PropertyChanged "Btn_Max"
    If Btn_Max = True Then
    btn2.Visible = True
    btn4.Visible = True
    Else
    btn2.Visible = False
    btn4.Visible = False
    End If
End Property

Public Property Get Btn_Min() As Boolean
Attribute Btn_Min.VB_Description = "Show minimize button"
    Btn_Min = m_Btn_Min
    If Btn_Min = True Then
    btn1.Visible = True
    Else
    btn1.Visible = False
    End If
End Property

Public Property Let Btn_Min(ByVal New_Btn_Min As Boolean)
    m_Btn_Min = New_Btn_Min
    PropertyChanged "Btn_Min"
    If Btn_Min = True Then
    btn1.Visible = True
    Else
    btn1.Visible = False
    End If
End Property

Public Property Get Btn_Close() As Boolean
Attribute Btn_Close.VB_Description = "Show close button"
    Btn_Close = m_Btn_Close
    If Btn_Close = True Then
    btn3.Visible = True
    Else
    btn3.Visible = False
    End If
End Property

Public Property Let Btn_Close(ByVal New_Btn_Close As Boolean)
    m_Btn_Close = New_Btn_Close
    PropertyChanged "Btn_Close"
    If Btn_Close = True Then
    btn3.Visible = True
    Else
    btn3.Visible = False
    End If
End Property

Private Sub btn2_Click()
RaiseEvent MaxClick
Btn2c
End Sub
Private Sub Btn2c()
btn2.Top = -300
btn4.Top = 45
End Sub
Private Sub Btn4c()
btn2.Top = 45
btn4.Top = -300
End Sub
Private Sub btn1_Click()
    RaiseEvent MinClick
End Sub

Private Sub btn3_Click()
    RaiseEvent CloseClick
End Sub
Public Function About() As String
Attribute About.VB_Description = "Show the about dialog"
Inf.Show
End Function
Private Sub ManualRefresh()
On Error GoTo ManualRefreshError
'refreshes co-ord's
tl.Left = 0
tl.Top = 0
bl.Top = UserControl.Height - bl.Height
bl.Left = 0
l.Left = 0
l.Top = tl.Height
l.Height = UserControl.Height - tl.Height - bl.Height
tr.Top = 0
tr.Left = UserControl.Width - tr.Width
t.Top = 0
t.Left = tl.Width
t.Width = UserControl.Width - tl.Width - tr.Width
br.Top = UserControl.Height - br.Height
br.Left = UserControl.Width - br.Width
b.Left = bl.Width
b.Top = UserControl.Height - b.Height
b.Width = UserControl.Width - bl.Width - br.Width
R.Left = UserControl.Width - R.Width
R.Top = tr.Height
R.Height = UserControl.Height - tr.Height - br.Height
Shape1.Left = R.Left - Shape1.Width
Shape1.Top = R.Top
Shape1.Height = l.Height
btn1.Top = 45
'btn2.Top = 45
btn3.Top = 45
'btn4.Top = 300
btn1.Left = UserControl.Width - 900
btn2.Left = UserControl.Width - 600
btn3.Left = UserControl.Width - 300
btn4.Left = UserControl.Width - 600
'hnd.Left = UserControl.Width - hnd.Width
'hnd.Top = UserControl.Height - hnd.Height
Exit Sub
ManualRefreshError:
ErrorCall "ManualRefreshError"
End Sub

'end grafxcode
'end grafxcode
'end grafxcode
'end grafxcode
'end grafxcode
'end grafxcode
'end grafxcode
'start bitbltcode
'start bitbltcode
'start bitbltcode
'start bitbltcode
'start bitbltcode
'start bitbltcode
'start bitbltcode

Private Sub ResizeImage(Image As Image, Picture As PictureBox)
'resizes images easily
On Error GoTo ResizeImageError
Image.Width = Picture.Width - 30
Image.Height = Picture.Height - 30

Exit Sub
ResizeImageError:
ErrorCall "ResizeImage"
End Sub
Private Sub Resize()
'Sets picture sizes for skin images
Picture2.Width = 712
Picture2.Height = 382

Picture3.Width = 67
Picture3.Height = 382

Picture4.Width = 1402
Picture4.Height = 382

Picture5.Width = 82
Picture5.Height = 1282

Picture6.Width = 82
Picture6.Height = 1282

Picture7.Width = 742
Picture7.Height = 82

Picture8.Width = 472
Picture8.Height = 82

Picture10.Width = 697
Picture10.Height = 82


Picture9.Width = 262
Picture9.Height = 247

Picture12.Width = 263
Picture12.Height = 247

Picture13.Width = 262
Picture13.Height = 247

Picture14.Width = 263
Picture14.Height = 247
End Sub
Private Sub BitBltC()
On Error GoTo BitBltCError
'splits picture
BitBlt Picture2.hdc, 0, 0, 45, 23, Picture1.hdc, 0, 0, SRCCOPY
BitBlt Picture3.hdc, 0, 0, 2, 23, Picture1.hdc, 46, 0, SRCCOPY
BitBlt Picture4.hdc, 0, 0, 91, 23, Picture1.hdc, 49, 0, SRCCOPY
BitBlt Picture5.hdc, 0, 0, 4, 83, Picture1.hdc, 0, 24, SRCCOPY
BitBlt Picture6.hdc, 0, 0, 4, 83, Picture1.hdc, 136, 24, SRCCOPY
BitBlt Picture7.hdc, 0, 0, 47, 4, Picture1.hdc, 0, 108, SRCCOPY
BitBlt Picture8.hdc, 0, 0, 29, 4, Picture1.hdc, 56, 108, SRCCOPY
BitBlt Picture10.hdc, 0, 0, 44, 4, Picture1.hdc, 96, 108, SRCCOPY
BitBlt Picture9.hdc, 0, 0, 15, 14, Picture1.hdc, 84, 26, SRCCOPY
'bitblt Picture12.hDC, 0, 0, 16, 14, Picture1.hDC, 99, 26, SRCCOPY
BitBlt Picture12.hdc, 0, 0, 16, 14, Picture1.hdc, 99, 40, SRCCOPY
BitBlt Picture13.hdc, 0, 0, 15, 14, Picture1.hdc, 115, 26, SRCCOPY
'bitblt Picture14.hDC, 0, 0, 16, 14, Picture1.hDC, 99, 40, SRCCOPY
BitBlt Picture14.hdc, 0, 0, 16, 14, Picture1.hdc, 99, 26, SRCCOPY
Exit Sub
BitBltCError:
ErrorCall "BitBltCError"
End Sub
Private Sub CopyImages()
'copies pictures => images
Image1.Picture = Picture2.Image
Image2.Picture = Picture3.Image
Image3.Picture = Picture4.Image
Image4.Picture = Picture5.Image
Image5.Picture = Picture6.Image
Image6.Picture = Picture7.Image
Image7.Picture = Picture8.Image
Image8.Picture = Picture10.Image
Image9.Picture = Picture9.Image
Image10.Picture = Picture12.Image
Image11.Picture = Picture13.Image
Image14.Picture = Picture14.Image
End Sub
Private Sub ResizeImages()
'resizes images to picture size
ResizeImage Image1, Picture2
ResizeImage Image2, Picture3
ResizeImage Image3, Picture4
ResizeImage Image4, Picture5
ResizeImage Image5, Picture6
ResizeImage Image6, Picture7
ResizeImage Image7, Picture8
ResizeImage Image8, Picture10
ResizeImage Image9, Picture9
End Sub

Public Sub LoadDefaultSkin()
ClearImages
Resize
BitBltC
Resize
CopyImages
ResizeImages
REload
End Sub
Private Sub LoadImage()
'loads an image into bitblt section
On Error GoTo Error:
    If FileReal(ImageFile) = True Then
    Picture11.Picture = LoadPicture(ImageFile)
    Else: ImageFile = ""
    ChangeImage
    End If
Exit Sub
Error:
ErrorCall "LoadImage"
ImageFile = ""
ChangeImage
Resume
End Sub
Private Sub LoadPictures()
'uses image to skin console
On Error GoTo Error
ClearImages
LoadImage
BitBlt Picture2.hdc, 0, 0, 45, 23, Picture11.hdc, 0, 0, SRCCOPY
BitBlt Picture3.hdc, 0, 0, 2, 23, Picture11.hdc, 46, 0, SRCCOPY
BitBlt Picture4.hdc, 0, 0, 91, 23, Picture11.hdc, 49, 0, SRCCOPY
BitBlt Picture5.hdc, 0, 0, 4, 83, Picture11.hdc, 0, 24, SRCCOPY
BitBlt Picture6.hdc, 0, 0, 4, 83, Picture11.hdc, 136, 24, SRCCOPY
BitBlt Picture7.hdc, 0, 0, 47, 4, Picture11.hdc, 0, 108, SRCCOPY
BitBlt Picture8.hdc, 0, 0, 29, 4, Picture11.hdc, 56, 108, SRCCOPY
BitBlt Picture10.hdc, 0, 0, 44, 4, Picture11.hdc, 96, 108, SRCCOPY
BitBlt Picture9.hdc, 0, 0, 15, 14, Picture11.hdc, 84, 26, SRCCOPY
BitBlt Picture12.hdc, 0, 0, 16, 14, Picture11.hdc, 99, 26, SRCCOPY
BitBlt Picture13.hdc, 0, 0, 15, 14, Picture11.hdc, 115, 26, SRCCOPY
BitBlt Picture14.hdc, 0, 0, 16, 14, Picture11.hdc, 99, 40, SRCCOPY
Resize
CopyImages
ResizeImages
REload
Exit Sub
Error:
ErrorCall "LoadPictures"
ImageFile = ""
LoadDefaultSkin
End Sub
Private Sub ClearImages()
'clears the pictures
Picture2.Cls
Picture3.Cls
Picture4.Cls
Picture5.Cls
Picture6.Cls
Picture7.Cls
Picture8.Cls
Picture9.Cls
Picture10.Cls
Picture12.Cls
Picture13.Cls
Picture14.Cls
End Sub
Public Property Get ImageFile() As String
Attribute ImageFile.VB_Description = "Skin Image"
    ImageFile = m_ImageFile
End Property

Public Property Let ImageFile(ByVal New_ImageFile As String)
    m_ImageFile = New_ImageFile
    PropertyChanged "ImageFile"
    If FileReal(ImageFile) = True Then
    ChangeImage
    Else
    LoadDefaultSkin
    End If
End Property

Private Sub t_Click()
    RaiseEvent TitlebarClick
End Sub
'Public Property Get Icon() As String
'    Icon = m_Icon
'End Property
'
'Public Property Let Icon(ByVal New_Icon As String)
'    m_Icon = New_Icon
'    PropertyChanged "Icon"
'ChangeIcon
'End Property

Private Sub ChangeIcon()
On Error GoTo LoadIconError
'    If FileReal(Icon) = True Then
'    Image12.Picture = LoadPicture(Icon)
'    Else
'    Image12.Picture = Image13.Picture
'    End If
Image12.Picture = Icon
Exit Sub
LoadIconError:
ErrorCall "LoadIconError"
'Icon = ""
Image12.Picture = Image13.Picture
End Sub

Private Sub t_DblClick()
    RaiseEvent TitlebarDblClick
End Sub

Private Sub t_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent TitleBarMouseUp(Button, Shift, X, Y)
End Sub

Private Sub t_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent TitleBarMouseDown(Button, Shift, X, Y)
If Movable = True Then
On Error Resume Next
MoveForm Parent
End If

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13
Public Function Mould() As String
Attribute Mould.VB_Description = "Uses Forms data to make Control like a replica form."
On Error GoTo Unmould
UserControl.Width = Parent.Width
UserControl.Height = Parent.Height
Me.Caption = Parent.Caption
Parent.BorderStyle = 0
Exit Function
Unmould:
MsgBox Error(Err), vbCritical, "Moulding Failed!"
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ResizeAtRuntime() As Boolean
Attribute ResizeAtRuntime.VB_Description = "Returns/sets whether users can change control size at Runtime"
    ResizeAtRuntime = m_ResizeAtRuntime
End Property

Public Property Let ResizeAtRuntime(ByVal New_ResizeAtRuntime As Boolean)
    m_ResizeAtRuntime = New_ResizeAtRuntime
    PropertyChanged "ResizeAtRuntime"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get Movable() As Boolean
Attribute Movable.VB_Description = "Sets whether holding the mouse in the title bitmap will enable the user to move the control."
    Movable = m_Movable
End Property

Public Property Let Movable(ByVal New_Movable As Boolean)
    m_Movable = New_Movable
    PropertyChanged "Movable"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,
Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "Set Icon for Control"
    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set m_Icon = New_Icon
    PropertyChanged "Icon"
    ChangeIcon
End Property

