VERSION 5.00
Begin VB.UserControl GraFxProgress 
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3285
   ScaleHeight     =   2550
   ScaleWidth      =   3285
   ToolboxBitmap   =   "GraFxProgress.ctx":0000
   Begin VB.Label lblstat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   120
      Picture         =   "GraFxProgress.ctx":0312
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2250
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   120
      Top             =   840
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "GraFxProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_Appearance = 1
Const m_def_FontColor = &HFFFFFF
Const m_def_ForeColor = &HFF0000
Const m_def_BackColor = 0
Const m_def_Value = "0"
Const m_def_Max = "100"
Const m_def_Min = "0"
'Const m_def_BorderStyle = 0
Const m_def_Caption = "Loading"
'Property Variables:
Dim m_Appearance As Variant
Dim m_FontColor As OLE_COLOR
Dim m_Font As Font
Dim m_ForeColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_Value As String
Dim m_Max As String
Dim m_Min As String
'Dim m_BorderStyle As Integer
Dim m_Caption As String
Dim m_Image As Picture
'Event Declarations:
Event Complete()
Attribute Complete.VB_Description = "Occurs when 100% happens"
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=7,0,0,0
'Public Property Get BorderStyle() As Integer
'    BorderStyle = m_BorderStyle
'End Property
'
'Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
'    m_BorderStyle = New_BorderStyle
'    PropertyChanged "BorderStyle"
'
'    If BorderStyle = 0 Then
'    Me.BorderStyle = 0
'    Else
'    Me.BorderStyle = 1
'    End If
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Loading
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Caption to display"
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    lblstat.Caption = Caption
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Set image if one"
    Set Image = m_Image
End Property

Public Property Set Image(ByVal New_Image As Picture)
    Set m_Image = New_Image
    PropertyChanged "Image"
    SetImage
    Rebuild
End Property
Private Sub SetImage()
On Error Resume Next
Image1.Picture = Image
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
'    m_BorderStyle = m_def_BorderStyle
    m_Caption = m_def_Caption
    'Set m_Image = LoadPicture("")
    Set m_Image = Image1.Picture
    m_Value = m_def_Value
    m_Max = m_def_Max
    m_Min = m_def_Min
    m_ForeColor = m_def_ForeColor
    m_BackColor = m_def_BackColor
    Set m_Font = Ambient.Font
    m_FontColor = m_def_FontColor
    m_Appearance = m_def_Appearance
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    Set m_Image = PropBag.ReadProperty("Image", Nothing)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_ForeColor = PropBag.ReadProperty("Forecolor", m_def_ForeColor)
    m_BackColor = PropBag.ReadProperty("Backcolor", m_def_BackColor)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_FontColor = PropBag.ReadProperty("FontColor", m_def_FontColor)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
End Sub

Private Sub UserControl_Resize()
Rebuild
End Sub


Private Sub UserControl_Show()
Rebuild
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Image", m_Image, Nothing)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Forecolor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Backcolor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("FontColor", m_FontColor, m_def_FontColor)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
End Sub
'
'Public Property Get Appearance() As AppearanceConstants
'    Appearance = m_Appearance
'End Property
'
'Public Property Let Appearance(ByVal NewValue As AppearanceConstants)
'    m_Appearance = NewValue
'
'    PropertyChanged "Appearance"
'End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Value() As String
Attribute Value.VB_Description = "Current Value"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As String)
    m_Value = New_Value
    PropertyChanged "Value"
    'Value = Val(Value)
    Rebuild
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,100
Public Property Get Max() As String
Attribute Max.VB_Description = "Maximum Value"
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As String)
    m_Max = New_Max
    PropertyChanged "Max"
    'Max = Val(Max)
    Rebuild
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Min() As String
Attribute Min.VB_Description = "Minimum Value"
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As String)
    m_Min = New_Min
    PropertyChanged "Min"
    'Min = Val(Min)
    Rebuild
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_Forecolor As OLE_COLOR)
    m_ForeColor = New_Forecolor
    PropertyChanged "Forecolor"
    Rebuild
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "Backcolor"
    Rebuild
End Property

Public Sub Rebuild()
'postion

Shape1.Top = 0
Shape2.Top = 0
Image1.Top = 0
Shape1.Left = 0
Shape2.Left = 0
Image1.Left = 0

Shape1.Width = UserControl.Width
Shape1.Height = UserControl.Height
Shape2.Height = UserControl.Height
Shape2.Width = 0
Image1.Height = UserControl.Height
Image1.Width = 0
lblstat.Top = (UserControl.Height / 2) - (lblstat.Height / 2)
lblstat.Left = (UserControl.Width / 2) - (lblstat.Width / 2)
lblstat.ForeColor = FontColor
Shape1.FillColor = BackColor
Shape2.FillColor = ForeColor
SetImage

Shape2.Width = Calc(Value)
Image1.Width = Calc(Value)
If GetPer(Value) = "100" Then RaiseEvent Complete

If Appearance = 0 Then
Shape2.Visible = True
Image1.Visible = False
Else
Shape2.Visible = False
Image1.Visible = True
End If

lblstat.Font.Bold = Font.Bold
lblstat.Font.Charset = Font.Charset
lblstat.Font.Italic = Font.Italic
lblstat.Font.Name = Font.Name
lblstat.Font.Size = Font.Size
lblstat.Font.Strikethrough = Font.Strikethrough
lblstat.Font.Underline = Font.Underline
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
    Rebuild
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,5
Public Property Get FontColor() As OLE_COLOR
Attribute FontColor.VB_Description = "Set Text Color"
    FontColor = m_FontColor
End Property

Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
    m_FontColor = New_FontColor
    PropertyChanged "FontColor"
    Rebuild
End Property
Private Function Calc(Inp As Integer)
On Error Resume Next
Dim Per As String

If Min = 0 Then
    If Inp = 0 Then
    Per = 0
    Else
    Per = Inp / Max * 100
    End If
Else
    If Inp = 0 Then
    Per = 0
    Else
    Per = Inp / Max - Min * 100
    End If
End If

Calc = UserControl.Width / 100 * Per
lblstat.Caption = Caption & " " & Per & "%"
End Function
Private Function GetPer(PerC As String) As String
On Error Resume Next
Dim PerA As String

If Min = 0 Then
    If PerC = 0 Then
    PerA = 0
    Else
    PerA = PerC / Max * 100
    End If
Else
    If PerC = 0 Then
    PerA = 0
    Else
    PerA = PerC / Max - Min * 100
    End If
End If

GetPer = PerA
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Appearance() As Variant
Attribute Appearance.VB_Description = "0=Color, 1=Image"
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Variant)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
    Rebuild
End Property
Public Function About() As String
Inf.Show
End Function
