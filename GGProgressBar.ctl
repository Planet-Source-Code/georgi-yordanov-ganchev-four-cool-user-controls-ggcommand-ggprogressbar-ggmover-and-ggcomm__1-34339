VERSION 5.00
Begin VB.UserControl GGProgressBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   ForeColor       =   &H8000000D&
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   239
   ToolboxBitmap   =   "GGProgressBar.ctx":0000
   Begin VB.PictureBox picProgress 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2340
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   975
      Begin VB.Label lblValue2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   0
         TabIndex        =   2
         Top             =   60
         Width           =   75
      End
   End
   Begin VB.Label lblValue 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1020
      TabIndex        =   0
      Top             =   60
      Width           =   75
   End
End
Attribute VB_Name = "GGProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum BorderStyleEnum
            [None]
            [Fixed Single]
End Enum

Public Enum OrientationEnum
            [Horizontal]
            [Vertical]
End Enum

Public Enum ScrollingEnum
            [Standart]
            [Smooth]
End Enum


'Default Property Values:
Const m_def_ValueSuffix = ""
Const m_def_ValuePrefix = ""
Const m_def_ShowValue = True
Const m_def_ProgressBlockWidth = 15
Const m_def_Scrolling = [Standart]
Const m_def_Min = 0
Const m_def_Max = 100
Const m_def_Value = 50
Const m_def_UseGradientColors = False
Const m_def_GradientFromColor = 65280
Const m_def_GradientToColor = 255
Const m_def_Orientation = [Horizontal]
'Property Variables:
Dim m_ToolTipText As String
Dim m_ValueSuffix As String
Dim m_ValuePrefix As String
Dim m_ShowValue As Boolean
Dim m_ProgressBlockWidth As Integer
Dim m_Scrolling As ScrollingEnum
Dim m_Min As Single
Dim m_Max As Single
Dim m_Value As Single
Dim m_UseGradientColors As Boolean
Dim m_GradientFromColor As OLE_COLOR
Dim m_GradientToColor As OLE_COLOR
Dim m_Orientation As OrientationEnum
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."

'Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

'GetSysColor - nIndex possible values
Private Const COLOR_SCROLLBAR = 0 'Scroll Bars
Private Const COLOR_BACKGROUND = 1 'Desktop
Private Const COLOR_ACTIVECAPTION = 2 'Active Title Bar
Private Const COLOR_INACTIVECAPTION = 3 'Inactive Title Bar
Private Const COLOR_MENU = 4 'Menu Bar
Private Const COLOR_WINDOW = 5 'Window Background
Private Const COLOR_WINDOWFRAME = 6 'Window Frame
Private Const COLOR_MENUTEXT = 7 'Menu Text
Private Const COLOR_WINDOWTEXT = 8 'Window Text
Private Const COLOR_CAPTIONTEXT = 9 'Active Title Bar Text
Private Const COLOR_ACTIVEBORDER = 10 'Active Border
Private Const COLOR_INACTIVEBORDER = 11 'Inactive Border
Private Const COLOR_APPWORKSPACE = 12 'Application Workspace
Private Const COLOR_HIGHLIGHT = 13 'Highlight
Private Const COLOR_HIGHLIGHTTEXT = 14 'Highlight Text
Private Const COLOR_BTNFACE = 15 'Button Face
Private Const COLOR_BTNSHADOW = 16 'Button Shadow
Private Const COLOR_GRAYTEXT = 17 'Disabled Text
Private Const COLOR_BTNTEXT = 18 'Button Text
Private Const COLOR_INACTIVECAPTIONTEXT = 19 'Inactive Title Bar Text
Private Const COLOR_BTNHIGHLIGHT = 20 'Button Highlight
Private Const COLOR_3DDKSHADOW = 21 'Button Dark Shadow
Private Const COLOR_3DLIGHT = 22 'Button Light Shadow
Private Const COLOR_INFOTEXT = 23 'ToolTip Text
Private Const COLOR_INFOBK = 24 'ToolTip
'If Window version >=5
Private Const COLOR_HOTLIGHT = 26 ' Hot Light
Private Const COLOR_GRADIENTACTIVECAPTION = 27 ' Gradient Active Title Bar
Private Const COLOR_GRADIENTINACTIVECAPTION = 28 'Gradient Inactive Title Bar

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
MsgBox "GGProgressBar user control - Georgi Yordanov Ganchev ©2000 - GogoX@Lycos.com - http://georgi-ganchev.tripod.com", vbInformation, "About GGProgressBar user control..."
End Sub

Private Sub DoScrolling(Scrolling As ScrollingEnum)
    
    If m_UseGradientColors Then Exit Sub
    
    Dim X As Integer
    
    UserControl.Cls
    
    Select Case Scrolling
    
    Case 0 'Standart
    
    If m_Orientation = Horizontal Then
    For X = 0 To UserControl.ScaleWidth Step m_ProgressBlockWidth + 2
    UserControl.Line (X + 1, 1)-(X + m_ProgressBlockWidth, UserControl.ScaleHeight - 2), , BF
    Next X
    End If
    
    If m_Orientation = Vertical Then
    For X = UserControl.ScaleHeight To 0 Step -(m_ProgressBlockWidth + 2)
    UserControl.Line (1, X - 1)-(UserControl.ScaleWidth - 2, X - m_ProgressBlockWidth), , BF
    Next X
    End If
    
    Case 1 'Smooth
    
    UserControl.Line (0, 0)-(UserControl.ScaleWidth, UserControl.ScaleHeight), , BF
    
    End Select

End Sub


Private Sub DrawValue(Value As Single)

On Error Resume Next
Dim X As Single

If m_Orientation = Horizontal Then
X = (Value - m_Min) * UserControl.ScaleWidth / (m_Max - m_Min)
picProgress.Left = X
If m_ShowValue = True Then
lblValue2.Left = lblValue.Left - picProgress.Left
lblValue2.Top = lblValue.Top - picProgress.Top
End If

UserControl.Refresh
End If

If m_Orientation = Vertical Then
X = (Value - m_Min) * UserControl.ScaleHeight / (m_Max - m_Min)
picProgress.Top = -X
If m_ShowValue = True Then
lblValue2.Top = lblValue.Top - picProgress.Top
lblValue2.Left = lblValue.Left - picProgress.Left
End If

UserControl.Refresh
End If

If m_ShowValue = True Then
 lblValue.Caption = m_ValuePrefix & Value & m_ValueSuffix
 lblValue2.Left = lblValue.Left - picProgress.Left
 lblValue2.Top = lblValue.Top - picProgress.Top
 lblValue2.Caption = lblValue.Caption
End If


End Sub

Private Sub ResizeMe()

If UserControl.Width < 90 Then UserControl.Size 90, UserControl.Height
If UserControl.Height < 90 Then UserControl.Size UserControl.Width, 90

picProgress.Move 0, 0
picProgress.Width = UserControl.ScaleWidth
picProgress.Height = UserControl.ScaleHeight


DrawValue m_Value
DrawGradient m_GradientFromColor, m_GradientToColor
DoScrolling m_Scrolling
Sub_FitValue

End Sub

Private Sub Sub_FitValue()

If m_ShowValue = False Then Exit Sub

lblValue.Left = (UserControl.ScaleWidth - lblValue.Width) / 2
lblValue.Top = (UserControl.ScaleHeight - lblValue.Height) / 2

lblValue2.Top = lblValue.Top
lblValue2.Left = lblValue.Left - picProgress.Left
lblValue2.Top = lblValue.Top - picProgress.Top

End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get Value() As Single
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Single)
    m_Value = New_Value
    If m_Value < m_Min Then m_Value = m_Min
    If m_Value > m_Max Then m_Value = m_Max
    DrawValue m_Value
    PropertyChanged "Value"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get UseGradientColors() As Boolean
    UseGradientColors = m_UseGradientColors
End Property

Public Property Let UseGradientColors(ByVal New_UseGradientColors As Boolean)
    m_UseGradientColors = New_UseGradientColors
    
    If m_UseGradientColors Then
    DrawGradient m_GradientFromColor, m_GradientToColor
    Else
    UserControl.Cls
    DoScrolling m_Scrolling
    DrawValue m_Value
    End If
    
    PropertyChanged "UseGradientColors"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get GradientFromColor() As OLE_COLOR
    GradientFromColor = m_GradientFromColor
End Property

Public Property Let GradientFromColor(ByVal New_GradientFromColor As OLE_COLOR)
'MsgBox "From=" & New_GradientFromColor
    m_GradientFromColor = Function_ConvertOleColorToLong(New_GradientFromColor)
'MsgBox "From=" & New_GradientFromColor
    DrawGradient m_GradientFromColor, m_GradientToColor
    PropertyChanged "GradientFromColor"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get GradientToColor() As OLE_COLOR
    GradientToColor = m_GradientToColor
End Property

Private Function Function_ConvertOleColorToLong(lngColor As OLE_COLOR) As Long

Dim lngOleColor As Long

OleTranslateColor lngColor, 0&, lngOleColor
Function_ConvertOleColorToLong = lngOleColor

End Function

Public Property Let GradientToColor(ByVal New_GradientToColor As OLE_COLOR)
    m_GradientToColor = Function_ConvertOleColorToLong(New_GradientToColor)
    DrawGradient m_GradientFromColor, m_GradientToColor
    PropertyChanged "GradientToColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Orientation() As OrientationEnum
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Orientation As OrientationEnum)
    m_Orientation = New_Orientation
    ResizeMe
    PropertyChanged "Orientation"
End Property

Private Sub lblValue_Click()
RaiseEvent Click
End Sub

Private Sub lblValue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X \ 15 + Int(lblValue.Left), Y \ 15 + Int(lblValue.Top))

End Sub


Private Sub lblValue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X \ 15 + Int(lblValue.Left), Y \ 15 + Int(lblValue.Top))

End Sub


Private Sub lblValue_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X \ 15 + Int(lblValue.Left), Y \ 15 + Int(lblValue.Top))

End Sub


Private Sub lblValue2_Click()
RaiseEvent Click
End Sub

Private Sub lblValue2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X \ 15 + Int(picProgress.Left) + 1 + Int(lblValue2.Left), Y \ 15 + Int(lblValue2.Top))

End Sub


Private Sub lblValue2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X \ 15 + Int(picProgress.Left) + 1 + Int(lblValue2.Left), Y \ 15 + Int(lblValue2.Top))

End Sub


Private Sub lblValue2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X \ 15 + Int(picProgress.Left) + 1 + Int(lblValue2.Left), Y \ 15 + Int(lblValue2.Top))

End Sub


Private Sub picProgress_Click()
    RaiseEvent Click

End Sub

Private Sub picProgress_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X + Int(picProgress.Left) + 1, Y)

End Sub

Private Sub picProgress_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X + Int(picProgress.Left) + 1, Y)

End Sub


Private Sub picProgress_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X + Int(picProgress.Left) + 1, Y)

End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_Value = m_def_Value
    m_UseGradientColors = m_def_UseGradientColors
    m_GradientFromColor = m_def_GradientFromColor
    m_GradientToColor = m_def_GradientToColor
    m_Orientation = m_def_Orientation
    m_Min = m_def_Min
    m_Max = m_def_Max
    m_Scrolling = m_def_Scrolling
    m_ProgressBlockWidth = m_def_ProgressBlockWidth
    
    m_ShowValue = m_def_ShowValue
    
    m_ValueSuffix = m_def_ValueSuffix
    m_ValuePrefix = m_def_ValuePrefix

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_UseGradientColors = PropBag.ReadProperty("UseGradientColors", m_def_UseGradientColors)
    
    m_GradientFromColor = PropBag.ReadProperty("GradientFromColor", m_def_GradientFromColor)
    m_GradientFromColor = Function_ConvertOleColorToLong(m_GradientFromColor)
    
    m_GradientToColor = PropBag.ReadProperty("GradientToColor", m_def_GradientToColor)
    m_GradientToColor = Function_ConvertOleColorToLong(m_GradientToColor)
    
    m_Orientation = PropBag.ReadProperty("Orientation", m_def_Orientation)
    m_Min = PropBag.ReadProperty("Min", m_def_Min)
    m_Max = PropBag.ReadProperty("Max", m_def_Max)

    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_Scrolling = PropBag.ReadProperty("Scrolling", m_def_Scrolling)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.BackColor = PropBag.ReadProperty("ProgressBackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ProgressForeColor", &H80000012)
    m_ProgressBlockWidth = PropBag.ReadProperty("ProgressBlockWidth", m_def_ProgressBlockWidth)
    picProgress.BackColor = PropBag.ReadProperty("ProgressRemainColor", &H8000000F)
    m_ShowValue = PropBag.ReadProperty("ShowValue", m_def_ShowValue)

    Sub_FitValue
    DrawValue m_Value
    DrawGradient m_GradientFromColor, m_GradientToColor
    DoScrolling m_Scrolling
    
    lblValue.ForeColor = PropBag.ReadProperty("ValueForeColor", &H80000012)
    lblValue2.ForeColor = lblValue.ForeColor
    Set lblValue.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set lblValue2.Font = lblValue.Font
    m_ValueSuffix = PropBag.ReadProperty("ValueSuffix", m_def_ValueSuffix)
    m_ValuePrefix = PropBag.ReadProperty("ValuePrefix", m_def_ValuePrefix)
End Sub

Private Sub UserControl_Resize()

ResizeMe


End Sub

Private Sub UserControl_Show()
ResizeMe

End Sub


'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("UseGradientColors", m_UseGradientColors, m_def_UseGradientColors)
    Call PropBag.WriteProperty("GradientFromColor", m_GradientFromColor, m_def_GradientFromColor)
    Call PropBag.WriteProperty("GradientToColor", m_GradientToColor, m_def_GradientToColor)
    Call PropBag.WriteProperty("Orientation", m_Orientation, m_def_Orientation)
    Call PropBag.WriteProperty("Min", m_Min, m_def_Min)
    Call PropBag.WriteProperty("Max", m_Max, m_def_Max)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Scrolling", m_Scrolling, m_def_Scrolling)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("ProgressBackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ProgressForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("ProgressBlockWidth", m_ProgressBlockWidth, m_def_ProgressBlockWidth)
    Call PropBag.WriteProperty("ProgressRemainColor", picProgress.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ShowValue", m_ShowValue, m_def_ShowValue)
    Call PropBag.WriteProperty("ValueForeColor", lblValue.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Font", lblValue.Font, Ambient.Font)
    Call PropBag.WriteProperty("ValueSuffix", m_ValueSuffix, m_def_ValueSuffix)
    Call PropBag.WriteProperty("ValuePrefix", m_ValuePrefix, m_def_ValuePrefix)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Min() As Single
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Min As Single)
    If New_Min >= m_Max Then Exit Property
    m_Min = New_Min
    
    If m_Value < m_Min Then
     m_Value = m_Min
    End If
    
    DoScrolling m_Scrolling
    DrawValue m_Value
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Max() As Single
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Max As Single)
    If New_Max <= m_Min Then Exit Property
    m_Max = New_Max
    If m_Value > m_Max Then
     m_Value = m_Max
    End If
    DoScrolling m_Scrolling
    DrawValue m_Value
    PropertyChanged "Max"

End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Scrolling() As ScrollingEnum
    Scrolling = m_Scrolling
End Property

Public Property Let Scrolling(ByVal New_Scrolling As ScrollingEnum)
    
    m_Scrolling = New_Scrolling
    DoScrolling m_Scrolling
    ResizeMe
    PropertyChanged "Scrolling"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As BorderStyleEnum
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleEnum)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property



'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get ProgressBackColor() As OLE_COLOR
Attribute ProgressBackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    ProgressBackColor = UserControl.BackColor
End Property

Public Property Let ProgressBackColor(ByVal New_ProgressBackColor As OLE_COLOR)
    UserControl.BackColor() = New_ProgressBackColor
    DrawValue m_Value
    DoScrolling m_Scrolling
    PropertyChanged "ProgressBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ProgressForeColor() As OLE_COLOR
Attribute ProgressForeColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    ProgressForeColor = UserControl.ForeColor
End Property

Public Property Let ProgressForeColor(ByVal New_ProgressForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ProgressForeColor
    DrawValue m_Value
    DoScrolling m_Scrolling
    PropertyChanged "ProgressForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ProgressBlockWidth() As Integer
    ProgressBlockWidth = m_ProgressBlockWidth
End Property

Public Property Let ProgressBlockWidth(ByVal New_ProgressBlockWidth As Integer)
    m_ProgressBlockWidth = New_ProgressBlockWidth
    If m_ProgressBlockWidth < 0 Then m_ProgressBlockWidth = 0
    DoScrolling m_Scrolling
    PropertyChanged "ProgressBlockWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picProgress,picProgress,-1,BackColor
Public Property Get ProgressRemainColor() As OLE_COLOR
Attribute ProgressRemainColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    ProgressRemainColor = picProgress.BackColor
End Property

Public Property Let ProgressRemainColor(ByVal New_ProgressRemainColor As OLE_COLOR)
    picProgress.BackColor() = New_ProgressRemainColor
    PropertyChanged "ProgressRemainColor"
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub DrawGradient(FromColor As OLE_COLOR, ToColor As OLE_COLOR)

On Error Resume Next
If Not m_UseGradientColors Then Exit Sub

UserControl.Cls



Dim X As Integer
Dim R1 As Single
Dim G1 As Single
Dim B1 As Single
Dim R2 As Single
Dim G2 As Single
Dim B2 As Single
Dim RX As Single
Dim GX As Single
Dim BX As Single
Dim RY As Single
Dim GY As Single
Dim BY As Single
Dim SW As Single
Dim SH As Single

R1 = GetRValue(FromColor)
G1 = GetGValue(FromColor)
B1 = GetBValue(FromColor)
R2 = GetRValue(ToColor)
G2 = GetGValue(ToColor)
B2 = GetBValue(ToColor)
SW = UserControl.ScaleWidth
SH = UserControl.ScaleHeight


If m_Orientation = Horizontal Then
RY = (R2 - R1) / SW
GY = (G2 - G1) / SW
BY = (B2 - B1) / SW

For X = 0 To SW
UserControl.Line (X, 0)-(X, UserControl.ScaleHeight), RGB(R1, G1, B1), BF
R1 = R1 + RY
G1 = G1 + GY
B1 = B1 + BY
Next X
End If

If m_Orientation = Vertical Then
RY = (R2 - R1) / SH
GY = (G2 - G1) / SH
BY = (B2 - B1) / SH

For X = SH To 0 Step -1
UserControl.Line (0, X)-(UserControl.ScaleWidth, X), RGB(R1, G1, B1), BF
R1 = R1 + RY
G1 = G1 + GY
B1 = B1 + BY
Next X
End If


End Sub

Private Function GetRValue(Color As OLE_COLOR) As Byte
GetRValue = Color And &HFF
End Function

Private Function GetGValue(Color As OLE_COLOR) As Byte
GetGValue = Color \ 256 And &HFF
End Function


Private Function GetBValue(Color As OLE_COLOR) As Byte
GetBValue = Color \ 65536
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get ShowValue() As Boolean
    ShowValue = m_ShowValue
End Property

Public Property Let ShowValue(ByVal New_ShowValue As Boolean)
    m_ShowValue = New_ShowValue
    lblValue.Visible = m_ShowValue
    lblValue.Caption = m_ValuePrefix & Value & m_ValueSuffix
    lblValue2.Visible = lblValue.Visible
    lblValue2.Caption = lblValue.Caption

    Sub_FitValue
    PropertyChanged "ShowValue"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblValue,lblValue,-1,ForeColor
Public Property Get ValueForeColor() As OLE_COLOR
Attribute ValueForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ValueForeColor = lblValue.ForeColor
End Property

Public Property Let ValueForeColor(ByVal New_ValueForeColor As OLE_COLOR)
    lblValue.ForeColor() = New_ValueForeColor
    lblValue2.ForeColor() = lblValue.ForeColor()
    PropertyChanged "ValueForeColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
    Set Font = lblValue.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblValue.Font = New_Font
    Set lblValue2.Font = lblValue.Font
    Sub_FitValue
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ValueSuffix() As String
    ValueSuffix = m_ValueSuffix
End Property

Public Property Let ValueSuffix(ByVal New_ValueSuffix As String)
    m_ValueSuffix = New_ValueSuffix
    lblValue.Caption = m_ValuePrefix & Value & m_ValueSuffix
    lblValue2.Caption = lblValue.Caption

    Sub_FitValue
    PropertyChanged "ValueSuffix"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ValuePrefix() As String
    ValuePrefix = m_ValuePrefix
End Property

Public Property Let ValuePrefix(ByVal New_ValuePrefix As String)
    m_ValuePrefix = New_ValuePrefix
    lblValue.Caption = m_ValuePrefix & Value & m_ValueSuffix
    lblValue2.Caption = lblValue.Caption

    Sub_FitValue
    PropertyChanged "ValuePrefix"
End Property

