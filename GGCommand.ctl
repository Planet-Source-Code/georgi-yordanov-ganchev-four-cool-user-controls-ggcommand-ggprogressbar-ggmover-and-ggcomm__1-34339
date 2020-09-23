VERSION 5.00
Begin VB.UserControl GGCommand 
   AutoRedraw      =   -1  'True
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1230
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   82
   ToolboxBitmap   =   "GGCommand.ctx":0000
   Begin VB.Image imgPicture 
      Height          =   330
      Left            =   45
      Top             =   90
      Width           =   330
   End
End
Attribute VB_Name = "GGCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 ' Enum Sastoiania na ramkata
Public Enum enum_ggcmd_BorderState
            [None] ' niama ramka
            [Raised] ' "povdignata" ramka
            [Inset] ' "natisnata" ramka
End Enum

 ' Enum Border3DStyle
Public Enum enum_ggcmd_Border3DStyle
            [Border3DStyle_None]
            [Border3DStyle_ClientEdge]
            [Border3DStyle_DlgModalFrame]
            [Border3DStyle_StaticEdge]
            [Border3DStyle_Frame]
End Enum

 ' Enum Podravniavane
Public Enum enum_ggcmd_Alignment
            [Left_Top]
            [Left_Middle]
            [Left_Bottom]
            [Center_Top]
            [Center_Middle]
            [Center_Bottom]
            [Right_Top]
            [Right_Middle]
            [Right_Bottom]
End Enum

 ' Enum ToolTip Icon
Public Enum enum_ggcmd_ToolTipIcon
    [ToolTipIcon_None]
    [ToolTipIcon_Info]
    [ToolTipIcon_Warning]
    [ToolTipIcon_Error]
End Enum

 ' Enum ToolTip Style
Public Enum enum_ggcmd_ToolTipStyle
    [ToolTipStyle_Standart]
    [ToolTipStyle_Balloon]
End Enum

Public Enum enum_ggcmd_GradientOrientation
       [GradientOrientation_Vertical]
       [GradientOrientation_Horizontal]
End Enum

'Property Variables:
Private pstrCaption As String
Private plngCaptionXOffset As Long
Private plngCaptionYOffset As Long

Private penumCaptionAlignment As enum_ggcmd_Alignment
Private penumPictureAlignment As enum_ggcmd_Alignment

Private pboolUseHoverProperties As Boolean
Private pboolFontBoldOnHover As Boolean
Private plngFontColorOnHover As OLE_COLOR

Private ppicPicture As Picture
Private plngPictureXOffset As Long
Private plngPictureYOffset As Long

Private pboolAlwaysShowBorder As Boolean

Private plngBorder3DStyle As enum_ggcmd_Border3DStyle

Private proplngGradientOrientation As Long
Private propboolUseGradientFill As Boolean
Private proplngGradientFromColor As OLE_COLOR
Private proplngGradientToColor As OLE_COLOR
Private propboolGradientInverseOnPress As Boolean

Private propboolDontReleaseCapture As Boolean

 ' ToolTip properties
Private pstrToolTip As String
Private pstrToolTipTitle As String
Private pboolToolTipCentered As Boolean
Private plngToolTipBackColor As OLE_COLOR
Private plngToolTipForeColor As OLE_COLOR
Private plngToolTipIcon As enum_ggcmd_ToolTipIcon
Private plngToolTipStyle As enum_ggcmd_ToolTipStyle
Private plngToolTipHwnd As Long

'Local variables
Private boolCaptured As Boolean
Private enumBorderState As enum_ggcmd_BorderState
Private boolMouseFocus As Boolean
Private sngCaptionXPos As Single
Private sngCaptionYPos As Single
Private sngPictureXPos As Single
Private sngPictureYPos As Single
Private byteOffset As Byte
Private boolMouseWasDown As Boolean

'Public events
Public Event Click()
Public Event DblClick()
Public Event MouseEnter()
Public Event MouseExit()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'API functions
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long


Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Const WM_USER = &H400
Private Const CW_USEDEFAULT = &H80000000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_FRAMECHANGED = &H20
Private Const HWND_TOPMOST = -1

''Tooltip Window Constants
Private Const TTS_NOPREFIX = &H2
Private Const TTF_TRANSPARENT = &H100
Private Const TTF_CENTERTIP = &H2
Private Const TTM_ADDTOOLA = (WM_USER + 4)
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_UPDATETIPTEXTA = (WM_USER + 12)
Private Const TTM_SETDELAYTIME = (WM_USER + 28)
Private Const TTM_SETMAXTIPWIDTH = (WM_USER + 24)
Private Const TTM_SETTIPBKCOLOR = (WM_USER + 19)
Private Const TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
Private Const TTM_SETTITLE = (WM_USER + 32)
Private Const TTS_BALLOON = &H40
Private Const TTS_ALWAYSTIP = &H1
Private Const TTF_SUBCLASS = &H10
Private Const TOOLTIPS_CLASSA = "tooltips_class32"
''Tooltip Window Types
Private Type TOOLINFO
    lSize As Long
    lFlags As Long
    lHwnd As Long
    lId As Long
    lpRect As RECT
    hInstance As Long
    lpStr As String
    lParam As Long
End Type

Private tpToolInfo As TOOLINFO

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Public Property Let Border3DStyle(New_Border3DStyle As enum_ggcmd_Border3DStyle)

plngBorder3DStyle = New_Border3DStyle

Call Sub_CalculateAlignment
Call Sub_PrintCaption

PropertyChanged "Border3DStyle"

End Property

Public Property Get Border3DStyle() As enum_ggcmd_Border3DStyle

Border3DStyle = plngBorder3DStyle

End Property


Public Property Let DontReleaseCapture(boolNewDontReleaseCapture As Boolean)

propboolDontReleaseCapture = boolNewDontReleaseCapture

End Property

Public Property Let GradientFromColor(lngNewGradientFromColor As OLE_COLOR)

If proplngGradientFromColor <> lngNewGradientFromColor Then
   proplngGradientFromColor = lngNewGradientFromColor
   Sub_PrintCaption
   PropertyChanged "GradientFromColor"
End If

End Property

Public Property Get GradientFromColor() As OLE_COLOR

GradientFromColor = proplngGradientFromColor

End Property


Public Property Let GradientInverseOnPress(boolNewGradientInverseOnPress As Boolean)

propboolGradientInverseOnPress = boolNewGradientInverseOnPress
Sub_PrintCaption
PropertyChanged "GradientInverseOnPress"

End Property

Public Property Get GradientInverseOnPress() As Boolean

GradientInverseOnPress = propboolGradientInverseOnPress

End Property

Public Property Let GradientOrientation(lngNewOrientation As enum_ggcmd_GradientOrientation)

If lngNewOrientation <> proplngGradientOrientation Then
   proplngGradientOrientation = lngNewOrientation
   GradientOrientation = lngNewOrientation
   Sub_CalculateAlignment
   Call Sub_PrintCaption
   PropertyChanged "GradientOrientation"
End If

End Property

Public Property Get GradientOrientation() As enum_ggcmd_GradientOrientation

GradientOrientation = proplngGradientOrientation

End Property

Public Property Get GradientToColor() As OLE_COLOR

GradientToColor = proplngGradientToColor

End Property

Public Property Let GradientToColor(lngNewGradientToColor As OLE_COLOR)

If proplngGradientToColor <> lngNewGradientToColor Then
   proplngGradientToColor = lngNewGradientToColor
   Sub_PrintCaption
   PropertyChanged "GradientToColor"
End If

End Property

Private Sub Sub_SetBorder3DStyle(lngBorderStyle As enum_ggcmd_Border3DStyle)

Select Case lngBorderStyle
       Case Border3DStyle_None
            SetWindowLong UserControl.hWnd, -20, 0
       Case Border3DStyle_ClientEdge
             'WS_EX_CLIENTEDGE=&h200
            SetWindowLong UserControl.hWnd, -20, &H200
       Case Border3DStyle_DlgModalFrame
             'WS_EX_DLGMODALFRAME=&h1
            SetWindowLong UserControl.hWnd, -20, &H1
       Case Border3DStyle_StaticEdge
             'WS_EX_STATICEDGE=&h20000
            SetWindowLong UserControl.hWnd, -20, &H20000
       Case Border3DStyle_Frame
             'WS_EX_CLIENTEDGE=&H200 Or WS_EX_DLGMODALFRAME=&H1
            SetWindowLong UserControl.hWnd, -20, &H201
End Select

SetWindowPos UserControl.hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED

End Sub

Public Property Let ToolTipBackColor(New_ToolTipBackColor As OLE_COLOR)

plngToolTipBackColor = New_ToolTipBackColor

PropertyChanged "ToolTipBackColor"

End Property

Public Property Let ToolTipForeColor(New_ToolTipForeColor As OLE_COLOR)

plngToolTipForeColor = New_ToolTipForeColor

PropertyChanged "ToolTipForeColor"

End Property


Public Property Get ToolTipBackColor() As OLE_COLOR

ToolTipBackColor = plngToolTipBackColor

End Property

Public Property Get ToolTipForeColor() As OLE_COLOR

ToolTipForeColor = plngToolTipForeColor

End Property


Public Property Get ToolTipCentered() As Boolean

ToolTipCentered = pboolToolTipCentered

End Property

Public Property Let ToolTipCentered(New_ToolTipCentered As Boolean)

pboolToolTipCentered = New_ToolTipCentered

PropertyChanged "ToolTipCentered"

End Property

Private Function ToolTipCreate() As Boolean
    
 ' This code is downloaded from www.planet-source-code.com
 ' File name is(was in 2001) - Awesome To3117410252001.zip
    
Dim lpRect As RECT
Dim lngToolTipStyle As Long
Dim lngOleColor As Long
    
If plngToolTipHwnd <> 0 Then
 Call ToolTipDestroy
End If
    
lngToolTipStyle = TTS_ALWAYSTIP Or TTS_NOPREFIX
    
 ''create baloon style if desired
If plngToolTipStyle = ToolTipStyle_Balloon Then lngToolTipStyle = lngToolTipStyle Or TTS_BALLOON
    
plngToolTipHwnd = CreateWindowEx(0&, _
                                 TOOLTIPS_CLASSA, _
                                 vbNullString, _
                                 lngToolTipStyle, _
                                 CW_USEDEFAULT, _
                                 CW_USEDEFAULT, _
                                 CW_USEDEFAULT, _
                                 CW_USEDEFAULT, _
                                 UserControl.hWnd, _
                                 0&, _
                                 App.hInstance, _
                                 0&)
 ''make our tooltip window a topmost window
SetWindowPos plngToolTipHwnd, _
             HWND_TOPMOST, _
             0&, _
             0&, _
             0&, _
             0&, _
             SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
                    
 ''get the rect of the parent control
GetClientRect plngToolTipHwnd, lpRect
        
 ''now set our tooltip info structure
With tpToolInfo
  ''if we want it centered, then set that flag
 If pboolToolTipCentered = True Then
  .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP
 Else
  .lFlags = TTF_SUBCLASS
 End If
 ''set the hwnd prop to our parent control's hwnd
 .lHwnd = UserControl.hWnd
 .lId = 0
 .hInstance = App.hInstance
 .lpStr = pstrToolTip
 .lpRect = lpRect
End With
        
 ''add the tooltip structure
SendMessage plngToolTipHwnd, TTM_ADDTOOLA, 0&, tpToolInfo
        
'If plngToolTipHwnd <> 0 Then
' SendMessage plngToolTipHwnd, TTM_UPDATETIPTEXTA, 0&, tpToolInfo
'End If
        
 ''if we want a title or we want an icon
If pstrToolTipTitle <> vbNullString Or plngToolTipIcon <> ToolTipIcon_None Then
 SendMessage plngToolTipHwnd, TTM_SETTITLE, CLng(plngToolTipIcon), ByVal pstrToolTipTitle
End If

OleTranslateColor plngToolTipForeColor, 0&, lngOleColor
SendMessage plngToolTipHwnd, TTM_SETTIPTEXTCOLOR, lngOleColor, 0&

OleTranslateColor plngToolTipBackColor, 0&, lngOleColor
SendMessage plngToolTipHwnd, TTM_SETTIPBKCOLOR, lngOleColor, 0&
    
End Function
Public Property Let AlwaysShowBorder(New_AlwaysShowBorder As Boolean)

pboolAlwaysShowBorder = New_AlwaysShowBorder
If pboolAlwaysShowBorder = True Then
 Sub_DrawBorder [Raised]
Else
 Sub_DrawBorder [None]
End If
PropertyChanged "AlwaysShowBorder"

End Property

Public Property Get AlwaysShowBorder() As Boolean

AlwaysShowBorder = pboolAlwaysShowBorder

End Property


Public Property Let Caption(New_Caption As String)

pstrCaption = New_Caption
Sub_CalculateAlignment
Sub_PrintCaption

PropertyChanged "Caption"

End Property

Public Property Let CaptionAlignment(New_CaptionAlignment As enum_ggcmd_Alignment)

penumCaptionAlignment = New_CaptionAlignment
Sub_CalculateAlignment
Sub_PrintCaption
PropertyChanged "CaptionAlignment"

End Property

Public Property Get CaptionAlignment() As enum_ggcmd_Alignment

CaptionAlignment = penumCaptionAlignment

End Property

Public Property Let CaptionXOffset(New_CaptionXOffset As Long)

plngCaptionXOffset = New_CaptionXOffset
Sub_CalculateAlignment
Sub_PrintCaption
PropertyChanged "CaptionXOffset"

End Property

Public Property Get CaptionXOffset() As Long

CaptionXOffset = plngCaptionXOffset

End Property


Public Property Let CaptionYOffset(New_CaptionYOffset As Long)

plngCaptionYOffset = New_CaptionYOffset
Sub_CalculateAlignment
Sub_PrintCaption
PropertyChanged "CaptionYOffset"


End Property

Public Property Get CaptionYOffset() As Long

CaptionYOffset = plngCaptionYOffset

End Property


Public Property Let FontBoldOnHover(New_FontBoldOnHover As Boolean)

pboolFontBoldOnHover = New_FontBoldOnHover
PropertyChanged "FontBoldOnHover"

End Property

Public Property Get FontBoldOnHover() As Boolean

FontBoldOnHover = pboolFontBoldOnHover

End Property

Public Property Let FontColorOnHover(New_FontColorOnHover As OLE_COLOR)

plngFontColorOnHover = New_FontColorOnHover
PropertyChanged "FontColorOnHover"

End Property

Public Property Get FontColorOnHover() As OLE_COLOR

FontColorOnHover = plngFontColorOnHover

End Property


Public Property Let ForeColor(New_ForeColor As OLE_COLOR)

UserControl.ForeColor() = New_ForeColor
''''' ' Za da se opravi buga pri smiana na Font.Bold i .ForeColor
'''''Sub_CalculateAlignment
Sub_PrintCaption
Sub_DrawBorder enumBorderState
PropertyChanged "ForeColor"

End Property

Public Property Get ForeColor() As OLE_COLOR

ForeColor = UserControl.ForeColor

End Property

Public Property Let PictureAlignment(New_PictureAlignment As enum_ggcmd_Alignment)

penumPictureAlignment = New_PictureAlignment
Sub_CalculateAlignment
Sub_PrintCaption
PropertyChanged "PictureAlignment"

End Property

Public Property Get PictureAlignment() As enum_ggcmd_Alignment

PictureAlignment = penumPictureAlignment

End Property

'Public Property Let PictureTiled(New_PictureTiled As Boolean)
'
'pboolPictureTiled = New_PictureTiled
'Call Sub_PrintCaption
'
'PropertyChanged "PictureTiled"
'
'End Property

'Public Property Get PictureTiled() As Boolean
'
'PictureTiled = pboolPictureTiled
'
'End Property
'
Public Property Let PictureXOffset(New_PictureXOffset As Long)

plngPictureXOffset = New_PictureXOffset
Sub_CalculateAlignment
Sub_PrintCaption
PropertyChanged "PictureXOffset"

End Property

Public Property Get PictureXOffset() As Long

PictureXOffset = plngPictureXOffset

End Property

Public Property Let PictureYOffset(New_PictureYOffset As Long)

plngPictureYOffset = New_PictureYOffset
Sub_CalculateAlignment
Sub_PrintCaption
PropertyChanged "PictureYOffset"

End Property

Public Property Get PictureYOffset() As Long

PictureYOffset = plngPictureYOffset

End Property


Private Sub Sub_CalculateAlignment(Optional byteCalculateFor As Byte = 0)

On Error Resume Next

Dim sngTextWidth As Single
Dim sngTextHeight As Single

Dim sngPictureWidth As Single
Dim sngPictureHeight As Single

Dim sngUCWidth As Single
Dim sngUCHeight As Single

Dim sngCenter As Single
Dim sngRight As Single
Dim sngMiddle As Single
Dim sngBottom As Single

Call Sub_SetBorder3DStyle(plngBorder3DStyle)

sngUCHeight = UserControl.ScaleHeight
sngUCWidth = UserControl.ScaleWidth

sngTextWidth = UserControl.TextWidth(pstrCaption)
sngTextHeight = UserControl.TextHeight(pstrCaption)

 ' Izchisliavane na poziciite po horizontalata za caption
sngCenter = (sngUCWidth - sngTextWidth) / 2
sngRight = sngUCWidth - sngTextWidth
 ' Izchisliavane na poziciite po vertikalata za caption
sngMiddle = (sngUCHeight - sngTextHeight) / 2
sngBottom = sngUCHeight - sngTextHeight

 ' Izchisliavane poziciata na Caption
Select Case penumCaptionAlignment
 Case [Left_Top]
  sngCaptionXPos = 1
  sngCaptionYPos = 1
 
 Case [Left_Middle]
  sngCaptionXPos = 1
  sngCaptionYPos = sngMiddle
 
 Case [Left_Bottom]
  sngCaptionXPos = 1
  sngCaptionYPos = sngBottom - 1
 
 Case [Center_Top]
  sngCaptionXPos = sngCenter
  sngCaptionYPos = 1
  
 Case [Center_Middle]
  sngCaptionXPos = sngCenter
  sngCaptionYPos = sngMiddle
 
 Case [Center_Bottom]
  sngCaptionXPos = sngCenter
  sngCaptionYPos = sngBottom - 1
 
 Case [Right_Top]
  sngCaptionXPos = sngRight - 2
  sngCaptionYPos = 1
  
 Case [Right_Middle]
  sngCaptionXPos = sngRight - 2
  sngCaptionYPos = sngMiddle
 
 Case [Right_Bottom]
  sngCaptionXPos = sngRight - 2
  sngCaptionYPos = sngBottom - 1
End Select

sngCaptionXPos = Int(sngCaptionXPos + plngCaptionXOffset)
sngCaptionYPos = Int(sngCaptionYPos + plngCaptionYOffset)


sngPictureWidth = ppicPicture.Width / Screen.TwipsPerPixelX / 1.75
sngPictureHeight = ppicPicture.Height / Screen.TwipsPerPixelY / 1.75

 ' Izchisliavane na poziciite po horizontalata za Picture
sngCenter = (sngUCWidth - sngPictureWidth) / 2
sngRight = sngUCWidth - sngPictureWidth
 ' Izchisliavane na poziciite po vertikalata za Picture
sngMiddle = (sngUCHeight - sngPictureHeight) / 2
sngBottom = sngUCHeight - sngPictureHeight

 ' Izchisliavane poziciata na Picture
Select Case penumPictureAlignment
 Case [Left_Top]
  sngPictureXPos = 1
  sngPictureYPos = 1
 
 Case [Left_Middle]
  sngPictureXPos = 1
  sngPictureYPos = sngMiddle
 
 Case [Left_Bottom]
  sngPictureXPos = 1
  sngPictureYPos = sngBottom - 1
 
 Case [Center_Top]
  sngPictureXPos = sngCenter
  sngPictureYPos = 1
  
 Case [Center_Middle]
  sngPictureXPos = sngCenter
  sngPictureYPos = sngMiddle
 
 Case [Center_Bottom]
  sngPictureXPos = sngCenter
  sngPictureYPos = sngBottom - 1
 
 Case [Right_Top]
  sngPictureXPos = sngRight - 1
  sngPictureYPos = 1
  
 Case [Right_Middle]
  sngPictureXPos = sngRight - 1
  sngPictureYPos = sngMiddle
 
 Case [Right_Bottom]
  sngPictureXPos = sngRight - 1
  sngPictureYPos = sngBottom - 1
End Select

sngPictureXPos = Int(sngPictureXPos + plngPictureXOffset)
sngPictureYPos = Int(sngPictureYPos + plngPictureYOffset)

End Sub



Public Property Get Caption() As String

Caption = pstrCaption

End Property


Public Property Get hWnd() As Long

hWnd = UserControl.hWnd

End Property


Private Sub Sub_DrawBorder(lngState As enum_ggcmd_BorderState, Optional boolDrawFocus As Boolean = False)

Dim sngUCSH As Single
Dim sngUCSW As Single
Dim lngLeftUpColor As Long
Dim lngRightDownColor As Long

 ' Opredeliane na cvetovete za liniite
 ' V zavisimost ot "State" na bordera
Select Case lngState
 Case 0 ' None - Liniite sa s cveta na .BackColor
  lngLeftUpColor = UserControl.BackColor
  lngRightDownColor = lngLeftUpColor
 Case 1 ' Raised - Liavata i gornata - beli - Diasnata i dolnata - tamno sivi
  lngLeftUpColor = vbWhite
  lngRightDownColor = vb3DShadow
 Case 2 ' Inset - Liavata i gornata - tamno sivi - Diasnata i dolnata - beli
  lngLeftUpColor = vb3DShadow
  lngRightDownColor = vbWhite
End Select

If pboolAlwaysShowBorder = True And lngState = None Then
 lngLeftUpColor = vbWhite
 lngRightDownColor = vb3DShadow
End If

sngUCSH = UserControl.ScaleHeight - 1
sngUCSW = UserControl.ScaleWidth - 1
 
'''If boolDrawFocus = True Then
''' UserControl.DrawStyle = vbDot
''' UserControl.Line (3, 3)-(sngUCSW - 3, sngUCSH - 3), , B
''' UserControl.DrawStyle = 0
'''End If
 
 ' Red na chertane - Liavo-Gore
UserControl.Line (0, sngUCSH)-(0, 0), lngLeftUpColor
UserControl.Line -(sngUCSW, 0), lngLeftUpColor
 ' Posle - Diasno-Doly
UserControl.Line (sngUCSW, -1)-(sngUCSW, sngUCSH), lngRightDownColor
UserControl.Line -(-1, sngUCSH), lngRightDownColor

enumBorderState = lngState


End Sub

Private Sub Sub_PrintCaption(Optional byteCaptionOffset As Byte = 0, _
                             Optional boolInverseGradient As Boolean = False)

On Error Resume Next
'Dim intI As Integer
'Dim intJ As Integer

UserControl.Cls

Sub_DrawGradient proplngGradientFromColor, proplngGradientToColor, boolInverseGradient

If TypeName(ppicPicture) = "Picture" Then
   If imgPicture.Left <> sngPictureXPos + byteCaptionOffset Or imgPicture.Top <> sngPictureYPos + byteCaptionOffset Then
      imgPicture.Move sngPictureXPos + byteCaptionOffset, sngPictureYPos + byteCaptionOffset
   End If
   If imgPicture.Picture <> ppicPicture Then
      Set imgPicture.Picture = ppicPicture
   End If
 'UserControl.PaintPicture ppicPicture, sngPictureXPos + byteCaptionOffset, sngPictureYPos + byteCaptionOffset
End If

UserControl.CurrentX = sngCaptionXPos + byteCaptionOffset
UserControl.CurrentY = sngCaptionYPos + byteCaptionOffset
UserControl.Print pstrCaption
byteOffset = byteCaptionOffset

If pboolAlwaysShowBorder = True Then
 Sub_DrawBorder [Raised]
End If


End Sub

Public Property Let ToolTip(New_ToolTip As String)

tpToolInfo.lpStr = New_ToolTip

If plngToolTipHwnd <> 0 Then
 SendMessage plngToolTipHwnd, TTM_UPDATETIPTEXTA, 0&, tpToolInfo
End If

pstrToolTip = New_ToolTip

PropertyChanged "ToolTip"

End Property

Public Property Get ToolTip() As String

ToolTip = pstrToolTip

End Property

Private Sub ToolTipDestroy()

DestroyWindow plngToolTipHwnd

End Sub

Public Property Let ToolTipIcon(New_ToolTipIcon As enum_ggcmd_ToolTipIcon)

plngToolTipIcon = New_ToolTipIcon

PropertyChanged "ToolTipIcon"

End Property

Public Property Get ToolTipIcon() As enum_ggcmd_ToolTipIcon

ToolTipIcon = plngToolTipIcon

End Property

Public Property Let ToolTipStyle(New_ToolTipStyle As enum_ggcmd_ToolTipStyle)

plngToolTipStyle = New_ToolTipStyle

PropertyChanged "ToolTipStyle"

End Property

Public Property Get ToolTipStyle() As enum_ggcmd_ToolTipStyle

ToolTipStyle = plngToolTipStyle

End Property

Public Property Get ToolTipTitle() As String

ToolTipTitle = pstrToolTipTitle

End Property



Public Property Let ToolTipTitle(New_ToolTipTitle As String)

pstrToolTipTitle = New_ToolTipTitle

PropertyChanged "ToolTipTitle"

End Property

Public Property Let UseGradientFill(boolNewUseGradientFill As Boolean)

If propboolUseGradientFill <> boolNewUseGradientFill Then
   propboolUseGradientFill = boolNewUseGradientFill
   UseGradientFill = propboolUseGradientFill
   Sub_CalculateAlignment
   Sub_PrintCaption
   PropertyChanged "UseGradientFill"
End If

End Property

Public Property Get UseGradientFill() As Boolean

UseGradientFill = propboolUseGradientFill

End Property

Public Property Let UseHoverProperties(New_UseHoverProperties As Boolean)

pboolUseHoverProperties = New_UseHoverProperties
PropertyChanged "UseHoverProperties"

End Property

Public Property Get UseHoverProperties() As Boolean

UseHoverProperties = pboolUseHoverProperties

End Property




Public Property Let Value(New_Value As Boolean)

If New_Value = True Then
 RaiseEvent Click
End If

End Property



Private Sub imgPicture_DblClick()

Call UserControl_DblClick

End Sub

Private Sub imgPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call UserControl_MouseDown(Button, Shift, X / 15, Y / 15)

End Sub


Private Sub imgPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call UserControl_MouseMove(Button, Shift, X / 15, Y / 15)

End Sub


Private Sub imgPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call UserControl_MouseUp(Button, Shift, X / 15, Y / 15)

End Sub


Private Sub UserControl_DblClick()

RaiseEvent DblClick

End Sub

Private Sub UserControl_GotFocus()

 ' Ako fokysa e vzet sled natiskane s mishkata
 ' da ne se pokazva bordera "Raised"

If boolMouseFocus = True Then
 boolMouseFocus = False
Else
 Sub_DrawBorder [Raised]
 Sub_DrawBorder None, True
End If

End Sub


Public Property Get MousePointer() As MousePointerConstants

 MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)

On Error Resume Next
 
UserControl.MousePointer() = New_MousePointer
PropertyChanged "MousePointer"

End Property

Private Function Fun_GetRValue(Color As OLE_COLOR) As Byte

On Error Resume Next

Fun_GetRValue = Color And &HFF

End Function

Private Function Fun_GetGValue(Color As OLE_COLOR) As Byte

On Error Resume Next

Fun_GetGValue = Color \ 256 And &HFF

End Function


Private Function Fun_GetBValue(Color As OLE_COLOR) As Byte

On Error Resume Next

Fun_GetBValue = Color \ 65536

End Function

Private Sub Sub_DrawGradient(FromColor As OLE_COLOR, _
                             ToColor As OLE_COLOR, _
                             Optional boolInverseGradient As Boolean = False)

On Error Resume Next

If propboolUseGradientFill = False Then Exit Sub

 ' Uncomment if .AutoRedraw=False
'UserControl.Cls



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
Dim lngFromColor As Long
Dim lngToColor As Long
Dim lngTemp As Long

OleTranslateColor FromColor, 0&, lngFromColor
OleTranslateColor ToColor, 0&, lngToColor

If boolInverseGradient = True Then
   lngTemp = lngFromColor
   lngFromColor = lngToColor
   lngToColor = lngTemp
End If

R1 = Fun_GetRValue(lngFromColor)
G1 = Fun_GetGValue(lngFromColor)
B1 = Fun_GetBValue(lngFromColor)
R2 = Fun_GetRValue(lngToColor)
G2 = Fun_GetGValue(lngToColor)
B2 = Fun_GetBValue(lngToColor)
SW = UserControl.ScaleWidth
SH = UserControl.ScaleHeight


If proplngGradientOrientation = GradientOrientation_Horizontal Then
   RY = (R2 - R1) / SW
   GY = (G2 - G1) / SW
   BY = (B2 - B1) / SW
   For X = 0 To SW
       UserControl.Line (X, 0)-(X, UserControl.ScaleHeight), RGB(R1, G1, B1) ', BF
       R1 = R1 + RY
       G1 = G1 + GY
       B1 = B1 + BY
   Next X
End If

If proplngGradientOrientation = GradientOrientation_Vertical Then
   RY = (R2 - R1) / SH
   GY = (G2 - G1) / SH
   BY = (B2 - B1) / SH
   For X = 0 To SH
       UserControl.Line (0, X)-(UserControl.ScaleWidth, X), RGB(R1, G1, B1) ', BF
       R1 = R1 + RY
       G1 = G1 + GY
       B1 = B1 + BY
   Next X
End If


End Sub


Private Sub UserControl_KeyPress(KeyAscii As Integer)

Exit Sub

If KeyAscii = 13 Then
 RaiseEvent Click
End If

End Sub


Private Sub UserControl_LostFocus()

Sub_PrintCaption
Sub_DrawBorder [None]

End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lngTemp As Long
Dim boolTemp As Boolean
Dim boolTempGradient As Boolean

 ' Saobshtavame na sabitieto "GotFocus"
 ' che fokysa e vzet sled natiskane s mishkata
 ' a ne s "TAB" klavisha
boolMouseFocus = True
boolMouseWasDown = True

 ' Parvo predizvikvam sabitieto
RaiseEvent MouseDown(Button, Shift, X, Y)

 ' Ako bytona e levia togava se chertae bordera
 ' I otnovo se vzima "Capture" zashtoto pri MouseDown toi se gybi
If Button = 1 Then
 lngTemp = UserControl.ForeColor
 boolTemp = UserControl.Font.Bold
 UserControl.ForeColor = plngFontColorOnHover
 UserControl.Font.Bold = pboolFontBoldOnHover
 If propboolUseGradientFill = True Then
    If propboolGradientInverseOnPress = True Then
       boolTempGradient = True
    End If
 End If
 Sub_CalculateAlignment
 Sub_PrintCaption 1, boolTempGradient
 UserControl.ForeColor = lngTemp
 UserControl.Font.Bold = boolTemp
 Sub_CalculateAlignment
 Sub_DrawBorder [Inset]
End If
SetCapture UserControl.hWnd
boolCaptured = True

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lngTemp As Long
Dim boolTemp As Boolean
Dim lngWindowFromPoint As Long
Dim tpPointApi As POINTAPI

tpPointApi.X = CLng(X)
tpPointApi.Y = CLng(Y)

ClientToScreen UserControl.hWnd, tpPointApi
lngWindowFromPoint = WindowFromPoint(tpPointApi.X, tpPointApi.Y)

If lngWindowFromPoint = UserControl.hWnd And X >= 0 And Y >= 0 And X <= UserControl.ScaleWidth And Y <= UserControl.ScaleHeight Then
 If boolCaptured = False Then

  If pstrToolTip <> "" Then
   Call ToolTipCreate
'   clsMyToolTip.TipText = pstrToolTip
'   clsMyToolTip.Style = TTBalloon
'   clsMyToolTip.Centered = True
'   clsMyToolTip.Title = "45yvw45y "
'   clsMyToolTip.ParentControl = Me.hwnd
'   clsMyToolTip.Create
  End If
  
  RaiseEvent MouseEnter
  SetCapture UserControl.hWnd
  If pboolUseHoverProperties = True Then
   lngTemp = UserControl.ForeColor
   boolTemp = UserControl.Font.Bold
   UserControl.ForeColor = plngFontColorOnHover
   UserControl.Font.Bold = pboolFontBoldOnHover
   Sub_CalculateAlignment
   Sub_PrintCaption byteOffset
   UserControl.Font.Bold = boolTemp
   UserControl.ForeColor = lngTemp
  End If
  Sub_DrawBorder [Raised]
  boolCaptured = True
 End If
Else
 If propboolDontReleaseCapture = False Then
    RaiseEvent MouseExit
    
    Call ToolTipDestroy
    
    boolMouseWasDown = False
    ReleaseCapture
    lngTemp = UserControl.ForeColor
    boolTemp = UserControl.Font.Bold
    UserControl.Font.Bold = pboolFontBoldOnHover
    UserControl.Font.Bold = boolTemp
    Sub_CalculateAlignment
    Sub_PrintCaption
    boolCaptured = False
    Sub_DrawBorder [None]
 End If
End If

RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

pstrCaption = UserControl.Extender.Name
penumCaptionAlignment = [Center_Middle]
penumPictureAlignment = [Left_Middle]
plngFontColorOnHover = vbBlue
pboolUseHoverProperties = True

plngToolTipBackColor = vbInfoBackground
plngToolTipForeColor = vbInfoText

proplngGradientOrientation = [GradientOrientation_Vertical]
propboolUseGradientFill = True
proplngGradientFromColor = vb3DHighlight
proplngGradientToColor = vb3DShadow
propboolGradientInverseOnPress = True

Set UserControl.Font = Ambient.Font

Set ppicPicture = LoadPicture("")

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lngTemp As Long
Dim boolTemp As Boolean

 ' Pri MouseUp ne se vzema focusa
boolMouseFocus = False

If Button = 1 And boolMouseWasDown = True Then
 RaiseEvent Click
End If
 
boolMouseWasDown = False
 
RaiseEvent MouseUp(Button, Shift, X, Y)

If Button = 1 Then
 lngTemp = UserControl.ForeColor
 boolTemp = UserControl.Font.Bold
 UserControl.ForeColor = plngFontColorOnHover
 UserControl.Font.Bold = pboolFontBoldOnHover
 Sub_CalculateAlignment
 Sub_PrintCaption
 UserControl.ForeColor = lngTemp
 UserControl.Font.Bold = boolTemp
 Sub_CalculateAlignment
 Sub_DrawBorder [Raised]
End If

SetCapture UserControl.hWnd
boolCaptured = True

End Sub

Private Sub UserControl_Paint()

'Sub_PrintCaption byteOffset

If enumBorderState <> [None] Then
 Sub_DrawBorder enumBorderState
End If

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

pstrCaption = PropBag.ReadProperty("Caption", UserControl.Extender.Name)
penumCaptionAlignment = PropBag.ReadProperty("CaptionAlignment", [Center_Middle])
plngCaptionXOffset = PropBag.ReadProperty("CaptionXOffset", 0)
plngCaptionYOffset = PropBag.ReadProperty("CaptionYOffset", 0)

pboolUseHoverProperties = PropBag.ReadProperty("UseHoverProperties", True)
pboolFontBoldOnHover = PropBag.ReadProperty("FontBoldOnHover", False)
plngFontColorOnHover = PropBag.ReadProperty("FontColorOnHover", vbBlue)

UserControl.BackColor = PropBag.ReadProperty("BackColor", vb3DFace)
UserControl.ForeColor = PropBag.ReadProperty("ForeColor", vbButtonText)

Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)

Set ppicPicture = PropBag.ReadProperty("Picture", Nothing)
penumPictureAlignment = PropBag.ReadProperty("PictureAlignment", [Left_Middle])
plngPictureXOffset = PropBag.ReadProperty("PictureXOffset", 0)
plngPictureYOffset = PropBag.ReadProperty("PictureYOffset", 0)

pboolAlwaysShowBorder = PropBag.ReadProperty("AlwaysShowBorder", False)

pstrToolTip = PropBag.ReadProperty("ToolTip", "")
pstrToolTipTitle = PropBag.ReadProperty("ToolTipTitle", "")
plngToolTipBackColor = PropBag.ReadProperty("ToolTipBackColor", vbInfoBackground)
plngToolTipForeColor = PropBag.ReadProperty("ToolTipForeColor", vbInfoText)
plngToolTipIcon = PropBag.ReadProperty("ToolTipIcon", ToolTipIcon_None)
plngToolTipStyle = PropBag.ReadProperty("ToolTipStyle", ToolTipStyle_Standart)
'pboolToolTipCentered = PropBag.ReadProperty("ToolTipCentered", False)

plngBorder3DStyle = PropBag.ReadProperty("Border3DStyle", Border3DStyle_None)

proplngGradientOrientation = PropBag.ReadProperty("GradientOrientation", [GradientOrientation_Vertical])
propboolUseGradientFill = PropBag.ReadProperty("UseGradientFill", True)
proplngGradientFromColor = PropBag.ReadProperty("GradientFromColor", vb3DHighlight)
proplngGradientToColor = PropBag.ReadProperty("GradientToColor", vb3DShadow)
propboolGradientInverseOnPress = PropBag.ReadProperty("GradientInverseOnPress", True)

UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)

UserControl.Enabled = PropBag.ReadProperty("Enabled", True)

End Sub

Private Sub UserControl_Resize()

Sub_CalculateAlignment
Sub_PrintCaption

End Sub


Private Sub UserControl_Show()

Sub_CalculateAlignment
Sub_PrintCaption

End Sub


Private Sub UserControl_Terminate()

Call ToolTipDestroy

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

Call PropBag.WriteProperty("Caption", pstrCaption, UserControl.Extender.Name)
Call PropBag.WriteProperty("CaptionAlignment", penumCaptionAlignment, [Center_Middle])
Call PropBag.WriteProperty("CaptionXOffset", plngCaptionXOffset, 0)
Call PropBag.WriteProperty("CaptionYOffset", plngCaptionYOffset, 0)

Call PropBag.WriteProperty("UseHoverProperties", pboolUseHoverProperties, True)
Call PropBag.WriteProperty("FontBoldOnHover", pboolFontBoldOnHover, False)
Call PropBag.WriteProperty("FontColorOnHover", plngFontColorOnHover, vbBlue)

Call PropBag.WriteProperty("BackColor", UserControl.BackColor, vb3DFace)
Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, vbButtonText)

Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)

Call PropBag.WriteProperty("Picture", ppicPicture, Nothing)
Call PropBag.WriteProperty("PictureAlignment", penumPictureAlignment, [Left_Middle])
Call PropBag.WriteProperty("PictureXOffset", plngPictureXOffset, 0)
Call PropBag.WriteProperty("PictureYOffset", plngPictureYOffset, 0)

Call PropBag.WriteProperty("AlwaysShowBorder", pboolAlwaysShowBorder, False)

Call PropBag.WriteProperty("ToolTip", pstrToolTip, "")
Call PropBag.WriteProperty("ToolTipTitle", pstrToolTipTitle, "")
Call PropBag.WriteProperty("ToolTipBackColor", plngToolTipBackColor, vbInfoBackground)
Call PropBag.WriteProperty("ToolTipForeColor", plngToolTipForeColor, vbInfoText)
Call PropBag.WriteProperty("ToolTipIcon", plngToolTipIcon, ToolTipIcon_None)
Call PropBag.WriteProperty("ToolTipStyle", plngToolTipStyle, ToolTipStyle_Standart)
'Call PropBag.WriteProperty("ToolTipCentered", pboolToolTipCentered, False)
Call PropBag.WriteProperty("Border3DStyle", plngBorder3DStyle, Border3DStyle_None)


Call PropBag.WriteProperty("GradientOrientation", proplngGradientOrientation, [GradientOrientation_Vertical])
Call PropBag.WriteProperty("UseGradientFill", propboolUseGradientFill, True)
Call PropBag.WriteProperty("GradientFromColor", proplngGradientFromColor, vb3DHighlight)
Call PropBag.WriteProperty("GradientToColor", proplngGradientToColor, vb3DShadow)
Call PropBag.WriteProperty("GradientInverseOnPress", propboolGradientInverseOnPress, True)

Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)

Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
 
 BackColor = UserControl.BackColor

End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
 
UserControl.BackColor() = New_BackColor
Sub_PrintCaption
Sub_DrawBorder enumBorderState
PropertyChanged "BackColor"

End Property


Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
 Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
 
Set UserControl.Font = New_Font
Sub_CalculateAlignment
Sub_PrintCaption
PropertyChanged "Font"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
 
Set Picture = ppicPicture
 
End Property

Public Property Set Picture(ByVal New_Picture As Picture)

On Error Resume Next

Set ppicPicture = New_Picture

Sub_CalculateAlignment
Sub_PrintCaption

PropertyChanged "Picture"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
 Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
 UserControl.Enabled() = New_Enabled
 PropertyChanged "Enabled"
End Property

