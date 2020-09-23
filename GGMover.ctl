VERSION 5.00
Begin VB.UserControl GGMover 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   405
   InvisibleAtRuntime=   -1  'True
   Picture         =   "GGMover.ctx":0000
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   27
   ToolboxBitmap   =   "GGMover.ctx":03D8
End
Attribute VB_Name = "GGMover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Moving As Boolean
Attribute Moving.VB_VarDescription = "This event occur each time when coordinates of object are changed and MovingEventPresent is set to True.."

Dim CoordinateXWhenMouseDown As Single
Dim CoordinateYWhenMouseDown As Single

Dim XCoordinate As Single
Dim YCoordinate As Single

Event Moving(Target As Object)
'Default Property Values:
Const m_def_MovingEventPresent = False
'Property Variables:
Dim m_MovingEventPresent As Boolean


Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
MsgBox "GGMover user control - Georgi Yordanov Ganchev ©2000 - GogoX@Lycos.com - http://georgi-ganchev.tripod.com", vbInformation, "GGMover user control"
End Sub


Public Sub CaptureObject(CoordinateX As Single, CoordinateY As Single)

CoordinateXWhenMouseDown = CoordinateX
CoordinateYWhenMouseDown = CoordinateY
Moving = True

End Sub


Public Sub ReleaseTargetObject()

Moving = False

End Sub


Public Sub MoveObject(Target As Object, XWhenMouseMove As Single, YWhenMouseMove As Single, Optional MoveByX As Boolean = True, Optional MoveByY As Boolean = True, Optional MinX As Long = 0, Optional MaxX As Long = 0, Optional MinY As Single = 0, Optional MaxY As Single = 0)

If Moving = True Then

If MoveByX Then XCoordinate = Target.Left + XWhenMouseMove - CoordinateXWhenMouseDown
If MoveByY Then YCoordinate = Target.Top + YWhenMouseMove - CoordinateYWhenMouseDown

If XCoordinate < MinX And (MinX = MaxX = 0) Then XCoordinate = MinX
If XCoordinate > MaxX And (MinX = MaxX = 0) Then XCoordinate = MaxX

If YCoordinate < MinY And (MinY = MaxY = 0) Then YCoordinate = MinY
If YCoordinate > MaxY And (MinY = MaxY = 0) Then YCoordinate = MaxY

If MoveByX Then Target.Left = XCoordinate
If MoveByY Then Target.Top = YCoordinate

If m_MovingEventPresent = True Then
 RaiseEvent Moving(Target)
End If


End If

End Sub

Private Sub UserControl_Resize()

UserControl.Size 405, 405

End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get MovingEventPresent() As Boolean
    MovingEventPresent = m_MovingEventPresent
End Property

Public Property Let MovingEventPresent(ByVal New_MovingEventPresent As Boolean)
    m_MovingEventPresent = New_MovingEventPresent
    PropertyChanged "MovingEventPresent"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_MovingEventPresent = m_def_MovingEventPresent
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_MovingEventPresent = PropBag.ReadProperty("MovingEventPresent", m_def_MovingEventPresent)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("MovingEventPresent", m_MovingEventPresent, m_def_MovingEventPresent)
End Sub

