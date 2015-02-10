VERSION 5.00
Begin VB.UserControl cBarrier 
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   885
   FillStyle       =   0  'Solid
   ScaleHeight     =   330
   ScaleWidth      =   885
   Begin VB.Shape Shape1 
      BorderStyle     =   2  'Dash
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   132
      Left            =   120
      Top             =   120
      Width           =   732
   End
End
Attribute VB_Name = "cBarrier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private propLength&, propHorizontal As Boolean

Public Property Let isHorizontal(Val As Boolean)
  propHorizontal = Val
  vLength = propLength
  mRedraw
End Property
Public Property Let vLength(Length&)
  propLength = Length&
  mRedraw
End Property
Public Property Let mPosX(x&)
  UserControl.Extender.Left = x - UserControl.Extender.Width / 2
  mRedraw
End Property
Public Property Let mPosY(y&)
  UserControl.Extender.Top = y - UserControl.Extender.Height / 2
  mRedraw
End Property
Public Sub mRedraw()
  Dim dimA!, dimB!
  dimA! = propLength * 1.2
  dimB! = 1.2
  If propHorizontal Then
    UserControl.Extender.Width = dimA
    UserControl.Extender.Height = dimB
  Else
    UserControl.Extender.Width = dimB
    UserControl.Extender.Height = dimA
  End If
  UserControl.ScaleTop = -0.1
  UserControl.ScaleHeight = 1.2
  UserControl.ScaleLeft = -0.1
  UserControl.ScaleWidth = 1.2
  Shape1.Left = 0
  Shape1.Width = 1
  Shape1.Top = 0
  Shape1.Height = 1
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Form1.oResizer.StartSizing UserControl.Extender
End Sub
Public Property Get Enabled() As Boolean
   Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
   UserControl.Enabled = NewValue
End Property

Public Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 9 Then
    mPosX = x
    mPosY = y
  End If
End Sub

