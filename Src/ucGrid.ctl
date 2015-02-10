VERSION 5.00
Begin VB.UserControl ucGrid 
   BackColor       =   &H0000C000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1740
   Enabled         =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   1740
   Begin VB.Line Line1 
      Index           =   3
      X1              =   960
      X2              =   960
      Y1              =   1200
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   480
      X2              =   480
      Y1              =   1200
      Y2              =   0
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   1560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   1560
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "ucGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Public Sub mReDraw(XDiv&, YDiv&)
  Dim X&, Y&, i&
  UserControl.ScaleMode = UserControl.Extender.Container.ScaleMode
  UserControl.Extender.Top = UserControl.Extender.Container.ScaleTop      'Fill the container
  UserControl.Extender.Left = UserControl.Extender.Container.ScaleLeft
  UserControl.Extender.Width = UserControl.Extender.Container.ScaleWidth
  UserControl.Extender.Height = Abs(UserControl.Extender.Container.ScaleHeight)
  UserControl.ScaleTop = UserControl.Extender.Container.ScaleTop      'Fill the container
  UserControl.ScaleLeft = UserControl.Extender.Container.ScaleLeft
  UserControl.ScaleWidth = UserControl.Extender.Container.ScaleWidth
  UserControl.ScaleHeight = UserControl.Extender.Container.ScaleHeight
  
  For i = Line1.UBound To 0 Step -1
    Line1(i).Visible = False
  Next
  i = 0
 ' If XDiv * YDiv <= 0 Then Exit Sub
  For X = XDiv To ScaleWidth Step XDiv
    If i > Line1.UBound Then Load Line1(i)
    Line1(i).Visible = True
    Line1(i).X1 = X
    Line1(i).X2 = X
    Line1(i).Y1 = 0
    Line1(i).Y2 = Abs(ScaleHeight)
    i = i + 1
  Next
  For Y = Abs(YDiv) To Abs(ScaleHeight) Step Abs(YDiv)
    If i > Line1.UBound Then Load Line1(i)
    Line1(i).Visible = True
    Line1(i).X1 = 0
    Line1(i).X2 = ScaleWidth
    Line1(i).Y1 = Y
    Line1(i).Y2 = Y
    i = i + 1
  Next
End Sub

