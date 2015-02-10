VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCollision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'The "RECT" type required by the IntersectRect API call
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'Our API calls
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

 
Public Function mIsColliding(c1 As Control, c2 As Control) As Boolean     'RT2015 returns true if two contros are colliding
  Dim rect1 As RECT
  Dim rect2 As RECT
  Dim udtTempRect As RECT     'The IntersectRect call will return a rectangle equal in size to the intersection between our two rectangles... but we don't really need this data here
    rect1.Top = c1.Top
    rect1.Left = c1.Left
    rect1.Bottom = c1.Top + c1.Height
    rect1.Right = c1.Left + c1.Width
    rect2.Top = c2.Top
    rect2.Left = c2.Left
    rect2.Bottom = c2.Top + c2.Height
    rect2.Right = c2.Left + c2.Width
    'Check for rectangle collision
    mIsColliding = IntersectRect(udtTempRect, rect1, rect2)

End Function
Public Function mIsControlOver(c1 As Control, X&, Y&) As Boolean    'RT2015 returns true if two contros are colliding
  Dim rect1 As RECT
  Dim rect2 As RECT
  Dim udtTempRect As RECT     'The IntersectRect call will return a rectangle equal in size to the intersection between our two rectangles... but we don't really need this data here
    rect1.Top = c1.Top
    rect1.Left = c1.Left
    rect1.Bottom = c1.Top + c1.Height
    rect1.Right = c1.Left + c1.Width
    rect2.Left = X
    rect2.Top = Y
    rect2.Bottom = Y + c1.Height
    rect2.Right = X + c1.Width
    'Check for rectangle collision
    mIsControlOver = IntersectRect(udtTempRect, rect1, rect2)
End Function
 
'' TIP: Did you know you can fire the click event programatically. dim c as control. set c=ClickButton. c.value=true will fire the event
