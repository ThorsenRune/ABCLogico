VERSION 5.00
Begin VB.Form fConsole 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   6780
   ClientLeft      =   8880
   ClientTop       =   2292
   ClientWidth     =   7152
   ControlBox      =   0   'False
   Icon            =   "fConsole.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1121.032
   ScaleMode       =   0  'User
   ScaleWidth      =   7152
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox fField 
      Align           =   1  'Align Top
      BackColor       =   &H0000C000&
      Height          =   5532
      Left            =   0
      Picture         =   "fConsole.frx":208EA
      ScaleHeight     =   274.2
      ScaleMode       =   2  'Point
      ScaleWidth      =   355.2
      TabIndex        =   0
      Top             =   0
      Width           =   7152
      Begin ABCLogico.ucGrid ucGrid1 
         Height          =   1332
         Left            =   2160
         TabIndex        =   1
         Top             =   1920
         Width           =   1452
         _ExtentX        =   2561
         _ExtentY        =   2350
      End
      Begin VB.Image oProbe 
         Height          =   600
         Left            =   0
         Picture         =   "fConsole.frx":2B1EB
         Stretch         =   -1  'True
         Top             =   3720
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Image oCat 
         Height          =   705
         Index           =   3
         Left            =   2160
         Picture         =   "fConsole.frx":2B6B5
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Image oCat 
         Height          =   705
         Index           =   2
         Left            =   1440
         Picture         =   "fConsole.frx":2CEC5
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Image oCat 
         Height          =   705
         Index           =   1
         Left            =   720
         Picture         =   "fConsole.frx":2E6B1
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Image oTarget 
         Height          =   700
         Index           =   0
         Left            =   0
         Picture         =   "fConsole.frx":2FEC8
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   700
      End
      Begin VB.Image oWall 
         Height          =   700
         Index           =   0
         Left            =   0
         Picture         =   "fConsole.frx":30303
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   700
      End
      Begin VB.Image oFood 
         Height          =   700
         Index           =   0
         Left            =   0
         Picture         =   "fConsole.frx":3145F
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   700
      End
      Begin VB.Image oDog 
         Height          =   700
         Index           =   0
         Left            =   0
         Picture         =   "fConsole.frx":32FC8
         Stretch         =   -1  'True
         Top             =   720
         Width           =   700
      End
      Begin VB.Image oCat 
         Height          =   700
         Index           =   0
         Left            =   0
         Picture         =   "fConsole.frx":336DB
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   700
      End
   End
   Begin ABCLogico.ucStatusBar ucStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   372
      Left            =   0
      TabIndex        =   2
      Top             =   6408
      Width           =   7152
      _ExtentX        =   10012
      _ExtentY        =   445
      Value           =   "Cat position"
   End
End
Attribute VB_Name = "fConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const vGridDivisions = 10     'Number of vertical cells
Private oCollisionCheck As New cCollision
Public Function nCellsX&()
  nCellsX& = vGridDivisions * fField.Width / fField.Height
  If nCellsX < 1 Then nCellsX = 1
End Function
Private Function nCellsY&()
  'Set number of height units to be proportional to field dimensions
  nCellsY& = vGridDivisions
End Function

Function oObjAt1(oTestObj, X&, Y&) As Control  ' return object of type otestobj at pos, null if not found
' oTestObj must be a control array
  Dim c As Control
  oProbe.Visible = True
  oProbe.Move X, Y + 1, oTestObj(0).Width, oTestObj(0).Height
  oProbe.Visible = True
  Set oObjAt1 = Nothing
  For Each c In oTestObj
    If c.Visible = False Then
    ElseIf oCollisionCheck.mIsColliding(oProbe, c) Then
      Set oObjAt1 = c
      Exit For
    End If
  Next
  oProbe.Visible = False
End Function
Public Function mCanMoveTo(X&, Y&) As Boolean
'Can the cat go to this position? i.e. no walls or border limits
  mCanMoveTo = False
  If X& >= nCellsX& Then
    Exit Function
  ElseIf X& < 0 Then
    Exit Function
  ElseIf Y& >= nCellsY& Then
    Exit Function
  ElseIf Y& < 0 Then
    Exit Function
  End If
  mCanMoveTo = oObjAt1(oWall, X&, Y&) Is Nothing      'NO wall here
End Function
Public Sub mCatGoto(X&, Y&)   'Move the cat
    vCatPosX = X: vCatPosY = Y
End Sub



Private Sub mMove2Free(Ctl As Control)  'Move to next free location
  Dim c As Control, i&
 
  For Each c In Controls
    If c Is Ctl Then      'Skip myself
    ElseIf c.Container Is Ctl.Container Then
      c.Left = Round(c.Left): c.Top = Round(c.Top)
      For i = 0 To 100
        If Not oCollisionCheck.mIsColliding(c, Ctl) Then    'Ok and exit
          Exit For
        End If
        Ctl.Left = Round(Ctl.Left + 1)
        If Ctl.Left >= Ctl.Container.ScaleWidth - Ctl.Width Then
          Ctl.Left = 1
          Ctl.Top = Round(Ctl.Top + 1)
        End If
        If Ctl.Top >= vGridDivisions Then
          Ctl.Left = 1
          Ctl.Top = 1
        End If
      Next
    End If
  Next
End Sub
Private Sub LoadObjects(c As Control, Count&)
'Make sure the control is loaded
  Dim i&, o As Control
  For i = 0 To Count& - 1
    Set o = oGetControlByIdx(c, i)
    mSetObjSize o
    o.Left = Round(Rnd() * (o.Container.ScaleWidth - o.Width))
    o.Top = Round(Rnd() * (vGridDivisions))
    o.Visible = True
    mMove2Free o
  Next
  For Each o In Controls 'Unload excessive controls
    If o.Name = c.Name Then
      If o.Index >= Count Then
        If Count > 0 Then Unload o
      End If
    End If
  Next
End Sub
Private Sub mSetObjSize(c As Control)
  c.Height = 1: c.Width = 1
End Sub
Public Sub ReGenerate()
  fField.Cls
  Dim X&, Y&
'  fField.ScaleWidth = vGridDivisions * fField.Width / fField.Height
  LoadObjects oCat(0), 4    'A catpicture for each of four directions to go
  vDirection = DirRight     'Put the cat on board
  LoadObjects oTarget(0), 1
  LoadObjects oFood(0), fSettings.oFoods.Text
  LoadObjects oDog(0), fSettings.oDogs.Text
  LoadObjects oWall(0), fSettings.oWalls.Text
  OnResize
End Sub
Public Sub mReset()         'Unhide all
  Me.Show
  Dim c As Control
    OnResize
    fConsole.Settings True, sFilePrg  'Load cat/dogs scenary
    vCatPosX = 0: vCatPosY = 0: vDirection = DirRight     'Reset cat
    For Each c In oDog
      c.Visible = True
    Next
    For Each c In oFood
      c.Visible = True
    Next
End Sub
 
 
Public Sub Settings(GetIt As Boolean, FileName$)
  If FileName$ = "" Then Stop 'FileName$ = sPathBin & "\DataFile.txt"
  mCtrlSetting GetIt, FileName$, fField, oCat, "CAT", 0
  mCtrlSetting GetIt, FileName$, fField, oDog, "DOG", 0
  mCtrlSetting GetIt, FileName$, fField, oFood, "FOOD", 0
  mCtrlSetting GetIt, FileName$, fField, oWall, "WALL", 0
  mCtrlSetting GetIt, FileName$, fField, oTarget, "TARGET", 0
  RunDebug 555
End Sub

Private Sub fField_DragDrop(Source As Control, X As Single, Y As Single)
  oResizer.StopDrag Source, fField, X, Y
  mAlign2Grid Source, fField
  If Source Is oCat(vDirection) Then vCatPosX = oCat(vDirection).Left: vCatPosY = oCat(vDirection).Top - 1
End Sub
Private Sub mAlign2Grid(Source As Control, Target As Control)
  Source.Top = CInt(Source.Top)
  Source.Left = CInt(Source.Left)
  Source.Width = 1
  Source.Height = 1
End Sub

Private Sub fField_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   fConsole.ucStatusBar1.Value = fField.ToolTipText & ": x=" & Fix(X) & ",  y=" & Fix(Y)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF2 Then frmMsg.msgEdit ActiveControl
End Sub

Private Sub Form_Load()
   Show
  oIniFile.WinState True, Me
  Me.Height = MDIForm1.ScaleHeight: Me.Top = 0
'''  Settings True, ""
  mCaptionTexts True, Me
  KeyPreview = True
End Sub
Private Sub OnResize()
    Dim c As Control
  If Me.WindowState = vbMinimized Then Exit Sub
  fField.Height = ucStatusBar1.Top

  Me.ScaleMode = vbPixels
  If Me.Height > MDIForm1.ScaleHeight Then Me.Height = MDIForm1.ScaleHeight
  If fField.Width < 1 Then Exit Sub
  fField.Scale (0, nCellsY&)-(nCellsX, 0)

  For Each c In Controls
    If Not CtrlArryByName(c.Name) Is Nothing Then
      mAlign2Grid c, c.Container
    End If
  Next
  'Draw a grid
  ucGrid1.mReDraw 1, 1
End Sub
Private Sub fField_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  oResizer.MouseMove X, Y
  frmMsg.mMsgOnMouse fField, Button, Shift
End Sub

Private Sub fField_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  oResizer.StopSizing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnMouse Me, Button, Shift
End Sub

Private Sub Form_Resize()
  OnResize
End Sub


Private Sub Form_Unload(Cancel As Integer)
''  Settings False, ""
  oIniFile.WinState False, Me
End Sub

Private Sub oCat_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusRefresh oCat(Index)
  oResizer.StartDrag oCat(Index), Button, X, Y
End Sub

Private Sub oCat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnMouse oCat, Button, Shift
End Sub

Private Sub oDog_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
oDog(Index).Enabled = False
End Sub

Private Sub oDog_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusRefresh oDog(Index)
  oResizer.StartDrag oDog(Index), Button, X, Y
End Sub

Private Sub oDog_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnMouse oDog(0), Button, Shift
End Sub

Private Sub oFood_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
   oFood(Index).Enabled = False
End Sub

Private Sub oFood_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusRefresh oFood(Index)
  oResizer.StartDrag oFood(Index), Button, X, Y
End Sub

Private Sub oFood_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 frmMsg.mMsgOnMouse oFood(0), Button, Shift
End Sub

Private Sub oTarget_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   StatusRefresh oTarget(Index)
   oResizer.StartDrag oTarget(Index), Button, X, Y
'   oTarget(Index).Enabled = False
End Sub
Private Sub oTarget_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnMouse oTarget(0), Button, Shift
End Sub

Private Sub oWall_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  StatusRefresh oWall(Index)
  oResizer.StartDrag oWall(Index), Button, X, Y

End Sub

Private Sub oWall_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnMouse oWall(0), Button, Shift
End Sub


