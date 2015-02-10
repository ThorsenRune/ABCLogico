VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   13695
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   5640
      Max             =   2000
      TabIndex        =   14
      Top             =   0
      Value           =   1000
      Width           =   3735
   End
   Begin VB.CommandButton oDebug 
      Caption         =   "STOP"
      Height          =   252
      Index           =   0
      Left            =   4200
      TabIndex        =   13
      Top             =   0
      Width           =   852
   End
   Begin VB.CommandButton oDebug 
      Caption         =   "STEP"
      Height          =   252
      Index           =   1
      Left            =   3120
      TabIndex        =   12
      Top             =   0
      Width           =   852
   End
   Begin VB.CommandButton oDebug 
      Caption         =   "GO"
      Height          =   252
      Index           =   2
      Left            =   2040
      TabIndex        =   11
      Top             =   0
      Width           =   852
   End
   Begin VB.CommandButton oAction 
      BackColor       =   &H00FFFF80&
      Caption         =   "Move"
      Height          =   375
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6720
      Width           =   975
   End
   Begin VB.PictureBox fProgram 
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3555
      ScaleWidth      =   1635
      TabIndex        =   9
      Top             =   360
      Width           =   1695
      Begin VB.Image oSP 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   0
         Picture         =   "Form1.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.CommandButton oAction 
      BackColor       =   &H00FFFF80&
      Caption         =   "Eat"
      Height          =   375
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton oAction 
      BackColor       =   &H00FFFF80&
      Caption         =   "Turn"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton oCondition 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Wall"
      Height          =   375
      Index           =   2
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton oCondition 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dangerous"
      Height          =   375
      Index           =   1
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton oCondition 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Eatable"
      Height          =   375
      Index           =   0
      Left            =   120
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5040
      UseMaskColor    =   -1  'True
      Width           =   975
   End
   Begin VB.PictureBox fField 
      BackColor       =   &H00C0E0FF&
      Height          =   6975
      Left            =   1920
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   480
      Width           =   11175
      Begin VB.Image oWall 
         Height          =   692
         Index           =   0
         Left            =   0
         Picture         =   "Form1.frx":04C0
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   1112
      End
      Begin VB.Image oFood 
         Height          =   692
         Index           =   0
         Left            =   0
         Picture         =   "Form1.frx":084E
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   1112
      End
      Begin VB.Image oDog 
         Height          =   692
         Index           =   0
         Left            =   0
         Picture         =   "Form1.frx":08D7
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1112
      End
      Begin VB.Image oCat 
         Height          =   692
         Left            =   0
         Picture         =   "Form1.frx":0FEA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1112
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "Actions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label Con4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Conditions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4680
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Enum eDir
  up = 0
  Left = 1
  down = 2
  Right = 3
End Enum
Const xMax = 20
Dim vDirection As eDir
Dim xPos&, yPos&, dx&, dy&
Public oResizer  As New cResizer
Public Instructions As New cSortedCollection
Enum eProgramstate
  eDB_stop = 0
  eDB_Step = 1
  eDB_Run = 2
End Enum
Private pProgramState As eProgramstate
Private oCollisionCheck As New cCollision
Dim pInstrPointer&
Dim Condition As Boolean          'flag for last test conditional execution
Enum eExeColor
  WillExec = vbGreen
  NoExec = vbRed
  CondTest = vbYellow
End Enum

 
Private Function decode$(c As Control)
  decode$ = c.Caption
End Function



Private Sub Sleep(secs!)
  Dim t!
    t = Timer + secs
    While t > Timer: DoEvents: Wend                 'Speed of execution
End Sub


Private Function vMaxY&()
  'Set number of height units to be proportional to field dimensions
 vMaxY& = (xMax / fField.Width) * fField.Height
End Function

Private Sub Form_DragDrop(Source As Control, x As Single, y As Single)
    'Copy this code into the dragDrop Event of the Target object
'   oResizer.StopDrag Source, Form1, X, Y
    If Source.Container Is fProgram Then    'It has been dragged away from the program
      Source.Visible = False 'Hide it for later deletion'
    End If
    mAlignCommands
 '   Source.ZOrder 0 'Add if you want the control moved to the front
End Sub
 
Function isKindOf(Obj1 As Variant, Obj2 As Variant) As Boolean
  Dim c(1) As Control
  isKindOf = False
  If Not IsObject(Obj1) Then
    Exit Function
  ElseIf Not IsObject(Obj2) Then
    Exit Function
  End If
  Set c(0) = Obj1
  Set c(1) = Obj2
  If c(0).Name = c(1).Name Then
    isKindOf = True
  Else
    isKindOf = False
  End If
End Function
Private Sub mExecute(c As Control, Condition As Boolean)
  If isKindOf(c, oCondition(0)) Then 'Test for conditions
    Select Case decode(c)
        Case "Eatable":
          Condition = isKindOf(oObjAt(fField, NextX, NextY), oFood(0))
        Case "Dangerous"
          Condition = (isKindOf(oObjAt(fField, NextX, NextY), oDog(0)))
        Case "Wall"
          Condition = Not mCanMoveTo(NextX, NextY)
        Case Else:
          Stop
      End Select
  ElseIf Condition Then     'Actions
    Select Case decode(c)
      Case "Move":
        mCatMove NextX, NextY     'Prepare next move
      Case "Turn":
        vDirection = (vDirection + 1) Mod 4
      Case "Eat":
        If isKindOf(oObjAt(fField, xPos, yPos), oFood(0)) Then
          oObjAt(fField, xPos, yPos).Visible = False
        End If
      Case Else:
        Stop
    End Select
  End If
  '      c.value = True     'Will fire the click event
End Sub
Function oObjAt(fField As Control, x&, y&)  ' return decode value for object at
  Dim c As Control
  oCat.Top = y: oCat.Left = x
  oCat.Visible = False
  For Each c In Controls
    If c.Visible = False Then
    ElseIf Not c.Container Is fField Then
    ElseIf c Is oCat Then
    ElseIf oCollisionCheck.mIsColliding(oCat, c) Then
      Set oObjAt = c
      Exit For
    End If
  Next
  oCat.Top = yPos: oCat.Left = xPos
  oCat.Visible = True
End Function
Private Function mCanMoveTo(x&, y&) As Boolean
  mCanMoveTo = False
  If x& >= xMax Then Exit Function
  If x& < 0 Then Exit Function
  If y& >= vMaxY& Then Exit Function
  If y& < 0 Then Exit Function
  mCanMoveTo = Not isKindOf(oObjAt(fField, x&, y&), oWall(0))
End Function
Private Sub mCatMove(x&, y&)
  If mCanMoveTo(x, y) Then
    oCat.Left = xPos&
    oCat.Top = yPos&
    xPos& = x&: yPos& = y&
  End If
End Sub
Private Function NextX&()
    Select Case vDirection
      Case Left: NextX = xPos& - 1
      Case Right:  NextX = xPos& + 1
      Case Else: NextX = xPos&
    End Select
End Function
Private Function NextY&()
    Select Case vDirection
      Case up: NextY = yPos& - 1
      Case down:  NextY = yPos& + 1
      Case Else: NextY = yPos&
    End Select
End Function
Private Function TestCollision&(TestObj As Control, OtherObj As Control) 'Return number of collisions
  Dim c As Control
    For Each c In Controls
      If TestObj.Name = OtherObj.Name Then
        If oCollisionCheck.mIsColliding(TestObj, c) Then TestCollision& = TestCollision& + 1
      End If
    Next
End Function
Private Sub mMove2Free(ctl As Control)  'Move to next free location
  Dim c As Control, i&
  For Each c In Controls
    If c Is ctl Then      'Skip myself
    ElseIf c.Container Is ctl.Container Then
      c.Left = Round(c.Left): c.Top = Round(c.Top)
      For i = 0 To 100
        If Not oCollisionCheck.mIsColliding(c, ctl) Then    'Ok and exit
          Exit For
        ElseIf ctl.Left + ctl.Width < ctl.Container.ScaleWidth Then
          ctl.Left = Round(ctl.Left + ctl.Width)
        ElseIf ctl.Top + ctl.Height < ctl.Container.ScaleHeight Then
          ctl.Left = 0
          ctl.Top = Round(ctl.Top + ctl.Height)
        Else
          ctl.Left = 0: ctl.Top = 0
        End If
      Next
    End If
  Next
End Sub
Private Sub LoadObjects(c As Control, Count&)
  Dim i&, o As Control
  mSetObjSize c
  For i = 0 To Count&
    Set o = oGetControlByIdx(c, i)
'    If o.Name = oBarrier(0).Name Then
'      o.Visible = True
'      o.vLength = 1 + Rnd() * 5
'      o.isHorizontal = (Rnd() > 0.5)
'      o.mPosX = Rnd() * xMax
'      o.mPosY = Rnd() * vMaxY&
'      o.Left = Rnd() * xMax - o.Width
'      o.Top = Rnd() * vMaxY& - o.Height
'      mMove2Free o
''    ElseIf o.Name = oDog(0).Name Then
''    ElseIf o.Name = oFood(0).Name Then
'    Else
      o.Left = Rnd() * xMax - o.Width
      o.Top = Rnd() * vMaxY& - o.Height
      mMove2Free o
'    End If
  Next
End Sub
Private Sub mSetObjSize(c As Control)
  c.Height = 1: c.Width = 1
End Sub

Private Sub Form_Load()
  fField.ScaleWidth = xMax
  fField.ScaleHeight = vMaxY&
  Show
  mSetObjSize oCat
  LoadObjects oFood(0), xMax / 5
  LoadObjects oDog(0), xMax / 5
  LoadObjects oWall(0), xMax
  vDirection = down
  xPos& = 0: yPos& = 0
  Condition = True
End Sub
Private Sub MainLoop()
  Static Instances&
  Dim oIP As Control
  If 0 < Instances Then Exit Sub
  Instances& = 1
  While pProgramState > eDB_stop
    mExecute Instructions.mItem(pInstrPointer), Condition
    If pProgramState = eDB_Step Then pProgramState = eDB_stop     'Single step
    pInstrPointer& = (pInstrPointer& + 1) Mod (Instructions.mUBound + 1) 'Advance program pointer
    If pInstrPointer = 0 Then Condition = True 'Unconditional execution of first command
    Set oIP = Instructions.mItem(pInstrPointer)
    oSP.Top = oIP.Top
    If isKindOf(oIP, oAction(0)) Then
      If Condition Then
        oIP.BackColor = eExeColor.WillExec
      Else
        oIP.BackColor = eExeColor.NoExec
      End If
    Else
        oIP.BackColor = eExeColor.CondTest
    End If
    Sleep HScroll1.value / 1000
  Wend
  Instances& = 0
End Sub

Private Sub CmdsEnable(ThisContainer As Control, value As Boolean)
'Disable/enable all controls. WA to let program container receive mouse moves directly
  Dim c As Control
  For Each c In Form1
    If c.Container Is ThisContainer Then
      If (ActiveControl Is c) Then
        c.Enabled = True
      Else
        c.Enabled = value
      End If
    End If
  Next
'  Dim i&
'    For i = 0 To Instructions.mUBound
'      If (ActiveControl Is Instructions.mItem(i)) Then
'      Else
'        Instructions.mItem(i).Enabled = value
'      End If
'    Next
    
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Not pProgramState = eDB_stop Then
    Cancel = True
    pProgramState = eDB_stop
  End If
End Sub

Private Sub oAction_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   oResizer.StartDrag oAction(Index), Button, x, y
     CmdsEnable fProgram, False

End Sub

Private Sub oCondition_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   'Copy this code into the MouseMOVE Event of the drag object
  If oResizer.StartDrag(oCondition(Index), Button, x, y) Then
    CmdsEnable fProgram, False
  End If
End Sub
Function isCommand(c As Control) As Boolean
Attribute isCommand.VB_Description = "Pinko pallino"
  If c.Container.Name <> fProgram.Name Then          'Not placed in the program
    isCommand = False
  ElseIf c.Visible = False Then                     'Not visible = not active
    isCommand = False
  ElseIf c.Name = oCondition(0).Name Then
    isCommand = True
  ElseIf c.Name = oAction(0).Name Then
    isCommand = True
  Else
    isCommand = False
  End If
End Function

Public Sub mAlignCommands()     'Sort of a compile command
  Dim margin&, grid&, Nr&, i&, vTab&
  Dim c As Control
  vTab& = oSP.Width         'Tabulator size
  margin = oCondition(0).Height / 10
  grid& = oCondition(0).Height + margin
  Nr& = 0         'Control number
  Instructions.mClear               'Make an ordered list of the controls
  For Each c In Form1.Controls
    If isCommand(c) Then
      Instructions.mAdd c, c.Top + (c.Left / ScaleWidth)  'create a fractional order index leftmost becomes prioritized
    End If
  Next
  For i = 0 To Instructions.mUBound     'Align them in order on the screen
    Set c = Instructions.mItem(i)
    c.Top = margin + i * grid
    c.Left = vTab&
    If c.Name = oAction(0).Name Then c.Left = c.Left + vTab&
  Next
  CmdsEnable fProgram, True
End Sub

Private Sub oDebug_Click(Index As Integer)
'0=stop,1=step,2=run
  pProgramState = Index
  MainLoop
End Sub

Private Sub fField_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  oResizer.MouseMove x, y
End Sub

Private Sub fField_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  oResizer.StopSizing
End Sub

Function oGetControlByIdx(Source As Control, Nr&) As Control     'make sure the control exist, otherwise load it
 'Loads a control if it does not exist
  Dim o As Object, c As Control
    If Source.Name = oCondition(0).Name Then
      Set o = oCondition
    ElseIf Source.Name = oAction(0).Name Then
      Set o = oAction
    ElseIf Source.Name = oFood(0).Name Then
      Set o = oFood
    ElseIf Source.Name = oDog(0).Name Then
      Set o = oDog
    ElseIf Source.Name = oWall(0).Name Then
      Set o = oWall
    Else
      Stop
    End If
    If o Is Nothing Then Stop
    While o.UBound < Nr
      Load o(o.UBound + 1)
      Set c = o(o.UBound)
      c.Top = Source.Top
      c.Visible = Source.Visible
    Wend
    Set oGetControlByIdx = o(Nr)
End Function

Function oMakeCopy(Source As Control) As Control
  Dim o As Object
    If Source.Name = oCondition(0).Name Then
      Set o = oCondition
    ElseIf Source.Name = oAction(0).Name Then
      Set o = oAction
    Else
      Stop
    End If
   If Not o Is Nothing Then
      Load o(o.UBound + 1)
      Set oMakeCopy = o(o.UBound)
      oMakeCopy.Caption = Source.Caption
      oMakeCopy.Visible = True
  End If
End Function
Private Sub fprogram_DragDrop(Source As Control, x As Single, y As Single)
    'Copy this code into the dragDrop Event of the Target object
    If Not Source.Container Is fProgram Then
      Set Source = oMakeCopy(Source)
    End If
    oResizer.StopDrag Source, fProgram, x, y
    mAlignCommands
End Sub

Private Sub fprogram_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Caption = x & ":" & y

End Sub
