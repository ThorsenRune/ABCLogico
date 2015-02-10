VERSION 5.00
Begin VB.Form fProgram 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   588
   ClientWidth     =   9228
   ControlBox      =   0   'False
   Icon            =   "fProgram.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   9228
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   732
      Index           =   4
      Left            =   1560
      ScaleHeight     =   732
      ScaleWidth      =   2052
      TabIndex        =   24
      Top             =   0
      Width           =   2055
      Begin VB.ComboBox cbProgName 
         Height          =   315
         ItemData        =   "fProgram.frx":594A
         Left            =   0
         List            =   "fProgram.frx":5963
         TabIndex        =   26
         Text            =   "A"
         Top             =   240
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         Height          =   732
         Index           =   1
         Left            =   1320
         ScaleHeight     =   684
         ScaleWidth      =   564
         TabIndex        =   25
         Top             =   0
         Width           =   615
         Begin VB.Image cmdSave 
            Height          =   588
            Left            =   0
            Picture         =   "fProgram.frx":597C
            Stretch         =   -1  'True
            ToolTipText     =   "Salva percorso e programma"
            Top             =   0
            Width           =   612
         End
      End
      Begin VB.Label label 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Programma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   1332
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   612
      Index           =   2
      Left            =   0
      ScaleHeight     =   564
      ScaleWidth      =   684
      TabIndex        =   21
      Top             =   0
      Width           =   732
      Begin VB.Image cmdHome 
         Height          =   588
         Left            =   0
         Picture         =   "fProgram.frx":5D27
         Stretch         =   -1  'True
         Top             =   0
         Width           =   612
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   612
      Index           =   1
      Left            =   4680
      ScaleHeight     =   564
      ScaleWidth      =   684
      TabIndex        =   20
      Top             =   0
      Width           =   732
      Begin VB.Image Image1 
         Height          =   585
         Left            =   0
         Picture         =   "fProgram.frx":6162
         Stretch         =   -1  'True
         Top             =   0
         Width           =   630
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   612
      Index           =   0
      Left            =   3840
      ScaleHeight     =   564
      ScaleWidth      =   684
      TabIndex        =   19
      Top             =   0
      Width           =   732
      Begin VB.Image cmdSettings 
         Height          =   588
         Left            =   0
         Picture         =   "fProgram.frx":7BB4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   636
      End
   End
   Begin VB.PictureBox fProgramListContainer 
      BackColor       =   &H00E0E0E0&
      Height          =   6012
      Left            =   1920
      ScaleHeight     =   5964
      ScaleWidth      =   4164
      TabIndex        =   2
      Top             =   840
      Width           =   4215
      Begin VB.PictureBox fProgramList 
         Height          =   3615
         Left            =   0
         ScaleHeight     =   3564
         ScaleWidth      =   2844
         TabIndex        =   4
         Top             =   840
         Width           =   2892
         Begin VB.Image oSP 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   0
            Picture         =   "fProgram.frx":8EB5
            Stretch         =   -1  'True
            Top             =   0
            Width           =   255
         End
         Begin VB.Label lLnNr 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lnr"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   480
            Width           =   240
         End
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3492
         Left            =   3120
         TabIndex        =   3
         Top             =   840
         Width           =   252
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Programma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame fElements 
      BorderStyle     =   0  'None
      Caption         =   "Elementi del programma"
      Height          =   7215
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   1935
      Begin ABCLogico.ucOpCode oInstr 
         Height          =   252
         Index           =   10
         Left            =   240
         TabIndex        =   23
         Tag             =   " "
         Top             =   5040
         Width           =   1092
         _ExtentX        =   1926
         _ExtentY        =   445
         BackColor       =   12640511
         Caption         =   "Call"
         vDescription    =   "Description"
         ToolTip         =   ""
         vOPCode         =   "Call"
         IfCond          =   "0"
         IsFixed         =   "1"
      End
      Begin VB.PictureBox Picture1 
         Height          =   732
         Index           =   3
         Left            =   480
         ScaleHeight     =   684
         ScaleWidth      =   684
         TabIndex        =   22
         Tag             =   " "
         Top             =   2760
         Width           =   732
         Begin VB.Image oBin 
            Height          =   708
            Left            =   0
            Picture         =   "fProgram.frx":9375
            Stretch         =   -1  'True
            Top             =   0
            Width           =   708
         End
      End
      Begin ABCLogico.ucOpCode oInstr 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Tag             =   " "
         Top             =   1320
         Width           =   1095
         _ExtentX        =   3196
         _ExtentY        =   868
         Caption         =   "Giù v"
         vDescription    =   "ToolTip"
         ToolTip         =   ""
         vOPCode         =   "Down"
         IfCond          =   "0"
         IsFixed         =   "1"
      End
      Begin ABCLogico.ucOpCode oInstr 
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Tag             =   " "
         Top             =   600
         Width           =   1092
         _ExtentX        =   3196
         _ExtentY        =   868
         Caption         =   "Avanti"
         vDescription    =   "ToolTip"
         ToolTip         =   ""
         vOPCode         =   "Move"
         IfCond          =   "0"
         IsFixed         =   "1"
      End
      Begin ABCLogico.ucOpCode oInstr 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Tag             =   " "
         Top             =   960
         Width           =   1095
         _ExtentX        =   3196
         _ExtentY        =   868
         Caption         =   "Su  ^"
         vDescription    =   "Description"
         vOPCode         =   "Up"
         IfCond          =   "0"
         IsFixed         =   "1"
      End
      Begin ABCLogico.ucOpCode oInstr 
         Height          =   255
         HelpContextID   =   2
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Tag             =   " "
         Top             =   1680
         Width           =   1095
         _ExtentX        =   3196
         _ExtentY        =   868
         Caption         =   "Destra -->"
         vDescription    =   "Description"
         ToolTip         =   ""
         vOPCode         =   "Right"
         IfCond          =   "0"
         IsFixed         =   "1"
      End
      Begin ABCLogico.ucOpCode oInstr 
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   12
         Tag             =   " "
         Top             =   2040
         Width           =   1095
         _ExtentX        =   3196
         _ExtentY        =   868
         Caption         =   "<-- Sinistra"
         vDescription    =   "Description"
         ToolTip         =   ""
         vOPCode         =   "Left"
         IfCond          =   "0"
         IsFixed         =   "1"
      End
      Begin ABCLogico.ucOpCode oInstr 
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   13
         Tag             =   " "
         Top             =   2400
         Width           =   1095
         _ExtentX        =   3196
         _ExtentY        =   868
         Caption         =   "Mangia"
         vDescription    =   "Description"
         vOPCode         =   "Eat"
         IfCond          =   "0"
         IsFixed         =   "1"
      End
      Begin ABCLogico.ucOpCode oInstr 
         Height          =   252
         Index           =   6
         Left            =   240
         TabIndex        =   14
         Tag             =   " "
         Top             =   4320
         Width           =   1092
         _ExtentX        =   3196
         _ExtentY        =   868
         BackColor       =   12648447
         Caption         =   "Cibo?"
         vDescription    =   "ToolTip"
         ToolTip         =   ""
         vOPCode         =   "Eatable"
         IfCond          =   "-1"
         IsFixed         =   "1"
      End
      Begin ABCLogico.ucOpCode oInstr 
         Height          =   252
         Index           =   7
         Left            =   240
         TabIndex        =   15
         Tag             =   " "
         ToolTipText     =   "Some good text"
         Top             =   3600
         Width           =   1092
         _ExtentX        =   3196
         _ExtentY        =   868
         BackColor       =   12648447
         Caption         =   "Cane?"
         vDescription    =   "ToolTip"
         ToolTip         =   ""
         vOPCode         =   "Dangerous"
         IfCond          =   "-1"
         IsFixed         =   "1"
      End
      Begin ABCLogico.ucOpCode oInstr 
         Height          =   252
         Index           =   8
         Left            =   240
         TabIndex        =   16
         Tag             =   " "
         Top             =   3960
         Width           =   1092
         _ExtentX        =   3196
         _ExtentY        =   868
         BackColor       =   12648447
         Caption         =   "Ostacolo"
         vDescription    =   "ToolTip"
         ToolTip         =   ""
         vOPCode         =   "Wall"
         IfCond          =   "-1"
         IsFixed         =   "1"
      End
      Begin ABCLogico.ucOpCode oInstr 
         Height          =   252
         HelpContextID   =   2
         Index           =   9
         Left            =   240
         TabIndex        =   17
         Tag             =   " "
         Top             =   4680
         Width           =   1092
         _ExtentX        =   3196
         _ExtentY        =   868
         BackColor       =   12640511
         Caption         =   "Ripeti"
         vDescription    =   "ToolTip"
         ToolTip         =   ""
         vOPCode         =   "ForNext"
         IfCond          =   "0"
         IsFixed         =   "1"
      End
      Begin ABCLogico.ucOpCode oProcedure 
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Tag             =   " "
         Top             =   5400
         Width           =   1092
         _ExtentX        =   3196
         _ExtentY        =   868
         BackColor       =   12640511
         Caption         =   "{      }"
         vDescription    =   "Procedure "
         ToolTip         =   ""
         vOPCode         =   "Sub"
         IfCond          =   "0"
         IsFixed         =   "1"
      End
      Begin VB.Label label 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Commandi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdStep 
      Caption         =   ">"
      Default         =   -1  'True
      DownPicture     =   "fProgram.frx":AE93
      Height          =   645
      Left            =   720
      Picture         =   "fProgram.frx":36F5B
      TabIndex        =   0
      Tag             =   "STEP"
      ToolTipText     =   "Fai un passo"
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "fProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nIfCond As Boolean          'flag for last test conditional execution
Dim propInstrPointer&
Public Instructions As New cSortedCollection
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal _
     hWndNewParent As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
     (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Sub MoveControl(Child As Control, nparent As Form)
  Dim retval As Long
  retval = SetParent(Child.hWnd, nparent.hWnd)
  'Retval is previous parents handle
End Sub
Public Sub mScrollInto(yPos&)
  'Scroll into
  If yPos < -fProgramList.Top Then
    fProgramList.Top = yPos
  ElseIf yPos > (0.7 * fProgramListContainer.ScaleHeight) Then
    fProgramList.Top = (0.7 * fProgramListContainer.ScaleHeight - yPos)
  End If
End Sub

Private Sub cbProgName_Click()
  sFilePrg$ = sPathData$ + "\ABCL__" & fSettings.vProgName & ".DAT"
  If ActiveControl Is cbProgName Then UserProgram_Load
End Sub

Private Sub cmdHome_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    frmMsg.mMsgByTagStr "Reset"
    mMainReset        'Reset the program
  End If
End Sub
Private Sub mRun()
  If pProgramState = eDB_Run Then
    ExecSpeed = ExecSpeed / 2
  ElseIf pProgramState = eDB_Break Then
    pProgramState = eDB_Run
  End If
End Sub
Public Sub UserProgram_Load()
    mClearAll
    sFilePrg$ = sPathData$ + "\ABCL__" & fSettings.vProgName & ".DAT"
    Settings True, sFilePrg            'Load programe
    mMainReset
    MDIForm1.Caption = "ABC Logico  " & "Program " & fSettings.vProgName & "   " & sLang & "                                             ver 2015/R.Thorsen"
End Sub

 

 
 
 

Private Sub cmdSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then
  mUserProg_Save
 End If
End Sub

Private Sub cmdSettings_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then fSettings.Show: fSettings.SetFocus
End Sub

Private Sub cmdSettings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnMouse cmdSettings, Button, Shift
End Sub

Private Sub mStep()
  Static T!
  If T + 1 > Timer Then
    mRun   'Double click
  ElseIf pProgramState = eDB_Break Then 'In breakmode Make a step
    pProgramState = eDB_Step
  ElseIf pProgramState = eDB_stop Then  'Restart program
    pProgramState = eDB_Step
    mMainSub        'Start program
  End If
  T = Timer
End Sub

Private Sub cmdStep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    mStep
  Else
    frmMsg.mMsgOnMouse cmdStep, Button, Shift, X, Y
  End If
End Sub

Private Sub fElements_DragDrop(Source As Control, X As Single, Y As Single)
  oBin_DragDrop Source, X, Y
End Sub

Private Sub fElements_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMsg.mMsgOnMouse fElements, Button, Shift, X, Y
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF2 Then frmMsg.msgEdit ActiveControl
 If KeyCode = vbKeyDelete Then oBin_DragDrop ActiveControl, 0, 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMsg.mMsgOnMouse "ProgDescr", Button, Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Not pProgramState = eDB_stop Then
    Cancel = True
    pProgramState = eDB_stop
  End If
  mUserProg_Save
End Sub
Private Sub Form_Resize()
  OnResize
  DivX = Width / MDIForm1.ScaleWidth
End Sub

Public Sub Form_Unload(Cancel As Integer)
  Dim f As Form
    oIniFile.WinState False, Me
    For Each f In Forms
      Unload f
    Next
End Sub

Private Sub fProgramList_DragDrop(Source As Control, X As Single, Y As Single)
  If Source.IsSub Then     'Only drop subs
    If Source.IsFixed Then Set Source = fProgram.oMakeCopy(Source)
    Set Source.Container = fProgramList
    Source.Visible = True
    Source.Caption = "s" & Source.Index
    Source.sOperand = "s" & Source.Index
  End If
  mAlignCommands
End Sub
Private Sub fProgramList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMsg.mMsgOnMouse fProgramList, Button, Shift
End Sub
Private Sub fProgramListContainer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     frmMsg.mMsgOnMouse fProgramList, Button, Shift
End Sub

Public Sub Settings(GetIt As Boolean, FileName$)
Dim i&
Dim c As Control
  Dim InstanceNr&
  If FileName$ = "" Then MsgBox "Missing filename": Stop
  mCtrlSetting GetIt, FileName$, fProgramList, oProcedure, "Procedure", 1   'Main procedure
  i = 1   'Iterate from first container
  Do
    mCtrlSetting GetIt, FileName$, oProcedure(i), oProcedure, "Procedure", 1   'Get sub procedures
    mCtrlSetting GetIt, FileName$, oProcedure(i), oInstr, "Instruction", 10    'Get instructions
    i = i + 1
  Loop Until i > oProcedure.UBound
End Sub
 

Private Sub cmdHome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnMouse cmdHome, Button, Shift, X, Y
End Sub

Private Sub Image1_Click()
  frmMsg.Show
  frmMsg.mMsgByTagStr "About"       'Start message
End Sub

Private Sub label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnMouse label(Index), Button, Shift, X, Y
End Sub

Private Sub lLnNr_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  If Source Is oSP Then 'Drag drop instruction pointer
    'pInstrPointer = Index - 1
    nIfCond = True   'Unconditional execution of the instruction
  End If
End Sub

Private Sub lLnNr_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmMsg.mMsgOnMouse lLnNr(0), Button, Shift
End Sub

Private Sub oInstr_Click(Index As Integer)

End Sub

 

Private Sub oInstr_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
  oResizer.DontDrop oInstr(Index)    'Disable the control to drop on its container
End Sub

Private Sub oInstr_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Triggered if user press a key when selected
If KeyCode = vbKeyDelete And 0 < Index Then
    oBin_DragDrop oInstr(Index), 0, 0
    KeyCode = vbKeyTab
  End If
End Sub

Private Sub oInstr_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Stop 'this is inactive
  frmMsg.mMsgOnMouse oInstr(Index), Button, Shift
  oResizer.DontDrop Nothing     'Release the control as no drag is ongoing
End Sub

Private Sub oBin_DblClick()
  If MsgBox("Cancella tutto", vbYesNo) = vbYes Then mClearAll
End Sub

Private Sub oBin_DragDrop(Source As Control, X As Single, Y As Single)
  If TypeOf Source Is ucOpCode Then
    If Source.IsFixed = False Then  'It has been dragged away from the program
        CtrlRemove Source
    End If
  End If
  'Source.Visible = False 'Hide it for later deletion'
  mAlignCommands
End Sub

Private Sub mSetRepeat(c As Control)
   If c.sOPCode = "ForNext" Then
    c.nCounter = Val(InputBox(c.ToolTipText & " ?", , c.nCounter))
    c.Caption = c.nCounter & " X"
  End If
End Sub

Private Sub oBin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 frmMsg.mMsgOnMouse oBin, Button, Shift
End Sub
 
Private Sub mClearThis(cc As Control)
'obsolete
End Sub
Public Sub mClearAll()
  CtrlArraySetCount oInstr(0).Name, 10      'Leave only the static controls
  CtrlArraySetCount oProcedure(0).Name, 0
  mAlignCommands
End Sub
Private Property Get nGridY&()
  nGridY& = oInstr(0).Height + oInstr(0).Height / 10
End Property
Public Sub mAlignCommands()     'Sort of a compile command
  Dim Margin&, Nr&, i&, vTab&, c As Control, yPos&
  vTab& = oSP.Width         'Tabulator size
  Margin = oInstr(0).Height / 10
  'Ensure one an only one main sub
  Set c = oGetControlByIdx(oProcedure(0), 1)
  c.Visible = True: c.Caption = "Main"
  c.IsFixed = False
  Set c.Container = fProgramList
  yPos = 0
  For Each c In oProcedure
    If c.Visible And c.Container Is fProgramList Then 'skip
      c.Top = yPos:  c.Left = lLnNr(0).Width:    c.Width = Abs(fProgramList.ScaleWidth - 1 * c.Left)
      c.mAlignCommands
      yPos& = c.Top + c.Height
    End If
    fProgramList.Height = yPos& + Margin * 10
  Next
  CmdsEnable Me, fProgramList, True
End Sub
Public Sub ShowCommands(nUserLevel)
  Dim c As Control
'Shows the usercommands corresponding to the current user level
  For Each c In Controls
    If TypeOf c Is ucOpCode Then
      If c.sOPCode = "ForNext" Then
        c.Visible = nUserLevel >= 1
      ElseIf c.nIfCond Then
        c.Visible = nUserLevel >= 2
      ElseIf c.sOPCode = "Call" Then
        c.Visible = nUserLevel >= 3
      Else    'Decide what to implement for variables
        frmMsg.mMsgByTagStr "VarsNotImplemented"
      End If
    End If
  Next
End Sub
Public Sub SetLnNr(LnNr&, MyYPos&)
'Keep track of current line number
  Static CurrLnIdx&
  If MyYPos < lLnNr(0).Height + lLnNr(0).Top Then   'Restart numbering
      CurrLnIdx = 0
  Else            'Next line number
    CurrLnIdx = CurrLnIdx + 1
    While lLnNr.UBound < CurrLnIdx: Load lLnNr(lLnNr.UBound + 1): Wend
  End If
  lLnNr(CurrLnIdx).Visible = True
  lLnNr(CurrLnIdx).Caption = LnNr&
  lLnNr(CurrLnIdx).Top = MyYPos&
  lLnNr(CurrLnIdx).Left = 0
End Sub
Private Sub Form_Load()
  Show

'  oIniFile.WinState True, Me
  Me.Height = MDIForm1.ScaleHeight: Me.Top = 0
  Me.Left = 0
  nIfCond = True
  mCaptionTexts True, Me
  KeyPreview = True
End Sub
Public Sub mMainReset()
  fConsole.mReset
  pProgramState = eDB_stop      'Stop program
  mAlignCommands
  mSetPointerByLn = 1
  If ExecSpeed! < 5 Then
    ExecSpeed! = 2000
  End If
End Sub
Private Sub mMainSub()
  oProcedure(1).mSubroutine
  pProgramState = eDB_stop
End Sub
Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    'Copy this code into the dragDrop Event of the Target object
'   oResizer.StopDrag Source, Form1, X, Y
 '   Source.ZOrder 0 'Add if you want the control moved to the front
   oBin_DragDrop Source, X, Y
End Sub
Public Sub mExecute(c As Control, nIfCond As Boolean)
'Execution of current instruction statement
  If pProgramState = eDB_Run Then Me.ZOrder
  If c.nIfCond Then 'Test for conditions
    Select Case DeCode(c)
        Case "Eatable":
          frmMsg.mMsgByTagStr "CheckFood"
          nIfCond = Not fConsole.oObjAt1(fConsole.oFood, NextX, NextY) Is Nothing
        Case "Dangerous"
          frmMsg.mMsgByTagStr "CheckDog"
          nIfCond = Not fConsole.oObjAt1(fConsole.oDog, NextX, NextY) Is Nothing
        Case "Wall"
          frmMsg.mMsgByTagStr "CheckWall"
          nIfCond = Not fConsole.mCanMoveTo(NextX, NextY)
        Case Else:
'          Stop
      End Select
  ElseIf nIfCond Then     'Actions
    Select Case DeCode(c)
      Case "Move":
        mMove

      Case "Right":  vDirection = DirRight
      Case "Left":  vDirection = DirLeft
      Case "Up":  vDirection = Dirup
      Case "Down":  vDirection = Dirdown
      Case "Eat":
        mEat
       Case Else:
        If c.IsSub Then     'Enter a subroutine
          c.mSubroutine
        ElseIf Not c.mCallObj Is Nothing Then
          c.mCallObj.mSubroutine
        Else
          Stop
        End If
    End Select
  Else
    nIfCond = True     'Reset condition leaving it valid for one step only
  End If
  '      c.value = True     'Will fire the click event
End Sub
Private Sub mMove()
  Dim o As Control, oSearch As Object
  Set oSearch = fConsole.oProbe
    If Not fConsole.oObjAt1(fConsole.oTarget, NextX, NextY) Is Nothing Then
        frmMsg.mMsgByTagStr "CatHome"
        pProgramState = eDB_stop
        Exit Sub
    ElseIf Not fConsole.oObjAt1(fConsole.oDog, NextX, NextY) Is Nothing Then
        frmMsg.mMsgByTagStr "MetDog"
        pProgramState = eDB_stop
        Exit Sub
    ElseIf Not fConsole.oObjAt1(fConsole.oFood, NextX, NextY) Is Nothing Then
        frmMsg.mMsgByTagStr "StepInFood"
        pProgramState = eDB_stop
        Exit Sub
    ElseIf fConsole.mCanMoveTo(NextX, NextY) = False Then   'Cant move
        frmMsg.mMsgByTagStr "CantGo"
    Else                  'All clear go
        fConsole.mCatGoto NextX, NextY
    End If
End Sub
Public Sub mEat()
  Dim o As Control
    Set o = fConsole.oObjAt1(fConsole.oFood, NextX, NextY)
        If Not o Is Nothing Then
          o.Visible = False
          frmMsg.mMsgByTagStr "Eating"
        Else
          frmMsg.mMsgByTagStr "NoEating"
        End If
End Sub
Private Function NextX&()
    Select Case vDirection
      Case DirLeft: NextX = vCatPosX - 1
      Case DirRight:  NextX = vCatPosX + 1
      Case Else: NextX = vCatPosX
    End Select
End Function

Private Function NextY&()
    Select Case vDirection
      Case Dirup: NextY = vCatPosY + 1
      Case Dirdown:  NextY = vCatPosY - 1
      Case Else: NextY = vCatPosY
    End Select
End Function

 


Private Sub oInstr_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 'this is inactive
End Sub


Private Function MinVal(a, b)
  If a < b Then
    MinVal = a
  Else
    MinVal = b
  End If
End Function
Private Function maxval(a, b)
    If b < a Then
    maxval = a
  Else
    maxval = b
  End If
End Function
Public Property Let mSetPointer(yPos&)        'Set the program execution pointer

  'set the pointer
  oSP.Top = yPos
  mScrollInto yPos
End Property
 
Public Property Let mSetPointerByLn(LnNr)        'Set the program execution pointer
  'set the pointer by line number
  oSP.Top = lLnNr(LnNr - 1).Top
  mScrollInto oSP.Top
End Property
Function oMakeCopy(Source As Control) As Control
  Dim o As Object
  Dim Obj As Control
  Dim ObjArr
  Dim u As ucOpCode
   Set o = CtrlArryByName(Source.Name)
   If Not o Is Nothing Then
'      If o(o.UBound - 1).Name <> Source.Name Then Stop
      Load o(o.UBound + 1)
      Set oMakeCopy = o.Item(o.UBound)
      oMakeCopy.Caption = Source.Caption
      oMakeCopy.ToolTipText = Source.ToolTipText
      oMakeCopy.Tag = Source.Tag
      oMakeCopy.Visible = True
  End If
  If TypeOf Source Is ucOpCode Then
      oMakeCopy.MyProperties = Source.MyProperties
      oMakeCopy.IsFixed = False         'Release the control as beeing in program
  End If

End Function


Private Sub oProcedure_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
  If 0 < Index Then   'Dont drop on procedure in the element selection list
    oProcedure(Index).mDragDrop Source, X, Y
    mSetRepeat Source
    Source.mSetCall
  End If
End Sub

Private Sub oProcedure_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Stop
End Sub

Private Sub oSP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  oSP.Drag
  oSP.ToolTipText = "Trascinare al instruzione che desideri eseguire"
End Sub
Private Sub oSP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnMouse oSP, Button, Shift
End Sub
 


Private Sub Picture1_Click(Index As Integer)
  'We put images in picturebox to trap the focus
End Sub

Private Sub VScroll1_Change()
  Dim a!, l!
   a = (fProgramList.Height - fProgramListContainer.Height)   'The oversize of the list
   If a > 0 Then
    l = (a * (VScroll1.Value / VScroll1.Max))            'Multiplied by the fraction of scroll
    fProgramList.Top = -l
  Else
    VScroll1.Visible = False
    fProgramList.Top = 0
  End If
End Sub

Private Sub OnResize()
  Dim Margin&, Y&
  If WindowState = vbMinimized Then Exit Sub
  Margin = 20
  If Visible Then
    If frmMsg.vDocked Then
      Me.ScaleMode = vbTwips
      frmMsg.Left = Me.Left + fElements.Left
      frmMsg.Top = Me.Top + (Me.Height - Me.ScaleHeight) + fElements.Top + fElements.Height '+ Margin
      frmMsg.Width = fElements.Width
      frmMsg.Height = Me.ScaleHeight - frmMsg.Top
    End If
    fProgramListContainer.Left = fElements.Width + Margin
    fProgramListContainer.Width = Abs(ScaleWidth - fProgramListContainer.Left)
    fProgramList.Container.Height = Me.ScaleHeight - 1.1 * fProgramList.Container.Top
 
    Y& = oProcedure(oProcedure.UBound).Top + oProcedure(oProcedure.UBound).Height
    fProgramListContainer.ScaleMode = fProgramList.ScaleMode

    VScroll1.Visible = fProgramList.Height > fProgramList.Container.ScaleHeight
    VScroll1.Top = 0
    VScroll1.Left = fProgramList.Width
    VScroll1.Height = fProgramListContainer.Height - VScroll1.Top
    If VScroll1.Visible Then
        fProgramList.Width = Abs(fProgramListContainer.ScaleWidth - VScroll1.Width)
    Else
      fProgramList.Top = 0
        fProgramList.Width = Abs(fProgramListContainer.ScaleWidth - Margin)
    End If
  End If
End Sub

Private Sub VScroll1_GotFocus()
 frmMsg.mMsgShow VScroll1
End Sub

Private Sub VScroll1_Scroll()
  VScroll1_Change
End Sub
