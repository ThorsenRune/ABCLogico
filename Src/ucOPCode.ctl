VERSION 5.00
Begin VB.UserControl ucOpCode 
   BackColor       =   &H00FFFF80&
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   ControlContainer=   -1  'True
   ForeColor       =   &H00008000&
   ScaleHeight     =   300
   ScaleWidth      =   3480
   ToolboxBitmap   =   "ucOPCode.ctx":0000
   Begin VB.Label lBrace 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "{"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   84
   End
   Begin VB.Label lBrace 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "}"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   84
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2652
   End
   Begin VB.Image Image1 
      Height          =   732
      Left            =   2640
      Picture         =   "ucOPCode.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   852
   End
End
Attribute VB_Name = "ucOpCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Usercontrol containing the operational code (OPcode) or instruction
'Contains information about
'   Type    =   Action/Condition
'   Caption  what is displayed
'   Description what is displayed in the help
'   ToolTip     the tool tip
'   Value       a for next start value

Option Explicit

 
'Variabili proprietà:
Dim props$()
Public Enum kProps         'Direction of moving elements
  kCondition            'If its a condition
  kOPCode               'The "OPCODE" i.e. move turn if then etc
  kCaption
  kDescription
  kToolTip
  kCounter                'For next until count
  kVariable               'Variable counter for next loop
  kVisible
  kHelpId                 'The identifier for the help entry
  kOperand                'The operand of the instruction. A call would have a label,could also be an equation?
  ktop
  kLeft
  kBackColor
  kContainer
  kFixed                  'Determines if the control is in the program or in the list of instructions the user can choose from
End Enum
 
'Dichiarazioni di eventi:
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Viene generato quando si preme un tasto mentre lo stato attivo si trova su un oggetto."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Viene generato quando si preme il pulsante del mouse mentre lo stato attivo si trova su un oggetto."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Viene generato quando si sposta il mouse."

Dim prop_GridY&
Private oaInstructions As New cSortedCollection
Private pInstrPointer&
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lPPoint As Point) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lPPoint As Point) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'Valori predefiniti proprietà:


 
' returns absolute screen position of a control that has an hWnd
' X and Y members are in pixels
Public Function PointToTargetY&(Target As Control, Y&)
  Dim P1 As Point
  Dim c As Control
    P1.X = 0: P1.Y = 0 ' Y
    P1.Y = ScaleY(Y, ScaleMode, vbPixels)
    ClientToScreen UserControl.hWnd, P1
   ' SetCursorPos P1.X, P1.Y    'Set mouse position (for debugging purpose)
    ScreenToClient Target.hWnd, P1            'Returns pixel position in client
    PointToTargetY = Target.ScaleY(P1.Y, vbPixels)
End Function

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  oResizer.MouseDown UserControl.Extender, Button, X, Y
  If mCallObj Is Nothing Then mSetCall
End Sub
Public Sub mSetCall()
    If sOPCode = "Call" Then
      sOperand = InputBox("Chiama procedure", , sOperand)
    End If
End Sub
Public Function mCallObj() As Control
Dim c As Control
  If sOPCode <> "Call" Then Exit Function
  If sOperand = "" Then Exit Function
  For Each c In Parent.Controls
    If TypeOf c Is ucOpCode Then
    If c.IsSub Then
      If c.sOperand = sOperand Then
        Set mCallObj = c
      End If
    End If
    End If
  Next
End Function
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  oResizer.MouseUp UserControl.Extender, Button, X, Y
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnMouse UserControl.Extender, Button, Shift
  oResizer.DontDrop Nothing     'Release the control as no drag is ongoing
  Label1.ToolTipText = Extender.ToolTipText
End Sub
Private Sub UserControl_Initialize()
Dim i&
 ReDim props$(15)
 For i = 0 To UBound(props)
  props(i) = "0"
 Next
 UserControl.BackStyle = 1
End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  oResizer.MouseUp UserControl.Extender, Button, X, Y
End Sub

'Carica i valori delle proprietà dalla posizione di memorizzazione.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Set Me.Picture1 = PropBag.ReadProperty("Image", Nothing)
 
  UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFF80)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", 1)
  Caption = PropBag.ReadProperty("Caption", "Label1")
  props(kProps.kDescription) = PropBag.ReadProperty("vDescription", "Description")
  props(kProps.kToolTip) = PropBag.ReadProperty("ToolTip", "ToolTip")
  props(kProps.kCounter) = PropBag.ReadProperty("vCounter", 0)
  Set Picture = PropBag.ReadProperty("Picture1", Nothing)
  props(kProps.kOPCode) = PropBag.ReadProperty("vOPCode", "OpCode")
  props(kProps.kCondition) = PropBag.ReadProperty("IfCond", 0)
  props(kProps.kFixed) = PropBag.ReadProperty("IsFixed", 1)
  'You can use the following to set a default for all these usercontrols
   ' ucOpCode_Defaults UserControl.Extender
End Sub

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFF80)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, 1)
  Call PropBag.WriteProperty("Caption", props(kProps.kCaption), "Label1")
  Call PropBag.WriteProperty("vDescription", props(kProps.kDescription), kDescription)
  Call PropBag.WriteProperty("ToolTip", props(kProps.kToolTip), "ToolTip")
  Call PropBag.WriteProperty("vCounter", props(kProps.kCounter), 0)
  Call PropBag.WriteProperty("Picture1", Picture, Nothing)
  Call PropBag.WriteProperty("vOPCode", props(kProps.kOPCode), 1)
  Call PropBag.WriteProperty("IfCond", props(kProps.kCondition))
  Call PropBag.WriteProperty("IsFixed", props(kProps.kFixed))
End Sub
Public Function Index&()
  Index = UserControl.Extender.Index
End Function

 
'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Restituisce o imposta un valore che determina se un oggetto è in grado di rispondere agli eventi generati dall'utente."
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property



 

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=13,0,0,0
Public Property Get sDescription() As String
  sDescription = props(kProps.kDescription)
End Property

Public Property Let sDescription(NewVal As String)
  props(kProps.kDescription) = NewVal
  PropertyChanged "sDescription"
End Property
Public Property Get nVariable() As Long
  nVariable = val(props(kProps.kVariable))
End Property

Public Property Let nVariable(NewVal As Long)
  props(kProps.kVariable) = NewVal
End Property
Public Property Get sOperand$()
  sOperand = props(kProps.kOperand)
End Property
Public Property Let sOperand(NewVal$)
  props(kProps.kOperand) = NewVal
  mReDraw
End Property
'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,0,0
Public Property Get nCounter() As Long
  nCounter = val(props(kProps.kCounter))
End Property

Public Property Let nCounter(NewVal As Long)
  props(kProps.kCounter) = NewVal
  PropertyChanged "nCounter"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=Image1,Image1,-1,Picture
Public Property Get Picture1() As Picture
Attribute Picture1.VB_Description = "Restituisce o imposta un elemento grafico da visualizzare in un controllo."
  Set Picture1 = Image1.Picture
End Property

Public Property Set Picture1(ByVal New_Picture1 As Picture)
  Set Image1.Picture = New_Picture1
  PropertyChanged "Picture1"
End Property
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Public Property Get sOPCode() As String
  sOPCode = props(kProps.kOPCode)
End Property
Public Property Let sOPCode(NewVar As String)
  props(kProps.kOPCode) = NewVar
  PropertyChanged " sOPCode"
End Property
Public Property Get nIfCond() As Boolean
  nIfCond = Str2Bool(props(kProps.kCondition))
End Property
Public Property Let nIfCond(NewVal As Boolean)
  props(kProps.kCondition) = Bool2Str(NewVal)
  PropertyChanged "nIfCond"
End Property
Public Sub Settings(GetIt As Boolean, FileName$, ObjId$)
    MyProperties = oIniFile.Setting(GetIt, FileName$, "Instructions", ObjId$, props)
End Sub
Property Get MyProperties()
  props(ktop) = Extender.Top
  props(kLeft) = Extender.Left
  props(kVisible) = UserControl.Extender.Visible
  props(kContainer) = ctrlID(UserControl.Extender.Container)
  props(kToolTip) = Label1.ToolTipText
  props(kBackColor) = BackColor
  MyProperties = props
End Property
Property Let MyProperties(Value)
  Dim c As Control
  props = Value
  'Visual properties should change immediately
  Caption = props(kCaption)       'Trigger the refresh of the caption
  Extender.ToolTipText = props(kToolTip)
  If UBound(Value) < ktop Then Exit Property
  If IsNumeric(props(ktop)) Then Extender.Top = val(props(ktop))
  If IsNumeric(props(kLeft)) Then Extender.Left = val(props(kLeft))
  If IsNumeric(props(kVisible)) Then UserControl.Extender.Visible = val(props(kVisible))
  For Each c In Parent.Controls
    If ctrlID(c) = props(kContainer) Then
      Set UserControl.Extender.Container = c
    End If
  Next
  Label1.ToolTipText = props(kToolTip)
  Extender.ToolTipText = props(kToolTip)
  BackColor = val(props(kBackColor))
End Property
'-----------------------------------------
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'Click and drag the control to its position
  oResizer.MouseDown UserControl.Extender, Button, X, Y
  CmdsEnable fProgram, fProgram.fProgramList, False
End Sub
 Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnMouse UserControl.Extender, Button, Shift

End Sub
Public Sub mSubroutine()       'Execute the group of commands governed by the
  Dim N&
  'instruction pointer pInstrPointer
     Static Instances&
'    If 0 < Instances Then Exit Sub Else Instances& = 1
    
    Dim oIP As Control, nIfCond As Boolean
    nIfCond = True 'Start off with unconditional execution of first command
    For pInstrPointer = 0 To oaInstructions.mUBound   'Loop for executing instructions
       Set oIP = oaInstructions.mItem(pInstrPointer)
       fProgram.mSetPointer = PointToTargetY&(fProgram.fProgramList, oIP.Top)
      If isKindOf(oIP, fProgram.oInstr(0)) Or (oIP.Name = Extender.Name) Then
        If nIfCond Then
          oIP.BackColor = eExeColor.WillExec
        Else
          oIP.BackColor = eExeColor.NoExec
        End If
      Else
          oIP.BackColor = eExeColor.CondTest
      End If
       fProgram.mSetPointer = PointToTargetY&(fProgram.fProgramList, oIP.Top)
      If pProgramState = eDB_Run Then Sleep ExecSpeed
      While pProgramState = eDB_Break: DoEvents: Wend  'Break state wait here until it is released by changing to step in click event
  'Here program state may have been changed by the user, take appropriate actions
      If pProgramState = eDB_Step Then pProgramState = eDB_Break
      If pProgramState = eDB_stop Or Not isCtrlLoaded(oIP) Then Exit For
      If oIP.sOPCode = "ForNext" And nIfCond Then
        pInstrPointer = pInstrPointer + 1       'Point to next instruction  and repeat number of times
        For N& = 1 To oIP.nCounter
          fProgram.mExecute oaInstructions.mItem(pInstrPointer), nIfCond
        Next
      Else
        fProgram.mExecute oIP, nIfCond
      End If
      If pProgramState = eDB_stop Then Exit For
      oIP.BackColor = eExeColor.Passive
      fProgram.mSetPointer = PointToTargetY&(fProgram.oSP.Container, oIP.Top + oIP.Height)
    Next
    Instances& = 0
 End Sub
Private Sub lBrace_Click(Index As Integer)
  UserControl.BackStyle = 1
End Sub
Public Sub mDragDrop(Source As Control, X As Single, Y As Single)
  Dim c As Control
  'Note that this event does not fire. You have to call it
    If IsSub = False Then     'Cant drop on an instruction, it shouldnt happen however
'      Stop
    ElseIf Source.IsFixed Then    'A control has been dragged from list of commands
       Set Source = fProgram.oMakeCopy(Source)
    End If
    If Source.IsSub Then Source.Caption = ""
    'Move it into position
    oResizer.StopDrag Source, UserControl.Extender, X, Y
    If Source Is UserControl.Extender Then     'Dropped on myself, cant do
'     Stop
    Else
      Set Source.Container = UserControl.Extender
    End If
    Parent.mAlignCommands
    Parent.mSetPointer = Source.Top
End Sub
Public Property Get IsSub() As Boolean
  IsSub = (sOPCode = "Sub")
End Property

Public Property Get Name$()
  Name = UserControl.Name
End Property
Private Property Let nGridY(MaxGridHeight&)
  If prop_GridY& < MaxGridHeight& Then prop_GridY& = MaxGridHeight& * 1.1
End Property
Private Property Get nGridY&()
  nGridY& = prop_GridY&
End Property
Public Sub mAlignCommands()     'Sort of a compile command
'Align codes within a subroutine (this usercontrol is assumed to be a issub=true)
  Dim Margin&, Nr&, i&, vTab&
  Dim c As Control
  Dim yPos&, indent& 'If previous instruction is a condition then indent by tab
  Dim vYTab&
    If Label1.Visible Then
      lBrace(0).Left = Label1.Width
    Else
      lBrace(0).Left = 0
    End If
    lBrace(0).Visible = True:
    lBrace(1).Visible = True: lBrace(1).Left = 0
    vTab& = lBrace(0).Width         'Tabulator size
    vYTab& = lBrace(0).Height
  UserControl.ScaleMode = Extender.Parent.ScaleMode
  Nr& = 0: indent = 0        'Control number
  oaInstructions.mClear               'Make an ordered list of the controls
  For Each c In ContainedControls
    If isCommand(c) Then
    'Add an instruction according to its vertical position on the container
      oaInstructions.mAdd c, c.Top + (c.Left / ScaleWidth)  'create a fractional order index leftmost becomes prioritized
      If c.IsSub Then
        c.mAlignCommands          'Recurse into subprocedures
      End If
    End If
  Next
  yPos& = vYTab& + Margin     'Thats after left brace
  For i = 0 To oaInstructions.mUBound     'Align them in order on the screen
    Set c = oaInstructions.mItem(i)
    fProgram.SetLnNr i + 1, PointToTargetY&(fProgram.fProgramList, yPos)
    c.Top = yPos:
    c.Left = vTab& + indent
    c.TabIndex = 100
    If c.IsSub Then
      c.Width = ScaleWidth - 2 * c.Left
    End If
    'If this is a condition then indent next instruction as beeing conditioned
     If c.nIfCond Then indent = indent + vTab Else indent = 0
    yPos = yPos + c.Height ' vYTab&
  Next
  lBrace(1).Top = yPos
  Extender.Height = yPos + vYTab&
  BackColor = UserControl.Extender.Container.BackColor
End Sub


Public Property Get Caption() As String
  Caption = props(kProps.kCaption)
End Property

Public Property Let Caption(New_Caption As String)
  props(kProps.kCaption) = New_Caption
  PropertyChanged "Caption"
  mReDraw
End Property
Public Sub mReDraw()
  Label1.AutoSize = True
  If sOPCode = "Call" Then
    Label1 = props(kProps.kCaption) & " " & sOperand
  Else
    Label1.Caption = props(kProps.kCaption)
  End If
  Label1.Visible = (Label1.Caption <> "")

End Sub


'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=0,0,0,0
Public Property Get IsFixed() As Boolean
  IsFixed = Str2Bool(props(kProps.kFixed))
End Property
Private Function Str2Bool(S$) As Boolean
  Str2Bool = (val(S) <> 0)
End Function
Private Function Bool2Str$(val As Boolean)
  If val Then
    Bool2Str$ = "1"
  Else
    Bool2Str$ = "0"
  End If
End Function
Public Property Let IsFixed(NewVal As Boolean)
  props(kProps.kFixed) = Bool2Str(NewVal)
  PropertyChanged "IsFixed"
End Property

