VERSION 5.00
Begin VB.UserControl UCProcedure 
   BackColor       =   &H00E0E0E0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   708
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   ForwardFocus    =   -1  'True
   HitBehavior     =   2  'Use Paint
   KeyPreview      =   -1  'True
   MaskColor       =   &H80000009&
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   360
   ScaleWidth      =   708
   ToolboxBitmap   =   "UCProcedure.ctx":0000
   Begin VB.Label lBrace 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "}"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   0
      Width           =   96
   End
   Begin VB.Label lBrace 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "{"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   96
   End
End
Attribute VB_Name = "UCProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Valori predefiniti proprietà:
Const m_def_Caption = ""
Const m_def_Enabled = 0
'Const m_def_Caption = ""
Const m_def_ToolTipText = ""
'Variabili proprietà:
Dim m_Caption As String
Dim m_Enabled As Boolean
'Dim m_Caption As String
Dim m_ToolTipText As String
Dim prop_GridY&
Private oaInstructions As New cSortedCollection
Private pInstrPointer&
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lPPoint As Point) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lPPoint As Point) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

 
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


 
 Public Sub mSubroutine()       'Execute the group of commands governed by the
  Dim N&
  'instruction pointer pInstrPointer
     Static Instances&
'    If 0 < Instances Then Exit Sub Else Instances& = 1
    
    Dim oIP As Control, vIfCondition As Boolean
    vIfCondition = True 'Start off with unconditional execution of first command
    For pInstrPointer = 0 To oaInstructions.mUBound   'Loop for executing instructions
       Set oIP = oaInstructions.mItem(pInstrPointer)
       fProgram.mSetPointer = PointToTargetY&(fProgram.fProgramList, oIP.Top)
      If isKindOf(oIP, fProgram.oInstr(0)) Or (oIP.Name = Extender.Name) Then
        If vIfCondition Then
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
      If oIP.vOPCode = "ForNext" And vIfCondition Then
        pInstrPointer = pInstrPointer + 1       'Point to next instruction  and repeat number of times
        For N& = 1 To oIP.vCounter
          fProgram.mExecute oaInstructions.mItem(pInstrPointer), vIfCondition
        Next
      Else
        fProgram.mExecute oIP, vIfCondition
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
    If Source.Container Is fProgram.fElements Then    'A control has been dragged from list of commands
       Set Source = fProgram.oMakeCopy(Source)
    End If
    'Move it into position
     oResizer.StopDrag Source, UserControl.Extender, X, Y
     If Source Is UserControl.Extender Then     'Dropped on myself, cant do
     Else
       Set Source.Container = UserControl.Extender
     End If
     fProgram.mAlignCommands
     fProgram.mSetPointer = Source.Top
End Sub
Public Property Get Name$()
  Name = UserControl.Name
End Property
Private Property Let vGridY(MaxGridHeight&)
  If prop_GridY& < MaxGridHeight& Then prop_GridY& = MaxGridHeight& * 1.1
End Property
Private Property Get vGridY&()
  vGridY& = prop_GridY&
End Property
Public Sub mAlignCommands()     'Sort of a compile command
  Dim Margin&, Nr&, i&, vTab&
  Dim c As Control
  Dim YPos&, indent& 'If previous instruction is a condition then indent by tab
  vTab& = lBrace(0).Width         'Tabulator size

  Nr& = 0: indent = 0        'Control number
  oaInstructions.mClear               'Make an ordered list of the controls
  For Each c In Parent.Controls
    If c.Container Is UserControl.Extender Then
      If fProgram.isCommand(c) Then
      'Add an instruction according to its vertical position on the container
      oaInstructions.mAdd c, c.Top + (c.Left / ScaleWidth)  'create a fractional order index leftmost becomes prioritized
      If c.IsSub Then
        c.mAlignCommands          'Recurse into subprocedures
      End If
      End If
    End If
  Next
  YPos& = lBrace(0).Height + Margin     'Thats after left brace
  For i = 0 To oaInstructions.mUBound     'Align them in order on the screen
    Set c = oaInstructions.mItem(i)
    fProgram.SetLnNr i + 1, PointToTargetY&(fProgram.fProgramList, YPos)
    c.Top = YPos:
    c.Left = vTab& + indent
    c.TabIndex = 100
    If c.IsSub Then
      c.Width = ScaleWidth - 2 * c.Left
    End If
    'If this is a condition then indent next instruction as beeing conditioned
     If c.vCondition Then indent = indent + vTab Else indent = 0
    YPos = YPos + c.Height + Margin
  Next
  lBrace(1).Top = YPos: lBrace(1).Left = 0
  Extender.Height = YPos + lBrace(0).Height
End Sub


Private Sub UserControl_Initialize()
  'Property must be controlcontainer = True
  UserControl.BackStyle = 1
End Sub


 

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=13,0,0,
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Restituisce o imposta il testo visualizzato quando il mouse viene posizionato per un breve intervallo di tempo sul controllo."
  ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
  m_ToolTipText = New_ToolTipText
  PropertyChanged "ToolTipText"
End Property

'Inizializza le proprietà di UserControl.
Private Sub UserControl_InitProperties()
 
  m_ToolTipText = m_def_ToolTipText
  m_Caption = m_def_Caption
  m_Enabled = m_def_Enabled
End Sub
 

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'Click and drag the control to its position
  oResizer.StartDrag UserControl.Extender, Button, X, Y
  CmdsEnable fProgram, fProgram.fProgramList, False
End Sub
 

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnHover UserControl.Extender, Button, Shift

End Sub

'Carica i valori delle proprietà dalla posizione di memorizzazione.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
  m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
  m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
  UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFF80)
End Sub

 

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'  Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
  Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
  Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
  Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
  Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFF80)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=13,0,0,
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Restituisce o imposta il testo visualizzato sulla barra del titolo o sotto l'icona di un oggetto."
  Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
  m_Caption = New_Caption
  PropertyChanged "Caption"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Restituisce o imposta un valore che determina se un oggetto è in grado di rispondere agli eventi generati dall'utente."
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  New_Enabled = True  'temp remove line
  m_Enabled = New_Enabled
  UserControl.Enabled = New_Enabled
  PropertyChanged "Enabled"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Restituisce o imposta il colore di sfondo utilizzato per la visualizzazione di testo e grafica in un oggetto."
  BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
  UserControl.BackColor() = New_BackColor
  PropertyChanged "BackColor"
End Property

