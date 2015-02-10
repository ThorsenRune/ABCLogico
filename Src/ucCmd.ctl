VERSION 5.00
Begin VB.UserControl ucCmd 
   ClientHeight    =   1968
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   ScaleHeight     =   1968
   ScaleWidth      =   3480
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   732
      Left            =   120
      Picture         =   "ucCmd.ctx":0000
      Stretch         =   -1  'True
      Top             =   720
      Width           =   852
   End
End
Attribute VB_Name = "ucCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Variabili proprietà:
Dim m_Enabled As Boolean
Dim m_ToolTipText As String

'Valori predefiniti proprietà:
Const m_def_Enabled = 0
Const m_def_ToolTipText = ""

Public Property Get Picture1() As IPictureDisp
  Set Picture1 = Image1.Picture
End Property

Public Property Set Picture1(ByVal p As IPictureDisp)
  Set Image1.Picture = p
  PropertyChanged "Picture"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=Command1,Command1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Restituisce o imposta il testo visualizzato sulla barra del titolo o sotto l'icona di un oggetto."
  Caption = Command1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
  Command1.Caption() = New_Caption
  PropertyChanged "Caption"
End Property

 
 

Private Sub Image1_Click()

End Sub

'Inizializza le proprietà di UserControl.
Private Sub UserControl_InitProperties()
  'Set m_Image = LoadPicture("")
  m_Enabled = m_def_Enabled
  m_ToolTipText = m_def_ToolTipText
End Sub

'Carica i valori delle proprietà dalla posizione di memorizzazione.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  Command1.Caption = PropBag.ReadProperty("Caption", "Command1")
  Set Me.Picture1 = PropBag.ReadProperty("Image", Nothing)
  m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
  m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
  Command1.Enabled = PropBag.ReadProperty("VisibleCmd", True)
End Sub

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Caption", Command1.Caption, "Command1")
  Call PropBag.WriteProperty("Image", Me.Picture1, Nothing)
  Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
  Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
  Call PropBag.WriteProperty("VisibleCmd", Command1.Enabled, True)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Restituisce o imposta un valore che determina se un oggetto è in grado di rispondere agli eventi generati dall'utente."
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  m_Enabled = New_Enabled
  PropertyChanged "Enabled"
End Property

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

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=Command1,Command1,-1,Enabled
Public Property Get VisibleCmd() As Boolean
Attribute VisibleCmd.VB_Description = "Restituisce o imposta un valore che determina se un oggetto è in grado di rispondere agli eventi generati dall'utente."
  VisibleCmd = Command1.Enabled
End Property

Public Property Let VisibleCmd(ByVal New_VisibleCmd As Boolean)
  Command1.Enabled() = New_VisibleCmd
  PropertyChanged "VisibleCmd"
End Property

