VERSION 5.00
Begin VB.UserControl ucStatusBar 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   252
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   252
   ScaleWidth      =   3840
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label1 
      Caption         =   "Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   1440
      TabIndex        =   1
      Top             =   0
      Width           =   2412
   End
End
Attribute VB_Name = "ucStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Valori predefiniti proprietà:
Const m_def_Alignment = 0
Const m_def_Tekst = "0"
Const m_def_Beskrivelse = "0"
'Const m_def_Value = 0
'Const m_def_Tekst = 0
'Const m_def_Beskrivelse = 0
'Variabili proprietà:
Dim m_Alignment As Integer
Dim m_Tekst As String
Dim m_Beskrivelse As String
'Dim m_Value As Boolean
'Dim m_Tekst As Variant
'Dim m_Beskrivelse As Variant
'
'
'
''AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
''MemberInfo=0,0,0,0
'Public Property Get Value() As Boolean
'  Value = m_Value
'End Property
'
'Public Property Let Value(ByVal New_Value As Boolean)
'  m_Value = New_Value
'  PropertyChanged "Value"
'End Property

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
'
''AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
''MemberInfo=14,0,0,0
'Public Property Get Tekst() As Variant
'  Tekst = m_Tekst
'End Property
'
'Public Property Let Tekst(ByVal New_Tekst As Variant)
'  m_Tekst = New_Tekst
'  PropertyChanged "Tekst"
'End Property
'
''AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
''MemberInfo=14,0,0,0
'Public Property Get Beskrivelse() As Variant
'  Beskrivelse = m_Beskrivelse
'End Property
'
'Public Property Let Beskrivelse(ByVal New_Beskrivelse As Variant)
'  m_Beskrivelse = New_Beskrivelse
'  PropertyChanged "Beskrivelse"
'End Property

 

'Inizializza le proprietà di UserControl.
Private Sub UserControl_InitProperties()
'  m_Value = m_def_Value
'  m_Tekst = m_def_Tekst
'  m_Beskrivelse = m_def_Beskrivelse
  m_Tekst = m_def_Tekst
  m_Beskrivelse = m_def_Beskrivelse
  m_Alignment = m_def_Alignment
End Sub

'Carica i valori delle proprietà dalla posizione di memorizzazione.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'  m_Value = PropBag.ReadProperty("Value", m_def_Value)
  Command1.Caption = PropBag.ReadProperty("Caption", "Command1")
'  m_Tekst = PropBag.ReadProperty("Tekst", m_def_Tekst)
'  m_Beskrivelse = PropBag.ReadProperty("Beskrivelse", m_def_Beskrivelse)
  Label1.Caption = PropBag.ReadProperty("Value", "Label1")
  m_Tekst = PropBag.ReadProperty("Tekst", m_def_Tekst)
  m_Beskrivelse = PropBag.ReadProperty("Beskrivelse", m_def_Beskrivelse)
  m_Alignment = PropBag.ReadProperty("Alignment", m_def_Alignment)
End Sub

Private Sub UserControl_Resize()
  If Command1.Visible Then
    Label1.Left = Command1.Left + Command1.Width
  Else
    Label1.Left = 0
  End If
  Label1.Width = UserControl.ScaleWidth - Label1.Left
End Sub

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'  Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
  Call PropBag.WriteProperty("Caption", Command1.Caption, "Command1")
'  Call PropBag.WriteProperty("Tekst", m_Tekst, m_def_Tekst)
'  Call PropBag.WriteProperty("Beskrivelse", m_Beskrivelse, m_def_Beskrivelse)
  Call PropBag.WriteProperty("Value", Label1.Caption, "Label1")
  Call PropBag.WriteProperty("Tekst", m_Tekst, m_def_Tekst)
  Call PropBag.WriteProperty("Beskrivelse", m_Beskrivelse, m_def_Beskrivelse)
  Call PropBag.WriteProperty("Alignment", m_Alignment, m_def_Alignment)
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Value() As String
Attribute Value.VB_Description = "Restituisce o imposta il testo visualizzato sulla barra del titolo o sotto l'icona di un oggetto."
  Value = Label1.Caption
End Property

Public Property Let Value(ByVal New_Value As String)
  Label1.Caption() = New_Value
  PropertyChanged "Value"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=13,0,0,0
Public Property Get Tekst() As String
  Tekst = m_Tekst
End Property

Public Property Let Tekst(ByVal New_Tekst As String)
  m_Tekst = New_Tekst
  PropertyChanged "Tekst"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=13,0,0,0
Public Property Get Beskrivelse() As String
  Beskrivelse = m_Beskrivelse
End Property

Public Property Let Beskrivelse(ByVal New_Beskrivelse As String)
  m_Beskrivelse = New_Beskrivelse
  PropertyChanged "Beskrivelse"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=7,0,0,0
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Restituisce o imposta il valore dell'allineamento di un controllo CheckBox o OptionButton o del testo di un controllo."
  Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
  m_Alignment = New_Alignment
  PropertyChanged "Alignment"
End Property
Property Get Name$()
  Name$ = Extender.Name
End Property
Property Get Container() As Object
  Set Container = Extender.Container
End Property
