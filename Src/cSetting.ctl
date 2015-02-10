VERSION 5.00
Begin VB.UserControl cSetting 
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3516
   ScaleHeight     =   960
   ScaleWidth      =   3516
   Begin VB.TextBox txtValue 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label label 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "Numero cani"
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
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "cSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 
 
Const m_def_mMyNumber = 0
'Variabili proprietà:
Dim m_mMyNumber As Long
Dim m_mMyImage As Object

 
'Carica i valori delle proprietà dalla posizione di memorizzazione.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  txtValue = PropBag.ReadProperty("mMyNumber", 1)
  Set m_mMyImage = PropBag.ReadProperty("mMyImage", Nothing)
  label.Caption = PropBag.ReadProperty("Caption", "Numero cani")
End Sub

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("mMyNumber", m_mMyNumber, txtValue)
  Call PropBag.WriteProperty("mMyImage", m_mMyImage, Nothing)
  Call PropBag.WriteProperty("Caption", label.Caption, "Numero cani")
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=8,0,0,0
Public Property Get Value() As Long
  Value = Val(txtValue)
End Property

Public Property Let Value(Value As Long)
  txtValue = Value
  PropertyChanged "Value"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MemberInfo=9,0,0,0
Public Property Get mMyImage() As Object
  Set mMyImage = m_mMyImage
End Property

Public Property Set mMyImage(ByVal New_mMyImage As Object)
  Set m_mMyImage = New_mMyImage
  PropertyChanged "mMyImage"
End Property

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=label,label,0,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Restituisce o imposta il testo visualizzato sulla barra del titolo o sotto l'icona di un oggetto."
  Caption = label.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
  label.Caption() = New_Caption
  PropertyChanged "Caption"
End Property

'Inizializza le proprietà di UserControl.
Private Sub UserControl_InitProperties()
  m_mMyNumber = m_def_mMyNumber
End Sub

Private Sub value_Change()
  
End Sub
