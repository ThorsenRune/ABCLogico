VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Myvalue
'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=UserControl,UserControl,-1,MaskColor
Public Property Get Value() As Long
Attribute Value.VB_Description = "Restituisce o imposta il colore che specifica le aree trasparenti in MaskPicture."
  Value = Myvalue
End Property

Public Property Let Value(ByVal New_Value As Long)
  Myvalue = New_Value
  PropertyChanged "Value"
End Property

'Carica i valori delle proprietà dalla posizione di memorizzazione.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  Myvalue = PropBag.ReadProperty("Value", -2147483633)
End Sub

'Scrive i valori delle proprietà nella posizione di memorizzazione.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("Value", Myvalue, -2147483633)
End Sub

