VERSION 5.00
Begin VB.UserControl ucHelp 
   BackColor       =   &H80000001&
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4488
   FillColor       =   &H80000005&
   ScaleHeight     =   3180
   ScaleWidth      =   4488
   Begin VB.CommandButton cmdSaveHelp 
      Caption         =   "Salva"
      Height          =   252
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   3972
   End
   Begin VB.TextBox txtMsg 
      Height          =   2532
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "ucHelp.ctx":0000
      Top             =   0
      Width           =   3972
   End
End
Attribute VB_Name = "ucHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private CurrentSection$, CurrentKey$, CurrentString$
Private oIniFile As New cIniFile
'put in MouseMove event   after some hover the help message will appear
'oHelp.mMsgOnHover 'control', Button , Shift , X, Y)
Public Sub mMsgOnHover(c As Control, Button As Integer, Shift As Integer, Optional X As Single, Optional Y As Single)      'Will show a help text for the object
  Static HoverObj As Object, HoverTime!
  If Screen.ActiveControl Is UserControl.Extender Then
    Exit Sub
  End If
  If Not c Is HoverObj Then     'Start hovering
    HoverTime! = Timer + 1    'After 1 sec the help will show up
    Set HoverObj = c
  End If
  If HoverTime < Timer Then
    mMsgShow c
  End If
End Sub
Public Function ifCtrlHasIndex(c As Object)
  ifCtrlHasIndex = False
  On Error GoTo errhnd:
  ifCtrlHasIndex = (0 <= c.Index)
errhnd:    On Error GoTo 0
End Function
Public Sub mMsgShow(c As Control)        'Just show the help message
  Dim N$
  If Not IsObject(c) Then
    Stop
  ElseIf c.Tag = "" Then      'Use either tag or control identifier as section name
    CurrentSection$ = "Help:" & c.Name
    If TypeOf c Is Form Then
    Else
      N = c.Container.Name
      CurrentSection$ = CurrentSection$ & N 'Add the name of the container
      If ifCtrlHasIndex(c) Then
        CurrentSection$ = CurrentSection$ & c.Index          'Attmpt to suffix an index
      End If
    End If
  Else
    CurrentSection$ = "Help:" & c.Tag
  End If
  CurrentKey = sLang
  CurrentString$ = oIniFile.Setting(True, CurrentSection$, CurrentKey$)
  If CurrentString = "" Then
    CurrentString$ = oIniFile.Setting(True, CurrentSection$, "UK")
  End If
  If CurrentString = "" Then
 
    If TypeOf c Is PictureBox Then
    ElseIf TypeOf c Is Image Then
    Else
    CurrentString$ = c.Caption
    End If
 
    If CurrentString = "" Then CurrentString = CurrentSection$
  End If
  txtMsg.FontSize = nFontSize&
  Tekst = CurrentString
End Sub
Public Sub mMsgPrintFmt(Tag$, Optional Nr)
  CurrentSection$ = "Text:" & Tag$
    CurrentKey = sLang
  CurrentString$ = oIniFile.Setting(True, CurrentSection$, CurrentKey$)
  If CurrentString = "" Then
    CurrentString$ = oIniFile.Setting(True, CurrentSection$, "UK")
  End If
  txtMsg.FontSize = nFontSize& * 1.2
  Tekst = CurrentString
  If Not IsMissing(Nr) Then CurrentSection$ = Replace(CurrentSection$, "%n", CStr(Nr))
  SetFocus
End Sub

Public Sub mMsg_Write(sNewMessage)      'Overwrite the help message with a new and save in ini file
  oIniFile.Setting False, CurrentSection$, CurrentKey$, sNewMessage
End Sub








Private Sub cmdSaveHelp_Click()
  mMsg_Write Tekst
  cmdSaveHelp.Visible = False
End Sub

Private Sub txtMsg_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyS And Shift = vbCtrlMask Then 'ctr+s will save the next
    If MsgBox("Salva nuovo testo", vbYesNo) = vbYes Then
       mMsg_Write Tekst
    End If
  End If
  cmdSaveHelp.Visible = True
End Sub
Private Sub txtMsg_KeyPress(KeyAscii As Integer)
  txtMsg.ToolTipText = "Salva con Ctrl+S"
End Sub

 
 
  
 

Private Sub UserControl_Initialize()
      oIniFile.SetupFile = App.Path + "\HelpTips.ini"
      propFontSize& = txtMsg.FontSize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  txtMsg.Text = PropBag.ReadProperty("Tekst", "Messaggi")
End Sub

Private Sub UserControl_Resize()
  If cmdSaveHelp.Visible Then
    txtMsg.Move 0, 0, ScaleWidth, ScaleHeight - cmdSaveHelp.Height
    cmdSaveHelp.Top = txtMsg.Height
  Else
    txtMsg.Move 0, 0, ScaleWidth, ScaleHeight
  End If
End Sub

 

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Tekst", txtMsg.Text, "Messaggi")
End Sub

'AVVISO: NON RIMUOVERE O MODIFICARE LE SEGUENTI RIGHE DI COMMENTO
'MappingInfo=txtMsg,txtMsg,-1,Text
Public Property Get Tekst() As String
Attribute Tekst.VB_Description = "Restituisce o imposta il testo contenuto nel controllo."
  Tekst = txtMsg.Text
End Property

Public Property Let Tekst(ByVal New_Tekst As String)
  txtMsg.Text() = New_Tekst
  cmdSaveHelp.Visible = False
  PropertyChanged "Tekst"
  UserControl_Resize
End Property

