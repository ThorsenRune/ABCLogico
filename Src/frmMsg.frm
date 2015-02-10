VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMsg 
   Caption         =   "Suggerimenti"
   ClientHeight    =   6000
   ClientLeft      =   2376
   ClientTop       =   2400
   ClientWidth     =   10656
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   10656
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin RichTextLib.RichTextBox rtbVisible 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8065
      _ExtentY        =   7006
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmMsg.frx":1A52
   End
   Begin VB.PictureBox fStatus 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   0
      ScaleHeight     =   228
      ScaleWidth      =   10632
      TabIndex        =   2
      Top             =   5748
      Width           =   10656
      Begin VB.CommandButton cmdHyperlink 
         Caption         =   "Hyperlink"
         Height          =   252
         Left            =   2040
         TabIndex        =   6
         Top             =   0
         Width           =   1092
      End
      Begin VB.CommandButton cmdSaveHelp 
         Caption         =   "Save"
         Height          =   252
         Left            =   0
         TabIndex        =   5
         ToolTipText     =   "Ctrl + s will save the current entry"
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdForward 
         Caption         =   ">"
         Height          =   252
         Left            =   3960
         TabIndex        =   4
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<"
         Height          =   252
         Left            =   3360
         TabIndex        =   3
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   5400
         TabIndex        =   7
         Top             =   0
         Width           =   492
      End
   End
   Begin RichTextLib.RichTextBox rtbHidden 
      Height          =   972
      Left            =   4920
      TabIndex        =   1
      Top             =   0
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1715
      _Version        =   393217
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmMsg.frx":1AD4
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public splash As Boolean
Public vDocked As Boolean
Private CurrentSection$, CurrentKey$, CurrentString$
Private oIniFile As New cIniFile
 
Private CurrentObj  'Reserved for later use of a help key (F1) to get help on a hovered item


Dim SrcPos&(3) 'Start-stop of entrykey , start,stop of entry text
Const leftMark = "{"
Const rightMark = "}" & vbCrLf
Dim vaEntryList&()
Private isFormLoaded As Boolean
'put in MouseMove event   after some hover the help message will appear
'frmMsg.mMsgOnMouse 'control', Button , Shift , X, Y)
Public Sub mMsgOnMouse(Var1, Button As Integer, Shift As Integer, Optional X As Single, Optional Y As Single)        'Will show a help text for the object
  Dim Obj As Object
  If Button = vbRightButton And Shift = vbCtrlMask Then      'Edit the caption for translation purposes
    msgEdit Var1
  ElseIf Screen.ActiveForm Is Me Then
    Exit Sub
  ElseIf Button = vbRightButton Then
    mMsgShow Var1
  ElseIf Button = vbLeftButton Then
  Else
  End If
'  If HoverTime < Timer Then mMsgShow Obj

End Sub
Private Function mContainedIn(c As Control) As Control
  Dim cc As Control
    For Each cc In c.Parent.Controls
      If cc.Container Is c Then
        Set mContainedIn = cc
      End If
    Next
End Function
Public Sub msgEdit(Var1)   'Takes a string,control or control array and lets the user edit the texts (translation)
  Dim s$, Obj As Object
  Dim cb As ComboBox
    If IsObject(Var1) Then Set Obj = Var1
    If IsCtrlArray(Obj) Then Set Obj = Obj(0) 'Select first element in control arrays
    If hasCaption(Obj) Then
      s$ = InputBox("New caption", , Obj.Caption)
      If s$ <> "" Then Obj.Caption = s$ 'Save the translated caption
    End If
    If TypeOf Obj Is PictureBox Then      'Propagate tooltiptexts
        If TypeOf mContainedIn(Obj) Is Image Then
          Set Obj = mContainedIn(Obj)   'Use the contained image
        End If
    End If
    If hasToolTip(Obj) Then
      s$ = InputBox("New tool tip", , Obj.ToolTipText)
      If s <> "" Then Obj.ToolTipText = s:
    End If
    If TypeOf Obj Is ComboBox Then
      Set cb = Obj
      s$ = InputBox("New text for entry", , cb.List(cb.ListIndex))
      If s <> "" Then cb.List(cb.ListIndex) = s:
    End If
    If s <> "" Then mCaptionTexts False, Obj  'Save the translated caption

End Sub
Private Sub setCurrentObj(Var1)
Static HoverTime!, Obj As Object
  If IsObject(Var1) Then Set Obj = Var1
  If Not Obj Is CurrentObj Then     'Select this object after some seconds of hovering (for later use of a help key (F1) you use the CurrentObj)
    HoverTime! = Timer + 1    'After 1 sec the help will show up
    If IsObject(Var1) Then
      Set CurrentObj = Var1
    Else
      CurrentObj = Var1
    End If
  End If
End Sub
Public Sub mMsgShow(Obj)         'Just show the help message
  Dim N$, c As Control
  If IsObject(Obj) Then
    If IsCtrlArray(Obj) Then
      Set c = Obj(0)
    Else
      Set c = Obj
    End If
    CurrentSection$ = ctrlID(c)
    CurrentKey = CurrentSection$
  Else
    CurrentKey = Obj
  End If
  If mEntryFind(CurrentKey, 0) = False Then
      mEntryAdd CurrentKey
  End If
  mEntryShow
  ZOrder 0
End Sub
Public Sub mMsgSplash()
  Show
  Sleep 1
  FormAlwaysOnTop Me, True
'  Hide
'  Show vbModal
End Sub
Public Sub mMsgByTagStr(Tag$, Optional Nr)
  CurrentSection$ = Tag$
  CurrentKey = CurrentSection$
  If mEntryFind(CurrentKey, 0) = False Then
      mEntryAdd CurrentKey
  End If
  If Not IsMissing(Nr) Then CurrentSection$ = Replace(CurrentSection$, "%n", CStr(Nr))
  If Visible Then SetFocus
  mEntryShow
End Sub

Private Sub mMsg_Write()      'Write the message to the shadow RTF and then to the whole RTF to disk
  If mEntryFind(CurrentKey$, 0) Then      'Ensure you have the right key
    rtbHidden.SelRTF = rtbVisible.TextRTF
    rtbHidden.SaveFile sFileHelp
  End If
  fStatus.Visible = False
  Form_Resize
End Sub
Private Sub cmdHyperlink_Click()
  HyperlinkCreate
End Sub

Private Sub cmdSaveHelp_Click()
  mMsg_Write
End Sub

Private Sub Form_Deactivate()
  Me.ZOrder 1       'Hide on deactivate
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not splash Then oIniFile.WinState False, Me
  splash = False
End Sub

Private Sub fStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnMouse Me, Button, Shift
End Sub

Private Sub rtbVisible_KeyDown(KeyCode As Integer, Shift As Integer)
  If fStatus.Visible = False Then
    fStatus.Visible = True
    Form_Resize
  ElseIf KeyCode = vbKeyS And Shift = vbCtrlMask Then 'ctr+s will save the next
       mMsg_Write
  End If
End Sub
Private Sub rtbVisible_KeyPress(KeyAscii As Integer)
  rtbVisible.ToolTipText = "Salva con Ctrl+S"
End Sub
Private Sub Form_Initialize()
     oIniFile.sMustFindFile sFileHelp
     isFormLoaded = False
     splash = True
End Sub
Private Sub Form_Resize()
 
  If rtbVisible.Visible = False Then
  ElseIf fStatus.Visible Then
    rtbVisible.Move 0, 0, ScaleWidth, fStatus.Top
  Else
    rtbVisible.Move 0, 0, ScaleWidth, ScaleHeight
  End If
End Sub
Private Sub Form_Load()
  If Not splash Then
    Show
    oIniFile.WinState True, Me
  End If
  vDocked = False      'Dont move help screen
  fStatus.Visible = False
  If oIniFile.sMustFindFile(sFileHelp) Then
    rtbHidden.LoadFile sFileHelp
    rtbVisible.LoadFile sFileHelp
  End If
  mCaptionTexts True, Me
  isFormLoaded = True
End Sub
Private Sub cmdBack_Click()
  If mEntryPrevious Then
    mEntryShow
  Else
    mEntryFind "", 0
  End If
End Sub
Private Function mEntryFind(entryKey$, Optional startPos& = -1) As Boolean
'Dim SrcPos&(5) left {, right },  end of entry

    mEntryFind = False
    If 0 <= startPos Then SrcPos&(0) = startPos
    SrcPos&(0) = rtbHidden.Find(leftMark + entryKey$, SrcPos&(0))  'Key start
    If SrcPos&(0) < 0 Then Exit Function
    SrcPos&(1) = rtbHidden.Find(rightMark, SrcPos&(0))
    If SrcPos&(1) < 0 Then Exit Function
    SrcPos&(2) = SrcPos&(1) + rtbHidden.SelLength       'Find start of text
    'Set the key
    rtbHidden.SelStart = SrcPos(0)
    rtbHidden.SelLength = SrcPos(1) - SrcPos(0)
    CurrentKey$ = Replace(rtbHidden.SelText, leftMark, "")
    Label1 = CurrentKey$
    'Mark the text
    SrcPos&(3) = rtbHidden.Find(leftMark, SrcPos&(2))   'Find end of entry
    If SrcPos&(3) < 0 Then SrcPos&(3) = Len(rtbHidden.Text) 'If no endmark select rest of text
        rtbHidden.SelStart = SrcPos(2)
        rtbHidden.SelLength = SrcPos(3) - SrcPos(2)
    mEntryFind = True

End Function
Private Function mEntryPrevious() As Boolean
  Dim i&(3)
    i(3) = rtbHidden.SelStart     'End of search
    i(2) = -1      ' Search from the start
   Do   'I2= this entry ,i1 the one before
      i(2) = rtbHidden.Find(leftMark, i(2) + 1, i(3))   'Next entry
      If i(2) < 0 Then Exit Do       'Found the end entry
      i(0) = i(1)
      i(1) = i(2)       'shif entries down
    Loop
    mEntryPrevious = mEntryFind("", i(0))
End Function
Private Sub cmdForward_Click()
  If mEntryFind("", rtbHidden.SelStart) Then
       mEntryShow
  Else    'Not found

  End If
End Sub
 
Public Sub mEntryShow()       'Shows the entry text
 '      rtbVisible.TextRTF = rtbHidden.TextRTF
   rtbVisible.TextRTF = rtbHidden.SelRTF
   fStatus.Visible = False
End Sub
Public Sub mEntryAdd(NewEntryKey$)
  If mEntryFind(CurrentKey$, 0) Then
    rtbHidden.SelStart = rtbHidden.SelStart + rtbHidden.SelLength
    rtbHidden.SelText = vbCrLf & leftMark & NewEntryKey$ & rightMark
  ElseIf mEntryFind(NewEntryKey, 0) = False Then
    rtbHidden.SelStart = Len(rtbHidden.Text)
    rtbHidden.SelText = vbCrLf & leftMark & NewEntryKey$ & rightMark
  End If
End Sub

 
 

Property Let Hyperlink(Value As Boolean)
  If Value Then
    rtbVisible.SelColor = vbBlue
    rtbVisible.SelUnderline = True
  Else
    HyperlinkSelect
    rtbVisible.SelColor = vbBlack
    rtbVisible.SelUnderline = False
  End If
End Property
Property Get Hyperlink() As Boolean
  If rtbVisible.SelColor = vbBlue And rtbVisible.SelUnderline = True Then
  Hyperlink = True
  Else
  Hyperlink = False
  End If

End Property
Sub HyperlinkSelect()
  Dim P&, l&
    If Not Hyperlink Then Exit Sub   'Expand selection to cover it all
    l = rtbVisible.SelLength
    P = rtbVisible.SelStart
    Do          'Find start of link
      rtbVisible.SelStart = P:      rtbVisible.SelLength = 1
      If Not Hyperlink Then P = P + 1: Exit Do
      If P < 1 Then Exit Do
      P = P - 1             'Step back
    Loop
    For l = 1 To 200            'Find end of link
      rtbVisible.SelStart = P + l
      rtbVisible.SelLength = 1
      If rtbVisible.SelLength = 0 Then        'Hidden text will give sellength =0
      ElseIf Not Hyperlink Then
        Exit For
      End If
    Next
    rtbVisible.SelStart = P
    rtbVisible.SelLength = l
End Sub

Private Sub HyperlinkCreate()
  While (Right(rtbVisible.SelText, 1) = " ")      'Remove trailing spaces
    rtbVisible.SelLength = rtbVisible.SelLength - 1:
  Wend
  Hyperlink = True                                'Mark as hyperlink
  mLinkText = InputBox("LinkTo", , mLinkText)
  If mLinkText = "" Then Hyperlink = False
End Sub
Private Sub rtbVisible_Click()
  Hyperlink_Follow
End Sub
 Private Sub Hyperlink_Follow()
'Links are circular. We want to simply jump to the next instance of the anchor
  Dim lnk$
  If Not Hyperlink Then Exit Sub
  lnk = "§" & mLinkText$ & "§"
    If 0 <= rtbVisible.Find(lnk, rtbVisible.SelStart + rtbVisible.SelLength + 1) Then   'Search rest of current text
     
    ElseIf 0 <= rtbHidden.Find(lnk, rtbHidden.SelStart + rtbHidden.SelLength) Then
      ThisEntry
      rtbVisible.Find lnk, 0
    ElseIf 0 <= rtbHidden.Find(lnk, 0) Then
      ThisEntry
      mShowEntry
      rtbVisible.Find (lnk)
    Else
'      Stop
    End If
    HyperlinkSelect
End Sub

Private Function ThisEntry() As Boolean
  Dim i&(3)
    i(3) = rtbHidden.SelStart     'End of search
    i(2) = -1      ' Search from the start
   Do   'I2= this entry ,i1 the one before
      i(2) = rtbHidden.Find(leftMark, i(2) + 1, i(3))   'Next entry
      If i(2) < 0 Then Exit Do       'Found the end entry
      i(0) = i(1)
      i(1) = i(2)       'shif entries down
    Loop
    ThisEntry = FindEntry("", i(1))
End Function
Private Function FindEntry(entryKey$, Optional startPos& = -1) As Boolean
'Dim SrcPos&(5) left {, right },  end of entry

    FindEntry = False
    If 0 <= startPos Then SrcPos&(0) = startPos
    SrcPos&(0) = rtbHidden.Find(leftMark + entryKey$, SrcPos&(0))  'Key start
    If SrcPos&(0) < 0 Then Exit Function
    SrcPos&(1) = rtbHidden.Find(rightMark, SrcPos&(0))
    If SrcPos&(1) < 0 Then Exit Function
    SrcPos&(2) = SrcPos&(1) + rtbHidden.SelLength       'Find start of text
    'Set the key
    rtbHidden.SelStart = SrcPos(0)
    rtbHidden.SelLength = SrcPos(1) - SrcPos(0)
    CurrentKey$ = Replace(rtbHidden.SelText, leftMark, "")
    'Mark the text
    SrcPos&(3) = rtbHidden.Find(leftMark, SrcPos&(2))   'Find end of entry
    If SrcPos&(3) < 0 Then SrcPos&(3) = Len(rtbHidden.Text) 'If no endmark select rest of text
        rtbHidden.SelStart = SrcPos(2)
        rtbHidden.SelLength = SrcPos(3) - SrcPos(2)
    FindEntry = True

End Function

Private Sub mShowEntry()       'Shows the entry text
    rtbVisible.TextRTF = rtbHidden.SelRTF
End Sub


Private Property Let mLinkText(Tekst$)
  Dim lnk$, P&
  Dim l&
    rtbVisible.SelLength = 1
    HyperlinkSelect   'Select the link
    P = rtbVisible.SelStart
    l& = rtbVisible.SelLength
    If 0 <= rtbVisible.Find("§" & mLinkText & "§", P, P + l) Then 'Replace text
      rtbVisible.SelText = "§" & Tekst$ & "§"
    Else 'Create new
      rtbVisible.SelStart = P + l
      rtbVisible.SelText = "§" & Tekst$ & "§"
    End If
    If 0 > rtbVisible.Find("§" & Tekst$ & "§", P) Then Stop
    vSelHidden = True
End Property
Private Property Get mLinkText$()
  Dim lnk$, P1&, P2&
  HyperlinkSelect
  lnk = rtbVisible.SelText
  P1 = InStr(lnk, "§")
  P2 = InStr(P1 + 1, lnk, "§")
  If P1 <= 0 Or P2 <= 0 Then Exit Property
  mLinkText = Mid(lnk, P1 + 1, P2 - P1 - 1)
End Property
Private Function mEntryThis() As Boolean
  Dim i&(3)
    i(3) = rtbHidden.SelStart     'End of search
    i(2) = -1      ' Search from the start
   Do   'I2= this entry ,i1 the one before
      i(2) = rtbHidden.Find(leftMark, i(2) + 1, i(3))   'Next entry
      If i(2) < 0 Then Exit Do       'Found the end entry
      i(0) = i(1)
      i(1) = i(2)       'shif entries down
    Loop
    mEntryThis = mEntryFind("", i(1))
End Function

Private Property Get vSelHidden() As Boolean  'The RTF does not implement the hidden property
'This is a workaround where we brutally insert the RTF code
  Dim Rs$
  Rs$ = rtbVisible.SelRTF
'   MsgBox RS
  If (0 < InStr(Rs, "\v\")) Or (0 < InStr(Rs, "\v ")) Then
    vSelHidden = True
  Else
    vSelHidden = False
  End If
End Property
Private Property Let vSelHidden(Value As Boolean)   'Sets the currently selected test to be hidden
  Dim Rs$, P&, l&, s$
    P = rtbVisible.SelStart: l = rtbVisible.SelLength:
    s = rtbVisible.SelText
    'RS$ = rtbVisible.SelRTF
    If Value = False Then
      rtbVisible.SelRTF = s
    Else
      Rs = "{\rtf1\v " & s & "\v0}"       'Codes for hidden text
      rtbVisible.SelRTF = Rs:
    End If
    rtbVisible.SelStart = P: rtbVisible.SelLength = l:
End Property

Private Sub rtbVisible_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnMouse rtbVisible, Button, Shift
End Sub
