Attribute VB_Name = "Shared"
Option Explicit
'Translate controls by leftclick+ctrlbutton
'Translate texts by rewriting and press save
Private propDirection As eDir
Public nFontSize&
Public DivX!, DivY!  'Screen division
Public Function sLang$(Optional SetLang$)
  Static s$
  If SetLang <> "" And (SetLang <> s) Then ' another language has been set
    'you should restart the programme if SetLang<>""
    s = oIniFile.Setting(False, sFileSettings$, "Settings", "Language", SetLang$)
    SetDataDir      'Get language dependent files
    mCaptionTexts True, fProgram
    mCaptionTexts True, fConsole
    MsgBox s
  ElseIf s$ = "" Then      'NO or another language has been set
    s = oIniFile.Setting(True, sFileSettings$, "Settings", "Language", SetLang$)
  End If
  sLang$ = s
End Function

Public Sub SetDataDir()
  sPathData = App.Path + "\.."
  ChDir sPathData
  sPathData = CurDir + "\data"
  sPathBin = CurDir + "\bin"
  If Dir(sPathData, vbDirectory) = "" Then MkDir sPathData
  If Dir(sPathBin, vbDirectory) = "" Then MkDir sPathBin
  sFileHelp = sPathBin + "\ABCL_Help_" & sLang & ".RTF"
  sFileText$ = sPathBin + "\ABCL_Text_" & sLang & ".TXT"
  sFileSettings$ = oIniFile.FileMyIni
End Sub
Public Sub mUserProg_Save()
  oIniFile.SectionDelete sFilePrg, "Items"      'Clear the settings to avoid cluttering
  fProgram.Settings False, sFilePrg
  fConsole.Settings False, sFilePrg
End Sub
Public Sub Main()
  DivX! = 0.4: DivY = 0.7
  SetDataDir
  fSettings.Hide       'Hide settings but load values
  fProgram.Show
  fConsole.Show
  vDirection = Dirdown
  frmMsg.vDocked = False
  frmMsg.mMsgByTagStr "About"       'Start message
  frmMsg.mMsgSplash                             'Show as modal
  fProgram.UserProgram_Load
End Sub
Public Sub Terminate()
  Unload fProgram
End Sub

Private Function GetDescrByCtrl(c As Control)
  Dim txt$
  Dim cc As Object
  If IsCtrlArray(c) Then txt$ = c(0).ToolTipText Else txt$ = c.ToolTipText
  If txt = "" Then
    Set cc = CtrlArryByName(c.Name)
    txt = cc(0).ToolTipText
  End If
GetDescrByCtrl = txt
End Function
Public Sub StatusRefresh(c As Control)
  Dim txt$
  txt = GetDescrByCtrl(c)
  fConsole.ucStatusBar1.Value = txt$ & ": x=" & Fix(c.Left + 0.5) & ",  y=" & Fix(c.Top - 0.5)
End Sub
Public Property Let vCatPosX(Val&)
  fConsole.oCat(vDirection).Left = Val&
End Property
Public Property Get vCatPosX&()
  vCatPosX& = fConsole.oCat(vDirection).Left
End Property
Public Property Let vCatPosY(Val&)
  fConsole.oCat(vDirection).Top = Val& + 1
End Property
Public Property Get vCatPosY&()
  vCatPosY = fConsole.oCat(vDirection).Top - 1
End Property
Public Property Let vDirection(vDir As eDir)
  Dim i&
  If fConsole.Visible = False Then Exit Property
    For i = 0 To fConsole.oCat.UBound     'Hide cats
        fConsole.oCat(i).Visible = False
    Next
    fConsole.oCat(vDir).Left = vCatPosX
    fConsole.oCat(vDir).Top = vCatPosY + 1
    fConsole.oCat(vDir).Visible = True
    propDirection = vDir
    StatusRefresh fConsole.oCat(vDir)
End Property
Public Property Get vDirection() As eDir
  vDirection = propDirection
End Property
'Get the control array by its member
Public Function CtrlArryByName(CName$) As Object
 If CName$ = fProgram.oInstr(0).Name Then
      Set CtrlArryByName = fProgram.oInstr
    ElseIf CName$ = fProgram.oProcedure(0).Name Then
      Set CtrlArryByName = fProgram.oProcedure
    ElseIf CName$ = fConsole.oFood(0).Name Then
      Set CtrlArryByName = fConsole.oFood
    ElseIf CName$ = fConsole.oDog(0).Name Then
      Set CtrlArryByName = fConsole.oDog
    ElseIf CName$ = fConsole.oWall(0).Name Then
      Set CtrlArryByName = fConsole.oWall
    ElseIf CName$ = fConsole.oCat(0).Name Then
      Set CtrlArryByName = fConsole.oCat
    ElseIf CName$ = fConsole.oTarget(0).Name Then
      Set CtrlArryByName = fConsole.oTarget
    Else
'      DBCheck "You forgot to declare this" & CName$
    End If
End Function
Public Sub CtrlArraySetCount(CName$, UBound1&)
  Dim CtrlColl As Variant, K As Control
  Dim i&
    Set CtrlColl = CtrlArryByName(CName$)
    If CtrlColl Is Nothing Then Stop      'Make sure the number of controls are exact
    For i = CtrlColl.UBound To UBound1& Step -1    'First remove contained controls
      Set K = CtrlColl(i)
      If isCtrlLoaded(K) Then
          K.Visible = False
          CtrlRemove K
      End If
    Next
    While CtrlColl.UBound < UBound1&:
        Load CtrlColl(CtrlColl.UBound + 1):
    Wend
End Sub
Public Sub CtrlRemove(K As Control)
  Dim c As Control
    For Each c In K.Parent.Controls
      If TypeOf c.Container Is ucOpCode Then
      If c.Container Is K Then
        Set c.Container = c.Parent
        CtrlRemove c
      End If
      End If
    Next
    On Error Resume Next
    Unload K
    On Error GoTo 0
End Sub
Public Function isCtrlLoaded(c) As Boolean
'Test if a control is loaded
  Dim cc As Control, f As Form
  For Each f In Forms
    For Each cc In f.Controls
      If c Is cc Then isCtrlLoaded = True: Exit Function
    Next
  Next
  isCtrlLoaded = False
End Function
Function oGetControlByIdx(Source As Control, Nr&) As Control     'make sure the control exist, otherwise load it
 'Loads a control if it does not exist
  Dim o As Object, c As Control
    Set o = CtrlArryByName(Source.Name)
    If o Is Nothing Then Stop
    While o.UBound < Nr
      Load o(o.UBound + 1)
      Set c = o(o.UBound)
      c.Top = Source.Top
      c.Visible = Source.Visible
      If TypeOf c Is ucOpCode Then c.IsFixed = False
    Wend
    On Error Resume Next      'Make sure it exist
      If Not o(Nr).Visible Then
        If Err.Number Then Load o(Nr)
      End If
    On Error GoTo 0
    Set oGetControlByIdx = o(Nr)
End Function

Public Function DeCode$(c As Control)
  If TypeOf c Is ucOpCode Then
 '   DeCode = fProgram.oInstr(1). sOPCode
    DeCode = c.sOPCode
  ElseIf isCtrlLoaded(c) Then
    DeCode$ = c.Tag
  End If
End Function
Function CtrlByName(Parent As Form, CName$) As Control    'make sure the control exist, otherwise load it
 'Loads a control if it does not exist
  Dim o As Object, c As Control
    For Each c In Parent.Controls
      If c.Name = CName Then
        Set CtrlByName = c
        Exit Function
      End If
    Next
End Function
Private Sub mCaptionTextSub(GetIt As Boolean, c As Object)
 Dim Tag$, s$
 Dim StrArr&()
     Tag$ = ctrlID$(c)
    If hasCaption(c) Then
         s$ = oIniFile.Setting(GetIt, sFileText, "Captions", Tag, c.Caption)
         If c.Caption <> s$ Then c.Caption = s$
    End If
    If hasToolTip(c) Then
         s$ = oIniFile.Setting(GetIt, sFileText, "ToolTips", Tag, c.ToolTipText)
         If c.ToolTipText <> s$ Then
          c.ToolTipText = s$
        End If
    End If
    If TypeOf c Is ComboBox Then      'Save list of entries
      ListContent(c) = oIniFile.Setting(GetIt, sFileText, c.Name, Tag, ListContent(c))
    End If
End Sub
Public Sub mCaptionTexts(GetIt As Boolean, ParentOrCtrl As Object)   'Get the texts for the captions
  Dim c As Control, Tag$, s$
  Dim OnlyC As Control
  If TypeOf ParentOrCtrl Is Control Then     'A control has been passed
     mCaptionTextSub GetIt, ParentOrCtrl
  ElseIf TypeOf ParentOrCtrl Is Form Then
      mCaptionTextSub GetIt, ParentOrCtrl
      For Each c In ParentOrCtrl.Controls
        mCaptionTextSub GetIt, c
      Next
  Else
    For Each c In ParentOrCtrl.Controls
      mCaptionTextSub GetIt, c
    Next
  End If
End Sub

Public Property Let ListContent(Combo1 As ComboBox, Lst)
    Dim i, DefI&
    DefI& = Combo1.ListIndex
        If IsArray(Lst) Then
          Combo1.Clear
          For i = 0 To UBound(Lst)
            If Lst(i) <> "" Then Combo1.AddItem Lst(i)
          Next
        Else
          Combo1.Clear
          Combo1.AddItem (Lst)
        End If
      If DefI < 0 Or DefI > Combo1.ListCount Then DefI = 0
      Combo1.ListIndex = DefI
End Property
Public Property Get ListContent(Combo1 As ComboBox)
    Dim i, l() As String
    If Combo1.ListCount > 0 Then
        ReDim l(Combo1.ListCount - 1)
            For i = 0 To Combo1.ListCount - 1
              l(i) = Combo1.List(i)
            Next
    End If
    ListContent = l
End Property
