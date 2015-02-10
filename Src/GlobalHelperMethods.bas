Attribute VB_Name = "GlobalHelperMethods"
Option Explicit
Public oCollisionCheck As New cCollision

Function isKindOf(Obj1 As Variant, Obj2 As Variant) As Boolean
'Determine if two objects are in the same collection
  Dim c(1) As Control
  isKindOf = False
  If Not isCtrlLoaded(Obj1) Then Exit Function
  If Not IsObject(Obj1) Then
    Exit Function
  ElseIf Not IsObject(Obj2) Then
    Exit Function
  End If
  Set c(0) = Obj1
  Set c(1) = Obj2
  If c(0) Is Nothing Then
      isKindOf = False
  ElseIf c(0).Name = c(1).Name Then
    isKindOf = True
  Else
    isKindOf = False
  End If
End Function
Public Sub Sleep(MilliSecs!)
  Dim T!
    T = Timer + MilliSecs / 1000
    While T > Timer: DoEvents: Wend                 'Speed of execution
End Sub

Public Sub CmdsEnable(f As Form, ThisContainer As Control, Value As Boolean)
'Disable/enable all controls. WA to let program container receive mouse moves directly
  Dim c As Control
  For Each c In f.Controls
    If TypeOf c Is Menu Then
    
    ElseIf c.Container Is ThisContainer Then
      If (f.ActiveControl Is c) Then
        c.Enabled = True
      ElseIf TypeOf c Is ucOpCode Then
        If c.IsSub Then 'can drop on this
          c.Enabled = True
        ElseIf c.sOPCode = "Call" Then
          c.Enabled = True      'Allow pseudodropping a procedure
        Else
          c.Enabled = Value
        End If
      Else
        c.Enabled = Value
      End If
    End If
  Next
End Sub
Public Sub mCtrlSetting(GetIt As Boolean, FN$, CContainer As Control, CColl As Object, FamilyName$, FromIdx$)
'CColl    pass the collection to save/retrieve
'The startindex can be used to avoid modification of existing controls
Dim StrArry$(), UBnd&, c As Control, CIdx&
  UBnd& = oIniFile.Setting(GetIt, FN$, "Items", FamilyName$ & "_Count", CColl.UBound)
  If GetIt Then   'Get controls
    CtrlArraySetCount CColl(0).Name, UBnd&          'Reserve the space for the array
    For Each c In CColl   'Try to read all controls
      If FromIdx$ <= c.Index Then   'Skip static controls
        CIdx& = c.Index
        StrArry$ = oIniFile.Setting(True, FN$, "Items", FamilyName$ & CIdx&, StrArry$)           'Read the data
        If c Is CContainer Then     'Skip myself
        ElseIf oIniFile.FoundKey Then   'Something was retrieved
          vCtrlStr(c) = StrArry$
          CColl(CIdx).Visible = True
          CColl(CIdx).ZOrder
        Else
        End If
      End If
    Next
  Else      'Save controls
    For Each c In CColl
      If c.Container Is CContainer Then       'Save controls in container
         CIdx& = c.Index
         oIniFile.Setting False, FN$, "Items", FamilyName$ & CIdx&, vCtrlStr(c)           'Write the data
      End If
    Next
  End If
End Sub

Private Property Get vCtrlStr(c)
 Dim s$()
  ReDim s$(8)
    If TypeOf c Is ucOpCode Then
      vCtrlStr = c.MyProperties
      Exit Property
    End If
    If isCtrlLoaded(c) Then
      s$(1) = c.Name
      s$(2) = c.Index
      s$(3) = c.Tag
      s$(4) = c.Container.Name
      If CtrlArryByName(c.Container.Name) Is Nothing Then
      Else
        s$(5) = c.Container.Index
      End If
      If c.Visible Then s$(6) = 1 Else s$(6) = 0
      s$(7) = Round(c.Top)
      s$(8) = Round(c.Left)
    End If
    vCtrlStr = s$
End Property
Private Property Let vCtrlStr(c, s)
  Dim aC, Cont, CCol
    If TypeOf c Is ucOpCode Then
      c.MyProperties = s
      Exit Property
    End If
    Set aC = CtrlArryByName(CStr(s(1)))  '      S$(1) = C.Name
    If aC Is Nothing Then Exit Property
    Set c = aC(CStr(s(2)))           '      S$(2) = C.Index
    c.Tag = CStr(s(3))
    Set Cont = CtrlByName(c.Parent, CStr(s(4))) '      S$(4) = C.Container.Name
    If Cont Is Nothing Then
      RunDebug "Did not find the control in parent?"
    ElseIf CtrlArryByName(Cont.Name) Is Nothing Then
       Set c.Container = Cont
'       DBCheck "Cont ok"
    Else
      Set CCol = CtrlArryByName(Cont.Name)
      Set c.Container = CCol(Val(s(5)))    '    Container index
    End If
    c.Visible = (s(6) = 1)      'we cant use cbool because the result depends on the environment language used
    c.Top = s(7)
    c.Left = s(8)
End Property
Public Sub RunDebug(Optional s$)   'Use the clipboard message from runtime crash as break condition
  Static Count&, Msg$
    Count& = Count& + 1
    Msg$ = "RunDebug  " & s$ & Count
    Clipboard.Clear:    Clipboard.SetText Msg
    If Msg = "message" Then Debug.Assert 0
End Sub
Function isCommand(c As Control) As Boolean
  If TypeOf c Is ucOpCode Then
    On Error Resume Next
    isCommand = Not c.IsFixed 'Not placed in the program
    On Error GoTo 0
  Else
    isCommand = False
  End If
End Function

Public Function hasCaption(Obj As Object) As Boolean
  hasCaption = False
  On Error Resume Next
    hasCaption = Obj.Caption <> ""
  On Error GoTo 0
End Function
Public Function hasToolTip(Obj As Object) As Boolean
  hasToolTip = False
  On Error Resume Next
    hasToolTip = Obj.ToolTipText <> ""
  If Err.Number = 0 Then hasToolTip = True
  On Error GoTo 0
End Function
Public Function ifCtrlHasIndex(c As Object)
  ifCtrlHasIndex = False
  On Error GoTo errhnd:
  ifCtrlHasIndex = (0 <= c.Index)
errhnd:    On Error GoTo 0
End Function
Public Function IsCtrlArray(Obj)
  IsCtrlArray = False
  On Error Resume Next
  IsCtrlArray = (Obj(0).Index = 0)
  On Error GoTo 0
End Function
Public Function hasParent(Obj As Object) As Boolean
  hasParent = False
  On Error Resume Next
    If Obj.Parent Is Nothing Then hasParent = False
  hasParent = (Err.Number = 0)
  On Error GoTo 0
End Function
Public Function ctrlID$(Obj As Object) ' returns a unique id for a control
  Dim Tag$
  If hasParent(Obj) Then Tag$ = Obj.Parent.Name & ":"
  Tag$ = Tag$ & Obj.Name
  If ifCtrlHasIndex(Obj) Then Tag$ = Tag$ & ":" & Obj.Index
  ctrlID$ = Tag
End Function
Public Sub ucOpCode_Defaults(Obj As ucOpCode)
 
  Static SetAll
  Dim i&
   i& = Obj.Index
    If i = 2 Then
      SetAll = Obj.BackColor
    Else
      Obj.BackColor = SetAll
    End If

  Select Case i
  Case 0:
    Obj.Caption = "Avanti"
    Obj.sOPCode = "Move"
  Case 1:
    Obj.Caption = "Su  ^"
    Obj.sOPCode = "Up"
  Case 2:
    Obj.Caption = "Giù v"
    Obj.sOPCode = "Down"
  Case 3:
    Obj.Caption = "Destra -->"
    Obj.sOPCode = "Right"
  Case 4:
    Obj.Caption = "<-- Sinistra"
    Obj.sOPCode = "Left"
  Case 5:
    Obj.Caption = "Mangia"
    Obj.sOPCode = "Eat"
    
  Case 6:
    Obj.Caption = "Cibo?"
    Obj.sOPCode = "Eatable"
    Obj.nIfCond = True: Obj.BackColor = &HC0FFFF
  Case 7:
    Obj.Caption = "Cane?"
    Obj.sOPCode = "Dangerous"
    Obj.nIfCond = True: Obj.BackColor = &HC0FFFF
    Case 8:
    Obj.Caption = "Ostacolo"
    Obj.sOPCode = "Wall"
    Obj.nIfCond = True: Obj.BackColor = &HC0FFFF
    Case 9:
    Obj.Caption = "Ripeti"
    Obj.sOPCode = "ForNext"
    Obj.nIfCond = True: Obj.BackColor = &HC0FFFF
    End Select
End Sub
