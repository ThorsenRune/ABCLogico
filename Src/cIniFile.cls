VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public FileMyIni$
Public FoundKey As Boolean
Const KeySeparator = vbCrLf + vbTab
Const ArrySep = ";"
Public Sub OpenIniFile(FileName$, Optional VarName$)
  Dim AV
      AV = Shell("Notepad " & FileName$, vbNormalFocus)
      AppActivate AV, True
End Sub
Public Function SectionDelete(FileName$, Section$)
  WritePrivateProfileString Section$, vbNullString, "", FileName$
End Function
Public Function sMustFindFile(FileName$)
  If Dir(FileName) = "" Then
      FileName = InputBox("Must find file:", , FileName)
  End If
  sMustFindFile = (Dir(FileName) <> "")
End Function

Public Sub WinState(GetIt As Boolean, Frm As Form, Optional WinStateFile$, Optional Section$)
  Dim c As Control, i&
  'Set/get the State of the form.
  If Section = "" Then Section$ = Frm.Name
  If WinStateFile = "" Then WinStateFile$ = sPathBin & "\" & App.EXEName + ".win"
  Frm.WindowState = Setting(GetIt, WinStateFile, Section$, "State", Frm.WindowState)
  If Frm.WindowState = vbNormal Then
    Frm.Top = Setting(GetIt, WinStateFile$, Section$, "TOP", Frm.Top)
    Frm.Left = Setting(GetIt, WinStateFile$, Section$, "LEFT", Frm.Left)
    Frm.Width = Setting(GetIt, WinStateFile$, Section$, "WIDTH", Frm.Width)
    Frm.Height = Setting(GetIt, WinStateFile$, Section$, "HEIGHT", Frm.Height)
  End If
  'Ensure visibility
'  If Frm.WindowState = 1 And Frm.Visible Then
    If Frm.Left + Frm.Width / 2 < 0 Then
      Frm.Left = 0
    ElseIf Frm.Left + Frm.Width / 2 > MDIForm1.ScaleWidth Then
      Frm.Left = MDIForm1.ScaleWidth - Frm.Width
    End If
    If Frm.Top < 0 Then
      Frm.Top = 0
    ElseIf Frm.Top > MDIForm1.ScaleHeight * 0.9 Then
      Frm.Top = MDIForm1.ScaleHeight * 0.5
    End If
 ' End If
End Sub
Private Sub Class_Initialize()
'Set basedirectories
    If FileMyIni$ = "" Then FileMyIni$ = sPathBin & "\" & App.EXEName + ".ini"
    sMustFindFile FileMyIni$
End Sub
 
Private Sub ReadKey(ByVal FileName$, ByVal keyName$, Data$, RecStart&, EndPos&)
 Dim l&, Fil&, s$
     '      If Not IsFileExisting(FileName$) Then Exit Sub
    Fil = FreeFile
    On Error GoTo clsfile:
    Open FileName$ For Input As #Fil
    On Error GoTo 0
    EndPos& = LOF(Fil):
    l& = Len(keyName)
    Do                            'Seek Key
      If EOF(Fil&) Then RecStart& = EndPos&:  GoTo clsfile
      RecStart& = Seek(Fil)
      Line Input #Fil, s
      If Left(s, l) = keyName$ Then 'Found
        Data = Trim(Mid(s, 1 + l + Len(KeySeparator)))
        Do
          If EOF(Fil) Then
            EndPos& = Seek(Fil)
            GoTo clsfile
          Else
            EndPos& = Seek(Fil) - 2     'Remove CRLF
          End If
          Line Input #Fil, s
          If Left(s, 1) = "{" Then GoTo clsfile   'End of record
          Data = Data + Mid(s, 2)                 'Assume a blank or tab at the start of each line in record
        Loop
      End If
    Loop
clsfile:
    On Error GoTo 0
    Close #Fil
    Data = Replace(Data, vbCrLf + vbTab, vbCrLf)
End Sub
Private Sub WriteKey(ByVal FileName$, ByVal keyName$, Data$)
  Dim Fil&, RecStart&, RecEnd&
  Dim NewLen&, OldLen&
      Data = Replace(Data, vbCrLf, vbCrLf + vbTab)
      ReadKey FileName$, keyName$, "", RecStart&, RecEnd&      'Find old record start,end
      Fil = FreeFile
      Open FileName$ For Binary As #Fil
      NewLen& = Len(keyName$ + KeySeparator) + Len(Data$)
      OldLen& = RecEnd - RecStart
      If LOF(Fil) = 0 Then      'New File
        Put #Fil, 1, CStr(keyName$ + KeySeparator + Data$)
      ElseIf NewLen& > OldLen& Then             'Put it at the end
         'Delete old record
        If 0 < OldLen& Then Put #Fil, RecStart&, CStr(String(OldLen&, " "))
        Put #Fil, LOF(Fil) + 1, CStr(vbCrLf + keyName$ + KeySeparator + Data$)
      Else
        Put #Fil, RecStart, CStr(keyName$ + KeySeparator + Data$)
        'Fill rest with space
        If NewLen& < OldLen& Then Put #Fil, , CStr(String(OldLen& - NewLen&, " "))
      End If
      Close #Fil
End Sub
 
 
 
Public Function Setting(GetIt As Boolean, FileName$, Section$, keyName$, Optional ByRef Value)
  'If GetIt=false then it saves the value in SetupFile$ else it loads it
  'saving value "" will delete the key. Value will be returned in Setting for use with controls
  ' the file looks like this
  '  [SECTION]
  '     Key=Value
  Const MaxLen = 1024
  Const CrLf = "{lf}"
  Dim i&
  Dim NC As Long, s$
'    If File_Path(FileName) = "" Then FileName = WinStateFile$ = GetExecPath$ + File_Title(App.EXEName) + ".ini"
  'Read Data
    If IsMissing(Value) Then Value = ""
    If FileName = "" Then FileName = FileMyIni
    If GetIt Then
     s$ = String(MaxLen, 0)     'Create a buffer
     'Retrieve the string
     NC = GetPrivateProfileString(Section$, keyName$, "", s$, MaxLen, FileName$)
     FoundKey = (0 < NC)
     If 0 < NC Then   'Key is found
        s = Left$(s$, NC)
        ValueConvert False, Value, s
     End If
   End If
   If Not GetIt Then
  'Write Data
    'the setting to the file under    Application title -> Keyname
      ValueConvert True, Value, s
      If Len(s) > 1024 Then Stop 'stop' "Settings string overflow"
      If s = "" Then s = vbNullString 'If no value is present then delete the key
      NC = WritePrivateProfileString(Section$, keyName$, s, FileName$)
      If NC = 0 Then MsgBox "file access problem:" & FileName$
    End If
    Setting = Value
End Function


 

Private Sub ValueConvert(ToString As Boolean, Value, s$)
'2006/Jun
  Const CrLf = "{lf}"
    Dim i&, V, VType As VariantTypeConstants
        VType = VarType(Value)
     If ToString Then
        If vbArray = (vbArray And VType) Then                       'Array
            VType = VType Xor vbArray
'            On Error Resume Next        'Test if control array
'            i = UBound(Value)
'            On Error GoTo 0
            Select Case VType
            Case vbString
                s = Replace(Replace(Join(Value, "|"), vbCrLf, CrLf), vbTab, "{TAB}")
            Case 2 To 7
                For i = LBound(Value) To UBound(Value)
                    s = s + "|" + Str(Value(i))
                Next
                Mid(s, 1) = " "
            Case Else
                Stop 'stop' "Undefined"
            End Select
        Else                                                         'Scalar
            'the setting to the file under    Application title -> Keyname
            Select Case VType
            Case vbString: s = Replace(Replace(Value, vbCrLf, CrLf), vbTab, "{TAB}")
            Case vbBoolean: s = Str(Abs(Value))
            Case vbByte, vbDouble, vbInteger, vbSingle, vbLong: s = Str(Value)
            Case Else: Stop ' "Undefined"
            End Select
        End If
        If s = "" Then s = vbNullString     'If no value is present then delete the key
     ElseIf s <> "" Then  'Convert string to Value
        s = Replace(s, CrLf, vbCrLf)
        If vbArray = (vbArray And VType) Then                     'Array
            VType = VType Xor vbArray
            Select Case VType
            Case vbString
                 Value = Split(Replace(Replace(s, "{TAB}", vbTab), vbCrLf, vbCrLf), "|")
            Case 2 To 7
                V = Split(s, "|")
                On Error Resume Next
                ReDim Value(UBound(V))
                On Error GoTo 0
                For i = LBound(Value) To UBound(Value)
                    Value(i) = Val(V(i))
                Next
            Case Else
                Stop ' "Undefined"
            End Select
        Else
            Select Case VType
            Case vbString: Value = Replace(Replace(s, "{TAB}", vbTab), vbCrLf, vbCrLf)
            Case vbBoolean: Value = (Val(s) <> 0)
            Case vbByte, vbDouble, vbSingle, vbInteger, vbLong:      Value = Val(s)
            Case Else: Stop ' "Undefined"
            End Select
        End If
     End If
End Sub
Private Sub ArrayConvert(ToString As Boolean, Value, s$, RowNr&)
'2006/Jun
  Const CrLf = "{lf}"
    Dim i&, V, VType As VariantTypeConstants, RDim&
    Dim IsVector As Boolean
    
        i = -1:    On Error Resume Next:      i = UBound(Value, 2):      On Error GoTo 0
        If i < 0 Then IsVector = True Else IsVector = False
        If ToString Then
          s = vbTab
          For i = 0 To UBound(Value, 1)
            If IsVector Then V = Value(i) Else V = Value(i, RowNr)
            Select Case VarType(V)
            Case vbString
                s = s + V + vbTab + "|"
            Case 2 To 7
                s = s + Str(V) + vbTab + "|"
            Case vbBoolean:
              If V Then
                s = s + "1" + vbTab + "|"
              Else
                s = s + "0" + vbTab + "|"
              End If
                
            Case Else
                Stop 'stop' "Undefined"
            End Select
          Next
        If s = "" Then s = vbNullString     'If no value is present then delete the key
     ElseIf s <> "" Then  'Convert string to Value
       s = Replace(s, CrLf, vbCrLf)         'Restore CrLf characters
       V = Split(s, vbTab + "|")
        If IsVector Then
          If UBound(Value) < UBound(V) - 1 Then ReDim Preserve Value(UBound(V))
        Else
          If UBound(Value, 1) < UBound(V, 1) - 1 Then ReDim Preserve Value(UBound(Value, 1), UBound(Value, 2))
          If UBound(Value, 2) < RowNr Then ReDim Preserve Value(UBound(Value, 1), RowNr)
        End If
       If IsVector Then VType = VarType(Value(0)) Else VType = VarType(Value(0, 0))
       For i = 0 To UBound(V) - 1
         If IsVector Then
            Select Case VType
              Case vbString:
                Value(i) = V(i)  'Replace(Replace(V(i), "{TAB}", vbTab), vbCrLf, vbCrLf)
              Case vbBoolean: Value(i) = (Val(V(i)) <> 0)
              Case vbByte, vbDouble, vbSingle, vbInteger, vbLong:
                              Value(i) = Val((V(i)))
            Case Else
                Stop ' "Undefined"
            End Select
         Else
            Select Case VType
              Case vbString:
                Value(i, RowNr) = V(i)  'Replace(Replace(V(i), "{TAB}", vbTab), vbCrLf, vbCrLf)
              Case vbBoolean: Value(i, RowNr) = (Val(V(i)) <> 0)
              Case vbByte, vbDouble, vbSingle, vbInteger, vbLong:
                              Value(i, RowNr) = Val((V(i)))
            Case Else
                Stop ' "Undefined"
            End Select
          End If
        Next
     End If
End Sub


