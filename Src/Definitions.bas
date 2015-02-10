Attribute VB_Name = "Definitions"
Option Explicit
Public oIniFile As New cIniFile    'File with settings and languages
Public oResizer  As New cResizer
'Global stuff that is visible thoughout the project
Public pProgramState As eProgramstate
Public ExecSpeed!
Public sFilePrg$    'The name of the currently developed program
Public sPathData$   'Path to data
Public sPathBin$    'Path to bin
Public sFileHelp$   'ABCL_Help_UK.RTF       language dependent help file RTF
Public sFileText$   'ABCL_Text_UK.INI       language depentent captions etc  INI
Public sFileSettings$   'Program settings
Public uc1 As ucOpCode
Public Enum eDir         'Direction of moving elements
  Dirup = 0
  DirLeft = 1
  Dirdown = 2
  DirRight = 3
End Enum
Enum eProgramstate  'State of the progrm
  eDB_stop = 0      '   program is stopped or not running
  eDB_Step = 1      '   Single step to next instruction
  eDB_Run = 2       '   Execute each instruction with a 'clock speed'
  eDB_Break = 3
End Enum
Enum eExeColor      'Color codes for instructions according to
  WillExec = vbGreen   '  Will execute based upon the condition
  NoExec = vbRed       '  skip because condition is not met
  CondTest = vbYellow  '  the condition to test
  Passive = &HE0E0E0    'Background of executed command
End Enum
Public Type Point
        X As Long
        Y As Long
End Type

 




