VERSION 5.00
Begin VB.Form fSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proprietà del progetto"
   ClientHeight    =   5772
   ClientLeft      =   4632
   ClientTop       =   1728
   ClientWidth     =   6036
   Icon            =   "fSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   6036
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   732
      Index           =   0
      Left            =   4800
      ScaleHeight     =   684
      ScaleWidth      =   804
      TabIndex        =   17
      Top             =   840
      Width           =   852
      Begin VB.Image cmdGenerate 
         Height          =   588
         Left            =   0
         Picture         =   "fSettings.frx":E332
         Stretch         =   -1  'True
         ToolTipText     =   "Salva percorso e programma"
         Top             =   0
         Width           =   612
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   372
      Index           =   6
      Left            =   0
      ScaleHeight     =   372
      ScaleWidth      =   4692
      TabIndex        =   13
      Top             =   4800
      Width           =   4692
      Begin VB.ComboBox cbMouseStyle 
         Height          =   288
         ItemData        =   "fSettings.frx":19D88
         Left            =   1320
         List            =   "fSettings.frx":19D92
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   0
         Width           =   3372
      End
      Begin VB.Label label 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   3492
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   372
      Index           =   5
      Left            =   0
      ScaleHeight     =   372
      ScaleWidth      =   4692
      TabIndex        =   10
      Top             =   3840
      Width           =   4692
      Begin VB.ComboBox cbUserLevel 
         Height          =   288
         ItemData        =   "fSettings.frx":19DB9
         Left            =   1320
         List            =   "fSettings.frx":19DC9
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   0
         Width           =   3372
      End
      Begin VB.Label label 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Livello"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1332
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   972
      Index           =   4
      Left            =   0
      ScaleHeight     =   972
      ScaleWidth      =   4572
      TabIndex        =   6
      Top             =   1080
      Width           =   4572
      Begin VB.ComboBox oDogs 
         Height          =   288
         ItemData        =   "fSettings.frx":19DDB
         Left            =   1320
         List            =   "fSettings.frx":19DF6
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1452
      End
      Begin VB.Image oDog 
         Height          =   705
         Index           =   1
         Left            =   120
         Picture         =   "fSettings.frx":19E11
         Stretch         =   -1  'True
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   972
      Index           =   3
      Left            =   120
      ScaleHeight     =   972
      ScaleWidth      =   4332
      TabIndex        =   5
      Top             =   120
      Width           =   4332
      Begin VB.ComboBox oFoods 
         Height          =   288
         ItemData        =   "fSettings.frx":1A524
         Left            =   1200
         List            =   "fSettings.frx":1A53F
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1452
      End
      Begin VB.Image oFood 
         Height          =   705
         Index           =   0
         Left            =   0
         Picture         =   "fSettings.frx":1A55A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   972
      Index           =   2
      Left            =   120
      ScaleHeight     =   972
      ScaleWidth      =   4572
      TabIndex        =   4
      Top             =   2040
      Width           =   4572
      Begin VB.ComboBox oWalls 
         Height          =   288
         ItemData        =   "fSettings.frx":1C0C3
         Left            =   1320
         List            =   "fSettings.frx":1C0DE
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   1452
      End
      Begin VB.Image oWall 
         Height          =   708
         Index           =   0
         Left            =   0
         Picture         =   "fSettings.frx":1C0F9
         Stretch         =   -1  'True
         Top             =   0
         Width           =   708
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   372
      Index           =   1
      Left            =   0
      ScaleHeight     =   372
      ScaleWidth      =   4692
      TabIndex        =   1
      Top             =   4320
      Width           =   4692
      Begin VB.ComboBox cbLang 
         Height          =   288
         ItemData        =   "fSettings.frx":1D255
         Left            =   1320
         List            =   "fSettings.frx":1D265
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   0
         Width           =   3372
      End
      Begin VB.Label label 
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
         Caption         =   "Lingua"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3492
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Annulla"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "fSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMyNumber&
Dim oMyImg As Control
Dim vMyText$
Public Sub Settings(GetIt As Boolean)
  nFontSize& = oIniFile.Setting(GetIt, sFileSettings$, "Settings", "Font", 12)
  If GetIt Then cbLang.Text = sLang Else sLang cbLang.Text
  oDogs.ListIndex = oIniFile.Setting(GetIt, sFileSettings$, "Settings", "nDogs", oDogs.ListIndex)
  oFoods.ListIndex = oIniFile.Setting(GetIt, sFileSettings$, "Settings", "Food", oFoods.ListIndex)
  oWalls.ListIndex = oIniFile.Setting(GetIt, sFileSettings$, "Settings", "nWalls", oWalls.ListIndex)
  fProgram.cbProgName.ListIndex = oIniFile.Setting(GetIt, sFileSettings$, "Settings", "ProgramNumber", fProgram.cbProgName.ListIndex)

  cbUserLevel.ListIndex = oIniFile.Setting(GetIt, sFileSettings$, "Settings", "UserLevel", cbUserLevel.ListIndex)
  cbMouseStyle.ListIndex = oIniFile.Setting(GetIt, sFileSettings$, "Settings", "DragStick", cbMouseStyle.ListIndex)
End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub


Public Function vProgName$()
  If fProgram.cbProgName.ListIndex < 0 Then fProgram.cbProgName.ListIndex = 0
  vProgName$ = fProgram.cbProgName.Text
End Function

Private Sub cbMouseStyle_Click()
    oResizer.nDragStick = cbMouseStyle.ListIndex = 1
End Sub

 

Private Sub cbUserLevel_Click()
    'Level of complexity for the user (currently unused)
  fProgram.ShowCommands cbUserLevel.ListIndex
End Sub

Private Sub cmdGenerate_Click()
 fConsole.ReGenerate
End Sub

Private Sub cmdSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then
  mUserProg_Save
 End If
End Sub

 
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Edit text for control
  If KeyCode = vbKeyF2 Then frmMsg.msgEdit ActiveControl
End Sub

Private Sub Form_Load()
  Show
  mCaptionTexts True, Me        'Get language specific texts for captions
  Settings True
  KeyPreview = True
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
  Settings False
End Sub

Private Sub label_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMsg.mMsgOnMouse label(Index), Button, Shift
End Sub

Private Sub OKButton_Click()
 Settings False
 Unload Me
End Sub

