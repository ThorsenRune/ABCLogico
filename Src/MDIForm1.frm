VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "ABC Logico 2015                                 R.Thorsen"
   ClientHeight    =   3192
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Resize()
  If DivX > 0.9 Then DivX = 0.9
  OnResize ScaleWidth * DivX, ScaleHeight * 0.7
End Sub
Public Sub OnResize(DivX, DivY)
  If WindowState = 1 Then Exit Sub
    fProgram.Left = 0: fProgram.Width = DivX
    fProgram.Top = 0: fProgram.Height = ScaleHeight
    
    fConsole.Left = fProgram.Width: fConsole.Width = Abs(ScaleWidth - fConsole.Left)
    fConsole.Top = 0: fConsole.Height = DivY
    
    frmMsg.Left = fConsole.Left: frmMsg.Width = Abs(ScaleWidth - fConsole.Left)
    frmMsg.Top = fConsole.Height: frmMsg.Height = ScaleHeight - fConsole.Height
End Sub

