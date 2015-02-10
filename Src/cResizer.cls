VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Class Originally intended for resizing windows but mostly used for Used for drag drop control
Public nDragStick As Boolean      'If set then the drag will start on a mouseclick rather than click-drag
Dim currentSizingObject As Object
Dim X1 As Single, Y1 As Single      'Click position in control when starting dragging
Public Sub StartSizing(Obj As Object)
  StopSizing
  Set currentSizingObject = Obj
  currentSizingObject.Enabled = False 'take away focus to container form
End Sub

Public Sub StopSizing()
  If Not currentSizingObject Is Nothing Then
    currentSizingObject.Enabled = True 'Reenable last used object
  End If
  Set currentSizingObject = Nothing
End Sub
Public Sub MouseMove(X As Single, Y As Single)
  If Not currentSizingObject Is Nothing Then
    currentSizingObject.UserControl_MouseMove 9, 0, X, Y
  End If

End Sub
Public Sub StopDrag(Source As Control, Target As Object, X As Single, Y As Single)
  Dim i1, i2
  If Source Is Nothing Then Exit Sub    'Check source validity
    'Check target for valid container type
    If Target Is Source Then        'Dont drop on yourself
    ElseIf TypeOf Target Is Form Or TypeOf Target Is Frame Or TypeOf Target Is PictureBox Or TypeOf Target Is ucOpCode Then
            'Set the container to target object
            On Error Resume Next
            Set Source.Container = Target
            On Error GoTo 0
            If Not Source.Container Is Target Then Exit Sub  'Error e.g. parent dragged onto child
            If TypeOf Target Is CommandButton Or TypeOf Target Is ucOpCode Then
              Source.Left = X - X1
              Source.Top = Y - Y1
            Else
              Source.Left = X - Target.ScaleX(X1, vbTwips, Target.ScaleMode)
              Source.Top = Y - Target.ScaleY(Y1, vbTwips, Target.ScaleMode)
            End If
            
           
        Else 'Target is not a Container
            'Is the target an object within itself
            If Not Source Is Target.Container Then
                'Set the container to target object's container
                Set Source.Container = Target.Container
                Source.Left = Target.Left + (X - X1)
                Source.Top = Target.Top + (Y - Y1)
            Else
                If X > X1 Then Source.Left = Source.Left + Target.Left + (X - X1) Else Source.Left = Source.Left - (X1 - X)
                If Y > Y1 Then Source.Top = Source.Top + Target.Top + (Y - Y1) Else Source.Top = Source.Top - (Y1 - Y)
           End If
        End If
'    Else 'Dropped on itself
'        If X > X1 Then Source.Left = Source.Left + (X - X1) Else Source.Left = Source.Left - (X1 - X)
'        If Y > Y1 Then Source.Top = Source.Top + (Y - Y1) Else Source.Top = Source.Top - (Y1 - Y)
'    End If
 
    Source.ZOrder 0 'Add if you want the control moved to the front
    
End Sub
Public Function MouseDown(Source As Control, Button As Integer, X As Single, Y As Single) As Boolean  'Drag on mouse down
  If Not nDragStick Then StartDrag Source, Button, X, Y
End Function
Public Function MouseUp(Source As Control, Button As Integer, X As Single, Y As Single) As Boolean    'Drag on mouse up
  If nDragStick Then StartDrag Source, Button, X, Y
End Function
Public Function StartDrag(Source As Control, Button As Integer, X As Single, Y As Single) As Boolean
    StartDrag = False       'Only true when dragging has started
    If Button = vbLeftButton Then    'Move but no buttom so drag is not what you want
      X1 = X: Y1 = Y 'Store the initial X and Y of the mouse on the drag object
      Source.Drag vbBeginDrag
      StartDrag = True
    Else
    End If
End Function
 
Public Function DontDrop(Target As Control)
'Call this when there is a dragover on a control that should not be
'     dropped upon
' Put a DontDrop calling controlobject in the dragover event and a DontDrop Nothing in the mousemove event
'
  Static LastControl As Control
    If isCtrlLoaded(LastControl) Then
      LastControl.Enabled = True
    End If
    If Target Is Screen.ActiveControl Then    'Its myself
    ElseIf Not Target Is Nothing Then
      Target.Enabled = False
      Set LastControl = Target
    End If
End Function
