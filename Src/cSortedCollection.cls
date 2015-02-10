VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSortedCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Maintaining a sorted list  key, object , sorted by key
'mAdd object,[key]         'if no key it will be assigned as the last
'mRemove
'mItem(index/key,[byKey])      'Returns Empty if not found
Private arrKey() As Variant
Private arrObj() As Variant
Private arrLengt&
Private currIdx&
Private Sub mArrSetLengt(newSize&)
  arrLengt& = newSize&
  If newSize < 0 Then newSize& = 0
    ReDim Preserve arrKey(newSize&)
    ReDim Preserve arrObj(newSize&)
End Sub
Public Sub mClear()   'Clear the collection
  mArrSetLengt -1
End Sub
Public Function mUBound&()
   mUBound& = arrLengt
End Function
Private Sub Swap(i1, i2)
  Dim tmpObj, tmpKey, Dir&
  tmpKey = arrKey(i1):   arrKey(i1) = arrKey(i2): arrKey(i2) = tmpKey
  If IsObject(arrObj(i1)) Then
    Set tmpObj = arrObj(i1):
    Set arrObj(i1) = arrObj(i2):
    Set arrObj(i2) = tmpObj
  Else
    tmpObj = arrObj(i1):   arrObj(i1) = arrObj(i2): arrObj(i2) = tmpObj
  End If
End Sub

Private Sub mBubbleThis(idx&)     'Bubble this into place, empty keys will be at the end
' Call this with a starting point and it will bubble the item up or down to its proper position
  Dim tmpObj, tmpKey, Dir&, i&
  If arrLengt < 1 Then Exit Sub 'Nothing to bubble
  If idx = 0 Then idx = 1
  Stop    'Check this routine
  For i = idx - 1 To 0 Step -1  'Bubble down
    If arrKey(i) > arrKey(i + 1) Then    'Swap two elements
      Swap i, i + 1
    Else
      Exit For          'Stop
    End If
  Next
  For i = idx To arrLengt - 1 'Bubble up
    If arrKey(i) > arrKey(i + 1) Then    'Swap two elements
      Swap i, i + 1
    Else
      Exit For          'Stop
    End If
  Next
End Sub
Public Sub mAdd(Obj, Key)
  Dim idx&, i&
    idx = arrLengt + 1
    mArrSetLengt (idx) 'Add an element
    If IsObject(Obj) Then
      Set arrObj(arrLengt) = Obj
    Else
      arrObj(arrLengt) = Obj
    End If
    arrKey(idx) = Key
    For i = idx - 1 To 0 Step -1
      If arrKey(i + 1) < arrKey(i) Then
        Swap i, i + 1
      Else
        Exit For          'Done
      End If
    Next

End Sub
Public Sub mRemove(Key)
  Dim idx&, i&
  idx = idxByKey(Key)
  If 0 <= idx Then
     For i = idx To arrLengt - 1      'Bubble up
      Swap i, i + 1
     Next
     mArrSetLengt arrLengt - 1        'And remove
  Else
    Stop  'not found
  End If
End Sub
Public Function mItem(idx&)      'Return the element by index
  Dim i&
  If IsObject(arrObj(idx)) Then
    Set mItem = arrObj(idx)
  Else
    mItem = arrObj(idx)
  End If
End Function
Private Function idxByKey&(Key)  'if not existing it will be the index of the element before
  Dim i&, i1&, i2&
  If VarType(Key) = vbInteger Then 'Get by index
    i = Key
  Else                              'Bisection search for key
    i1 = 0: i2 = arrLengt
    Do
      i = (i1 + i2) \ 2         'Floor of inteval midpoint
      If (arrKey(i) < Key) Then
        i1 = i + 1 'the key is after i
      Else
        i2 = i
      End If
    Loop Until (i2 - i1) < 1    'Until interval is zero
  End If
  idxByKey& = i           'Return the index of the element
End Function
Private Sub Class_Initialize()
  arrLengt& = -1
  ReDim arrKey(0)
  ReDim arrObj(0)
End Sub
