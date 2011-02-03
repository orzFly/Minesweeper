VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ctlIntegerStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim col As New Collection

Public Sub Push(ByVal value As Integer)
    col.Add value
End Sub

Public Function Pop()
    If Count = 0 Then Pop = 0: Exit Function
    Pop = col.Item(Count)
    col.Remove Count
End Function

Public Function Peek()
    If Count = 0 Then Peek = 0: Exit Function
    Peek = col.Item(Count)
End Function

Public Property Get Count() As Long
    Count = col.Count
End Property

Public Sub Clear()
    Set col = New Collection
End Sub
