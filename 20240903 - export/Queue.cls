VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pQueue As Collection

Public Property Get Count()
    Count = pQueue.Count
End Property

Public Function Enqueue(newItem As Variant)
    pQueue.Add newItem
End Function

Public Function Dequeue() As Variant
    If Count > 0 Then
        Dequeue = pQueue(1)
        pQueue.Remove 1
    End If
End Function

Public Function Peek() As Variant
    If Count > 0 Then
        Peek = pQueue(1)
    End If
End Function

Public Sub Clear()
    Set pQueue = New Collection
End Sub

Private Sub Class_Initialize()
    Set pQueue = New Collection
End Sub
