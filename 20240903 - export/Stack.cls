VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim pStack As Collection

Public Function Pop() As Variant
    If Count > 0 Then
        Pop = pStack.Item(Count)
        pStack.Remove Count
    End If
End Function

Public Function Push(newItem As Variant) As Variant
    pStack.Add newItem
    Push = pStack.Item(Count)
End Function

Public Function Peek() As Variant
    If Count > 0 Then
        Peek = pStack(1)
    End If
End Function

Public Property Get Count()
    Count = pStack.Count
End Property

Public Sub Clear()
    Set pStack = New Collection
End Sub

Public Function ToString() As String
    Dim elm As Variant
    Dim stackString As StringBuilder
    Set stackString = New StringBuilder
    
    For Each elm In pStack
        stackString.Append CStr(elm)
        stackString.Append vbCrLf
    Next elm
    ToString = stackString.ToString
End Function

Private Sub Class_Initialize()
    Set pStack = New Collection
End Sub
