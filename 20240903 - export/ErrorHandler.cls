VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private errStack As Stack

Public Sub Push(CurrentSub As String)
    errStack.Push CurrentSub
End Sub

Public Sub PushRaise(ErrNumber As Long, ErrSource As String, CurrentSub As String, ErrDescription As String)
    errStack.Push CurrentSub
    Err.Raise ErrNumber, ErrSource, ErrDescription
End Sub

Public Function ToString() As String
    ToString = errStack.ToString
End Function

Public Sub DisplayErrorMsg(ErrSource As String, ErrDescription As String)
    Dim CompleteMessage As StringBuilder
    
    With CompleteMessage
        .Append "The following unexpected runtime error occurred:"
        .Append vbCrLf
        .Append ErrDescription
        .Append vbCrLf
        .Append vbCrLf
        .Append "Error Location is:"
        .Append vbCrLf
        .Append errStack.ToString
    End With
                    
    MsgBox CompleteMessage.ToString, vbCritical, "Error"
    errStack.Clear
End Sub

Private Sub Class_Initialize()
    Set errStack = New Stack
End Sub
