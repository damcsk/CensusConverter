Attribute VB_Name = "mdlFactory"
Option Explicit

Public Function getLog(Optional doLogFile As Boolean = False) As Logger
    If log Is Nothing Then
        Set log = New Logger
        log.CreateLogger False
    End If
    Set getLog = log

    If doLogFile Then
        
    End If
End Function

Public Function GetDb() As WorkbookDB
    If wbDb Is Nothing Then
        Set wbDb = New WorkbookDB
    End If
    Set GetDb = wbDb
End Function
