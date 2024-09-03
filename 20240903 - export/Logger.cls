VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pRunID As String
Private pLog As Queue
Private pDoLogFile As Boolean

Public Sub Add(entry As String, Optional tabCount As Integer = 0, Optional priority As Integer = 5, _
Optional sSource As String = "SourceNotProvided")
On Error GoTo EH
    Dim logEntry As Scripting.Dictionary
    Set logEntry = New Scripting.Dictionary
    
    logEntry.Add "entry", entry
    logEntry.Add "tab", tabCount
    logEntry.Add "priority", priority
    
    pLog.Enqueue logEntry
    
    If pDoLogFile Then AddToFile CStr(priority), sSource, entry
EH:
    Exit Sub
End Sub

Public Sub Clear()
    Set pLog = New Queue
End Sub

Public Sub AddToFile(ByVal sType As String, ByVal sSource As String, ByVal sDetails As String)
On Error GoTo EH
    Dim sFilename As String
    Dim fileIsOpen As Boolean
    fileIsOpen = False
    sFilename = ThisWorkbook.Path & "\LOG_CC.txt" '"C:\temp\logging.txt" &
    
    If LenB(Dir(sFilename)) = 0 Then
        Open sFilename For Output As #1
        Close #1
    End If
    
    ' Archive file at certain size
    If FileLen(sFilename) > 20000 Then
        Dim sArchiveName As String
        sArchiveName = Replace(sFilename, ".txt", Format(Now, "ddmmyyyy hhmmss.txt"))
        FileCopy sFilename, sArchiveName
        Kill sFilename
        Open sFilename For Output As #1
        Close #1
    End If
    
    ' Open the file to write
    Dim filenumber As Variant
    filenumber = FreeFile
    Open sFilename For Append As #filenumber
    fileIsOpen = True
    
    Print pRunID, CStr(Now) & "," & sType & "," & sSource _
                                & "," & sDetails & "," & Application.UserName
EH:
    If fileIsOpen Then Close #filenumber
    Exit Sub
End Sub

Public Function GetRunID() As String
    GetRunID = pRunID
End Function

Public Sub CreateLogger(ByVal doLogFile As Boolean)
    pDoLogFile = doLogFile
    pRunID = Format(Now, "yyyyMMddhhmmss")
    Set pLog = New Queue
End Sub
