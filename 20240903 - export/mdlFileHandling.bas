Attribute VB_Name = "mdlFileHandling"
Option Explicit

' For an unused test below
Private Const BIF_NONEWFOLDERBUTTON As Long = &H200
Private Const BIF_BROWSEINCLUDEFILES As Long = &H4000

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' A folder picker dialogue box
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function subChooseFolder(strFolderLocation As String) As String
    On Error GoTo EH
    Dim strPath As String
    Dim openFileDialog As Office.FileDialog
    Dim strFolderName As String
    Set openFileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    strPath = strFolderLocation

    With openFileDialog
        .AllowMultiSelect = False
        ' Set the title of the dialog box.
        .Title = "Select a folder then hit OK or Cancel"
        .InitialFileName = strPath & "\"
        
        ' Show the dialog box. If the .Show method returns True, the
        ' user picked at least one file. If the .Show method returns
        ' False, the user clicked Cancel.
        If .Show = True Then
            strFolderName = .SelectedItems(1)
        End If
    End With
    subChooseFolder = strFolderName
    Set openFileDialog = Nothing
    Exit Function
    
EH:
    MsgBox "Error selecting folder"
    On Error GoTo -1
    Exit Function
End Function


Public Function subChooseFile(strFolderLocation As String, strFileName As String) As String
    On Error GoTo EH:
    
    Dim strPath As String
    Dim openFileDialog As Office.FileDialog
    Dim strFolderName As String
    Set openFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    strPath = strFolderLocation

    With openFileDialog
        .AllowMultiSelect = False
        ' Set the title of the dialog box.
        .Title = vbNullString
        
        If LenB(strFileName) = 0 Then
            .InitialFileName = strPath
        Else
            .InitialFileName = strPath & "\" & strFileName & ".xlsx"
        End If
        
        ' Clear out the current filters, and add our own.
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx;*.xls;*.xlsm"
        
        ' Show the dialog box. If the .Show method returns True, the
        ' user picked at least one file. If the .Show method returns
        ' False, the user clicked Cancel.
        If .Show = True Then
            strFolderName = .SelectedItems(1)
        End If
    End With
    If Not LenB(strFolderName) = 0 Then
        subChooseFile = strFolderName
    Else
        subChooseFile = strFolderLocation
    End If
    Set openFileDialog = Nothing
    
EH:
    MsgBox "Error selecting file"
    On Error GoTo -1
    Exit Function
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'subSaveFile
'
'Figures out a name for the indicated
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function strSaveFileName(strPath As String, strName As String, _
strSaveType As String, intExt As Integer) As String
    Dim currentName
    Dim newName
    Dim nameVersion
    Dim fileObject As Object
    
    currentName = strPath & ("\") & strName
    newName = currentName & " " & strSaveType & strFileFormatIntToStr(intExt)
    
    Set fileObject = CreateObject("Scripting.FileSystemObject")
    nameVersion = 1
    While fileObject.FileExists(newName)
        newName = currentName & "(" & CStr(nameVersion) & ")" & "  " & _
         strSaveType & strFileFormatIntToStr(intExt)
        nameVersion = nameVersion + 1
    Wend
    strSaveFileName = newName
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'strFileFormatIntToStr
'
'This function converts the indicated numerical file format and converts it to
'the string equivalent (with a period in the beginning)
'
'RETURNS: String equivalent of indicated fileFormat (in integer form)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function strFileFormatIntToStr(fileFormat As Integer)
    Select Case fileFormat
        Case 6
            strFileFormatIntToStr = ".csv"
        Case 51
            strFileFormatIntToStr = ".xlsx"
        Case 52
            strFileFormatIntToStr = ".xlsm"
        Case 56
            strFileFormatIntToStr = ".xls"
        Case xlHtml
            strFileFormatIntToStr = ".html"
        Case 200
            strFileFormatIntToStr = ".rfp"
        Case 201
            strFileFormatIntToStr = ".xml"
        Case Else
            strFileFormatIntToStr = ".txt"
    End Select
End Function

Public Function getFileFormat(wb As Workbook) As Integer
    Select Case wb.fileFormat
        Case xlOpenXMLWorkbookMacroEnabled '52 xlsm
            getFileFormat = 52
        Case xlExcel8 '56 xls
            getFileFormat = 56
        Case xlWorkbookDefault '51 xlsx
            getFileFormat = 51
    End Select
End Function

Public Function isFileOpen(FileName As String)
    Dim ff As Long, Errno As Long
    
    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    Errno = Err
    On Error GoTo 0
    
    Select Case Errno
        Case 0:
            isFileOpen = False
        Case 70:
            isFileOpen = True
        Case Else:
            Error Errno
    End Select
End Function


' Get the files specified
Function GetFiles(folderPath As String, Optional extFilter As String = vbNullString) As Variant
    Dim file As Variant, files As Variant
    Dim NumFiles As Long, Idx As Long
    
    ' Collect files only and ignore folders
    ' Loop through each file in the specified folder
    file = Dir(folderPath & extFilter)
    NumFiles = 0
    Do While LenB(file) <> 0
        NumFiles = NumFiles + 1
        If NumFiles = 1 Then
            ReDim files(1 To 1)
        Else
            ReDim Preserve files(1 To NumFiles)
        End If
        files(NumFiles) = file
        file = Dir()
    Loop
    GetFiles = files
End Function

Sub ProcessFiles(MyFolder As String)
    Dim MyFile As Variant
    Dim files As Variant
    Dim Idx As Long
    
    files = GetFiles(MyFolder)
    If Not IsEmpty(files) Then
        For Idx = LBound(files) To UBound(files)
            MyFile = files(Idx)
            ' Process the file here
            Debug.Print MyFile
        Next Idx
    Else
        Debug.Print "No files found for Dir(""" & MyFolder & """)"
    End If
End Sub


Public Function GetFilenameFromPath(ByVal strPath As String) As String
' Returns the rightmost characters of a string upto but not including the rightmost '\'
' e.g. 'c:\winnt\win.ini' returns 'win.ini'

    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function


Public Sub RemoveFileNameIllegalChars(ByRef textToStrip As String)
    Dim illegalFileCharacters As Variant
    illegalFileCharacters = Array("\", "/", "|", "<", ">", """", "*", ":", "?", ".", vbCr, vbLf)
    
    RemoveFromString textToStrip, illegalFileCharacters
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' A folder picker that would also show the files in the folder.
'
' Non-priority, to review later
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Shell_Browse_Folder_or_File()
    Dim WShell As Object
    Dim WShellFolder As Object
    Dim WShellFolderItem As Object
    Dim flags As Long
    Dim startFolder As Variant 'must be Variant
   
    Set WShell = CreateObject("Shell.Application")
   
    startFolder = ThisWorkbook.Path
    flags = BIF_BROWSEINCLUDEFILES Or BIF_NONEWFOLDERBUTTON
   
    Set WShellFolder = WShell.BrowseForFolder(0, "Select Folder", flags, startFolder)
    If WShellFolder Is Nothing Then
        MsgBox "Dialogue cancelled"
    Else
        Set WShellFolderItem = WShellFolder.Self
        If Not WShellFolderItem Is Nothing Then
            If WShellFolderItem.IsFolder Then
                MsgBox "Selected folder: " & WShellFolderItem.Path
            Else
                MsgBox "Selected file: " & WShellFolderItem.Path
            End If
        End If
    End If
   
End Sub



