Attribute VB_Name = "mdlMain"
Option Explicit

Private fp As String                    'This workbook's filepath
Private wb As Workbook                  'This workbook
Public ERR_STACK As ErrorHandler
Public log As Logger
Public wbDb As WorkbookDB

Public Sub Main(Optional eventNo As Integer = 0)
On Error GoTo ErrHandler
    Dim folderPath As String
    Dim log As Logger
    Set log = mdlFactory.getLog
    
    mdlMisc.OptimizedMode True ' Turns off all resource intensive features of excel to drastically speed up program
    SetupRun
    
    Set wb = ThisWorkbook
    fp = ThisWorkbook.Path
    
    Select Case eventNo
        Case 1
        'mdlRenewalRenamer.subRenewalRename folderPath, True, True, True
    End Select
    
    log.Add "---------------------" & vbLf & "Run ID: " & log.GetRunID & vbLf & "----------------------" & vbLf
    log.Add "-------------------------" & vbCrLf & "END RUN ID: " & log.GetRunID & vbCrLf & "--------------------------"
    
'''''''''''''''
'Error Handling
'''''''''''''''
ExitHub: 'Can act as a finally block.
    ThisWorkbook.Sheets("Main").Activate
    mdlMisc.OptimizedMode False
    Cleanup
    Exit Sub
    
ErrHandler:
    ThisWorkbook.Sheets("Main").Activate
    mdlMisc.OptimizedMode False
    
    ERR_STACK.Push "mdlMain.Main"
    ERR_STACK.DisplayErrorMsg Err.Source, Err.Description
    
    Cleanup
    On Error GoTo -1
End Sub

Public Sub SetupRun()
    Set ERR_STACK = New ErrorHandler
End Sub

Public Sub Cleanup()
    ' Cleanup code
    Set ERR_STACK = New ErrorHandler
End Sub

Public Sub testing_btn()
    'UserForm1.Show
    Dim strTemp As String
    strTemp = "Hello!!??" & vbNewLine
    strTemp = Trim$(strTemp)
    'mdlMisc.RemoveFileNameIllegalChars strTemp
    strTemp = vbNullString
    strTemp = ""
End Sub
Public Sub test()
    Main 0
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Various unused
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' For formatting picker:
'xlDialogFormatNumber
'xlDialogAlignment
'xlDialogFormatFont
'xlDialogBorder
'xlDialogPatterns
'xlDialogCellProtection



'    folderPath = mdlFileHandling.subChooseFolder(ThisWorkbook.Path)
'    mdlErr.Logger "---------------------" & vbLf & "Run ID: " & RUN_ID & vbLf & "----------------------" & vbLf
'    If folderPath <> "" Then
'        mdlErr.Logger "Folder path selected: " & folderPath
'    Else
'        mdlErr.Logger "User did not select a folder. No actions were taken."
'    End If
'    mdlErr.Logger "-------------------------" & vbCrLf & "END RUN ID: " & RUN_ID & vbCrLf & "--------------------------"
'    ThisWorkbook.Save

    
Sub ColorDialog()
    'Create variables for the color codes
    Dim FullColorCode As Long
    Dim RGBRed As Integer
    Dim RGBGreen As Integer
    Dim RGBBlue As Integer
    
    'Get the color code from the cell named "RGBColor"
    FullColorCode = Range("A1").Interior.Color
    
    'Get the RGB value for each color (possible values 0 - 255)
    RGBRed = FullColorCode Mod 256
    RGBGreen = (FullColorCode \ 256) Mod 256
    RGBBlue = FullColorCode \ 65536
    
    'Open the ColorPicker dialog box, applying the RGB color as the default
    If Application.Dialogs(xlDialogEditColor).Show _
        (1, RGBRed, RGBGreen, RGBBlue) = True Then
    
        'Set the variable RGBColorCode equal to the value
        'selected the DialogBox
        FullColorCode = ActiveWorkbook.Colors(1)
        
        'Set the color of the cell named "RGBColor"
        Range("A1").Interior.Color = FullColorCode
    
    Else
        'Do nothing if the user selected cancel
    End If

End Sub

