Attribute VB_Name = "mdlMisc"
Option Explicit
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernal32" (ByRef lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernal32" (ByRef lpFrequency As Currency) As Long

' VBA optimization notes:
' - Run in optimized mode
' - Check for empty strings using "If LenB(strText) = 0 Then" instead of "If strText = "" Then"
' - Use Stringbuilder instead of &
' - Use vbNullString instead of ""
' - Use string versions of methods instead of variant versions for the following methods:
'        Left$(), Mid$(), Right$(), Chr$(), ChrW$(),
'        UCase$(), LCase$(), LTrim$(), RTrim$(), Trim$(),
'        Space$(), String$(), Format$(), Hex$(), Oct$(),
'        Str$(), Error$
' - Check if string to replace is in string to search before running replace. Using InStrB() is possibly good for this
' - InStrB is 10 times faster than using InStr, so possibly good to check for existence with InStrB and THEN use InStr if it returns true for in string.
' - Use string compare, especially in vbBinaryCompare.
'      + if you need the comparison to be case-insensitive, use UCase$ or LCase$
' - Use InStrB if you don't care about the index of the found string
' - Use ByRef and not ByVal for String whenever possible (to avoid the overhead of copied strings)
' - Don't use "Dim x as new MyClass", use the two line version instead
' - Always use option Explicit
' - Do not use vbTextCompare with StrComp(), InStr(), and Replace(); in fact, avoid vbTextCompare as possible
' - If you are going to need the value of a property more than once, assign it to a variable. Variables are generally 10 to 20 times faster than properties.
' - Don't Replace, InStr, or Split using vbNewLine or vbCrLf. You have to use individual vbLf and vbCr.
Private silencedFlag As Boolean

Public Sub RemoveSpacesAndSpecialChars(ByRef textToStrip As String, Optional ByVal keepHashSign As Boolean = False)
    Dim charsToRemove As Variant
    If keepHashSign Then
        charsToRemove = Array(" ", "-", "\", "/", "!", "@", "$", "%", "^", "&", "*", "(", ")", _
            """", ":", "{", "[", "]", "}", "|", "`", "+", "=", ";", ",", vbCr, vbLf)
    Else
        charsToRemove = Array(" ", "-", "\", "/", "!", "@", "$", "%", "^", "&", "*", "(", ")", _
            """", ":", "{", "[", "]", "}", "|", "`", "+", "=", ";", ",", vbCr, vbLf, "#")
    End If
    
    RemoveFromString textToStrip, charsToRemove
End Sub

Public Sub RemoveFromString(ByRef textToStrip As String, ByRef charsToRemove As Variant)
    Dim char As Variant

    For Each char In charsToRemove
        If InStrB(textToStrip, char) <> 0 Then
            textToStrip = Replace(textToStrip, char, vbNullString)
        End If
    Next char
End Sub

'''''''''''''''''''''''''''''
'Only digits. Will delete any
'value that isn't a digit as soon
'as it's typed. Gives the illusion
'that typing anything other than
'digits is disabled.
'''''''''''''''''''''''''''''
Public Sub removeNonDigits(controlCurrent As Control)
    If Not mdlMisc.subIsDigitsOnly(controlCurrent.value) And Len(controlCurrent.value) > 0 Then
        controlCurrent.value = Left(controlCurrent.value, _
        Len(controlCurrent.value) - 1)
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'These functions do exactly as they say: caculate the age or DOB given the other
'
'The optional current date is if you want to give a theoretical age from a past
'or future date.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function calculateAge(dateBorn As Date, Optional currentDate As Date = 0) As Integer
    currentDate = Date
    calculateAge = DateDiff("yyyy", dateBorn, currentDate) + _
     (currentDate < DateSerial(Year(currentDate), Month(dateBorn), _
      Day(dateBorn)))
End Function

Function calculateDOB(ageGiven As Integer) As Date
    calculateDOB = DateAdd("yyyy", -ageGiven, Date)
End Function

Function IsUserFormLoaded(ByVal UFName As String) As Boolean
    Dim UForm As Object
     
    IsUserFormLoaded = False
    For Each UForm In VBA.UserForms
        If UForm.Name = UFName Then
            IsUserFormLoaded = True
            Exit For
        End If
    Next
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine checks all the header cells in the array against each possible
'string for the column we are looking for (the possible strings are
'deliminated by a semicolon).
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function intFindCol(colNames As String, arrToSearch As _
Variant, Optional invertedAxis As Boolean = False) As Integer
    Dim currentCol As Integer
    Dim currentColName As Variant
    Dim index As Integer
    Dim headerLength As Integer

    headerLength = UBound(arrToSearch, IIf(invertedAxis, 1, 2))
    
    'Check each potential string
    For Each currentColName In Split(colNames, ";")
        'Check each cell in the header row
        For index = 1 To headerLength
            'Exit when a match is found
            Dim temp As Variant
            temp = arrToSearch(IIf(invertedAxis, index, 1), IIf(invertedAxis, 1, index))
            
            If LCase$(arrToSearch(IIf(invertedAxis, index, 1), IIf(invertedAxis, 1, index))) = LCase$(currentColName) Then
                intFindCol = index
                Exit Function
            End If
        Next index
    Next currentColName
    intFindCol = -1
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Simple function for mixed data type dictionaries.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDictElm(dict As Scripting.Dictionary, key As Variant) As Variant
    Select Case VarType(dict(key))
        Case vbObject
            Set GetDictElm = dict(key)
        Case Else
            GetDictElm = dict(key)
    End Select
End Function

'Uses Range.Find to get a range of all find results within a worksheet
' Same as Find All from search dialog box
' Parameters:
'   * Same as native .Find function
'   * iDoEvents parameter: performs a DoEvents between each iteration (to keep excel from hanging in long searches)
' Notes:
'   * With Lookin= xlValues, hidden cells are not searched.
'   * What parameter has a 255 character limitation (native Excel limitation)
' Returns: a range with all matched cells found
Function FindAll(rng As Range, ByVal What As Variant, Optional LookIn As XlFindLookIn = xlFormulas, _
Optional LookAt As XlLookAt = xlWhole, Optional SearchOrder As XlSearchOrder = xlByColumns, _
Optional SearchDirection As XlSearchDirection = xlNext, Optional MatchCase As Boolean = False, _
Optional MatchByte As Boolean = False, Optional SearchFormat As Boolean = False, _
Optional iDoEvents As Boolean = False) As Range
    Dim NextResult As Range, result As Range, area As Range
    Dim FirstMatch As String
    
    If Len(What) > 255 Then Err.Raise 1, "FindAll", "Parameter 'What' must not have more than 255 characters"
    For Each area In rng.Areas
        FirstMatch = vbNullString
        With area
            Set NextResult = .Find(What:=What, After:=.Cells(.Cells.Count), LookIn:=LookIn, _
                                    LookAt:=LookAt, SearchOrder:=SearchOrder, SearchDirection:=SearchDirection, MatchCase:=MatchCase, MatchByte:=MatchByte, SearchFormat:=SearchFormat)
            
            If Not NextResult Is Nothing Then
                FirstMatch = NextResult.Address
                Do
                    If result Is Nothing Then
                        Set result = NextResult
                    Else
                        Set result = Union(result, NextResult)
                    End If
                    Set NextResult = .FindNext(NextResult)
                    
                    If iDoEvents Then DoEvents
                Loop While Not NextResult Is Nothing And NextResult.Address <> FirstMatch
            End If
        End With
    Next
    
    Set FindAll = result
End Function

''Usage of FindAll:
'Sub testUsageOfFindAll()
'  Dim SearchRange As Range, SearchResults As Range, rng As Range
'    Set SearchRange = MyWorksheet.UsedRange
'    Set SearchResults = FindAll(SearchRange, "Search this")
'
'    If SearchResults Is Nothing Then
'        'No match found
'    Else
'        For Each rng In SearchResults
'            'Loop for each match
'        Next
'    End If
'End Sub


' Converts column integer to string of letters
Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function


'Sort an array of objects using a given property 'propName'
Sub SortObjects(list, propName As String)
    Dim First As Long, Last As Long, i As Long, j As Long, vTmp, oTmp As Object, arrComp()
    First = LBound(list)
    Last = UBound(list)
    'fill the "compare" array...
    ReDim arrComp(First To Last)
    For i = First To Last
        arrComp(i) = CallByName(list(i), propName, VbGet)
    Next i
    'now sort by comparing on `arrComp` not `list`
    For i = First To Last - 1
        For j = i + 1 To Last
            If arrComp(i) > arrComp(j) Then
                vTmp = arrComp(j)          'swap positions in the "comparison" array
                arrComp(j) = arrComp(i)
                arrComp(i) = vTmp
                Set oTmp = list(j)             '...and in the original array
                Set list(j) = list(i)
                Set list(i) = oTmp
            End If
        Next j
    Next i
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Checks if a given range has any numbers in the text/value
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function HaveNumbers(oRng As Range) As Boolean
    On Error GoTo EH
    HaveNumbers = HasNumbers(oRng.Text)
Exit Function
EH:
    ERR_STACK.PushRaise Err.Number, Err.Source, "mdlRenewalRenamer.HaveNumbers", Err.Description
End Function

Function HasNumbers(oStr As String) As Boolean
    On Error GoTo EH
    
    Dim bHasNumbers As Boolean, i As Long
    bHasNumbers = False
    For i = 1 To Len(oStr)
        Dim strTemp As String
        strTemp = Mid$(oStr, i, 1)
        If IsNumeric(strTemp) Then
            bHasNumbers = True
            Exit For
        End If
    Next
    HasNumbers = bHasNumbers

    Exit Function
EH:
    ERR_STACK.PushRaise Err.Number, Err.Source, "mdlRenewalRenamer.HasNumbers", Err.Description
End Function

'Checks if Value in string is only composed of digits.
'
'PRE : N/A
'POST: returns true if Value is only composed of digits
'      returns false if Value has characters that are not digits
Public Function subIsDigitsOnly(value As String) As Boolean
    subIsDigitsOnly = Len(value) > 0 And Not value Like "*[!0-9.]*"
End Function

Public Function subIsAlphaOnly(value As String) As Boolean
    subIsAlphaOnly = Len(value) > 0 And Not value Like "*[!A-Z]*"
End Function


' Turns off several excel options to greatly speed up processing:
'   - Turns off ScreenUpdating
'   - Sets Calculation to Manual instead of automatic,
'   - Turns off Events
'   - Turns of the display of Alerts
'
' (screen updating and calculation are the biggest resource hogs)
' Return settings to normal with subResume in a finally block
Public Sub OptimizedMode(ByVal enable As Boolean)
    With Excel.Application
        .EnableEvents = Not enable
        .Calculation = IIf(enable, xlCalculationManual, xlCalculationAutomatic)
        .ScreenUpdating = Not enable
        .EnableAnimations = Not enable
        .DisplayStatusBar = Not enable
        .PrintCommunication = Not enable
    End With
    If enable And (Not silencedFlag) Then
        silencedFlag = True
    Else
        silencedFlag = False
    End If
End Sub
