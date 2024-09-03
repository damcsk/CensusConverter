Attribute VB_Name = "TEST_WStoARR"
Option Explicit


Public Sub startTest()
    Dim strBuilder As StringBuilder
    Dim timeMe As clsTimer
    Dim runTime As Variant
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim newSetting As DbDictCrossTab
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("dbColEmp")
    Set timeMe = New clsTimer
    Set strBuilder = New StringBuilder
    
    Set newSetting = New DbDictCrossTab
    newSetting.Init wsObject:=ws
    Set newSetting = Nothing
    
    Dim dict1 As Scripting.Dictionary
    Dim dict2 As Scripting.Dictionary
    Dim testDict As Scripting.Dictionary
    Dim testVar As Variant
    Dim testItem As Integer
    Set dict1 = New Scripting.Dictionary
    Set dict2 = New Scripting.Dictionary
    testItem = 5
    dict1.Add 1, dict2
    dict1.Add 2, testItem
    mdlMisc.GetDictElm dict1, 2
    
    
    Dim testRng As Range
    Dim testArr As Variant
    Dim arrRows As Long
    Dim arrCols As Long
    Dim newWb As Workbook
    Dim newWs As Worksheet
    Set newWb = Application.Workbooks.Add
    Set newWs = newWb.Worksheets(1)
    Set testRng = ws.Range("A1:G33")
    testArr = testRng.Value2
    arrRows = UBound(testArr, 1) - LBound(testArr, 1) + 1
    arrCols = UBound(testArr, 2) - LBound(testArr, 2) + 1
    
    strBuilder.Append "Results of direct copy-paste: "
    timeMe.StartTimer
    newWs.Range(newWs.Cells(1, 1), newWs.Cells(arrRows, arrCols)).Value2 = testArr
    runTime = timeMe.ElapsedTime
    strBuilder.Append runTime
    strBuilder.Append vbNewLine
    
    strBuilder.Append "Results of for-looping: "
    timeMe.StartTimer
    For i = 1 To arrRows
        For j = 1 To arrCols
            newWs.Cells(i, j).Value2 = testArr(i, j)
        Next j
    Next i
    runTime = timeMe.ElapsedTime
    strBuilder.Append runTime
    strBuilder.Append vbNewLine
    
    MsgBox strBuilder.ToString
    
End Sub

Private Function testConversion(tableRange As Range, wsSrc As Worksheet, startRow As Long) As Variant
    Dim timer As clsTimer
    Dim tempHolder As Variant
    Set timer = New clsTimer
    
    timer.StartTimer
    Set tempHolder = wsSrc.Range(wsSrc.Cells(startRow, 1), wsSrc.Cells(wsSrc.Rows.Count, wsSrc.Columns.Count).End(xlUp))
    MsgBox timer.ElapsedTime
    
    Set testConversion = tempHolder
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This subroutine takes the worksheet and puts the used range starting from the indicated cell to the last used
'cell into an array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function arrConvertRange(ws As Worksheet, col1 As Integer, _
row1 As Integer, maxRows As Integer) As Variant
    Dim rngLastCell As Range
    Dim rngInRange As Range
    
    ws.Activate
On Error GoTo RangeArrayErr
    Set rngInRange = Range(Cells(row1, col1), Cells(maxRows, maxRows))
    Set rngLastCell = rngGetLastCell(rngInRange, xlByColumns + xlByRows, True)
    
    'Set the census to the array
    arrConvertRange = Range(Cells(row1, col1), Cells(rngLastCell.Row, _
     rngLastCell.Column))
Exit Function

RangeArrayErr:

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 06/17/2024: This was added a long time ago, and I don't remember the logic
'           anymore. However, this is by far the best version of this type
'           of function I've used.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function rngGetLastCell(InRange As Range, SearchOrder As XlSearchOrder, _
 Optional ProhibitEmptyFormula As Boolean = False) As Range
    Dim ws As Worksheet
    Dim r As Range
    Dim LastCell As Range
    Dim LastR As Range
    Dim LastC As Range
    Dim SearchRange As Range
    Dim LookIn As XlFindLookIn
    Dim RR As Range
    
    Set ws = InRange.Worksheet
    
    If ProhibitEmptyFormula = False Then
        LookIn = xlFormulas
    Else
        LookIn = xlValues
    End If
    
    Select Case SearchOrder
        Case XlSearchOrder.xlByColumns, XlSearchOrder.xlByRows, _
                XlSearchOrder.xlByColumns + XlSearchOrder.xlByRows
            ' OK
        Case Else
            Err.Raise 5
            Exit Function
    End Select
    
    With ws
        If InRange.Cells.Count = 1 Then
            Set RR = .UsedRange
        Else
           Set RR = InRange
        End If
        Set r = RR(RR.Cells.Count)
        
        If SearchOrder = xlByColumns Then
            Set LastCell = RR.Find(What:="*", After:=r, LookIn:=LookIn, _
                    LookAt:=xlPart, SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, MatchCase:=False)
        ElseIf SearchOrder = xlByRows Then
            Set LastCell = RR.Find(What:="*", After:=r, LookIn:=LookIn, _
                    LookAt:=xlPart, SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, MatchCase:=False)
        ElseIf SearchOrder = xlByColumns + xlByRows Then
            Set LastC = RR.Find(What:="*", After:=r, LookIn:=LookIn, _
                    LookAt:=xlPart, SearchOrder:=xlByColumns, _
                    SearchDirection:=xlPrevious, MatchCase:=False)
            Set LastR = RR.Find(What:="*", After:=r, LookIn:=LookIn, _
                    LookAt:=xlPart, SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, MatchCase:=False)
            On Error GoTo NoContent
            Set LastCell = Application.Intersect(LastR.EntireRow, LastC.EntireColumn)
        Else
            Err.Raise 5
            Exit Function
        End If
    End With
    
    Set rngGetLastCell = LastCell

NoContent:
    Set LastCell = InRange.Cells(2, 2)
    Resume Next
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This function will determine the ranges of the databases and shove them
'into arrays.
'
'pre : meta information should be set first
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function varCreateArray(wsDB As Worksheet, dicKeys As Scripting.Dictionary, _
strLengthKey As String) As Variant
    Dim intLastRow As Long
    Dim intLastCol As Long
    Dim rngDB As Range
    Dim varDB As Variant
    Dim key As Variant

    wsDB.Activate
    
    'Find last row and last column
    intLastRow = wsDB.Cells(Rows.Count, dicKeys.Item(strLengthKey)).End(xlUp).Row
    intLastCol = 1
    For Each key In dicKeys.Keys
        If dicKeys.Item(key) > intLastCol Then
            intLastCol = dicKeys.Item(key)
        End If
    Next key
    
    'Take range and put in array
    Set rngDB = Range(Cells(2, 1), Cells(intLastRow, intLastCol))
    varCreateArray = rngDB
End Function

