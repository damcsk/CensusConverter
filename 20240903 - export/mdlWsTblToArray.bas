Attribute VB_Name = "mdlWsTblToArray"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' mdlWsToArray
'
' This module will contain the various functions used to obtain table ranges from
' worksheets.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function BuildArrayFromTable(ws As Worksheet) As Variant
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BuildTableRange
'
' Will build a table based on the provided first column and first row, and the
' optional max column and max row.
'
' This one is only for when the column and row that has the most items is already known.
'
' It defaults to using the first row and the first column as the maxCol and maxRow
'
' :param ws:        The source worksheet
' :param startCol:  The column the table starts on
' :param startRow:  The row the table start on
' :param maxCol:    The column with the most elements
' :param maxRow:    The row with the most elements
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildTableRange(ws As Worksheet, startCol As Long, startRow As Long, _
Optional maxCol As Long = -1, Optional maxRow As Long = -1) As Range
    Dim lastRow As Long, lastCol As Long
    
    If maxCol = -1 Then maxCol = startCol
    If maxRow = -1 Then maxRow = startRow
    
    lastRow = ws.Cells(maxRow, maxCol).End(xlDown).Row
    lastCol = ws.Cells(maxRow, maxCol).End(xlToRight).Column
    Set BuildTableRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(lastRow, lastCol))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' rngGetLastCellSimple
'
' Simple version. Does exactly the same as rngGetLastCell without specifying
' a start range. It also defaults to looking in xlValue instead of xlFormula.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function rngGetLastCellSimple(ws As Worksheet, Optional startCol As Long = 1, _
Optional startRow As Long = 1) As Range
    Dim LastR As Range
    Dim LastC As Range
    
    Set LastC = ws.Cells.Find(What:="*", After:=ws.Cells(startRow, startCol), LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    Set LastR = ws.Cells.Find(What:="*", After:=ws.Cells(startRow, startCol), LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    Set rngGetLastCellSimple = Application.Intersect(LastR.EntireRow, LastC.EntireColumn)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' rngGetLastCell
'
' Don't let this fool you, it's not as complicated as it looks. It's Mostly
' validation of passed arguments.
'
' The InRange is mostly for speeding things up. The main part is only one line
' of code found in the last conditional, so if you want to use defaults
' just use rngGetLastCellSimple.
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
    
    ' Set the worksheet of the range
    Set ws = InRange.Worksheet
    
    ' Count formula cells that evaluate to empty?
    If ProhibitEmptyFormula = False Then
        LookIn = xlFormulas
    Else
        LookIn = xlValues
    End If
    
    ' What last cell to use: column, row, or both
    Select Case SearchOrder
        Case XlSearchOrder.xlByColumns, XlSearchOrder.xlByRows, _
                XlSearchOrder.xlByColumns + XlSearchOrder.xlByRows
            ' OK
        Case Else
            Err.Raise 5
            Exit Function
    End Select
    
    With ws
        ' If there's only one cell, that's the range.
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
