VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbDictLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDict As Scripting.Dictionary
Private pWsSrc As Worksheet
Private pStartRow As Long
Private pStartCol As Long
Private pInverted As Boolean
Private pMaxAttributeCount As Long

Private pChanges As Scripting.Dictionary
Private pChangedMaxAttributeCount As Long

Public Sub DiscardChanges()
    Set pChanges = New Scripting.Dictionary
    pChangedMaxAttributeCount = pMaxAttributeCount
End Sub

' Initialize the class and dictionary
Private Sub Class_Initialize()
    Set pDict = New Scripting.Dictionary
    Set pChanges = New Scripting.Dictionary
    pStartRow = 1
    pStartCol = 1
    pInverted = False
End Sub

' Initialize the LookupTableDictionary
Public Sub Init(Optional wsName As String = "", Optional wsIndex As Long = 0, _
    Optional wsObject As Worksheet = Nothing, Optional tblRange As Range = Nothing, _
    Optional startRow As Long = 1, Optional startCol As Long = 1, _
    Optional invert As Boolean = False)
On Error GoTo EH
    pStartRow = startRow
    pStartCol = startCol
    pInverted = invert

    ' This will set the table range property of the class.
    If Not tblRange Is Nothing Then
        Set pWsSrc = tblRange.Worksheet
        pStartRow = tblRange.Row
        pStartCol = tblRange.Column
    Else
        ' A settings build must have at least a range or a worksheet to work
        ' Checks if the user supplied the worksheet object first, then tries the index,
        ' then tries the worksheet name. If none are provided, raises an error.
        If Not wsObject Is Nothing Then
            Set pWsSrc = wsObject
        ElseIf wsIndex > 0 Then
            Set pWsSrc = ThisWorkbook.Worksheets(wsIndex)
        ElseIf Len(wsName) > 0 Then
            Set pWsSrc = ThisWorkbook.Worksheets(wsName)
        Else
            Err.Raise vbObjectError + 1, "MissingLookupSheet", "A lookup table's worksheet must be specified by ref, name, or index in order to build."
        End If
        
        Set tblRange = pWsSrc.Range(pWsSrc.Cells(pStartRow, pStartCol), mdlWsTblToArray.rngGetLastCellSimple(pWsSrc))
    End If

    BuildDictionary tblRange

    Exit Sub
EH:
    ERR_STACK.PushRaise Err.Number, Err.Source, "DbDictLookup.Init", Err.Description
End Sub

' Build the dictionary from the table range
Private Sub BuildDictionary(tblRange As Range)
On Error GoTo EH
    Dim cell As Range
    Dim data As Variant
    Dim dataArr() As Variant
    Dim itemArr() As Variant
    Dim i As Long, j As Long, itemArrayCount As Long, keyCount As Long
    
    ' Can't build a db without a table.
    If tblRange Is Nothing Then
        Err.Raise vbObjectError + 513, "NoDictionaryToBuild", "A lookup table must have the Table Range set before being built."
    End If
     
    ' Assign range to array for faster processing
    dataArr = tblRange.Value2
    
    If pInverted Then
        pMaxAttributeCount = UBound(dataArr, 1) - 1
        keyCount = UBound(dataArr, 2)
    Else
        pMaxAttributeCount = UBound(dataArr, 2) - 1
        keyCount = UBound(dataArr, 1)
    End If
    
    For i = 2 To keyCount
        itemArrayCount = 0
        ReDim itemArr(1 To pMaxAttributeCount)
        
        For j = 2 To (pMaxAttributeCount + 1)
            If pInverted Then
                data = dataArr(j, i)
            Else
                data = dataArr(i, j)
            End If
            
            If Len(data) > 0 Then
                itemArr(j - 1) = data
                itemArrayCount = itemArrayCount + 1
            Else
                Exit For
            End If
        Next j
        
        ReDim Preserve itemArr(1 To itemArrayCount)
        
        If pInverted Then
            pDict(dataArr(1, i)) = itemArr
        Else
            pDict(dataArr(i, 1)) = itemArr
        End If
    Next
    Erase dataArr
EH:
    ERR_STACK.PushRaise Err.Number, Err.Source, "DbDictLookup.BuildDictionary", Err.Description
End Sub

' CRUD Operations

' Create or Update a key-value pair
Public Sub SetItem(key As Variant, values() As Variant)
    pDict(key) = values
    pChanges(key) = values
End Sub

' Read a value by key
Public Function GetItem(key As Variant) As Variant
    If pDict.Exists(key) Then
        GetItem = pDict(key)
    Else
        Err.Raise vbObjectError + 1, "LookupTableDictionary", "Key not found"
    End If
End Function

' Double check pMaxAttributeCount
Public Sub AddToItemArray(key As Variant, val As Variant)
    Dim itemArr() As Variant
    Dim changedItemArr() As Variant
    Dim i As Long
    
    If Not (pDict.Exists(key) Or pChanges.Exists(key)) Then
        Err.Raise vbObjectError + 1, "LookupTableDictionary", "Key not found"
    End If
    
    If pChanges.Exists(key) Then
        Set changedItemArr = pChanges(key)
        ReDim Preserve changedItemArr(1 To (UBound(changedItemArr) + 1))
        
        
    Else
        itemArr = pDict(key)
        ReDim changedItemArr(1 To (UBound(itemArr) + 1))
        pChanges.Add key, changedItemArr
        
        For i = 1 To UBound(itemArr)
            changedItemArr(i) = itemArr(i)
        Next i
    End If
    
    changedItemArr(UBound(changedItemArr)) = val
End Sub

' Double check pMaxAttributeCount
Public Sub RemoveFromItemArray(key As Variant)

End Sub

' Delete a key-value pair
Public Sub DeleteItem(key As Variant)
    If pDict.Exists(key) Then
        pDict.Remove key
        pChanges.Add key, vbNull
    Else
        Err.Raise vbObjectError + 1, "LookupTableDictionary", "Key not found"
    End If
End Sub

' Apply changes to the worksheet
Public Sub CommitChanges()

End Sub

' Clear dictionary and changes
Public Sub ClearDict()
    pDict.RemoveAll
    pChanges.RemoveAll
End Sub

' Get the internal dictionary (optional)
Public Function GetDictionary() As Scripting.Dictionary
    Set GetDictionary = pDict
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' WriteToWorksheet:
'
' Writes the dictionary to the worksheet. Defaults to including any
' changed data.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub WriteToWorksheet(Optional useChanged As Boolean = True)
    Dim tblRange As Range
    Dim i As Long, j As Long
    Dim key As Variant
    Dim tempArr() As Variant
    Dim tempValue As Variant
    Dim result() As Variant
    Dim rowCount As Long, colCount As Long
    
    ' Get row and col counts. TODO: make sure pChangedMaxAttributeCount is accurate
    rowCount = pDict.Count
    
    If useChanged Then
        colCount = pChangedMaxAttributeCount
    Else
        colCount = pMaxAttributeCount
    End If
    
    ' Resize result array to fit new table dimensions
    ReDim result(1 To rowCount, 1 To colCount + 1)
    
    For Each key In pDict.Keys
        ' Place the key in the first column of the resulting lookup table
        result(i, 1) = key
        
        ' For every column, input the value
        For j = 0 To UBound(dict(key))
            ' Use the updated values if they exist in the changes dict
            ElseIf useChanged And pChanges.Exists(key) Then
                tempValue = pChanges(key)(j)
            Else
                tempValue = pDict(key)(j)
            End If
            
            ' If the value is empty, reached the last column and should move to next row
            If Len(tempValue) = 0 Then
                Exit For
            Else
                result(i, j + 2) = tempValue
            End If
        Next j
        
        ' Manually increment row index
        i = i + 1
    Next key
    
    ' Get the current table range and clear it's contents
    Set tblRange = pWsSrc.Range(pWsSrc.Cells(pStartRow, pStartCol), mdlWsTblToArray.rngGetLastCellSimple(pWsSrc))
    tblRange.ClearContents
    
    ' Set the range to the results array
    pWsSrc.Range(pWsSrc.Cells(pStartRow, pStartCol)).Resize(rowCount, colCount + 1).value = result
    
    ' If using changed values, update and clear committed changes
    If useChanged Then
        pMaxAttributeCount = pChangedMaxAttributeCount
        Set pChanged = New Scripting.Dictionary
    End If
    
    ' Prepare the main dictionary to be rebuilt
    ClearDict
    
    ' Set up table range to re-create dict
    Set tblRange = pWsSrc.Range(pWsSrc.Cells(pStartRow, pStartCol), mdlWsTblToArray.rngGetLastCellSimple(pWsSrc))
    BuildDictionary tblRange
    
    Erase result
End Sub
