VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbDictCrossTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' DbDictCrossTab:
'
' A dictionary-based database built from a 2 axis table. Outer dictionary keys are
' in the first row, inner dictionary keys are in the first column.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The main data structure for the class. Takes the form of dictionaries in a dictionary.
' Private or meta values are indicated by "_" at the start of the item's key.
Private pDbDict As Scripting.Dictionary

' The header keys by column dictionary allows you to find the header key based on the
' column index it's on in the worksheet
Private pHeaderKeysByColDict As Scripting.Dictionary
' The rows by attribute key dictionary allows you to find the row index based on the
' attribute's dictionary key. This allows the user to write back to the table
Private pRowsByAttributeKeyDict As Scripting.Dictionary
' The changed values dictionary will track changes between the current db's dictionary
' and the underlying spreadsheet table.
Private pChangedValuesDict As Scripting.Dictionary

' The DB src.
Private pWsSrc As Worksheet
Private pStartRow As Long
Private pStartCol As Long
Private pTableRange As Range
Private pInverted As Boolean

Private Sub Class_Initialize()
    Set pDbDict = New Scripting.Dictionary
    Set pChangedValuesDict = New Scripting.Dictionary
    pStartRow = 1
    pStartCol = 1
    pInverted = False
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Init:
'
' - Will build a settings dictionary based on the provided arguments.
' - Must have at least a worksheet or a range to pull from.
' - Worksheet can be indicated by name, index, or by ref
' - Range can be indicated by ws/startrow/startcol or by ref
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(Optional wsName As String = "", Optional wsIndex As Long = 0, _
    Optional wsObject As Worksheet = Nothing, Optional tblRange As Range = Nothing, _
    Optional startRow As Long = 1, Optional startCol As Long = 1, _
    Optional invert As Boolean = False)
On Error GoTo EH
    pInverted = invert
    pStartRow = startRow
    pStartCol = startCol
    
    ' This will set the table range property of the class.
    If Not tblRange Is Nothing Then
        Set pTableRange = tblRange
        Set pWsSrc = pTableRange.Worksheet
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
            Err.Raise vbObjectError + 513, "MissingCrossTabSheet", "A crosstab table's worksheet must be specified by ref, name, or index in order to build."
        End If
        
        ' Set pTableRange = BuildTableRange()
        ' TODO: Test to see if one of the mdlWsTblToArray methods is better
        Set pTableRange = pWsSrc.Range(pWsSrc.Cells(pStartRow, pStartCol), _
            pWsSrc.Cells(pWsSrc.Rows.Count, pWsSrc.Columns.Count).End(xlUp))
    End If
    
    BuildDictionary
    
    Exit Sub
EH:
    ERR_STACK.PushRaise Err.Number, Err.Source, "DbDictCrossTab.Build", Err.Description
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BuildDictionary:
'
' - Builds a dictionary of dictionaries based on a crosstab table.
' - Default orientation of the table is items are on the x-axis (columns),
'   while the item properties are on the y-axis (rows).
' - Invert property swaps the item and property axises.
'
' :pre: pTableRange must be set
' :post: pHeaderKeysByColDict and pDbDict are built.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildDictionary()
On Error GoTo EH
    ' Temp Index/Key/Count trackers
    Dim itemKey As String
    Dim itemIndex As Long
    Dim itemCount As Integer
    Dim attributeKey As Variant
    Dim attributeIndex As Long
    Dim attributeCount As Integer
    Dim attributeValue As Variant
    
    ' Temp objects for building the dict
    Dim tempArr As Variant
    Dim itemDict As Scripting.Dictionary
    Dim attributeKeysByRowDict As Scripting.Dictionary
    
    ' Initialize class dict
    Set pHeaderKeysByColDict = New Scripting.Dictionary
    Set pDbDict = New Scripting.Dictionary
    Set pRowsByAttributeKeyDict = New Scripting.Dictionary
    
    ' Can't build a db with no table.
    If pTableRange Is Nothing Then
        Err.Raise vbObjectError + 513, "NoDictionaryToBuild", "A crosstab table must have the Table Range set before being built."
    End If
    
    ' Shove the ws table into an array, which is orders of magnitude faster to work with.
    ' TODO: Again, test to make sure there isn't a max range on this.
    tempArr = pTableRange.Value2
    
    ' Sets the row and column counts.
    If pInverted Then
        attributeCount = UBound(tempArr, 2)
        itemCount = UBound(tempArr, 1)
    Else
        attributeCount = UBound(tempArr, 1)
        itemCount = UBound(tempArr, 2)
    End If
    
    ' Table cache prep.
    For attributeIndex = 2 To attributeCount
        ' The attribute key is the x-axis if pInverted (columns)
        If pInverted Then
            attributeKey = tempArr(1, attributeIndex)
        Else
            attributeKey = tempArr(attributeIndex, 1)
        End If
        
        ' Err on dupe keys, ignore blank keys completely.
        If itemDict.Exists(attributeKey) Then
            'If properly managed, the db should not have duplicate attribute keys
            Err.Raise vbObjectError + 513, "DuplicateAttributeKey", "The attribute key axis has a duplicate value at index " + attributeIndex + "."
        ElseIf Len(attributeKey) > 0 Then
            ' See class variable declarations for purpose
            pRowsByAttributeKeyDict.Add attributeKey, attributeIndex
            ' Used to get the attribute row in O(1) time complexity
            attributeKeysByRowDict.Add attributeIndex, attributeKey
        End If
    Next attributeIndex
    
    ' Dictionary build - Outer loop is the item loop
    For itemIndex = 2 To itemCount
        ' If pInverted the item keys are on the y-axis (rows)
        If pInverted Then
            itemKey = tempArr(itemIndex, 1)
        Else
            itemKey = tempArr(1, itemIndex)
        End If
        
        If pDbDict.Exists(itemKey) Then
            ' If properly managed, the db should not have duplicate item keys
           Err.Raise vbObjectError + 513, "DuplicateItemKey", "The item key axis has a duplicate value at index " + itemIndex + "."
        ElseIf Len(itemKey) > 0 Then
            ' Initialize item dictionary for item attributes
            Set itemDict = New Scripting.Dictionary
            
            ' The item's worksheet index is preserved here. Worksheet index also takes into account start row/col
            If pInverted Then
                itemDict.Add "_metaItemIndex", itemIndex + pStartRow
            Else
                itemDict.Add "_metaItemIndex", itemIndex + pStartCol
            End If
            pHeaderKeysByColDict.Add itemIndex, itemKey
            
            ' Inner Loop is the item's attribute loop
            For Each attributeKey In pRowsByAttributeKeyDict.Keys()
                ' Attribute key dupes have already been filtered out.
                attributeIndex = pRowsByAttributeKeyDict(attributeKey)
                
                ' Attribute value cell depends on inversion
                If pInverted Then
                    attributeValue = tempArr(itemIndex, attributeIndex)
                Else
                    attributeValue = tempArr(attributeIndex, itemIndex)
                End If
                
                ' Set attribute for item.
                itemDict.Add attributeKey, attributeValue
            Next attributeKey
            
            ' Add item to items dict
            pDbDict.Add itemKey, itemDict
        End If
    Next itemIndex
    
    Exit Sub
EH:
    ERR_STACK.PushRaise Err.Number, Err.Source, "DbDictCrossTab.BuildDictionary", Err.Description
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetDbDict:
'
' Returns the underlying dictionary of the DB
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDbDict() As Scripting.Dictionary
    Set GetDbDict = pDbDict
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetDataField:
'
' Returns the entire data field dictionary, along with it's attributes.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetItem(colKey As Variant) As Scripting.Dictionary
    Set GetItem = pDbDict(colKey)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetDataFieldAttribute:
'
' Returns the value of the data field attribute, based on the provided keys.
' Use this instead of just using the main dictionary. The main dictionary has
' key fields which need to return an array from the overflow table instead
' of a value.
'
' TODO: returns the array from the lookup db if the attributeKey has "_Key"
'       as the last four letters of the key
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetItemAttribute(colKey As Variant, attributeKey As Variant) As Variant
    Dim tempDict As Scripting.Dictionary
    Dim valHolder As Variant
    Set tempDict = GetDataField(colKey)
    Select Case VarType(tempDict)
        Case vbObject
            Set valHolder = tempDict(attributeKey)
        Case Else
            valHolder = GetItem(colKey).Item(attributeKey)
    End Select
    
    If LCase$(Right$(CStr(attributeKey), 4)) = "_Key" Then
        ' TODO get lookup array
    End If
    
    Select Case VarType(valHolder)
        Case vbObject
            Set GetItemAttribute = valHolder
        Case Else
            GetItemAttribute = valHolder
    End Select
End Function

Public Function GetDataField(colKey As Variant)
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Build:
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function UpdateAndSaveDataFieldAttribute() As Boolean

End Function

Public Function UpdateDataFieldAttribute() As Boolean

End Function

Public Sub SaveChanges()

End Sub

Public Sub AddItem()

End Sub

Public Sub AddAttribute()

End Sub

Public Function Remove(key As Variant) As Boolean

End Function

Public Function Edit(key As Variant, value As Variant) As Boolean
    
End Function
