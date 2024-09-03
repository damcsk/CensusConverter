VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const INT_EMP = 1
Private Const INT_SP = 2
Private Const INT_DEP = 3

Private dicPersonProperties As Scripting.Dictionary
' Each Property will be a dictionary with the following
' cell_og_row
' cell_og_col
' cell_og_color     (?)
' cell_new_row      (?)
' cell_new_col      (?)
' cell_new_color
' value_og
' value
'

' isUnique
' canDoDepGroups

'Flags and meta-info about the class
Private pCensusIndex As Integer
Private pSpouseKey As Integer
Private pEmpKey As Integer
Private pChildCountKey As Integer
Private pDepSLFlag As Boolean
Private pDepMLFlag As Boolean
Private pSpSLFlag As Boolean
Private pDepsNeeded As Boolean
Private pSpNeeded As Boolean
Private pDepGroup As Scripting.Dictionary

'The instance variables.
Private pType As String
Private pEmpIndex As Integer
Private pFirstName As String
Private pLastName As String
Private pGender As String
Private pDOBAge As Variant
Private pDOB As Date
Private pAge As Integer
Private pHomeZip As Long
Private pWorkZip As Long
Private pWorkState As String
Private pMed As String
Private pDen As String
Private pVis As String
Private pLife As String
Private pSTD As String
Private pLTD As String
Private pMedicare As String
Private pDepCount As Integer
Private pSpDOBAge As Variant
Private pSpDOB As Date
Private pSpAge As Integer
Private pSpGen As String
Private pSpMedicare As String
Private pSpTobacco As String

Private pTobacco As String
Private pEmpStatus As String
Private pSalary As Long
Private pSalaryType As String
Private pOccupation As String
Private pDOH As String
Private pWorkCity As String
Private pWorkCounty As String
Private pDepsGT21 As Integer
Private pErrSummary As Scripting.Dictionary

'The error flags for the instance variables (for marking errors when
'writing back to the census table).
Private pTypeErr As Integer


Private Sub Class_Initialize()
    Set pDepGroup = New Scripting.Dictionary
    pChildCountKey = 0
    pDepSLFlag = False
    pSpSLFlag = False
    pDepsNeeded = False
    pSpNeeded = False
End Sub

Private Function CreateProperty() As Scripting.Dictionary
    Dim dicNewProperty As Scripting.Dictionary
    Set dicNewProperty = New Scripting.Dictionary
    
    dicNewProperty.Add "og_arr_row", vbNullString
    dicNewProperty.Add "og_arr_col", vbNullString
    dicNewProperty.Add "new_arr_row", vbNullString
    dicNewProperty.Add "new_arr_col", vbNullString
    dicNewProperty.Add "cell_color", vbNullString
    dicNewProperty.Add "value_og", vbNullString
    dicNewProperty.Add "value", vbNullString
    
    Set CreateProperty = dicNewProperty
End Function

Public Function GetProperty(key As String) As Variant
    
End Function

Public Function GetPropertyValue(key As String) As Variant
    
End Function

Public Sub AddProperty()
    
End Sub

Public Sub AddPropertyFromRange(rng As Range)
    
End Sub

Public Sub SetProperty()

End Sub

Public Function isEqual() As Boolean
    
End Function

Public Sub SetPropertyAttribute(prop_key As String, attr_key As String, val As Variant)

End Sub
