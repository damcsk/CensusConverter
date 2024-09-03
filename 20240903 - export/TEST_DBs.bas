Attribute VB_Name = "TEST_DBs"
Private pDict As Scripting.Dictionary


Private Sub test_Lookup()
    Dim lookupTbl As DbDictLookup
    Dim arr As Variant
    
    Set lookupTbl = New DbDictLookup
    lookupTbl.Init wsName:="dbLookupTable"
    arr = lookupTbl.GetItem("Tbl208")
    MsgBox ""
End Sub

