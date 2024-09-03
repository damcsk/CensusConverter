VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SettingsCols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Inherits from SettingsSuper

' Will have three versions
' The full db version with all possible columns as the keys - will always be built, even if the other two are built
' The settings declared version, with only the columns that are indicated by their enabled flag in the full db version
