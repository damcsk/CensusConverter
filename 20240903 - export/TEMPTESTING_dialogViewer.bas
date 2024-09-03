Attribute VB_Name = "TEMPTESTING_dialogViewer"
' Possible useful:
'
' Possible for formatting selection?
' xlDialogApplyStyle)
'
' Formatting:
' xlDialogAlignment

'
'Application.CommandBars("Worksheet Menu Bar").Controls("Format").Controls("Cells...").Execute



Sub dialogViewer()
    Dim test As Variant
    Dim temp As Boolean
    
    With Application
    
    
'        MsgBox "Showing xlDialogActivate"
'        .Dialogs(xlDialogActivate).Show
'        MsgBox "Showing xlDialogActiveCellFont"
'        .Dialogs(xlDialogActiveCellFont).Show
'        'MsgBox "Showing xlDialogAddChartAutoformat"
'        '.Dialogs(xlDialogAddChartAutoformat).Show
'        MsgBox "Showing xlDialogAddinManager"
'        .Dialogs(xlDialogAddinManager).Show
'        MsgBox "Showing xlDialogArrangeAll"
'        .Dialogs(xlDialogArrangeAll).Show

        MsgBox "Showing xlDialogBorder"
        .Dialogs(xlDialogBorder).Show
        MsgBox "Showing xlDialogCalculation"
        .Dialogs(xlDialogCalculation).Show
        MsgBox "Showing xlDialogCellProtection"
        .Dialogs(xlDialogCellProtection).Show
        MsgBox "Showing xlDialogChangeLink"
        .Dialogs(xlDialogChangeLink).Show
        MsgBox "Showing xlDialogChartAddData"
        .Dialogs(xlDialogChartAddData).Show
        MsgBox "Showing xlDialogChartLocation"
        .Dialogs(xlDialogChartLocation).Show
        MsgBox "Showing xlDialogChartOptionsDataLabelMultiple"
        .Dialogs(xlDialogChartOptionsDataLabelMultiple).Show
        MsgBox "Showing xlDialogChartOptionsDataLabels"
        .Dialogs(xlDialogChartOptionsDataLabels).Show
        MsgBox "Showing xlDialogChartOptionsDataTable"
        .Dialogs(xlDialogChartOptionsDataTable).Show
        MsgBox "Showing xlDialogChartSourceData"
        .Dialogs(xlDialogChartSourceData).Show
        MsgBox "Showing xlDialogChartTrend"
        .Dialogs(xlDialogChartTrend).Show
        MsgBox "Showing xlDialogChartType"
        .Dialogs(xlDialogChartType).Show
        MsgBox "Showing xlDialogChartWizard"
        .Dialogs(xlDialogChartWizard).Show
        MsgBox "Showing xlDialogCheckboxProperties"
        .Dialogs(xlDialogCheckboxProperties).Show
        MsgBox "Showing xlDialogClear"
        .Dialogs(xlDialogClear).Show
        MsgBox "Showing xlDialogColorPalette"
        .Dialogs(xlDialogColorPalette).Show
        MsgBox "Showing xlDialogColumnWidth"
        .Dialogs(xlDialogColumnWidth).Show
        MsgBox "Showing xlDialogCombination"
        .Dialogs(xlDialogCombination).Show
        MsgBox "Showing xlDialogConditionalFormatting"
        .Dialogs(xlDialogConditionalFormatting).Show
        MsgBox "Showing xlDialogConsolidate"
        .Dialogs(xlDialogConsolidate).Show
        MsgBox "Showing xlDialogCopyChart"
        .Dialogs(xlDialogCopyChart).Show
        MsgBox "Showing xlDialogCopyPicture"
        .Dialogs(xlDialogCopyPicture).Show
        MsgBox "Showing xlDialogCreateList"
        .Dialogs(xlDialogCreateList).Show
        MsgBox "Showing xlDialogCreateNames"
        .Dialogs(xlDialogCreateNames).Show
        MsgBox "Showing xlDialogCreatePublisher"
        .Dialogs(xlDialogCreatePublisher).Show
        MsgBox "Showing xlDialogCreateRelationship"
        .Dialogs(xlDialogCreateRelationship).Show
        MsgBox "Showing xlDialogCustomizeToolbar"
        .Dialogs(xlDialogCustomizeToolbar).Show
        MsgBox "Showing xlDialogCustomViews"
        .Dialogs(xlDialogCustomViews).Show
        MsgBox "Showing xlDialogDataDelete"
        .Dialogs(xlDialogDataDelete).Show
        MsgBox "Showing xlDialogDataLabel"
        .Dialogs(xlDialogDataLabel).Show
        MsgBox "Showing xlDialogDataLabelMultiple"
        .Dialogs(xlDialogDataLabelMultiple).Show
        MsgBox "Showing xlDialogDataSeries"
        .Dialogs(xlDialogDataSeries).Show
        MsgBox "Showing xlDialogDataValidation"
        .Dialogs(xlDialogDataValidation).Show
        MsgBox "Showing xlDialogDefineName"
        .Dialogs(xlDialogDefineName).Show
        MsgBox "Showing xlDialogDefineStyle"
        .Dialogs(xlDialogDefineStyle).Show
        MsgBox "Showing xlDialogDeleteFormat"
        .Dialogs(xlDialogDeleteFormat).Show
        MsgBox "Showing xlDialogDeleteName"
        .Dialogs(xlDialogDeleteName).Show
        MsgBox "Showing xlDialogDemote"
        .Dialogs(xlDialogDemote).Show
        MsgBox "Showing xlDialogDisplay"
        .Dialogs(xlDialogDisplay).Show
        MsgBox "Showing xlDialogDocumentInspector"
        .Dialogs(xlDialogDocumentInspector).Show
        MsgBox "Showing xlDialogEditboxProperties"
        .Dialogs(xlDialogEditboxProperties).Show
        MsgBox "Showing xlDialogEditColor"
        .Dialogs(xlDialogEditColor).Show
        MsgBox "Showing xlDialogEditDelete"
        .Dialogs(xlDialogEditDelete).Show
        MsgBox "Showing xlDialogEditionOptions"
        .Dialogs(xlDialogEditionOptions).Show
        MsgBox "Showing xlDialogEditSeries"
        .Dialogs(xlDialogEditSeries).Show
        MsgBox "Showing xlDialogErrorbarX"
        .Dialogs(xlDialogErrorbarX).Show
        MsgBox "Showing xlDialogErrorbarY"
        .Dialogs(xlDialogErrorbarY).Show
        MsgBox "Showing xlDialogErrorChecking"
        .Dialogs(xlDialogErrorChecking).Show
        MsgBox "Showing xlDialogEvaluateFormula"
        .Dialogs(xlDialogEvaluateFormula).Show
        MsgBox "Showing xlDialogExternalDataProperties"
        .Dialogs(xlDialogExternalDataProperties).Show
        MsgBox "Showing xlDialogExtract"
        .Dialogs(xlDialogExtract).Show
        MsgBox "Showing xlDialogFileDelete"
        .Dialogs(xlDialogFileDelete).Show
        MsgBox "Showing xlDialogFileSharing"
        .Dialogs(xlDialogFileSharing).Show
        MsgBox "Showing xlDialogFillGroup"
        .Dialogs(xlDialogFillGroup).Show
        MsgBox "Showing xlDialogFillWorkgroup"
        .Dialogs(xlDialogFillWorkgroup).Show
        MsgBox "Showing xlDialogFilter"
        .Dialogs(xlDialogFilter).Show
        MsgBox "Showing xlDialogFilterAdvanced"
        .Dialogs(xlDialogFilterAdvanced).Show
        MsgBox "Showing xlDialogFindFile"
        .Dialogs(xlDialogFindFile).Show
        MsgBox "Showing xlDialogFont"
        .Dialogs(xlDialogFont).Show
        MsgBox "Showing xlDialogFontProperties"
        .Dialogs(xlDialogFontProperties).Show
        MsgBox "Showing xlDialogFormatAuto"
        .Dialogs(xlDialogFormatAuto).Show
        MsgBox "Showing xlDialogFormatChart"
        .Dialogs(xlDialogFormatChart).Show
        MsgBox "Showing xlDialogFormatCharttype"
        .Dialogs(xlDialogFormatCharttype).Show
        MsgBox "Showing xlDialogFormatFont"
        .Dialogs(xlDialogFormatFont).Show
        MsgBox "Showing xlDialogFormatLegend"
        .Dialogs(xlDialogFormatLegend).Show
        MsgBox "Showing xlDialogFormatMain"
        .Dialogs(xlDialogFormatMain).Show
        MsgBox "Showing xlDialogFormatMove"
        .Dialogs(xlDialogFormatMove).Show
        MsgBox "Showing xlDialogFormatNumber"
        .Dialogs(xlDialogFormatNumber).Show
        MsgBox "Showing xlDialogFormatOverlay"
        .Dialogs(xlDialogFormatOverlay).Show
        MsgBox "Showing xlDialogFormatSize"
        .Dialogs(xlDialogFormatSize).Show
        MsgBox "Showing xlDialogFormatText"
        .Dialogs(xlDialogFormatText).Show
        MsgBox "Showing xlDialogFormulaFind"
        .Dialogs(xlDialogFormulaFind).Show
        MsgBox "Showing xlDialogFormulaGoto"
        .Dialogs(xlDialogFormulaGoto).Show
        MsgBox "Showing xlDialogFormulaReplace"
        .Dialogs(xlDialogFormulaReplace).Show
        MsgBox "Showing xlDialogFunctionWizard"
        .Dialogs(xlDialogFunctionWizard).Show
        MsgBox "Showing xlDialogGallery3dArea"
        .Dialogs(xlDialogGallery3dArea).Show
        MsgBox "Showing xlDialogGallery3dBar"
        .Dialogs(xlDialogGallery3dBar).Show
        MsgBox "Showing xlDialogGallery3dColumn"
        .Dialogs(xlDialogGallery3dColumn).Show
        MsgBox "Showing xlDialogGallery3dLine"
        .Dialogs(xlDialogGallery3dLine).Show
        MsgBox "Showing xlDialogGallery3dPie"
        .Dialogs(xlDialogGallery3dPie).Show
        MsgBox "Showing xlDialogGallery3dSurface"
        .Dialogs(xlDialogGallery3dSurface).Show
        MsgBox "Showing xlDialogGalleryArea"
        .Dialogs(xlDialogGalleryArea).Show
        MsgBox "Showing xlDialogGalleryBar"
        .Dialogs(xlDialogGalleryBar).Show
        MsgBox "Showing xlDialogGalleryColumn"
        .Dialogs(xlDialogGalleryColumn).Show
        MsgBox "Showing xlDialogGalleryCustom"
        .Dialogs(xlDialogGalleryCustom).Show
        MsgBox "Showing xlDialogGalleryDoughnut"
        .Dialogs(xlDialogGalleryDoughnut).Show
        MsgBox "Showing xlDialogGalleryLine"
        .Dialogs(xlDialogGalleryLine).Show
        MsgBox "Showing xlDialogGalleryPie"
        .Dialogs(xlDialogGalleryPie).Show
        MsgBox "Showing xlDialogGalleryRadar"
        .Dialogs(xlDialogGalleryRadar).Show
        MsgBox "Showing xlDialogGalleryScatter"
        .Dialogs(xlDialogGalleryScatter).Show
        MsgBox "Showing xlDialogGoalSeek"
        .Dialogs(xlDialogGoalSeek).Show
        MsgBox "Showing xlDialogGridlines"
        .Dialogs(xlDialogGridlines).Show
        MsgBox "Showing xlDialogImportTextFile"
        .Dialogs(xlDialogImportTextFile).Show
        MsgBox "Showing xlDialogInsert"
        .Dialogs(xlDialogInsert).Show
        MsgBox "Showing xlDialogInsertHyperlink"
        .Dialogs(xlDialogInsertHyperlink).Show
        MsgBox "Showing xlDialogInsertObject"
        .Dialogs(xlDialogInsertObject).Show
        MsgBox "Showing xlDialogInsertPicture"
        .Dialogs(xlDialogInsertPicture).Show
        MsgBox "Showing xlDialogInsertTitle"
        .Dialogs(xlDialogInsertTitle).Show
        MsgBox "Showing xlDialogLabelProperties"
        .Dialogs(xlDialogLabelProperties).Show
        MsgBox "Showing xlDialogListboxProperties"
        .Dialogs(xlDialogListboxProperties).Show
        MsgBox "Showing xlDialogMacroOptions"
        .Dialogs(xlDialogMacroOptions).Show
        MsgBox "Showing xlDialogMailEditMailer"
        .Dialogs(xlDialogMailEditMailer).Show
        MsgBox "Showing xlDialogMailLogon"
        .Dialogs(xlDialogMailLogon).Show
        MsgBox "Showing xlDialogMailNextLetter"
        .Dialogs(xlDialogMailNextLetter).Show
        MsgBox "Showing xlDialogMainChart"
        .Dialogs(xlDialogMainChart).Show
        MsgBox "Showing xlDialogMainChartType"
        .Dialogs(xlDialogMainChartType).Show
        MsgBox "Showing xlDialogManageRelationships"
        .Dialogs(xlDialogManageRelationships).Show
        MsgBox "Showing xlDialogMenuEditor"
        .Dialogs(xlDialogMenuEditor).Show
        MsgBox "Showing xlDialogMove"
        .Dialogs(xlDialogMove).Show
        MsgBox "Showing xlDialogMyPermission"
        .Dialogs(xlDialogMyPermission).Show
        MsgBox "Showing xlDialogNameManager"
        .Dialogs(xlDialogNameManager).Show
        MsgBox "Showing xlDialogNew"
        .Dialogs(xlDialogNew).Show
        MsgBox "Showing xlDialogNewName"
        .Dialogs(xlDialogNewName).Show
        MsgBox "Showing xlDialogNewWebQuery"
        .Dialogs(xlDialogNewWebQuery).Show
        MsgBox "Showing xlDialogNote"
        .Dialogs(xlDialogNote).Show
        MsgBox "Showing xlDialogObjectProperties"
        .Dialogs(xlDialogObjectProperties).Show
        MsgBox "Showing xlDialogObjectProtection"
        .Dialogs(xlDialogObjectProtection).Show
        MsgBox "Showing xlDialogOpen"
        .Dialogs(xlDialogOpen).Show
        MsgBox "Showing xlDialogOpenLinks"
        .Dialogs(xlDialogOpenLinks).Show
        MsgBox "Showing xlDialogOpenMail"
        .Dialogs(xlDialogOpenMail).Show
        MsgBox "Showing xlDialogOpenText"
        .Dialogs(xlDialogOpenText).Show
        MsgBox "Showing xlDialogOptionsCalculation"
        .Dialogs(xlDialogOptionsCalculation).Show
        MsgBox "Showing xlDialogOptionsChart"
        .Dialogs(xlDialogOptionsChart).Show
        MsgBox "Showing xlDialogOptionsEdit"
        .Dialogs(xlDialogOptionsEdit).Show
        MsgBox "Showing xlDialogOptionsGeneral"
        .Dialogs(xlDialogOptionsGeneral).Show
        MsgBox "Showing xlDialogOptionsListsAdd"
        .Dialogs(xlDialogOptionsListsAdd).Show
        MsgBox "Showing xlDialogOptionsME"
        .Dialogs(xlDialogOptionsME).Show
        MsgBox "Showing xlDialogOptionsTransition"
        .Dialogs(xlDialogOptionsTransition).Show
        MsgBox "Showing xlDialogOptionsView"
        .Dialogs(xlDialogOptionsView).Show
        MsgBox "Showing xlDialogOutline"
        .Dialogs(xlDialogOutline).Show
        MsgBox "Showing xlDialogOverlay"
        .Dialogs(xlDialogOverlay).Show
        MsgBox "Showing xlDialogOverlayChartType"
        .Dialogs(xlDialogOverlayChartType).Show
        MsgBox "Showing xlDialogPageSetup"
        .Dialogs(xlDialogPageSetup).Show
        MsgBox "Showing xlDialogParse"
        .Dialogs(xlDialogParse).Show
        MsgBox "Showing xlDialogPasteNames"
        .Dialogs(xlDialogPasteNames).Show
        MsgBox "Showing xlDialogPasteSpecial"
        .Dialogs(xlDialogPasteSpecial).Show
        MsgBox "Showing xlDialogPatterns"
        .Dialogs(xlDialogPatterns).Show
        MsgBox "Showing xlDialogPermission"
        .Dialogs(xlDialogPermission).Show
        MsgBox "Showing xlDialogPhonetic"
        .Dialogs(xlDialogPhonetic).Show
        MsgBox "Showing xlDialogPivotCalculatedField"
        .Dialogs(xlDialogPivotCalculatedField).Show
        MsgBox "Showing xlDialogPivotCalculatedItem"
        .Dialogs(xlDialogPivotCalculatedItem).Show
        MsgBox "Showing xlDialogPivotClientServerSet"
        .Dialogs(xlDialogPivotClientServerSet).Show
        MsgBox "Showing xlDialogPivotFieldGroup"
        .Dialogs(xlDialogPivotFieldGroup).Show
        MsgBox "Showing xlDialogPivotFieldProperties"
        .Dialogs(xlDialogPivotFieldProperties).Show
        MsgBox "Showing xlDialogPivotFieldUngroup"
        .Dialogs(xlDialogPivotFieldUngroup).Show
        MsgBox "Showing xlDialogPivotShowPages"
        .Dialogs(xlDialogPivotShowPages).Show
        MsgBox "Showing xlDialogPivotSolveOrder"
        .Dialogs(xlDialogPivotSolveOrder).Show
        MsgBox "Showing xlDialogPivotTableOptions"
        .Dialogs(xlDialogPivotTableOptions).Show
        MsgBox "Showing xlDialogPivotTableSlicerConnections"
        .Dialogs(xlDialogPivotTableSlicerConnections).Show
        MsgBox "Showing xlDialogPivotTableWhatIfAnalysisSettings"
        .Dialogs(xlDialogPivotTableWhatIfAnalysisSettings).Show
        MsgBox "Showing xlDialogPivotTableWizard"
        .Dialogs(xlDialogPivotTableWizard).Show
        MsgBox "Showing xlDialogPlacement"
        .Dialogs(xlDialogPlacement).Show
        MsgBox "Showing xlDialogPrint"
        .Dialogs(xlDialogPrint).Show
        MsgBox "Showing xlDialogPrinterSetup"
        .Dialogs(xlDialogPrinterSetup).Show
        MsgBox "Showing xlDialogPrintPreview"
        .Dialogs(xlDialogPrintPreview).Show
        MsgBox "Showing xlDialogPromote"
        .Dialogs(xlDialogPromote).Show
        MsgBox "Showing xlDialogProperties"
        .Dialogs(xlDialogProperties).Show
        MsgBox "Showing xlDialogPropertyFields"
        .Dialogs(xlDialogPropertyFields).Show
        MsgBox "Showing xlDialogProtectDocument"
        .Dialogs(xlDialogProtectDocument).Show
        MsgBox "Showing xlDialogProtectSharing"
        .Dialogs(xlDialogProtectSharing).Show
        MsgBox "Showing xlDialogPublishAsWebPage"
        .Dialogs(xlDialogPublishAsWebPage).Show
        MsgBox "Showing xlDialogPushbuttonProperties"
        .Dialogs(xlDialogPushbuttonProperties).Show
        MsgBox "Showing xlDialogRecommendedPivotTables"
        .Dialogs(xlDialogRecommendedPivotTables).Show
        MsgBox "Showing xlDialogReplaceFont"
        .Dialogs(xlDialogReplaceFont).Show
        MsgBox "Showing xlDialogRoutingSlip"
        .Dialogs(xlDialogRoutingSlip).Show
        MsgBox "Showing xlDialogRowHeight"
        .Dialogs(xlDialogRowHeight).Show
        MsgBox "Showing xlDialogRun"
        .Dialogs(xlDialogRun).Show
        MsgBox "Showing xlDialogSaveAs"
        .Dialogs(xlDialogSaveAs).Show
        MsgBox "Showing xlDialogSaveCopyAs"
        .Dialogs(xlDialogSaveCopyAs).Show
        MsgBox "Showing xlDialogSaveNewObject"
        .Dialogs(xlDialogSaveNewObject).Show
        MsgBox "Showing xlDialogSaveWorkbook"
        .Dialogs(xlDialogSaveWorkbook).Show
        MsgBox "Showing xlDialogSaveWorkspace"
        .Dialogs(xlDialogSaveWorkspace).Show
        MsgBox "Showing xlDialogScale"
        .Dialogs(xlDialogScale).Show
        MsgBox "Showing xlDialogScenarioAdd"
        .Dialogs(xlDialogScenarioAdd).Show
        MsgBox "Showing xlDialogScenarioCells"
        .Dialogs(xlDialogScenarioCells).Show
        MsgBox "Showing xlDialogScenarioEdit"
        .Dialogs(xlDialogScenarioEdit).Show
        MsgBox "Showing xlDialogScenarioMerge"
        .Dialogs(xlDialogScenarioMerge).Show
        MsgBox "Showing xlDialogScenarioSummary"
        .Dialogs(xlDialogScenarioSummary).Show
        MsgBox "Showing xlDialogScrollbarProperties"
        .Dialogs(xlDialogScrollbarProperties).Show
        MsgBox "Showing xlDialogSearch"
        .Dialogs(xlDialogSearch).Show
        MsgBox "Showing xlDialogSelectSpecial"
        .Dialogs(xlDialogSelectSpecial).Show
        MsgBox "Showing xlDialogSendMail"
        .Dialogs(xlDialogSendMail).Show
        MsgBox "Showing xlDialogSeriesAxes"
        .Dialogs(xlDialogSeriesAxes).Show
        MsgBox "Showing xlDialogSeriesOptions"
        .Dialogs(xlDialogSeriesOptions).Show
        MsgBox "Showing xlDialogSeriesOrder"
        .Dialogs(xlDialogSeriesOrder).Show
        MsgBox "Showing xlDialogSeriesShape"
        .Dialogs(xlDialogSeriesShape).Show
        MsgBox "Showing xlDialogSeriesX"
        .Dialogs(xlDialogSeriesX).Show
        MsgBox "Showing xlDialogSeriesY"
        .Dialogs(xlDialogSeriesY).Show
        MsgBox "Showing xlDialogSetBackgroundPicture"
        .Dialogs(xlDialogSetBackgroundPicture).Show
        MsgBox "Showing xlDialogSetManager"
        .Dialogs(xlDialogSetManager).Show
        MsgBox "Showing xlDialogSetMDXEditor"
        .Dialogs(xlDialogSetMDXEditor).Show
        MsgBox "Showing xlDialogSetPrintTitles"
        .Dialogs(xlDialogSetPrintTitles).Show
        MsgBox "Showing xlDialogSetTupleEditorOnColumns"
        .Dialogs(xlDialogSetTupleEditorOnColumns).Show
        MsgBox "Showing xlDialogSetTupleEditorOnRows"
        .Dialogs(xlDialogSetTupleEditorOnRows).Show
        MsgBox "Showing xlDialogSetUpdateStatus"
        .Dialogs(xlDialogSetUpdateStatus).Show
        MsgBox "Showing xlDialogShowDetail"
        .Dialogs(xlDialogShowDetail).Show
        MsgBox "Showing xlDialogShowToolbar"
        .Dialogs(xlDialogShowToolbar).Show
        MsgBox "Showing xlDialogSize"
        .Dialogs(xlDialogSize).Show
        MsgBox "Showing xlDialogSlicerCreation"
        .Dialogs(xlDialogSlicerCreation).Show
        MsgBox "Showing xlDialogSlicerPivotTableConnections"
        .Dialogs(xlDialogSlicerPivotTableConnections).Show
        MsgBox "Showing xlDialogSlicerSettings"
        .Dialogs(xlDialogSlicerSettings).Show
        MsgBox "Showing xlDialogSort"
        .Dialogs(xlDialogSort).Show
        MsgBox "Showing xlDialogSortSpecial"
        .Dialogs(xlDialogSortSpecial).Show
        MsgBox "Showing xlDialogSparklineInsertColumn"
        .Dialogs(xlDialogSparklineInsertColumn).Show
        MsgBox "Showing xlDialogSparklineInsertLine"
        .Dialogs(xlDialogSparklineInsertLine).Show
        MsgBox "Showing xlDialogSparklineInsertWinLoss"
        .Dialogs(xlDialogSparklineInsertWinLoss).Show
        MsgBox "Showing xlDialogSplit"
        .Dialogs(xlDialogSplit).Show
        MsgBox "Showing xlDialogStandardFont"
        .Dialogs(xlDialogStandardFont).Show
        MsgBox "Showing xlDialogStandardWidth"
        .Dialogs(xlDialogStandardWidth).Show
        MsgBox "Showing xlDialogStyle"
        .Dialogs(xlDialogStyle).Show
        MsgBox "Showing xlDialogSubscribeTo"
        .Dialogs(xlDialogSubscribeTo).Show
        MsgBox "Showing xlDialogSubtotalCreate"
        .Dialogs(xlDialogSubtotalCreate).Show
        MsgBox "Showing xlDialogSummaryInfo"
        .Dialogs(xlDialogSummaryInfo).Show
        MsgBox "Showing xlDialogTable"
        .Dialogs(xlDialogTable).Show
        MsgBox "Showing xlDialogTabOrder"
        .Dialogs(xlDialogTabOrder).Show
        MsgBox "Showing xlDialogTextToColumns"
        .Dialogs(xlDialogTextToColumns).Show
        MsgBox "Showing xlDialogUnhide"
        .Dialogs(xlDialogUnhide).Show
        MsgBox "Showing xlDialogUpdateLink"
        .Dialogs(xlDialogUpdateLink).Show
        MsgBox "Showing xlDialogVbaInsertFile"
        .Dialogs(xlDialogVbaInsertFile).Show
        MsgBox "Showing xlDialogVbaMakeAddin"
        .Dialogs(xlDialogVbaMakeAddin).Show
        MsgBox "Showing xlDialogVbaProcedureDefinition"
        .Dialogs(xlDialogVbaProcedureDefinition).Show
        MsgBox "Showing xlDialogView3d"
        .Dialogs(xlDialogView3d).Show
        MsgBox "Showing xlDialogWebOptionsBrowsers"
        .Dialogs(xlDialogWebOptionsBrowsers).Show
        MsgBox "Showing xlDialogWebOptionsEncoding"
        .Dialogs(xlDialogWebOptionsEncoding).Show
        MsgBox "Showing xlDialogWebOptionsFiles"
        .Dialogs(xlDialogWebOptionsFiles).Show
        MsgBox "Showing xlDialogWebOptionsFonts"
        .Dialogs(xlDialogWebOptionsFonts).Show
        MsgBox "Showing xlDialogWebOptionsGeneral"
        .Dialogs(xlDialogWebOptionsGeneral).Show
        MsgBox "Showing xlDialogWebOptionsPictures"
        .Dialogs(xlDialogWebOptionsPictures).Show
        MsgBox "Showing xlDialogWindowMove"
        .Dialogs(xlDialogWindowMove).Show
        MsgBox "Showing xlDialogWindowSize"
        .Dialogs(xlDialogWindowSize).Show
        MsgBox "Showing xlDialogWorkbookAdd"
        .Dialogs(xlDialogWorkbookAdd).Show
        MsgBox "Showing xlDialogWorkbookCopy"
        .Dialogs(xlDialogWorkbookCopy).Show
        MsgBox "Showing xlDialogWorkbookInsert"
        .Dialogs(xlDialogWorkbookInsert).Show
        MsgBox "Showing xlDialogWorkbookMove"
        .Dialogs(xlDialogWorkbookMove).Show
        MsgBox "Showing xlDialogWorkbookName"
        .Dialogs(xlDialogWorkbookName).Show
        MsgBox "Showing xlDialogWorkbookNew"
        .Dialogs(xlDialogWorkbookNew).Show
        MsgBox "Showing xlDialogWorkbookOptions"
        .Dialogs(xlDialogWorkbookOptions).Show
        MsgBox "Showing xlDialogWorkbookProtect"
        .Dialogs(xlDialogWorkbookProtect).Show
        MsgBox "Showing xlDialogWorkbookTabSplit"
        .Dialogs(xlDialogWorkbookTabSplit).Show
        MsgBox "Showing xlDialogWorkbookUnhide"
        .Dialogs(xlDialogWorkbookUnhide).Show
        MsgBox "Showing xlDialogWorkgroup"
        .Dialogs(xlDialogWorkgroup).Show
        MsgBox "Showing xlDialogWorkspace"
        .Dialogs(xlDialogWorkspace).Show
        MsgBox "Showing xlDialogZoom"
        .Dialogs(xlDialogZoom).Show

    End With
End Sub
