' ============================================================================
' Sorteeri detailid - Sort assembly components by material into folders
' 
' Categorizes assembly parts into browser folders based on material name,
' then creates model states for BOM filtering.
'
' Usage: Run from an open assembly
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/SortingLib.vb"

Imports System.Collections.Generic
Imports System.Windows.Forms
Imports Inventor

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    
    ' Enable immediate logging
    UtilsLib.SetLogger(Logger)
    
    ' Validate document
    If app.ActiveDocument Is Nothing Then
        UtilsLib.LogError("Sorteeri detailid: No active document")
        MessageBox.Show("Aktiivne dokument puudub.", "Sorteeri detailid")
        Exit Sub
    End If
    
    If app.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        UtilsLib.LogError("Sorteeri detailid: Active document is not an assembly")
        MessageBox.Show("Aktiivseks dokumendiks peab olema koost (.iam).", "Sorteeri detailid")
        Exit Sub
    End If
    
    Dim asmDoc As AssemblyDocument = CType(app.ActiveDocument, AssemblyDocument)
    UtilsLib.LogInfo("Sorteeri detailid: Starting for " & asmDoc.DisplayName)
    
    ' ========================================================================
    ' CONFIGURATION: Folder patterns (material regex -> folder name)
    ' ========================================================================
    Dim folderPatterns As New Dictionary(Of String, List(Of String))
    
    ' Puit - wood materials
    folderPatterns.Add("Puit", New List(Of String)({".*vineer.*", ".*PLP.*", ".*Kask.*", ".*Okas.*"}))
    
    ' Papp - cardboard/HDF materials
    folderPatterns.Add("Papp", New List(Of String)({"HDF", "Kartong"}))
    
    ' Poroloon - foam materials
    folderPatterns.Add("Poroloon", New List(Of String)({"RG.*", "HR.*", "Dryfeel.*"}))
    
    ' Metall - metal materials
    folderPatterns.Add("Metall", New List(Of String)({".*alumiinium.*", ".*teras.*"}))
    
    ' ========================================================================
    ' CONFIGURATION: Model state / design view definitions
    ' Nothing = all folders enabled, otherwise list of included folder names
    ' ========================================================================
    Dim stateDefs As New Dictionary(Of String, List(Of String))
    
    ' Kõik - all parts visible/unsuppressed
    stateDefs.Add("Kõik", Nothing)
    
    ' Karkass papiga - frame with cardboard (wood + metal + cardboard)
    stateDefs.Add("Karkass papiga", New List(Of String)({"Puit", "Metall", "Papp"}))
    
    ' Karkass papita - frame without cardboard (wood + metal only)
    stateDefs.Add("Karkass papita", New List(Of String)({"Puit", "Metall"}))
    
    ' ========================================================================
    ' EXECUTION
    ' ========================================================================
    SortingLib.Run(asmDoc, folderPatterns, stateDefs)
    
    ' Summary
    UtilsLib.LogInfo("Sorteeri detailid: Completed")
    
    ' Refresh view
    app.ActiveView.Update()
End Sub
