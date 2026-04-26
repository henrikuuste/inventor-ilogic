' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Nimeta detailid - Rename assembly occurrences with descriptive names
' 
' Renames occurrences using pattern: "<Description> (<Part Number>):<instance>"
' This makes parts easier to identify in the assembly browser.
'
' Usage: 
' - Run from an open assembly
' - If occurrences are selected, only those are renamed
' - If nothing is selected, all top-level occurrences are renamed
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/OccurrenceNamingLib.vb"

Imports System.Windows.Forms
Imports Inventor

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    
    ' Enable immediate logging
    UtilsLib.SetLogger(Logger)
    
    ' Validate document
    If app.ActiveDocument Is Nothing Then
        UtilsLib.LogError("Nimeta detailid: No active document")
        MessageBox.Show("Aktiivne dokument puudub.", "Nimeta detailid")
        Exit Sub
    End If
    
    If app.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        UtilsLib.LogError("Nimeta detailid: Active document is not an assembly")
        MessageBox.Show("Aktiivseks dokumendiks peab olema koost (.iam).", "Nimeta detailid")
        Exit Sub
    End If
    
    Dim asmDoc As AssemblyDocument = CType(app.ActiveDocument, AssemblyDocument)
    UtilsLib.LogInfo("Nimeta detailid: Starting for " & asmDoc.DisplayName)
    
    ' Check if there are selected occurrences
    Dim selectedCount As Integer = OccurrenceNamingLib.GetSelectedOccurrenceCount(asmDoc)
    Dim renamedCount As Integer = 0
    
    If selectedCount > 0 Then
        ' Rename only selected occurrences
        UtilsLib.LogInfo("Nimeta detailid: Renaming " & selectedCount & " selected occurrence(s)")
        renamedCount = OccurrenceNamingLib.RenameSelectedOccurrences(asmDoc)
    Else
        ' Rename all occurrences
        Dim totalCount As Integer = asmDoc.ComponentDefinition.Occurrences.Count
        UtilsLib.LogInfo("Nimeta detailid: Renaming all " & totalCount & " occurrence(s)")
        renamedCount = OccurrenceNamingLib.RenameAllOccurrences(asmDoc)
    End If
    
    ' Summary
    If selectedCount > 0 Then
        UtilsLib.LogInfo("Nimeta detailid: Renamed " & renamedCount & " of " & selectedCount & " selected occurrence(s)")
    Else
        UtilsLib.LogInfo("Nimeta detailid: Renamed " & renamedCount & " occurrence(s)")
    End If
    
    ' Refresh view
    app.ActiveView.Update()
End Sub
