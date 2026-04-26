' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' Ekspordi BOM - Export Bill of Materials to an Excel template
'
' Uses the active model state, prompts for a template and save path.
' Template: last row with 2+ cells containing "{{" is the per-part mapping
' row; rows above the header line get assembly iProperties. See Templates/BOM_Export_README.md
'
' Requires: AddReference "System.Data"  (DataTable for {{=...}} math)
' ============================================================================

AddReference "System.Data"
AddVbFile "Lib/RuntimeLib.vb"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/BOMExportLib.vb"

Sub Main()
    If Not AppRuntime.Initialize(ThisApplication) Then Return
    
    UtilsLib.SetLogger(Logger)
    Dim app As Inventor.Application = ThisApplication

    If app.ActiveDocument Is Nothing Then
        Logger.Error("Ekspordi BOM: No active document")
        System.Windows.Forms.MessageBox.Show("Aktiivne dokument puudub.", "Ekspordi BOM")
        Return
    End If

    If app.ActiveDocument.DocumentType <> Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
        Logger.Error("Ekspordi BOM: Active document is not an assembly")
        System.Windows.Forms.MessageBox.Show("Aktiivseks dokumendiks peab olema koost (.iam).", "Ekspordi BOM")
        Return
    End If

    Try
        Dim asmDoc As Inventor.AssemblyDocument = CType(app.ActiveDocument, Inventor.AssemblyDocument)
        ' Structured BOM = respects structured ordering/item numbering for active model state
        BOMExportLib.ExportWithDialog(asmDoc, False)
    Catch ex As System.Exception
        Logger.Error("Ekspordi BOM: " & ex.ToString())
        System.Windows.Forms.MessageBox.Show("Viga BOM-i eksportimisel:" & vbCrLf & ex.Message, "Ekspordi BOM")
    End Try
End Sub

' --- Optional: batch (uncomment, set paths, run from a copy of the rule) ---
' Dim ex As New System.Collections.Generic.List(Of ExportConfig)()
' ex.Add(New ExportConfig() With { .Name = "Puit", .ModelState = "Puit", .TemplatePath = "C:\Data\BOM_Puit.xlsx", .OutputPath = "C:\Out\out_puit.xlsx" })
' BOMExportLib.ExportBatch(asmDoc, ex, True)
