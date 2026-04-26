' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' TestDwgExport - Test DWG/DXF export with 2010 format at 1:1 scale
' 
' Tests CAMDrawingLib functions:
' - CreateDrawingFromTemplate
' - DetermineBaseViewOrientation
' - ExportToDwgOrDxf
'
' Usage: Open a part document, then run this rule.
'        Creates a drawing and exports to DWG and DXF.
'
' Note: Creates temporary files in the same folder as the part.
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/FileSearchLib.vb"
AddVbFile "Lib/CAMDrawingLib.vb"

Imports Inventor
Imports System.Collections.Generic

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    Dim doc As Document = app.ActiveDocument
    Logger.Info("TestDwgExport: Starting DWG/DXF export tests...")
    
    ' Validate document type
    If doc Is Nothing Then
        Logger.Error("TestDwgExport: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "TestDwgExport")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("TestDwgExport: Active document is not a part file.")
        MessageBox.Show("See reegel töötab ainult detaili failidega (.ipt).", "TestDwgExport")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Logger.Info("TestDwgExport: Part: " & partDoc.DisplayName)
    
    ' Get output folder
    Dim outputFolder As String = System.IO.Path.GetDirectoryName(partDoc.FullDocumentName)
    Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(partDoc.FullDocumentName) & "_CAM_Test"
    
    Logger.Info("TestDwgExport: Output folder: " & outputFolder)
    Logger.Info("TestDwgExport: Base name: " & baseName)
    
    ' ========================================================================
    ' Test 1: Create drawing with view (using library)
    ' ========================================================================
    Logger.Info("TestDwgExport: Test 1 - Create drawing with view...")
    
    Dim drawDoc As DrawingDocument = CAMDrawingLib.CreateDrawingFromTemplate(app)
    If drawDoc Is Nothing Then
        Logger.Error("TestDwgExport: Failed to create drawing")
        MessageBox.Show("Joonise loomine ebaõnnestus.", "TestDwgExport")
        Exit Sub
    End If
    
    Dim sheet As Sheet = drawDoc.ActiveSheet
    
    ' Add a view at 1:1 scale with determined orientation
    Dim baseOrientation As ViewOrientationTypeEnum = CAMDrawingLib.DetermineBaseViewOrientation(partDoc)
    Try
        Dim frontView As DrawingView = sheet.DrawingViews.AddBaseView( _
            partDoc, _
            app.TransientGeometry.CreatePoint2d(sheet.Width / 2, sheet.Height / 2), _
            1.0, _
            baseOrientation, _
            DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
            Nothing, Nothing)
        Logger.Info("TestDwgExport: View created at 1:1 scale with orientation: " & baseOrientation.ToString())
    Catch ex As Exception
        Logger.Error("TestDwgExport: Failed to create view: " & ex.Message)
    End Try
    
    ' ========================================================================
    ' Test 2: Export to DWG (using library)
    ' ========================================================================
    Logger.Info("TestDwgExport: Test 2 - Export to DWG (2010 format)...")
    
    Dim dwgPath As String = System.IO.Path.Combine(outputFolder, baseName & ".dwg")
    Dim dwgExportSuccess As Boolean = False
    
    CAMDrawingLib.ExportToDwgOrDxf(app, drawDoc, dwgPath, "DWG")
    If System.IO.File.Exists(dwgPath) Then
        Dim fileInfo As New System.IO.FileInfo(dwgPath)
        dwgExportSuccess = True
        Logger.Info("TestDwgExport: DWG export SUCCESS")
        Logger.Info("TestDwgExport: File: " & dwgPath)
        Logger.Info("TestDwgExport: Size: " & fileInfo.Length & " bytes")
    Else
        Logger.Error("TestDwgExport: DWG file not created")
    End If
    
    ' ========================================================================
    ' Test 3: Export to DXF (using library)
    ' ========================================================================
    Logger.Info("TestDwgExport: Test 3 - Export to DXF (2010 format)...")
    
    Dim dxfPath As String = System.IO.Path.Combine(outputFolder, baseName & ".dxf")
    Dim dxfExportSuccess As Boolean = False
    
    CAMDrawingLib.ExportToDwgOrDxf(app, drawDoc, dxfPath, "DXF")
    If System.IO.File.Exists(dxfPath) Then
        Dim fileInfo As New System.IO.FileInfo(dxfPath)
        dxfExportSuccess = True
        Logger.Info("TestDwgExport: DXF export SUCCESS")
        Logger.Info("TestDwgExport: File: " & dxfPath)
        Logger.Info("TestDwgExport: Size: " & fileInfo.Length & " bytes")
    Else
        Logger.Error("TestDwgExport: DXF file not created")
    End If
    
    ' ========================================================================
    ' Test 4: Check exported file info
    ' ========================================================================
    Logger.Info("TestDwgExport: Test 4 - Exported file info...")
    
    If dwgExportSuccess Then
        Logger.Info("TestDwgExport: DWG path: " & dwgPath)
    End If
    If dxfExportSuccess Then
        Logger.Info("TestDwgExport: DXF path: " & dxfPath)
    End If
    
    ' ========================================================================
    ' Summary
    ' ========================================================================
    Logger.Info("TestDwgExport: ========================================")
    Logger.Info("TestDwgExport: TEST SUMMARY")
    Logger.Info("TestDwgExport: ========================================")
    Logger.Info("TestDwgExport: Part: " & partDoc.DisplayName)
    Logger.Info("TestDwgExport: Base orientation: " & baseOrientation.ToString())
    Logger.Info("TestDwgExport: DWG export: " & If(dwgExportSuccess, "PASS", "FAIL"))
    Logger.Info("TestDwgExport: DXF export: " & If(dxfExportSuccess, "PASS", "FAIL"))
    Logger.Info("TestDwgExport: ========================================")
    
    MessageBox.Show("DWG/DXF export test completed." & vbCrLf & vbCrLf & _
                    "Part: " & partDoc.DisplayName & vbCrLf & _
                    "Orientation: " & baseOrientation.ToString() & vbCrLf & vbCrLf & _
                    "DWG export: " & If(dwgExportSuccess, "PASS", "FAIL") & vbCrLf & _
                    "DXF export: " & If(dxfExportSuccess, "PASS", "FAIL") & vbCrLf & vbCrLf & _
                    If(dwgExportSuccess, "DWG: " & dwgPath & vbCrLf, "") & _
                    If(dxfExportSuccess, "DXF: " & dxfPath & vbCrLf, "") & vbCrLf & _
                    "Check iLogic log for details." & vbCrLf & _
                    "Drawing is left open for inspection.", "TestDwgExport")
End Sub
