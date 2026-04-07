' ============================================================================
' TestDrawingViewPlacement - Test drawing view creation and positioning API
' 
' Tests CAMDrawingLib functions:
' - DetermineBaseViewOrientation (sheet metal, BB_ThicknessAxis, default)
' - AddAllViews (T-layout with 6 orthographic views)
' - CreateDrawingFromTemplate
'
' Usage: Open a part document, then run this rule.
'        Creates a new drawing with all 6 orthographic views.
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
    Logger.Info("TestDrawingViewPlacement: Starting view placement tests...")
    
    ' Validate document type
    If doc Is Nothing Then
        Logger.Error("TestDrawingViewPlacement: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "TestDrawingViewPlacement")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("TestDrawingViewPlacement: Active document is not a part file.")
        MessageBox.Show("See reegel töötab ainult detaili failidega (.ipt).", "TestDrawingViewPlacement")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Logger.Info("TestDrawingViewPlacement: Part: " & partDoc.DisplayName)
    
    ' Get part bounding box for logging
    Dim partBox As Box = partDoc.ComponentDefinition.RangeBox
    Dim xSize As Double = (partBox.MaxPoint.X - partBox.MinPoint.X) * 10
    Dim ySize As Double = (partBox.MaxPoint.Y - partBox.MinPoint.Y) * 10
    Dim zSize As Double = (partBox.MaxPoint.Z - partBox.MinPoint.Z) * 10
    
    Logger.Info("TestDrawingViewPlacement: Part size (mm): X=" & FormatNumber(xSize, 2) & _
                " Y=" & FormatNumber(ySize, 2) & " Z=" & FormatNumber(zSize, 2))
    
    ' ========================================================================
    ' Test 1: Check base view orientation detection (using library)
    ' ========================================================================
    Logger.Info("TestDrawingViewPlacement: Test 1 - Base view orientation detection...")
    
    Dim isSheetMetal As Boolean = CAMDrawingLib.IsSheetMetal(partDoc)
    Dim hasFlatPattern As Boolean = CAMDrawingLib.HasFlatPattern(partDoc)
    Dim baseOrientation As ViewOrientationTypeEnum = CAMDrawingLib.DetermineBaseViewOrientation(partDoc)
    
    Logger.Info("TestDrawingViewPlacement: Sheet metal: " & isSheetMetal)
    Logger.Info("TestDrawingViewPlacement: Has flat pattern: " & hasFlatPattern)
    Logger.Info("TestDrawingViewPlacement: Base orientation: " & baseOrientation.ToString())
    
    ' ========================================================================
    ' Test 2: Create drawing from template (using library)
    ' ========================================================================
    Logger.Info("TestDrawingViewPlacement: Test 2 - Create drawing from template...")
    
    Dim drawDoc As DrawingDocument = CAMDrawingLib.CreateDrawingFromTemplate(app)
    If drawDoc Is Nothing Then
        Logger.Error("TestDrawingViewPlacement: Failed to create drawing")
        MessageBox.Show("Joonise loomine ebaõnnestus.", "TestDrawingViewPlacement")
        Exit Sub
    End If
    
    Dim sheet As Sheet = drawDoc.ActiveSheet
    Logger.Info("TestDrawingViewPlacement: Sheet size: " & FormatNumber(sheet.Width * 10, 1) & " x " & _
                FormatNumber(sheet.Height * 10, 1) & " mm")
    
    ' ========================================================================
    ' Test 3: Calculate and set sheet size (using library)
    ' ========================================================================
    Logger.Info("TestDrawingViewPlacement: Test 3 - Calculate sheet size...")
    
    Dim dimSpaceMm As Double = Math.Max(30, Math.Max(xSize, Math.Max(ySize, zSize)) * 0.08)
    Dim sheetSize() As Double = CAMDrawingLib.CalculateSheetSize(partDoc, dimSpaceMm)
    CAMDrawingLib.ResizeSheet(sheet, sheetSize(0), sheetSize(1))
    ' ========================================================================
    ' Test 4: Add all views (using library)
    ' ========================================================================
    Logger.Info("TestDrawingViewPlacement: Test 4 - Add all views...")
    
    Dim views As List(Of DrawingView) = CAMDrawingLib.AddAllViews(sheet, partDoc, app)
    Logger.Info("TestDrawingViewPlacement: Views created: " & views.Count)
    
    ' Log view details
    For i As Integer = 0 To views.Count - 1
        Dim view As DrawingView = views(i)
        Logger.Info("TestDrawingViewPlacement:   View " & (i + 1) & ": " & view.Name & _
                    " Position=(" & FormatNumber(view.Position.X * 10, 1) & "," & _
                    FormatNumber(view.Position.Y * 10, 1) & ") mm Scale=" & view.Scale)
    Next
    
    ' ========================================================================
    ' Test 5: Verify scale
    ' ========================================================================
    Logger.Info("TestDrawingViewPlacement: Test 5 - Verify scale...")
    
    Dim allScalesCorrect As Boolean = True
    For Each view As DrawingView In views
        If Math.Abs(view.Scale - 1.0) > 0.001 Then
            Logger.Warn("TestDrawingViewPlacement: View " & view.Name & " scale is NOT 1:1")
            allScalesCorrect = False
        End If
    Next
    
    If allScalesCorrect Then
        Logger.Info("TestDrawingViewPlacement: All views are 1:1 scale - PASS")
    End If
    
    ' ========================================================================
    ' Summary
    ' ========================================================================
    Logger.Info("TestDrawingViewPlacement: ========================================")
    Logger.Info("TestDrawingViewPlacement: TEST SUMMARY")
    Logger.Info("TestDrawingViewPlacement: ========================================")
    Logger.Info("TestDrawingViewPlacement: Part: " & partDoc.DisplayName)
    Logger.Info("TestDrawingViewPlacement: Sheet metal: " & isSheetMetal)
    Logger.Info("TestDrawingViewPlacement: Flat pattern: " & hasFlatPattern)
    Logger.Info("TestDrawingViewPlacement: Base orientation: " & baseOrientation.ToString())
    Logger.Info("TestDrawingViewPlacement: Views created: " & views.Count)
    Logger.Info("TestDrawingViewPlacement: All 1:1 scale: " & allScalesCorrect)
    Logger.Info("TestDrawingViewPlacement: ========================================")
    
    MessageBox.Show("View placement test completed." & vbCrLf & vbCrLf & _
                    "Base orientation: " & baseOrientation.ToString() & vbCrLf & _
                    "Views created: " & views.Count & vbCrLf & _
                    "All 1:1 scale: " & allScalesCorrect & vbCrLf & vbCrLf & _
                    "Check iLogic log for details." & vbCrLf & _
                    "Drawing is left open for inspection.", "TestDrawingViewPlacement")
End Sub
