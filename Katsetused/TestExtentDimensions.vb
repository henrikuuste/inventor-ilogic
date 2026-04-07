' ============================================================================
' TestExtentDimensions - Test DrawingCurves iteration and extent dimensions
' 
' Tests CAMDrawingLib functions:
' - CreateDrawingFromTemplate
' - AddAllViews (multiple views with proper spacing)
' - AddExtentDimensions (places dimensions OUTSIDE view bounds)
' - AddExtentDimensionsToAllViews (adds dims to all views)
' - CalculateSheetSizeFromViews (accounting for dimension spacing)
' - FitSheetToViews (moves views and resizes sheet)
'
' Usage: Open a part document, then run this rule.
'        Creates a drawing with multiple views and extent dimensions.
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/CAMDrawingLib.vb"

Imports Inventor
Imports System.Collections.Generic

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    Dim doc As Document = app.ActiveDocument
    Logger.Info("TestExtentDimensions: Starting extent dimension tests...")
    
    ' Validate document type
    If doc Is Nothing Then
        Logger.Error("TestExtentDimensions: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "TestExtentDimensions")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("TestExtentDimensions: Active document is not a part file.")
        MessageBox.Show("See reegel töötab ainult detaili failidega (.ipt).", "TestExtentDimensions")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Logger.Info("TestExtentDimensions: Part: " & partDoc.DisplayName)
    
    ' ========================================================================
    ' Test 1: Create drawing and add all views
    ' ========================================================================
    Logger.Info("TestExtentDimensions: Test 1 - Create drawing with all views...")
    
    Dim drawDoc As DrawingDocument = CAMDrawingLib.CreateDrawingFromTemplate(app)
    If drawDoc Is Nothing Then
        Logger.Error("TestExtentDimensions: Failed to create drawing")
        MessageBox.Show("Joonise loomine ebaõnnestus.", "TestExtentDimensions")
        Exit Sub
    End If
    
    Dim sheet As Sheet = drawDoc.ActiveSheet
    
    ' Add all views (front, projected, flat pattern if sheet metal)
    ' The library now uses dimSpace that accounts for dimension space
    Dim views As List(Of DrawingView) = CAMDrawingLib.AddAllViews(sheet, partDoc, app)
    Logger.Info("TestExtentDimensions: Views created: " & views.Count)
    
    ' Log view info
    For i As Integer = 0 To views.Count - 1
        Dim view As DrawingView = views(i)
        Logger.Info("TestExtentDimensions:   " & view.Name & ": " & _
                    FormatNumber(view.Width * 10, 1) & " x " & FormatNumber(view.Height * 10, 1) & " mm at (" & _
                    FormatNumber(view.Position.X * 10, 1) & ", " & FormatNumber(view.Position.Y * 10, 1) & ")")
    Next
    
    ' ========================================================================
    ' Test 2: Verify DrawingCurves iteration on first view
    ' ========================================================================
    Logger.Info("TestExtentDimensions: Test 2 - Verify DrawingCurves iteration on first view...")
    
    Dim firstView As DrawingView = views(0)
    Dim curveCount As Integer = 0
    Dim minX As Double = Double.MaxValue
    Dim maxX As Double = Double.MinValue
    Dim minY As Double = Double.MaxValue
    Dim maxY As Double = Double.MinValue
    
    Try
        For Each curve As DrawingCurve In firstView.DrawingCurves
            curveCount += 1
            
            ' Check start point
            Try
                Dim startPt As Point2d = curve.StartPoint
                If startPt IsNot Nothing Then
                    If startPt.X < minX Then minX = startPt.X
                    If startPt.X > maxX Then maxX = startPt.X
                    If startPt.Y < minY Then minY = startPt.Y
                    If startPt.Y > maxY Then maxY = startPt.Y
                End If
            Catch : End Try
            
            ' Check end point
            Try
                Dim endPt As Point2d = curve.EndPoint
                If endPt IsNot Nothing Then
                    If endPt.X < minX Then minX = endPt.X
                    If endPt.X > maxX Then maxX = endPt.X
                    If endPt.Y < minY Then minY = endPt.Y
                    If endPt.Y > maxY Then maxY = endPt.Y
                End If
            Catch : End Try
            
            ' Check mid point (for arcs)
            Try
                Dim midPt As Point2d = curve.MidPoint
                If midPt IsNot Nothing Then
                    If midPt.X < minX Then minX = midPt.X
                    If midPt.X > maxX Then maxX = midPt.X
                    If midPt.Y < minY Then minY = midPt.Y
                    If midPt.Y > maxY Then maxY = midPt.Y
                End If
            Catch : End Try
        Next
        
        Logger.Info("TestExtentDimensions: " & firstView.Name & " - Found " & curveCount & " curves")
        Logger.Info("TestExtentDimensions: X range: " & FormatNumber(minX * 10, 2) & " to " & FormatNumber(maxX * 10, 2) & " mm")
        Logger.Info("TestExtentDimensions: Y range: " & FormatNumber(minY * 10, 2) & " to " & FormatNumber(maxY * 10, 2) & " mm")
        Logger.Info("TestExtentDimensions: Extent: " & FormatNumber((maxX - minX) * 10, 2) & " x " & FormatNumber((maxY - minY) * 10, 2) & " mm")
    Catch ex As Exception
        Logger.Error("TestExtentDimensions: Error iterating curves: " & ex.Message)
    End Try
    
    ' ========================================================================
    ' Test 3: Add extent dimensions to ALL VIEWS
    ' ========================================================================
    Logger.Info("TestExtentDimensions: Test 3 - AddExtentDimensionsToAllViews...")
    
    Dim dimCountBefore As Integer = sheet.DrawingDimensions.Count
    
    ' Add dimensions to all views - view spacing should account for this
    CAMDrawingLib.AddExtentDimensionsToAllViews(sheet, app)
    Dim dimCountAfter As Integer = sheet.DrawingDimensions.Count
    Dim dimsCreated As Integer = dimCountAfter - dimCountBefore
    
    Logger.Info("TestExtentDimensions: Dimensions before: " & dimCountBefore)
    Logger.Info("TestExtentDimensions: Dimensions after: " & dimCountAfter)
    Logger.Info("TestExtentDimensions: Dimensions created: " & dimsCreated)
    
    ' Expected: 2 dimensions per view (horizontal + vertical)
    Dim expectedDims As Integer = views.Count * 2
    Dim dimPass As Boolean = dimsCreated >= expectedDims
    Logger.Info("TestExtentDimensions: Expected at least " & expectedDims & " dimensions - " & If(dimPass, "PASS", "FAIL"))
    
    ' ========================================================================
    ' Test 4: Calculate sheet size using library function
    ' ========================================================================
    Logger.Info("TestExtentDimensions: Test 4 - Calculate sheet size with dimension spacing...")
    
    ' Use library function that properly accounts for dimension space
    Dim dimOffset As Double = CAMDrawingLib.DIMENSION_OFFSET * 10  ' mm
    Dim borderPadding As Double = 15  ' mm
    
    Dim requiredSize() As Double = CAMDrawingLib.CalculateSheetSizeFromViews(views, dimOffset, borderPadding)
    Dim requiredWidth As Double = requiredSize(0)
    Dim requiredHeight As Double = requiredSize(1)
    
    Logger.Info("TestExtentDimensions: Required sheet size: " & _
                FormatNumber(requiredWidth, 1) & " x " & FormatNumber(requiredHeight, 1) & " mm")
    
    ' ========================================================================
    ' Test 5: Move views and resize sheet using library function
    ' ========================================================================
    Logger.Info("TestExtentDimensions: Test 5 - FitSheetToViews...")
    
    CAMDrawingLib.FitSheetToViews(sheet, views, app, dimOffset, borderPadding)
    Dim finalWidth As Double = sheet.Width * 10
    Dim finalHeight As Double = sheet.Height * 10
    
    Logger.Info("TestExtentDimensions: Final sheet size: " & FormatNumber(finalWidth, 1) & " x " & _
                FormatNumber(finalHeight, 1) & " mm")
    
    ' ========================================================================
    ' Test 6: Verify dimensions still exist after resize
    ' ========================================================================
    Logger.Info("TestExtentDimensions: Test 6 - Verify dimensions after resize...")
    
    Dim dimCountFinal As Integer = sheet.DrawingDimensions.Count
    Logger.Info("TestExtentDimensions: Dimensions after resize: " & dimCountFinal)
    
    Dim dimPreserved As Boolean = dimCountFinal >= dimsCreated
    Logger.Info("TestExtentDimensions: Dimensions preserved: " & If(dimPreserved, "PASS", "FAIL"))
    
    ' ========================================================================
    ' Summary
    ' ========================================================================
    Dim passed As Boolean = dimPass AndAlso dimPreserved
    
    Logger.Info("TestExtentDimensions: ========================================")
    Logger.Info("TestExtentDimensions: TEST SUMMARY")
    Logger.Info("TestExtentDimensions: ========================================")
    Logger.Info("TestExtentDimensions: Part: " & partDoc.DisplayName)
    Logger.Info("TestExtentDimensions: Views created: " & views.Count)
    Logger.Info("TestExtentDimensions: First view curves: " & curveCount)
    Logger.Info("TestExtentDimensions: First view extent: " & FormatNumber((maxX - minX) * 10, 1) & " x " & FormatNumber((maxY - minY) * 10, 1) & " mm")
    Logger.Info("TestExtentDimensions: Dimensions created: " & dimsCreated & " (expected: " & expectedDims & ")")
    Logger.Info("TestExtentDimensions: Dimension offset: " & dimOffset & " mm")
    Logger.Info("TestExtentDimensions: Final sheet size: " & FormatNumber(finalWidth, 0) & " x " & FormatNumber(finalHeight, 0) & " mm")
    Logger.Info("TestExtentDimensions: Test result: " & If(passed, "PASS", "FAIL"))
    Logger.Info("TestExtentDimensions: ========================================")
    
    MessageBox.Show("Extent dimension test completed." & vbCrLf & vbCrLf & _
                    "Part: " & partDoc.DisplayName & vbCrLf & _
                    "Views created: " & views.Count & vbCrLf & _
                    "Dimensions created: " & dimsCreated & " (all views)" & vbCrLf & _
                    "Dimension offset: " & dimOffset & " mm" & vbCrLf & _
                    "Final sheet size: " & FormatNumber(finalWidth, 0) & " x " & FormatNumber(finalHeight, 0) & " mm" & vbCrLf & vbCrLf & _
                    "Result: " & If(passed, "PASS", "FAIL") & vbCrLf & vbCrLf & _
                    "Check iLogic log for details.", "TestExtentDimensions")
End Sub
