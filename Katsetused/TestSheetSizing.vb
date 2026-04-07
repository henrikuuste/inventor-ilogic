' ============================================================================
' TestSheetSizing - Test sheet sizing based on actual view bounds
' 
' Tests CAMDrawingLib functions:
' - CreateDrawingFromTemplate
' - AddAllViews
' - CalculateSheetSizeFromViews (calculates size based on view bounds + dimension space + border)
' - ResizeSheet
'
' The correct workflow is:
' 1. Create drawing with large initial sheet
' 2. Add all views
' 3. Calculate sheet size from actual view bounds + padding
' 4. Resize sheet to fit
'
' Usage: Open a part document, then run this rule.
' ============================================================================

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/CAMDrawingLib.vb"

Imports Inventor
Imports System.Collections.Generic

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    UtilsLib.SetLogger(Logger)
    Dim doc As Document = app.ActiveDocument
    Logger.Info("TestSheetSizing: Starting sheet size tests...")
    
    ' Validate document type
    If doc Is Nothing Then
        Logger.Error("TestSheetSizing: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "TestSheetSizing")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("TestSheetSizing: Active document is not a part file.")
        MessageBox.Show("See reegel töötab ainult detaili failidega (.ipt).", "TestSheetSizing")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Logger.Info("TestSheetSizing: Part: " & partDoc.DisplayName)
    
    ' Log part info
    Dim isSheetMetal As Boolean = CAMDrawingLib.IsSheetMetal(partDoc)
    Dim hasFlatPattern As Boolean = CAMDrawingLib.HasFlatPattern(partDoc)
    Logger.Info("TestSheetSizing: Sheet metal: " & isSheetMetal & ", Flat pattern: " & hasFlatPattern)
    
    ' ========================================================================
    ' Test 1: Create drawing from template
    ' ========================================================================
    Logger.Info("TestSheetSizing: Test 1 - Create drawing from template...")
    
    Dim drawDoc As DrawingDocument = CAMDrawingLib.CreateDrawingFromTemplate(app)
    If drawDoc Is Nothing Then
        Logger.Error("TestSheetSizing: Failed to create drawing")
        MessageBox.Show("Joonise loomine ebaõnnestus.", "TestSheetSizing")
        Exit Sub
    End If
    
    Dim sheet As Sheet = drawDoc.ActiveSheet
    
    Dim initialWidth As Double = sheet.Width * 10
    Dim initialHeight As Double = sheet.Height * 10
    Logger.Info("TestSheetSizing: Initial sheet size: " & FormatNumber(initialWidth, 1) & " x " & _
                FormatNumber(initialHeight, 1) & " mm")
    
    ' ========================================================================
    ' Test 2: Add all views (views need to exist before we can size sheet to fit)
    ' ========================================================================
    Logger.Info("TestSheetSizing: Test 2 - Add all views...")
    
    Dim views As List(Of DrawingView) = CAMDrawingLib.AddAllViews(sheet, partDoc, app)
    Logger.Info("TestSheetSizing: Views created: " & views.Count)
    
    ' Log view positions and sizes
    For i As Integer = 0 To views.Count - 1
        Dim view As DrawingView = views(i)
        Logger.Info("TestSheetSizing:   " & view.Name & ": " & _
                    FormatNumber(view.Width * 10, 1) & " x " & FormatNumber(view.Height * 10, 1) & " mm at (" & _
                    FormatNumber(view.Position.X * 10, 1) & ", " & FormatNumber(view.Position.Y * 10, 1) & ")")
    Next
    
    ' ========================================================================
    ' Test 3: Calculate sheet size from view bounds (using library)
    ' ========================================================================
    Logger.Info("TestSheetSizing: Test 3 - Calculate sheet size from view bounds...")
    
    ' Use default dimension offset (25mm) and 10mm border padding
    Dim requiredSize() As Double = CAMDrawingLib.CalculateSheetSizeFromViews(views)
    Dim requiredWidth As Double = requiredSize(0)
    Dim requiredHeight As Double = requiredSize(1)
    
    Logger.Info("TestSheetSizing: Required sheet size (with dim space + border): " & _
                FormatNumber(requiredWidth, 1) & " x " & FormatNumber(requiredHeight, 1) & " mm")
    
    ' ========================================================================
    ' Test 4: Move views to fit new sheet (MUST do this BEFORE resizing!)
    ' ========================================================================
    Logger.Info("TestSheetSizing: Test 4 - Move views to fit new sheet size...")
    
    ' Get dimension offset for bounds calculation (same as used in CalculateSheetSizeFromViews)
    Dim dimOffsetCm As Double = CAMDrawingLib.DIMENSION_OFFSET  ' 2.5 cm = 25 mm
    
    ' Calculate current view bounds INCLUDING dimension space
    Dim minX As Double = Double.MaxValue
    Dim minY As Double = Double.MaxValue
    Dim maxX As Double = Double.MinValue
    Dim maxY As Double = Double.MinValue
    
    For Each view As DrawingView In views
        Dim vLeft As Double = view.Position.X - view.Width / 2
        Dim vRight As Double = view.Position.X + view.Width / 2 + dimOffsetCm  ' Add dim space on right
        Dim vBottom As Double = view.Position.Y - view.Height / 2 - dimOffsetCm  ' Add dim space below
        Dim vTop As Double = view.Position.Y + view.Height / 2
        
        If vLeft < minX Then minX = vLeft
        If vRight > maxX Then maxX = vRight
        If vBottom < minY Then minY = vBottom
        If vTop > maxY Then maxY = vTop
    Next
    
    ' Calculate offset to move views to be centered on the NEW sheet size
    Dim viewsCenterX As Double = (minX + maxX) / 2
    Dim viewsCenterY As Double = (minY + maxY) / 2
    Dim newSheetCenterX As Double = (requiredWidth / 10) / 2   ' Convert mm to cm
    Dim newSheetCenterY As Double = (requiredHeight / 10) / 2
    
    Dim offsetX As Double = newSheetCenterX - viewsCenterX
    Dim offsetY As Double = newSheetCenterY - viewsCenterY
    
    Logger.Info("TestSheetSizing: Dim offset: " & dimOffsetCm * 10 & " mm")
    Logger.Info("TestSheetSizing: Moving views by (" & FormatNumber(offsetX * 10, 1) & ", " & _
                FormatNumber(offsetY * 10, 1) & ") mm to center on new sheet")
    
    ' Move all views
    For Each view As DrawingView In views
        Dim newPos As Point2d = app.TransientGeometry.CreatePoint2d( _
            view.Position.X + offsetX, _
            view.Position.Y + offsetY)
        view.Position = newPos
    Next
    
    ' ========================================================================
    ' Test 5: Resize sheet to fit views (now that views are within bounds)
    ' ========================================================================
    Logger.Info("TestSheetSizing: Test 5 - Resize sheet to fit views...")
    
    CAMDrawingLib.ResizeSheet(sheet, requiredWidth, requiredHeight)
    Dim afterWidth As Double = sheet.Width * 10
    Dim afterHeight As Double = sheet.Height * 10
    
    Dim widthMatch As Boolean = Math.Abs(afterWidth - requiredWidth) < 1
    Dim heightMatch As Boolean = Math.Abs(afterHeight - requiredHeight) < 1
    Dim resizePass As Boolean = widthMatch AndAlso heightMatch
    
    Logger.Info("TestSheetSizing: After resize: " & FormatNumber(afterWidth, 1) & " x " & _
                FormatNumber(afterHeight, 1) & " mm - " & If(resizePass, "PASS", "FAIL"))
    
    ' ========================================================================
    ' Summary
    ' ========================================================================
    Logger.Info("TestSheetSizing: ========================================")
    Logger.Info("TestSheetSizing: TEST SUMMARY")
    Logger.Info("TestSheetSizing: ========================================")
    Logger.Info("TestSheetSizing: Part: " & partDoc.DisplayName)
    Logger.Info("TestSheetSizing: Sheet metal: " & isSheetMetal)
    Logger.Info("TestSheetSizing: Views created: " & views.Count)
    Logger.Info("TestSheetSizing: Required size (dim space + border): " & FormatNumber(requiredWidth, 0) & " x " & FormatNumber(requiredHeight, 0) & " mm")
    Logger.Info("TestSheetSizing: Final sheet size: " & FormatNumber(afterWidth, 0) & " x " & FormatNumber(afterHeight, 0) & " mm")
    Logger.Info("TestSheetSizing: Resize: " & If(resizePass, "PASS", "FAIL"))
    Logger.Info("TestSheetSizing: ========================================")
    
    MessageBox.Show("Sheet sizing test completed." & vbCrLf & vbCrLf & _
                    "Part: " & partDoc.DisplayName & vbCrLf & _
                    "Sheet metal: " & isSheetMetal & vbCrLf & _
                    "Views created: " & views.Count & vbCrLf & vbCrLf & _
                    "Required size (dim space + border): " & FormatNumber(requiredWidth, 0) & " x " & FormatNumber(requiredHeight, 0) & " mm" & vbCrLf & _
                    "Final sheet size: " & FormatNumber(afterWidth, 0) & " x " & FormatNumber(afterHeight, 0) & " mm" & vbCrLf & vbCrLf & _
                    "Resize: " & If(resizePass, "PASS", "FAIL") & vbCrLf & vbCrLf & _
                    "Check iLogic log for details.", "TestSheetSizing")
End Sub
