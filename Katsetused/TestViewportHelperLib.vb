' TestViewportHelperLib.vb - Test script for ViewportHelperLib functionality
' Run this to verify Phase 3 of the Unified UI Library implementation
'
' Requirements: Open a part or assembly document before running
'
' Tests:
' 1. Highlight objects in viewport
' 2. Transient point markers
' 3. Transient line markers
' 4. Preview work features
' 5. Cleanup

AddVbFile "Lib/UILib.vb"
AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/ViewportHelperLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    ' Validate document
    If doc Is Nothing Then
        MessageBox.Show(StringsLib.MSG_NO_ACTIVE_DOCUMENT, "TestViewportHelperLib")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject AndAlso _
       doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MessageBox.Show(StringsLib.MSG_REQUIRES_ASSEMBLY_OR_PART, "TestViewportHelperLib")
        Exit Sub
    End If
    
    Logger.Info("TestViewportHelperLib: Starting ViewportHelperLib tests...")
    Logger.Info("TestViewportHelperLib: Document type: " & doc.DocumentType.ToString())
    
    ' Initialize the library
    ViewportHelperLib.Initialize(app)
    Logger.Info("TestViewportHelperLib: Library initialized")
    
    ' Run tests via non-modal dialog
    ShowTestDialog(app, doc)
    
    ' Final cleanup
    ViewportHelperLib.Cleanup()
    Logger.Info("TestViewportHelperLib: All tests completed and cleaned up.")
End Sub

Sub ShowTestDialog(app As Inventor.Application, doc As Document)
    Dim frm As System.Windows.Forms.Form = UILib.CreateForm("Test: Viewport Helpers", 450, 350)
    Dim content As System.Windows.Forms.TableLayoutPanel = UILib.CreateContentPanel()
    Dim buttons As System.Windows.Forms.FlowLayoutPanel = UILib.CreateButtonPanel()
    
    UILib.AddSectionHeader(content, "Transient Graphics Tests")
    
    ' Add test buttons
    ' Store app in form tag for button handlers (module vars don't persist)
    frm.Tag = app
    
    Dim btnAddPoints As System.Windows.Forms.Button = UILib.CreateButton("Add Point Markers", 140)
    AddHandler btnAddPoints.Click, Sub(s, e)
        Dim appRef As Inventor.Application = CType(frm.Tag, Inventor.Application)
        TestPointMarkers(appRef)
    End Sub
    UILib.AddFullWidthRow(content, btnAddPoints)
    
    Dim btnAddLines As System.Windows.Forms.Button = UILib.CreateButton("Add Line Markers", 140)
    AddHandler btnAddLines.Click, Sub(s, e)
        Dim appRef As Inventor.Application = CType(frm.Tag, Inventor.Application)
        TestLineMarkers(appRef)
    End Sub
    UILib.AddFullWidthRow(content, btnAddLines)
    
    Dim btnClearMarkers As System.Windows.Forms.Button = UILib.CreateButton("Clear Markers", 140)
    AddHandler btnClearMarkers.Click, Sub(s, e)
        Dim appRef As Inventor.Application = CType(frm.Tag, Inventor.Application)
        ViewportHelperLib.ClearMarkers(appRef)
        Logger.Info("TestViewportHelperLib: Markers cleared")
    End Sub
    UILib.AddFullWidthRow(content, btnClearMarkers)
    
    UILib.AddSectionHeader(content, "Preview Features")
    
    Dim btnAddWorkPoint As System.Windows.Forms.Button = UILib.CreateButton("Add Preview WorkPoint", 160)
    AddHandler btnAddWorkPoint.Click, Sub(s, e)
        Dim appRef As Inventor.Application = CType(frm.Tag, Inventor.Application)
        TestPreviewWorkPoint(appRef, appRef.ActiveDocument)
    End Sub
    UILib.AddFullWidthRow(content, btnAddWorkPoint)
    
    Dim btnDeletePreviews As System.Windows.Forms.Button = UILib.CreateButton("Delete Preview Features", 160)
    AddHandler btnDeletePreviews.Click, Sub(s, e)
        Dim count As Integer = ViewportHelperLib.GetPreviewFeatureCount()
        ViewportHelperLib.DeletePreviewFeatures()
        Logger.Info("TestViewportHelperLib: Deleted " & count.ToString() & " preview features")
    End Sub
    UILib.AddFullWidthRow(content, btnDeletePreviews)
    
    Dim btnCommitPreviews As System.Windows.Forms.Button = UILib.CreateButton("Commit (Keep) Previews", 160)
    AddHandler btnCommitPreviews.Click, Sub(s, e)
        Dim count As Integer = ViewportHelperLib.GetPreviewFeatureCount()
        ViewportHelperLib.CommitPreviewFeatures()
        Logger.Info("TestViewportHelperLib: Committed " & count.ToString() & " features (no longer tracked, will remain in model)")
        Logger.Info("TestViewportHelperLib: Note: After commit, those features cannot be deleted via this tool")
    End Sub
    UILib.AddFullWidthRow(content, btnCommitPreviews)
    
    UILib.AddSectionHeader(content, "Highlight Tests")
    
    Dim btnHighlightInfo As System.Windows.Forms.Button = UILib.CreateButton("Highlight Info...", 140)
    AddHandler btnHighlightInfo.Click, Sub(s, e)
        MessageBox.Show( _
            "Highlight tests require pre-selected objects." & vbCrLf & vbCrLf & _
            "1. Close this dialog" & vbCrLf & _
            "2. Select faces/edges in the viewport" & vbCrLf & _
            "3. Run 'Highlight SelectSet' button", _
            "Highlight Test Info")
    End Sub
    UILib.AddFullWidthRow(content, btnHighlightInfo)
    
    Dim btnHighlightSel As System.Windows.Forms.Button = UILib.CreateButton("Highlight SelectSet", 160)
    AddHandler btnHighlightSel.Click, Sub(s, e)
        Dim appRef As Inventor.Application = CType(frm.Tag, Inventor.Application)
        TestHighlightSelectSet(appRef, appRef.ActiveDocument)
    End Sub
    UILib.AddFullWidthRow(content, btnHighlightSel)
    
    Dim btnClearHighlights As System.Windows.Forms.Button = UILib.CreateButton("Clear Highlights", 140)
    AddHandler btnClearHighlights.Click, Sub(s, e)
        ViewportHelperLib.ClearHighlights()
        Logger.Info("TestViewportHelperLib: Highlights cleared")
    End Sub
    UILib.AddFullWidthRow(content, btnClearHighlights)
    
    ' Close button
    Dim btnClose As System.Windows.Forms.Button = UILib.CreateButton(StringsLib.BTN_CLOSE)
    AddHandler btnClose.Click, Sub(s, e)
        frm.Close()
    End Sub
    buttons.Controls.Add(btnClose)
    
    frm.Controls.Add(content)
    frm.Controls.Add(buttons)
    
    UILib.FinalizeForm(frm)
    
    ' Show as non-modal so user can interact with viewport
    UILib.ShowNonModal(frm)
End Sub

Sub TestPointMarkers(app As Inventor.Application)
    Logger.Info("TestViewportHelperLib: Adding point markers using ClientGraphics...")
    
    ' Add points in a pattern around origin (units are cm)
    ' Using orange color for visibility
    ' app is passed to each call because module vars don't persist
    ViewportHelperLib.AddPointMarkerAt(app, 0, 0, 0, "orange")   ' Origin
    ViewportHelperLib.AddPointMarkerAt(app, 5, 0, 0, "orange")   ' +X
    ViewportHelperLib.AddPointMarkerAt(app, 0, 5, 0, "orange")   ' +Y
    ViewportHelperLib.AddPointMarkerAt(app, 0, 0, 5, "orange")   ' +Z
    ViewportHelperLib.AddPointMarkerAt(app, 5, 5, 0, "orange")   ' XY diagonal
    ViewportHelperLib.AddPointMarkerAt(app, 5, 0, 5, "orange")   ' XZ diagonal
    ViewportHelperLib.AddPointMarkerAt(app, 0, 5, 5, "orange")   ' YZ diagonal
    ViewportHelperLib.AddPointMarkerAt(app, 5, 5, 5, "orange")   ' XYZ corner
    
    Logger.Info("TestViewportHelperLib: 8 ORANGE point markers created at 0-5cm from origin")
    Logger.Info("TestViewportHelperLib: Each point is a small 3D cross, visible in viewport")
    Logger.Info("TestViewportHelperLib: Zoom to origin area to see markers")
    Logger.Info("TestViewportHelperLib: Click 'Clear Markers' to remove")
End Sub

Sub TestLineMarkers(app As Inventor.Application)
    Logger.Info("TestViewportHelperLib: Adding line markers using ClientGraphics...")
    
    ' Draw coordinate axes in different colors
    ' app is passed to each call because module vars don't persist
    ViewportHelperLib.AddLineMarkerAt(app, 0, 0, 0, 10, 0, 0, "red")    ' X axis - red
    ViewportHelperLib.AddLineMarkerAt(app, 0, 0, 0, 0, 10, 0, "green")  ' Y axis - green
    ViewportHelperLib.AddLineMarkerAt(app, 0, 0, 0, 0, 0, 10, "blue")   ' Z axis - blue
    
    ' Draw a triangle in yellow
    ViewportHelperLib.AddLineMarkerAt(app, 0, 0, 0, 5, 0, 0, "yellow")
    ViewportHelperLib.AddLineMarkerAt(app, 5, 0, 0, 2.5, 4, 0, "yellow")
    ViewportHelperLib.AddLineMarkerAt(app, 2.5, 4, 0, 0, 0, 0, "yellow")
    
    Logger.Info("TestViewportHelperLib: Created colored axes: X=red, Y=green, Z=blue (10cm)")
    Logger.Info("TestViewportHelperLib: Created yellow triangle at origin")
    Logger.Info("TestViewportHelperLib: Zoom to origin area to see markers")
    Logger.Info("TestViewportHelperLib: Click 'Clear Markers' to remove")
End Sub

Sub TestPreviewWorkPoint(app As Inventor.Application, doc As Document)
    Logger.Info("TestViewportHelperLib: Creating preview work point...")
    
    Try
        Dim compDef As Object = Nothing
        
        If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            compDef = CType(doc, PartDocument).ComponentDefinition
        ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            compDef = CType(doc, AssemblyDocument).ComponentDefinition
        End If
        
        ' Create a point at 3,3,3 cm
        Dim tg As TransientGeometry = app.TransientGeometry
        Dim pt As Object = tg.CreatePoint(3, 3, 3)
        
        Dim wp As Object = ViewportHelperLib.CreatePreviewWorkPoint(compDef, pt, "_Preview_WorkPoint_")
        
        If wp IsNot Nothing Then
            Logger.Info("TestViewportHelperLib: Preview work point created at (3, 3, 3) cm")
            Logger.Info("TestViewportHelperLib: Click 'Delete Preview Features' to remove, or 'Commit' to keep")
        Else
            Logger.Warn("TestViewportHelperLib: Failed to create preview work point")
        End If
    Catch ex As Exception
        Logger.Error("TestViewportHelperLib: Error creating work point - " & ex.Message)
    End Try
End Sub

Sub TestHighlightSelectSet(app As Inventor.Application, doc As Document)
    Logger.Info("TestViewportHelperLib: Highlighting current selection...")
    
    Try
        Dim selSet As SelectSet = doc.SelectSet
        
        If selSet.Count = 0 Then
            Logger.Warn("TestViewportHelperLib: Nothing selected. Select objects in viewport first.")
            MessageBox.Show("Vali esmalt objekte vaateaknas.", "Highlight Test")
            Exit Sub
        End If
        
        ViewportHelperLib.ClearHighlights()
        
        Dim highlightedCount As Integer = 0
        For i As Integer = 1 To selSet.Count
            ViewportHelperLib.Highlight(selSet.Item(i))
            highlightedCount += 1
        Next
        
        Logger.Info("TestViewportHelperLib: Highlighted " & highlightedCount.ToString() & " objects")
        Logger.Info("TestViewportHelperLib: Note: Highlight is separate from selection.")
        Logger.Info("TestViewportHelperLib: If highlight disappears when selection clears, this may be Inventor behavior")
        Logger.Info("TestViewportHelperLib: For persistent markers, use Preview Work Features instead")
    Catch ex As Exception
        Logger.Error("TestViewportHelperLib: Error highlighting - " & ex.Message)
    End Try
End Sub
