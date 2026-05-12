' TestClientGraphicsCompDef.vb - Test ClientGraphics on ComponentDefinition
' This tests the non-InteractionEvents approach that works in add-ins

Sub Main()
    Logger.Info("=== ClientGraphics on ComponentDefinition Test ===")
    
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing Then
        MessageBox.Show("Open a part or assembly first")
        Exit Sub
    End If
    
    Dim compDef As ComponentDefinition = Nothing
    
    If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        compDef = CType(doc, PartDocument).ComponentDefinition
    ElseIf doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
        compDef = CType(doc, AssemblyDocument).ComponentDefinition
    Else
        MessageBox.Show("Open a part or assembly")
        Exit Sub
    End If
    
    Const GRAPHICS_ID As String = "_TestCG_"
    
    Try
        ' Clean up any existing
        Try
            doc.GraphicsDataSetsCollection.Item(GRAPHICS_ID).Delete()
            Logger.Info("Deleted existing GraphicsDataSets")
        Catch
        End Try
        Try
            compDef.ClientGraphicsCollection.Item(GRAPHICS_ID).Delete()
            Logger.Info("Deleted existing ClientGraphics")
        Catch
        End Try
        
        ' Step 1: Create GraphicsDataSets on DOCUMENT
        Logger.Info("Creating GraphicsDataSets on Document...")
        Dim dataSets As GraphicsDataSets = doc.GraphicsDataSetsCollection.Add(GRAPHICS_ID)
        Logger.Info("GraphicsDataSets created: " & dataSets.Count.ToString())
        
        ' Step 2: Create coordinate set
        Logger.Info("Creating CoordinateSet...")
        Dim coordSet As GraphicsCoordinateSet = dataSets.CreateCoordinateSet(1)
        Logger.Info("CoordinateSet created")
        
        ' Step 3: Put coordinates - spiral pattern (same as working test)
        Dim oPointCoords(90) As Double
        Dim dRadius As Double = 1
        Dim dAngle As Double = 0
        
        For i As Integer = 0 To 29
            oPointCoords(i * 3) = dRadius * Math.Cos(dAngle)
            oPointCoords(i * 3 + 1) = dRadius * Math.Sin(dAngle)
            oPointCoords(i * 3 + 2) = i / 2.0
            dRadius = dRadius + 0.25
            dAngle = dAngle + (Math.PI / 6)
        Next
        
        coordSet.PutCoordinates(oPointCoords)
        Logger.Info("Coordinates set (30 points in spiral)")
        
        ' Step 4: Create ClientGraphics on COMPONENTDEFINITION
        Logger.Info("Creating ClientGraphics on ComponentDefinition...")
        Dim clientGraphics As ClientGraphics = compDef.ClientGraphicsCollection.Add(GRAPHICS_ID)
        Logger.Info("ClientGraphics created")
        
        ' Step 5: Create GraphicsNode
        Logger.Info("Creating GraphicsNode...")
        Dim node As GraphicsNode = clientGraphics.AddNode(1)
        Logger.Info("GraphicsNode created")
        
        ' Step 6: Create LineStripGraphics
        Logger.Info("Creating LineStripGraphics...")
        Dim lineStrip As LineStripGraphics = node.AddLineStripGraphics()
        lineStrip.CoordinateSet = coordSet
        lineStrip.LineWeight = 5
        Logger.Info("LineStripGraphics created and configured")
        
        ' Step 7: Add color
        Logger.Info("Adding color...")
        Dim colorSet As GraphicsColorSet = dataSets.CreateColorSet(1)
        colorSet.Add(1, 0, 255, 0)  ' Green
        lineStrip.ColorSet = colorSet
        Logger.Info("Color set to green")
        
        ' Step 8: Update view
        Logger.Info("Updating view...")
        app.ActiveView.Update()
        Logger.Info("View updated")
        
        ' Also try Camera.Fit
        Logger.Info("Fitting camera...")
        Dim camera As Camera = app.ActiveView.Camera
        camera.Fit()
        camera.Apply()
        Logger.Info("Camera fit applied")
        
        ' Show message
        MessageBox.Show("Look for a GREEN spiral near the origin." & vbCrLf & _
                        "Check the log for any issues." & vbCrLf & vbCrLf & _
                        "Click OK to delete the graphics.", _
                        "ComponentDefinition ClientGraphics Test")
        
        ' Cleanup
        clientGraphics.Delete()
        dataSets.Delete()
        Logger.Info("Graphics deleted")
        
    Catch ex As Exception
        Logger.Error("Error: " & ex.Message)
        Logger.Error("Stack: " & ex.StackTrace)
        MessageBox.Show("Error: " & ex.Message, "Test Failed")
    End Try
    
    Logger.Info("=== Test Complete ===")
End Sub
