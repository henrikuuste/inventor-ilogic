' TestClientGraphicsDirect.vb - Direct test of ClientGraphics
' Minimal test based on working forum example
' Run this to verify if overlay graphics work in this Inventor installation

Sub Main()
    Logger.Info("=== Direct ClientGraphics Test ===")
    
    Dim app As Inventor.Application = ThisApplication
    
    Try
        ' Create InteractionEvents
        Logger.Info("Creating InteractionEvents...")
        Dim oIE As InteractionEvents = app.CommandManager.CreateInteractionEvents()
        oIE.Start()
        Logger.Info("InteractionEvents started")
        
        ' Get InteractionGraphics
        Dim oIG As InteractionGraphics = oIE.InteractionGraphics
        Logger.Info("Got InteractionGraphics")
        
        ' Get data sets
        Dim oDataSets As GraphicsDataSets = oIG.GraphicsDataSets
        Logger.Info("Got GraphicsDataSets")
        
        ' Create a coordinate set
        Dim oCoordSet As GraphicsCoordinateSet = oDataSets.CreateCoordinateSet(1)
        Logger.Info("Created CoordinateSet")
        
        ' Create coordinates for a simple spiral (from working example)
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
        
        oCoordSet.PutCoordinates(oPointCoords)
        Logger.Info("Coordinates set (spiral pattern)")
        
        ' Get overlay client graphics
        Dim oClientGraphics As ClientGraphics = oIG.OverlayClientGraphics
        Logger.Info("Got OverlayClientGraphics")
        
        ' Create a graphics node
        Dim oLineNode As GraphicsNode = oClientGraphics.AddNode(1)
        Logger.Info("Created GraphicsNode")
        
        ' Create LineStripGraphics (connected lines)
        Dim oLineStrip As LineStripGraphics = oLineNode.AddLineStripGraphics()
        oLineStrip.CoordinateSet = oCoordSet
        Logger.Info("Created LineStripGraphics")
        
        ' Add color
        Dim oColorSet As GraphicsColorSet = oDataSets.CreateColorSet(1)
        oColorSet.Add(1, 255, 0, 0)  ' Red
        oLineStrip.ColorSet = oColorSet
        Logger.Info("Set color to red")
        
        ' Update the overlay
        oIG.UpdateOverlayGraphics(app.ActiveView)
        Logger.Info("Called UpdateOverlayGraphics")
        
        ' Also try regular view update
        app.ActiveView.Update()
        Logger.Info("Called ActiveView.Update")
        
        ' Show message - graphics should be visible now
        MessageBox.Show("Look for a red spiral near the origin (0,0,0)." & vbCrLf & _
                        "The spiral extends about 8cm in X/Y and 15cm in Z." & vbCrLf & vbCrLf & _
                        "Click OK to clear the graphics.", _
                        "ClientGraphics Test")
        
        ' Stop interaction events
        oIE.Stop()
        Logger.Info("InteractionEvents stopped")
        
    Catch ex As Exception
        Logger.Error("Error: " & ex.Message)
        MessageBox.Show("Error: " & ex.Message, "ClientGraphics Test Failed")
    End Try
    
    Logger.Info("=== Test Complete ===")
End Sub
