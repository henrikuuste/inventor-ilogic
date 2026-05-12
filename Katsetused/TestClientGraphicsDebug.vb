' TestClientGraphicsDebug.vb - Comprehensive ClientGraphics debugging
' Tests various scenarios to find what works and what doesn't

AddVbFile "Lib/UILib.vb"
AddVbFile "Lib/StringsLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse (doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject AndAlso _
                               doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject) Then
        MessageBox.Show("Open a part or assembly first")
        Exit Sub
    End If
    
    Logger.Info("========================================")
    Logger.Info("ClientGraphics Debug Test Suite")
    Logger.Info("========================================")
    
    ShowTestDialog(app, doc)
End Sub

Sub ShowTestDialog(app As Inventor.Application, doc As Document)
    Dim frm As System.Windows.Forms.Form = UILib.CreateForm("ClientGraphics Debug", 500, 550)
    Dim content As System.Windows.Forms.TableLayoutPanel = UILib.CreateContentPanel()
    
    Dim compDef As ComponentDefinition = doc.ComponentDefinition
    
    ' Shared state via Form.Tag (workaround for no module variables)
    frm.Tag = New Dictionary(Of String, Object)()
    Dim state As Dictionary(Of String, Object) = CType(frm.Tag, Dictionary(Of String, Object))
    state("app") = app
    state("doc") = doc
    state("compDef") = compDef
    state("coordIndex") = 0
    state("test3Counter") = 0
    
    ' Test 1: Direct in Main (like working test)
    UILib.AddSectionHeader(content, "Test 1: Direct Creation (Reference)")
    
    Dim btn1 As System.Windows.Forms.Button = UILib.CreateButton("Create Direct (Should Work)", 220)
    AddHandler btn1.Click, Sub(s, e)
        Test1_DirectCreation(state)
    End Sub
    UILib.AddFullWidthRow(content, btn1)
    
    ' Test 2: Using state dictionary (simulates module variables)
    UILib.AddSectionHeader(content, "Test 2: Persistent State")
    
    Dim btn2a As System.Windows.Forms.Button = UILib.CreateButton("Setup Graphics", 150)
    AddHandler btn2a.Click, Sub(s, e)
        Test2a_SetupGraphics(state)
    End Sub
    UILib.AddFullWidthRow(content, btn2a)
    
    Dim btn2b As System.Windows.Forms.Button = UILib.CreateButton("Add Point (State)", 150)
    AddHandler btn2b.Click, Sub(s, e)
        Test2b_AddPointState(state)
    End Sub
    UILib.AddFullWidthRow(content, btn2b)
    
    Dim btn2c As System.Windows.Forms.Button = UILib.CreateButton("Add Line (State)", 150)
    AddHandler btn2c.Click, Sub(s, e)
        Test2c_AddLineState(state)
    End Sub
    UILib.AddFullWidthRow(content, btn2c)
    
    ' Test 3: Fresh creation each time
    UILib.AddSectionHeader(content, "Test 3: Fresh Each Time")
    
    Dim btn3 As System.Windows.Forms.Button = UILib.CreateButton("Create Fresh Graphics", 180)
    AddHandler btn3.Click, Sub(s, e)
        Test3_FreshCreation(state)
    End Sub
    UILib.AddFullWidthRow(content, btn3)
    
    ' Test 4: Add to existing
    UILib.AddSectionHeader(content, "Test 4: Add to Existing")
    
    Dim btn4 As System.Windows.Forms.Button = UILib.CreateButton("Add to Test1 Graphics", 180)
    AddHandler btn4.Click, Sub(s, e)
        Test4_AddToExisting(state)
    End Sub
    UILib.AddFullWidthRow(content, btn4)
    
    ' Cleanup
    UILib.AddSectionHeader(content, "Cleanup")
    
    Dim btnCleanup As System.Windows.Forms.Button = UILib.CreateButton("Clear All Graphics", 150)
    AddHandler btnCleanup.Click, Sub(s, e)
        CleanupAllGraphics(state)
    End Sub
    UILib.AddFullWidthRow(content, btnCleanup)
    
    Dim btnClose As System.Windows.Forms.Button = UILib.CreateButton(StringsLib.BTN_CLOSE)
    AddHandler btnClose.Click, Sub(s, e)
        CleanupAllGraphics(state)
        frm.Close()
    End Sub
    UILib.AddFullWidthRow(content, btnClose)
    
    frm.Controls.Add(content)
    UILib.FinalizeForm(frm)
    UILib.ShowNonModal(frm)
End Sub

' ============================================================
' TEST 1: Direct creation (like working standalone test)
' ============================================================
Sub Test1_DirectCreation(state As Dictionary(Of String, Object))
    Logger.Info("")
    Logger.Info("--- TEST 1: Direct Creation (LineStrip - like working spiral) ---")
    
    Dim app As Inventor.Application = CType(state("app"), Inventor.Application)
    Dim doc As Document = CType(state("doc"), Document)
    Dim compDef As ComponentDefinition = CType(state("compDef"), ComponentDefinition)
    
    Const ID As String = "_Test1_"
    
    Try
        ' Cleanup first
        Try : doc.GraphicsDataSetsCollection.Item(ID).Delete() : Catch : End Try
        Try : compDef.ClientGraphicsCollection.Item(ID).Delete() : Catch : End Try
        
        ' Create everything locally
        Dim dataSets As GraphicsDataSets = doc.GraphicsDataSetsCollection.Add(ID)
        Logger.Info("Test1: DataSets created")
        
        Dim coordSet As GraphicsCoordinateSet = dataSets.CreateCoordinateSet(1)
        Logger.Info("Test1: CoordSet created")
        
        ' Create a triangle using LineStripGraphics (4 points, 3 lines)
        Dim coords(11) As Double  ' 4 points x 3 coords
        coords(0) = 0 : coords(1) = 0 : coords(2) = 0     ' Point 1: origin
        coords(3) = 10 : coords(4) = 0 : coords(5) = 0    ' Point 2: +X
        coords(6) = 5 : coords(7) = 8 : coords(8) = 0     ' Point 3: top
        coords(9) = 0 : coords(10) = 0 : coords(11) = 0   ' Point 4: back to origin
        coordSet.PutCoordinates(coords)
        Logger.Info("Test1: Coords set for triangle (origin, +10X, top, back)")
        
        Dim cg As ClientGraphics = compDef.ClientGraphicsCollection.Add(ID)
        Logger.Info("Test1: ClientGraphics created")
        
        Dim node As GraphicsNode = cg.AddNode(1)
        Logger.Info("Test1: Node created")
        
        ' Create color set for red
        Dim colorSet As GraphicsColorSet = dataSets.CreateColorSet(2)
        colorSet.Add(1, 255, 50, 50)  ' Red
        
        ' Use LineStripGraphics like the working spiral test
        Dim ls As LineStripGraphics = node.AddLineStripGraphics()
        ls.CoordinateSet = coordSet
        ls.ColorSet = colorSet
        ls.LineWeight = 5
        Logger.Info("Test1: LineStripGraphics created (RED triangle)")
        
        app.ActiveView.Update()
        Logger.Info("Test1: View updated")
        Logger.Info("Test1: SUCCESS - Look for triangle at origin (10cm base)")
        
    Catch ex As Exception
        Logger.Error("Test1 FAILED: " & ex.Message)
    End Try
End Sub

' ============================================================
' TEST 2: Using state dictionary (simulates module variables)
' ============================================================
Sub Test2a_SetupGraphics(state As Dictionary(Of String, Object))
    Logger.Info("")
    Logger.Info("--- TEST 2a: Setup State Graphics ---")
    
    Dim app As Inventor.Application = CType(state("app"), Inventor.Application)
    Dim doc As Document = CType(state("doc"), Document)
    Dim compDef As ComponentDefinition = CType(state("compDef"), ComponentDefinition)
    
    Const ID As String = "_Test2_"
    
    Try
        ' Cleanup
        Try : doc.GraphicsDataSetsCollection.Item(ID).Delete() : Catch : End Try
        Try : compDef.ClientGraphicsCollection.Item(ID).Delete() : Catch : End Try
        
        ' Create and store in state
        Dim dataSets As GraphicsDataSets = doc.GraphicsDataSetsCollection.Add(ID)
        state("dataSets") = dataSets
        Logger.Info("Test2a: DataSets created and stored")
        
        Dim cg As ClientGraphics = compDef.ClientGraphicsCollection.Add(ID)
        state("clientGraphics") = cg
        Logger.Info("Test2a: ClientGraphics created and stored")
        
        Dim node As GraphicsNode = cg.AddNode(1)
        state("graphicsNode") = node
        Logger.Info("Test2a: Node created and stored")
        
        ' Create a color set for subsequent items (green)
        Dim colorSet As GraphicsColorSet = dataSets.CreateColorSet(99)
        colorSet.Add(1, 50, 255, 50)  ' Green
        state("colorSet") = colorSet
        
        state("coordIndex") = 0
        Logger.Info("Test2a: Setup complete (GREEN) - now click 'Add Point (State)'")
        
    Catch ex As Exception
        Logger.Error("Test2a FAILED: " & ex.Message)
    End Try
End Sub

Sub Test2b_AddPointState(state As Dictionary(Of String, Object))
    Logger.Info("")
    Logger.Info("--- TEST 2b: Add Point Using State ---")
    
    Dim app As Inventor.Application = CType(state("app"), Inventor.Application)
    
    If Not state.ContainsKey("dataSets") OrElse Not state.ContainsKey("graphicsNode") Then
        Logger.Error("Test2b: State not set up! Click 'Setup Graphics' first")
        Return
    End If
    
    Dim dataSets As GraphicsDataSets = CType(state("dataSets"), GraphicsDataSets)
    Dim node As GraphicsNode = CType(state("graphicsNode"), GraphicsNode)
    Dim colorSet As GraphicsColorSet = If(state.ContainsKey("colorSet"), CType(state("colorSet"), GraphicsColorSet), Nothing)
    
    Logger.Info("Test2b: dataSets IsNothing=" & (dataSets Is Nothing).ToString())
    Logger.Info("Test2b: node IsNothing=" & (node Is Nothing).ToString())
    
    Try
        Dim idx As Integer = CInt(state("coordIndex")) + 1
        state("coordIndex") = idx
        
        Dim coordSet As GraphicsCoordinateSet = dataSets.CreateCoordinateSet(idx)
        Logger.Info("Test2b: CoordSet created with index " & idx.ToString())
        
        ' Draw a small cross using LineStripGraphics (more visible than PointGraphics)
        Dim x As Double = 3, y As Double = 0, z As Double = 0
        Dim size As Double = 0.5
        
        ' Horizontal line
        Dim coords1(5) As Double
        coords1(0) = x - size : coords1(1) = y : coords1(2) = z
        coords1(3) = x + size : coords1(4) = y : coords1(5) = z
        coordSet.PutCoordinates(coords1)
        
        Dim ls1 As LineStripGraphics = node.AddLineStripGraphics()
        ls1.CoordinateSet = coordSet
        ls1.LineWeight = 3
        If colorSet IsNot Nothing Then ls1.ColorSet = colorSet
        
        ' Vertical line
        Dim idx2 As Integer = CInt(state("coordIndex")) + 1
        state("coordIndex") = idx2
        Dim coordSet2 As GraphicsCoordinateSet = dataSets.CreateCoordinateSet(idx2)
        Dim coords2(5) As Double
        coords2(0) = x : coords2(1) = y - size : coords2(2) = z
        coords2(3) = x : coords2(4) = y + size : coords2(5) = z
        coordSet2.PutCoordinates(coords2)
        
        Dim ls2 As LineStripGraphics = node.AddLineStripGraphics()
        ls2.CoordinateSet = coordSet2
        ls2.LineWeight = 3
        If colorSet IsNot Nothing Then ls2.ColorSet = colorSet
        
        Logger.Info("Test2b: Cross added at (3,0,0) GREEN")
        
        app.ActiveView.Update()
        Logger.Info("Test2b: View updated - look for cross at (3,0,0)")
        
    Catch ex As Exception
        Logger.Error("Test2b FAILED: " & ex.Message)
    End Try
End Sub

Sub Test2c_AddLineState(state As Dictionary(Of String, Object))
    Logger.Info("")
    Logger.Info("--- TEST 2c: Add Line Using State ---")
    
    Dim app As Inventor.Application = CType(state("app"), Inventor.Application)
    
    If Not state.ContainsKey("dataSets") OrElse Not state.ContainsKey("graphicsNode") Then
        Logger.Error("Test2c: State not set up! Click 'Setup Graphics' first")
        Return
    End If
    
    Dim dataSets As GraphicsDataSets = CType(state("dataSets"), GraphicsDataSets)
    Dim node As GraphicsNode = CType(state("graphicsNode"), GraphicsNode)
    Dim colorSet As GraphicsColorSet = If(state.ContainsKey("colorSet"), CType(state("colorSet"), GraphicsColorSet), Nothing)
    
    Try
        Dim idx As Integer = CInt(state("coordIndex")) + 1
        state("coordIndex") = idx
        
        Dim coordSet As GraphicsCoordinateSet = dataSets.CreateCoordinateSet(idx)
        Logger.Info("Test2c: CoordSet created with index " & idx.ToString())
        
        Dim coords(5) As Double
        coords(0) = 0 : coords(1) = 0 : coords(2) = 0
        coords(3) = 5 : coords(4) = 5 : coords(5) = 0
        coordSet.PutCoordinates(coords)
        Logger.Info("Test2c: Coords set for line (0,0,0) to (5,5,0)")
        
        Dim ls As LineStripGraphics = node.AddLineStripGraphics()
        ls.CoordinateSet = coordSet
        ls.LineWeight = 5
        If colorSet IsNot Nothing Then ls.ColorSet = colorSet
        Logger.Info("Test2c: LineStripGraphics added (GREEN)")
        
        app.ActiveView.Update()
        Logger.Info("Test2c: View updated - look for line from origin to (5,5,0)")
        
    Catch ex As Exception
        Logger.Error("Test2c FAILED: " & ex.Message)
    End Try
End Sub

' ============================================================
' TEST 3: Fresh creation each button click
' ============================================================
Sub Test3_FreshCreation(state As Dictionary(Of String, Object))
    Logger.Info("")
    Logger.Info("--- TEST 3: Fresh Creation Each Time ---")
    
    Dim app As Inventor.Application = CType(state("app"), Inventor.Application)
    Dim doc As Document = CType(state("doc"), Document)
    Dim compDef As ComponentDefinition = CType(state("compDef"), ComponentDefinition)
    
    Dim counter As Integer = CInt(state("test3Counter")) + 1
    state("test3Counter") = counter
    
    Dim ID As String = "_Test3_" & counter.ToString()
    
    Try
        ' Try to cleanup existing with same ID first
        Try : compDef.ClientGraphicsCollection.Item(ID).Delete() : Catch : End Try
        Try : doc.GraphicsDataSetsCollection.Item(ID).Delete() : Catch : End Try
        
        Dim dataSets As GraphicsDataSets = doc.GraphicsDataSetsCollection.Add(ID)
        Logger.Info("Test3: DataSets created with ID=" & ID)
        
        Dim coordSet As GraphicsCoordinateSet = dataSets.CreateCoordinateSet(1)
        
        ' Create color set (yellow)
        Dim colorSet As GraphicsColorSet = dataSets.CreateColorSet(2)
        colorSet.Add(1, 255, 255, 0)
        
        ' Create a small line strip (more visible than point)
        ' Position based on counter, draw a small "+" shape
        Dim x As Double = counter * 3
        Dim coords(5) As Double  ' 2 points for a line
        coords(0) = x - 1 : coords(1) = 0 : coords(2) = 0
        coords(3) = x + 1 : coords(4) = 0 : coords(5) = 0
        coordSet.PutCoordinates(coords)
        Logger.Info("Test3: Line at X=" & x.ToString())
        
        Dim cg As ClientGraphics = compDef.ClientGraphicsCollection.Add(ID)
        Dim node As GraphicsNode = cg.AddNode(1)
        
        Dim ls As LineStripGraphics = node.AddLineStripGraphics()
        ls.CoordinateSet = coordSet
        ls.ColorSet = colorSet
        ls.LineWeight = 3
        
        app.ActiveView.Update()
        Logger.Info("Test3: SUCCESS - small line at X=" & x.ToString())
        
    Catch ex As Exception
        Logger.Error("Test3 FAILED: " & ex.Message)
    End Try
End Sub

' ============================================================
' TEST 4: Add to existing Test1 graphics
' ============================================================
Sub Test4_AddToExisting(state As Dictionary(Of String, Object))
    Logger.Info("")
    Logger.Info("--- TEST 4: Add to Existing Test1 Graphics ---")
    
    Dim app As Inventor.Application = CType(state("app"), Inventor.Application)
    Dim doc As Document = CType(state("doc"), Document)
    Dim compDef As ComponentDefinition = CType(state("compDef"), ComponentDefinition)
    
    Const ID As String = "_Test1_"
    
    Try
        ' Try to get existing
        Dim dataSets As GraphicsDataSets = Nothing
        Dim cg As ClientGraphics = Nothing
        
        Try
            dataSets = doc.GraphicsDataSetsCollection.Item(ID)
            cg = compDef.ClientGraphicsCollection.Item(ID)
        Catch
            Logger.Error("Test4: Test1 graphics don't exist. Run Test1 first.")
            Return
        End Try
        
        Logger.Info("Test4: Found existing Test1 graphics")
        
        ' Get the node (assuming it's node 1)
        Dim node As GraphicsNode = cg.Item(1)
        Logger.Info("Test4: Got node from existing graphics")
        
        ' Add a new point
        Dim idx As Integer = dataSets.Count + 1
        Dim coordSet As GraphicsCoordinateSet = dataSets.CreateCoordinateSet(idx)
        
        Dim coords(2) As Double
        coords(0) = -3 : coords(1) = 0 : coords(2) = 0  ' At (-3,0,0)
        coordSet.PutCoordinates(coords)
        
        Dim pg As PointGraphics = node.AddPointGraphics()
        pg.CoordinateSet = coordSet
        pg.PointRenderStyle = PointRenderStyleEnum.kCrossPointStyle
        
        app.ActiveView.Update()
        Logger.Info("Test4: SUCCESS - cross at (-3,0,0)")
        
    Catch ex As Exception
        Logger.Error("Test4 FAILED: " & ex.Message)
    End Try
End Sub

' ============================================================
' Cleanup
' ============================================================
Sub CleanupAllGraphics(state As Dictionary(Of String, Object))
    Logger.Info("")
    Logger.Info("--- Cleanup All Graphics ---")
    
    Dim app As Inventor.Application = CType(state("app"), Inventor.Application)
    Dim doc As Document = CType(state("doc"), Document)
    Dim compDef As ComponentDefinition = CType(state("compDef"), ComponentDefinition)
    
    ' Clear state
    state.Remove("dataSets")
    state.Remove("clientGraphics")
    state.Remove("graphicsNode")
    state("coordIndex") = 0
    state("test3Counter") = 0
    
    Dim cleaned As Integer = 0
    
    ' Delete by known IDs
    Dim idsToDelete As New List(Of String)
    idsToDelete.Add("_Test1_")
    idsToDelete.Add("_Test2_")
    idsToDelete.Add("_TestCG_")
    idsToDelete.Add("_DebugCG_")
    For i As Integer = 1 To 20
        idsToDelete.Add("_Test3_" & i.ToString())
    Next
    
    For Each id As String In idsToDelete
        ' Delete ClientGraphics first
        Try
            compDef.ClientGraphicsCollection.Item(id).Delete()
            cleaned += 1
            Logger.Info("Cleanup: Deleted CG '" & id & "'")
        Catch
            ' Doesn't exist, that's OK
        End Try
        
        ' Then delete DataSets
        Try
            doc.GraphicsDataSetsCollection.Item(id).Delete()
            cleaned += 1
            Logger.Info("Cleanup: Deleted DS '" & id & "'")
        Catch
            ' Doesn't exist, that's OK
        End Try
    Next
    
    app.ActiveView.Update()
    Logger.Info("Cleanup: Total deleted = " & cleaned.ToString())
End Sub
