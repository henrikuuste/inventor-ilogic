' ViewportHelperLib.vb - Viewport Helper Display Library
' Provides highlight management, transient graphics, and preview work features
'
' TRANSIENT MARKERS (ClientGraphics):
' - AddPointMarkerAt(x, y, z) - Add a point marker at coordinates (cm)
' - AddLineMarkerAt(x1, y1, z1, x2, y2, z2) - Add a line marker (cm)
' - ClearMarkers() - Remove all transient markers
'
' PREVIEW WORK FEATURES:
' - CreatePreviewWorkPoint() - Add a work point that can be deleted/committed
' - DeletePreviewFeatures() - Remove all preview features
' - CommitPreviewFeatures() - Keep features, stop tracking
'
' Note: Uses Object types for Inventor API types (late binding)
' because library modules don't have direct access to Inventor types.
'
' KNOWN LIMITATIONS:
' - HighlightSet may clear when selection changes (Inventor behavior)

Public Module ViewportHelperLib
    
    ' ============================================================
    ' HIGHLIGHT MANAGEMENT
    ' ============================================================
    
    Private m_HighlightSet As Object  ' HighlightSet
    Private m_App As Object           ' Inventor.Application
    
    ''' <summary>
    ''' Initializes the viewport helper with the application instance.
    ''' Call once at rule start.
    ''' </summary>
    Public Sub Initialize(app As Object)
        m_App = app
        ClearHighlights()
    End Sub
    
    ''' <summary>
    ''' Highlights an object in the viewport.
    ''' </summary>
    Public Sub Highlight(obj As Object)
        If m_App Is Nothing Then Return
        If obj Is Nothing Then Return
        
        Try
            If m_HighlightSet Is Nothing Then
                m_HighlightSet = m_App.ActiveDocument.HighlightSets.Add()
            End If
            m_HighlightSet.AddItem(obj)
            m_App.ActiveView.Update()
        Catch
            ' Object may not support highlighting
        End Try
    End Sub
    
    ''' <summary>
    ''' Highlights multiple objects.
    ''' </summary>
    Public Sub HighlightMany(objects As IEnumerable)
        If objects Is Nothing Then Return
        For Each obj In objects
            Highlight(obj)
        Next
    End Sub
    
    ''' <summary>
    ''' Clears all highlights.
    ''' </summary>
    Public Sub ClearHighlights()
        Try
            If m_HighlightSet IsNot Nothing Then
                m_HighlightSet.Clear()
                If m_App IsNot Nothing Then
                    m_App.ActiveView.Update()
                End If
            End If
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Removes the highlight set entirely.
    ''' Call on rule exit.
    ''' </summary>
    Public Sub Cleanup()
        Try
            If m_HighlightSet IsNot Nothing Then
                m_HighlightSet.Delete()
                m_HighlightSet = Nothing
            End If
        Catch
        End Try
        
        CleanupClientGraphics()
        DeletePreviewFeatures()
    End Sub
    
    ' ============================================================
    ' CLIENT GRAPHICS - TRANSIENT MARKERS
    ' Uses ComponentDefinition.ClientGraphicsCollection (works in iLogic)
    ' Uses LineStripGraphics for visibility (PointGraphics can be hard to see)
    ' 
    ' NOTE: Module variables don't persist across button clicks in non-modal dialogs
    ' when using AddVbFile libraries. So we re-fetch by ID each time.
    ' ============================================================
    
    Private Const GRAPHICS_ID As String = "_ViewportHelper_CG_"
    
    ''' <summary>
    ''' Gets or creates the ClientGraphics infrastructure.
    ''' Returns tuple: (GraphicsDataSets, GraphicsNode, nextIndex, app)
    ''' Returns Nothing values if setup fails.
    ''' app parameter required because module variables don't persist in AddVbFile libraries.
    ''' </summary>
    Private Function GetOrCreateGraphics(app As Object) As Object()
        If app Is Nothing Then Return {Nothing, Nothing, 0, Nothing}
        
        Dim result(3) As Object  ' {dataSets, node, nextIndex, app}
        
        Try
            Dim doc As Object = app.ActiveDocument
            If doc Is Nothing Then Return {Nothing, Nothing, 0, Nothing}
            
            Dim compDef As Object = doc.ComponentDefinition
            If compDef Is Nothing Then Return {Nothing, Nothing, 0, Nothing}
            
            Dim dataSets As Object = Nothing
            Dim cg As Object = Nothing
            Dim node As Object = Nothing
            Dim nextIndex As Integer = 1
            
            ' Try to get existing
            Try
                dataSets = doc.GraphicsDataSetsCollection.Item(GRAPHICS_ID)
                cg = compDef.ClientGraphicsCollection.Item(GRAPHICS_ID)
                node = cg.Item(1)  ' Node 1
                nextIndex = dataSets.Count + 1
            Catch
                ' Doesn't exist, create new
                dataSets = doc.GraphicsDataSetsCollection.Add(GRAPHICS_ID)
                cg = compDef.ClientGraphicsCollection.Add(GRAPHICS_ID)
                node = cg.AddNode(1)
                nextIndex = 1
            End Try
            
            result(0) = dataSets
            result(1) = node
            result(2) = nextIndex
            result(3) = app
            Return result
        Catch
            Return {Nothing, Nothing, 0, Nothing}
        End Try
    End Function
    
    ''' <summary>
    ''' Creates a color set with the specified RGB values.
    ''' </summary>
    Private Function CreateColorSet(dataSets As Object, index As Integer, r As Integer, g As Integer, b As Integer) As Object
        Try
            Dim colorSet As Object = dataSets.CreateColorSet(index)
            colorSet.Add(1, r, g, b)
            Return colorSet
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Adds a transient point marker at the specified location.
    ''' Uses a small cross shape made of LineStripGraphics for visibility.
    ''' app: Inventor.Application object (required - module vars don't persist)
    ''' point should be an Inventor.Point object.
    ''' Optional colorPreset: "red", "green", "blue", "yellow", "cyan", "magenta", "orange", "white"
    ''' </summary>
    Public Sub AddPointMarker(app As Object, point As Object, Optional colorPreset As String = "cyan")
        AddPointMarkerAt(app, point.X, point.Y, point.Z, colorPreset)
    End Sub
    
    ''' <summary>
    ''' Adds a transient point marker at the specified coordinates (in cm).
    ''' Draws a 3D cross for visibility (lines in X, Y, and Z directions).
    ''' app: Inventor.Application object (required - module vars don't persist)
    ''' Optional colorPreset: "red", "green", "blue", "yellow", "cyan", "magenta", "orange", "white"
    ''' </summary>
    Public Sub AddPointMarkerAt(app As Object, x As Double, y As Double, z As Double, Optional colorPreset As String = "cyan")
        Dim gfx() As Object = GetOrCreateGraphics(app)
        Dim dataSets As Object = gfx(0)
        Dim node As Object = gfx(1)
        Dim idx As Integer = CInt(gfx(2))
        
        If dataSets Is Nothing OrElse node Is Nothing Then Return
        
        Try
            ' Get color RGB
            Dim r As Integer, g As Integer, b As Integer
            GetColorRGB(colorPreset, r, g, b)
            
            ' Create color set
            Dim colorSet As Object = CreateColorSet(dataSets, idx, r, g, b)
            idx += 1
            
            Dim markerSize As Double = 0.5  ' 5mm cross
            
            ' X direction line
            Dim coordSet1 As Object = dataSets.CreateCoordinateSet(idx) : idx += 1
            Dim coords1(5) As Double
            coords1(0) = x - markerSize : coords1(1) = y : coords1(2) = z
            coords1(3) = x + markerSize : coords1(4) = y : coords1(5) = z
            coordSet1.PutCoordinates(coords1)
            
            Dim ls1 As Object = node.AddLineStripGraphics()
            ls1.CoordinateSet = coordSet1
            ls1.LineWeight = 3
            If colorSet IsNot Nothing Then ls1.ColorSet = colorSet
            
            ' Y direction line
            Dim coordSet2 As Object = dataSets.CreateCoordinateSet(idx) : idx += 1
            Dim coords2(5) As Double
            coords2(0) = x : coords2(1) = y - markerSize : coords2(2) = z
            coords2(3) = x : coords2(4) = y + markerSize : coords2(5) = z
            coordSet2.PutCoordinates(coords2)
            
            Dim ls2 As Object = node.AddLineStripGraphics()
            ls2.CoordinateSet = coordSet2
            ls2.LineWeight = 3
            If colorSet IsNot Nothing Then ls2.ColorSet = colorSet
            
            ' Z direction line
            Dim coordSet3 As Object = dataSets.CreateCoordinateSet(idx) : idx += 1
            Dim coords3(5) As Double
            coords3(0) = x : coords3(1) = y : coords3(2) = z - markerSize
            coords3(3) = x : coords3(4) = y : coords3(5) = z + markerSize
            coordSet3.PutCoordinates(coords3)
            
            Dim ls3 As Object = node.AddLineStripGraphics()
            ls3.CoordinateSet = coordSet3
            ls3.LineWeight = 3
            If colorSet IsNot Nothing Then ls3.ColorSet = colorSet
            
            app.ActiveView.Update()
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Adds a transient line between two points.
    ''' app: Inventor.Application object (required - module vars don't persist)
    ''' startPoint and endPoint should be Inventor.Point objects.
    ''' Optional colorPreset: "red", "green", "blue", "yellow", "cyan", "magenta", "orange", "white"
    ''' </summary>
    Public Sub AddLineMarker(app As Object, startPoint As Object, endPoint As Object, Optional colorPreset As String = "cyan")
        AddLineMarkerAt(app, startPoint.X, startPoint.Y, startPoint.Z, _
                        endPoint.X, endPoint.Y, endPoint.Z, colorPreset)
    End Sub
    
    ''' <summary>
    ''' Adds a transient line between two coordinate sets (in cm).
    ''' app: Inventor.Application object (required - module vars don't persist)
    ''' Optional colorPreset: "red", "green", "blue", "yellow", "cyan", "magenta", "orange", "white"
    ''' </summary>
    Public Sub AddLineMarkerAt(app As Object, x1 As Double, y1 As Double, z1 As Double, _
                               x2 As Double, y2 As Double, z2 As Double, _
                               Optional colorPreset As String = "cyan")
        Dim gfx() As Object = GetOrCreateGraphics(app)
        Dim dataSets As Object = gfx(0)
        Dim node As Object = gfx(1)
        Dim idx As Integer = CInt(gfx(2))
        
        If dataSets Is Nothing OrElse node Is Nothing Then Return
        
        Try
            ' Get color RGB
            Dim r As Integer, g As Integer, b As Integer
            GetColorRGB(colorPreset, r, g, b)
            
            ' Create color set
            Dim colorSet As Object = CreateColorSet(dataSets, idx, r, g, b)
            idx += 1
            
            ' Create coordinate set
            Dim coordSet As Object = dataSets.CreateCoordinateSet(idx)
            
            ' Put coordinates (2 points = 1 line)
            Dim coords(5) As Double
            coords(0) = x1 : coords(1) = y1 : coords(2) = z1
            coords(3) = x2 : coords(4) = y2 : coords(5) = z2
            coordSet.PutCoordinates(coords)
            
            ' Create line strip graphics (supports ColorSet)
            Dim lineStrip As Object = node.AddLineStripGraphics()
            lineStrip.CoordinateSet = coordSet
            lineStrip.LineWeight = 3
            If colorSet IsNot Nothing Then lineStrip.ColorSet = colorSet
            
            app.ActiveView.Update()
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Gets RGB values for a color preset name.
    ''' </summary>
    Private Sub GetColorRGB(colorPreset As String, ByRef r As Integer, ByRef g As Integer, ByRef b As Integer)
        Select Case colorPreset.ToLower()
            Case "red" : r = 255 : g = 50 : b = 50
            Case "green" : r = 50 : g = 255 : b = 50
            Case "blue" : r = 50 : g = 150 : b = 255
            Case "yellow" : r = 255 : g = 255 : b = 0
            Case "cyan" : r = 0 : g = 255 : b = 255
            Case "magenta" : r = 255 : g = 0 : b = 255
            Case "orange" : r = 255 : g = 165 : b = 0
            Case "white" : r = 255 : g = 255 : b = 255
            Case Else : r = 0 : g = 255 : b = 255  ' Default cyan
        End Select
    End Sub
    
    ''' <summary>
    ''' Clears all transient markers.
    ''' app: Inventor.Application object (required - module vars don't persist)
    ''' </summary>
    Public Sub ClearMarkers(app As Object)
        Try
            If app Is Nothing Then Return
            Dim doc As Object = app.ActiveDocument
            If doc Is Nothing Then Return
            Dim compDef As Object = doc.ComponentDefinition
            
            ' Delete by ID (more reliable)
            Try : compDef.ClientGraphicsCollection.Item(GRAPHICS_ID).Delete() : Catch : End Try
            Try : doc.GraphicsDataSetsCollection.Item(GRAPHICS_ID).Delete() : Catch : End Try
            
            app.ActiveView.Update()
        Catch
        End Try
    End Sub
    
    Private Sub CleanupClientGraphics()
        ClearMarkers(m_App)
    End Sub
    
    ' ============================================================
    ' REAL FEATURE PREVIEW (UCS, Work Planes)
    ' ============================================================
    
    Private m_PreviewFeatures As New List(Of Object)
    Private m_PreviewCounter As Integer = 0
    
    ''' <summary>
    ''' Generates a unique preview name.
    ''' </summary>
    Private Function GetUniquePreviewName(baseName As String) As String
        m_PreviewCounter += 1
        Return baseName & "_" & m_PreviewCounter.ToString()
    End Function
    
    ''' <summary>
    ''' Creates a preview UCS that can be updated and deleted.
    ''' asmDoc should be an AssemblyDocument, matrix should be a Matrix.
    ''' Returns a UserCoordinateSystem object.
    ''' </summary>
    Public Function CreatePreviewUCS(asmDoc As Object, name As String, matrix As Object) As Object
        Try
            Dim ucs As Object = asmDoc.ComponentDefinition.UserCoordinateSystems.Add(matrix)
            ucs.Name = GetUniquePreviewName(name)
            m_PreviewFeatures.Add(ucs)
            m_App.ActiveView.Update()
            Return ucs
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Updates a preview UCS position/orientation.
    ''' ucs should be a UserCoordinateSystem, matrix should be a Matrix.
    ''' </summary>
    Public Sub UpdatePreviewUCS(ucs As Object, matrix As Object)
        Try
            ucs.Transformation = matrix
            m_App.ActiveView.Update()
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Creates a temporary work plane for preview.
    ''' Returns a WorkPlane object.
    ''' </summary>
    Public Function CreatePreviewWorkPlane(compDef As Object, planeInput As Object, name As String) As Object
        Try
            Dim wp As Object = compDef.WorkPlanes.AddByPlaneAndOffset(planeInput, 0)
            wp.Name = GetUniquePreviewName(name)
            wp.Visible = True
            m_PreviewFeatures.Add(wp)
            m_App.ActiveView.Update()
            Return wp
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Creates a temporary work point for preview.
    ''' point should be an Inventor.Point object.
    ''' Returns a WorkPoint object.
    ''' </summary>
    Public Function CreatePreviewWorkPoint(compDef As Object, point As Object, name As String) As Object
        Try
            Dim wp As Object = compDef.WorkPoints.AddFixed(point)
            wp.Name = GetUniquePreviewName(name)
            wp.Visible = True
            m_PreviewFeatures.Add(wp)
            m_App.ActiveView.Update()
            Return wp
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Adds a preview work point at the specified coordinates (in cm).
    ''' Convenient wrapper that creates the point internally.
    ''' Works reliably with non-modal forms.
    ''' </summary>
    Public Function AddPreviewPointAt(x As Double, y As Double, z As Double) As Object
        If m_App Is Nothing Then Return Nothing
        
        Try
            Dim doc As Object = m_App.ActiveDocument
            Dim compDef As Object = doc.ComponentDefinition
            Dim tg As Object = m_App.TransientGeometry
            Dim pt As Object = tg.CreatePoint(x, y, z)
            Return CreatePreviewWorkPoint(compDef, pt, "_Marker_")
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Adds a preview work axis (line) between two points.
    ''' Uses a construction work axis. Coordinates in cm.
    ''' Works reliably with non-modal forms.
    ''' </summary>
    Public Function AddPreviewLineAt(x1 As Double, y1 As Double, z1 As Double, _
                                     x2 As Double, y2 As Double, z2 As Double) As Object
        If m_App Is Nothing Then Return Nothing
        
        Try
            Dim doc As Object = m_App.ActiveDocument
            Dim compDef As Object = doc.ComponentDefinition
            Dim tg As Object = m_App.TransientGeometry
            
            Dim pt1 As Object = tg.CreatePoint(x1, y1, z1)
            Dim pt2 As Object = tg.CreatePoint(x2, y2, z2)
            Dim line As Object = tg.CreateLine(pt1, tg.CreateUnitVector(x2 - x1, y2 - y1, z2 - z1))
            
            Dim wa As Object = compDef.WorkAxes.AddFixed(line)
            wa.Name = GetUniquePreviewName("_LineMarker_")
            wa.Visible = True
            m_PreviewFeatures.Add(wa)
            m_App.ActiveView.Update()
            Return wa
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Returns the number of tracked preview features.
    ''' </summary>
    Public Function GetPreviewFeatureCount() As Integer
        Return m_PreviewFeatures.Count
    End Function
    
    ''' <summary>
    ''' Deletes all preview features (UCS, work planes, etc.).
    ''' Call on cancel or cleanup.
    ''' </summary>
    Public Sub DeletePreviewFeatures()
        For Each feature In m_PreviewFeatures.ToArray()
            Try
                feature.Delete()
            Catch
            End Try
        Next
        m_PreviewFeatures.Clear()
        
        If m_App IsNot Nothing Then
            Try
                m_App.ActiveView.Update()
            Catch
            End Try
        End If
    End Sub
    
    ''' <summary>
    ''' Commits preview features (keeps them, removes from cleanup list).
    ''' Features remain in model but are no longer tracked for deletion.
    ''' Call on OK/confirm.
    ''' </summary>
    Public Sub CommitPreviewFeatures()
        m_PreviewFeatures.Clear()
        ' Note: features remain in model, just no longer tracked
    End Sub
    
    ''' <summary>
    ''' Renames a preview feature (useful when committing).
    ''' </summary>
    Public Sub RenamePreviewFeature(feature As Object, newName As String)
        Try
            feature.Name = newName
        Catch
        End Try
    End Sub
    
    ' ============================================================
    ' VIEW UPDATES
    ' ============================================================
    
    ''' <summary>
    ''' Forces a viewport refresh.
    ''' </summary>
    Public Sub RefreshView()
        Try
            If m_App IsNot Nothing Then
                m_App.ActiveView.Update()
            End If
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Checks if the library has been initialized.
    ''' </summary>
    Public Function IsInitialized() As Boolean
        Return m_App IsNot Nothing
    End Function
    
End Module
