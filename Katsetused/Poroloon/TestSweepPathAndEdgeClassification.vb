' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' POC-2: Edge Classification with User Selection
'
' User selects the two end faces (profile cross-sections), then the script
' classifies all edges based on vertex membership:
' - Transverse: all vertices on one end face
' - Longitudinal: connects a start-face vertex to an end-face vertex
' - Side: everything else
'
' Run on: active part document (foam part with sweep feature)
' Output: iLogic log window with classified edge measurements
' ============================================================================

AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"

Imports Inventor
Imports System.Collections.Generic
Imports System.Windows.Forms

Sub Main()
    UtilsLib.SetLogger(Logger)
    
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("POC-2: Ava detaili dokument (.ipt)")
        MessageBox.Show("Ava detaili dokument (.ipt)", "POC-2")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    Dim tg As TransientGeometry = app.TransientGeometry
    
    Logger.Info("POC-2: === Edge Classification with User Selection ===")
    Logger.Info("POC-2: Part: " & partDoc.DisplayName)
    
    ' Find first solid body for reference
    Dim body As SurfaceBody = Nothing
    For Each b As SurfaceBody In compDef.SurfaceBodies
        Try
            If b.Volume(0.01) > 0 Then
                body = b
                Exit For
            End If
        Catch
        End Try
    Next
    
    If body Is Nothing Then
        Logger.Error("POC-2: No solid body found")
        MessageBox.Show("Detailis pole ühtegi tahket keha.", "POC-2")
        Exit Sub
    End If
    
    Logger.Info("POC-2: Body: '" & body.Name & "' | Faces: " & body.Faces.Count & " | Edges: " & body.Edges.Count)
    
    ' ========================================================================
    ' User selects the two end faces
    ' ========================================================================
    Dim startFace As Face = Nothing
    Dim endFace As Face = Nothing
    
    ' Pick first end face
    Logger.Info("POC-2: Waiting for user to select first end face...")
    Try
        Dim picked As Object = app.CommandManager.Pick( _
            SelectionFilterEnum.kPartFaceFilter, _
            "Vali ESIMENE otspind (profiili ristlõige) - ESC tühistamiseks")
        If picked IsNot Nothing AndAlso TypeOf picked Is Face Then
            startFace = CType(picked, Face)
        End If
    Catch
        Logger.Info("POC-2: User cancelled first face selection")
        Exit Sub
    End Try
    
    If startFace Is Nothing Then
        Logger.Info("POC-2: No first face selected")
        Exit Sub
    End If
    
    Dim startGeomType As String = GetFaceGeometryType(startFace)
    Dim startEdgeCount As Integer = 0
    Try : startEdgeCount = startFace.EdgeLoops.Item(1).Edges.Count : Catch : End Try
    Logger.Info("POC-2: First end face selected: " & startGeomType & " with " & startEdgeCount & " edges")
    
    ' Pick second end face
    Logger.Info("POC-2: Waiting for user to select second end face...")
    Try
        Dim picked As Object = app.CommandManager.Pick( _
            SelectionFilterEnum.kPartFaceFilter, _
            "Vali TEINE otspind (profiili ristlõige) - ESC tühistamiseks")
        If picked IsNot Nothing AndAlso TypeOf picked Is Face Then
            endFace = CType(picked, Face)
        End If
    Catch
        Logger.Info("POC-2: User cancelled second face selection")
        Exit Sub
    End Try
    
    If endFace Is Nothing Then
        Logger.Info("POC-2: No second face selected")
        Exit Sub
    End If
    
    Dim endGeomType As String = GetFaceGeometryType(endFace)
    Dim endEdgeCount As Integer = 0
    Try : endEdgeCount = endFace.EdgeLoops.Item(1).Edges.Count : Catch : End Try
    Logger.Info("POC-2: Second end face selected: " & endGeomType & " with " & endEdgeCount & " edges")
    
    ' ========================================================================
    ' Classify edges based on selected faces
    ' ========================================================================
    Logger.Info("POC-2: --- Classifying edges ---")
    
    ' Get vertices on each end face
    Dim startVertices As New HashSet(Of String)
    Dim endVertices As New HashSet(Of String)
    
    For Each el As EdgeLoop In startFace.EdgeLoops
        For Each e As Edge In el.Edges
            If e.StartVertex IsNot Nothing Then startVertices.Add(VertexKey(e.StartVertex))
            If e.StopVertex IsNot Nothing Then startVertices.Add(VertexKey(e.StopVertex))
        Next
    Next
    For Each el As EdgeLoop In endFace.EdgeLoops
        For Each e As Edge In el.Edges
            If e.StartVertex IsNot Nothing Then endVertices.Add(VertexKey(e.StartVertex))
            If e.StopVertex IsNot Nothing Then endVertices.Add(VertexKey(e.StopVertex))
        Next
    Next
    
    Logger.Info("POC-2:   Start face vertices: " & startVertices.Count)
    Logger.Info("POC-2:   End face vertices: " & endVertices.Count)
    
    ' Classify all edges on the body
    Dim longitudinal As New List(Of Edge)
    Dim transverseStart As New List(Of Edge)
    Dim transverseEnd As New List(Of Edge)
    Dim sideEdges As New List(Of Edge)
    
    For Each edge As Edge In body.Edges
        Dim sv As String = If(edge.StartVertex IsNot Nothing, VertexKey(edge.StartVertex), "")
        Dim ev As String = If(edge.StopVertex IsNot Nothing, VertexKey(edge.StopVertex), "")
        
        Dim startOnStart As Boolean = startVertices.Contains(sv)
        Dim startOnEnd As Boolean = endVertices.Contains(sv)
        Dim endOnStart As Boolean = startVertices.Contains(ev)
        Dim endOnEnd As Boolean = endVertices.Contains(ev)
        
        If (startOnStart AndAlso endOnEnd) OrElse (startOnEnd AndAlso endOnStart) Then
            longitudinal.Add(edge)
        ElseIf (startOnStart AndAlso endOnStart) Then
            transverseStart.Add(edge)
        ElseIf (startOnEnd AndAlso endOnEnd) Then
            transverseEnd.Add(edge)
        Else
            sideEdges.Add(edge)
        End If
    Next
    
    Logger.Info("POC-2: --- Classification Results ---")
    Logger.Info("POC-2:   Longitudinal: " & longitudinal.Count)
    Logger.Info("POC-2:   Transverse (start): " & transverseStart.Count)
    Logger.Info("POC-2:   Transverse (end): " & transverseEnd.Count)
    Logger.Info("POC-2:   Side/other: " & sideEdges.Count)
    
    ' Measure longitudinal edges
    If longitudinal.Count > 0 Then
        Logger.Info("POC-2: --- Longitudinal Edge Measurements ---")
        Dim minLen As Double = Double.MaxValue
        Dim maxLen As Double = Double.MinValue
        Dim sumLen As Double = 0
        
        For i As Integer = 0 To longitudinal.Count - 1
            Dim edge As Edge = longitudinal(i)
            Try
                Dim arcLen As Double = MeasureEdgeArcLength(edge)
                Dim chordLen As Double = MeasureEdgeChordLength(edge, tg)
                Dim geomType As String = GetEdgeGeometryType(edge)
                Dim ratio As Double = If(chordLen > 0.0001, arcLen / chordLen, 0)
                
                If arcLen < minLen Then minLen = arcLen
                If arcLen > maxLen Then maxLen = arcLen
                sumLen += arcLen
                
                Logger.Info("POC-2:     Long-" & (i + 1) & ": " & geomType & _
                            " | Arc=" & FormatMm(arcLen) & "mm" & _
                            " | Chord=" & FormatMm(chordLen) & "mm" & _
                            " | Ratio=" & ratio.ToString("0.000"))
            Catch ex As Exception
                Logger.Warn("POC-2:     Long-" & (i + 1) & ": FAILED - " & ex.Message)
            End Try
        Next
        
        Dim avgLen As Double = sumLen / longitudinal.Count
        Logger.Info("POC-2: --- Longitudinal Summary ---")
        Logger.Info("POC-2:   Min: " & FormatMm(minLen) & "mm")
        Logger.Info("POC-2:   Max: " & FormatMm(maxLen) & "mm")
        Logger.Info("POC-2:   Avg: " & FormatMm(avgLen) & "mm")
        Logger.Info("POC-2:   Diff (max-min): " & FormatMm(maxLen - minLen) & "mm (thickness effect on arc length)")
    Else
        Logger.Warn("POC-2: No longitudinal edges found - check face selection")
    End If
    
    ' Log transverse edges for reference
    If transverseStart.Count > 0 Then
        Logger.Info("POC-2: --- Transverse (Start Face) Edge Measurements ---")
        For i As Integer = 0 To Math.Min(transverseStart.Count - 1, 7)
            Try
                Dim arcLen As Double = MeasureEdgeArcLength(transverseStart(i))
                Logger.Info("POC-2:     Trans-S" & (i + 1) & ": " & GetEdgeGeometryType(transverseStart(i)) & _
                            " | Arc=" & FormatMm(arcLen) & "mm")
            Catch
            End Try
        Next
    End If
    
    Logger.Info("POC-2: === Done ===")
End Sub

Function MeasureEdgeArcLength(edge As Edge) As Double
    Dim evaluator As CurveEvaluator = edge.Evaluator
    Dim minParam As Double = 0, maxParam As Double = 0
    Call evaluator.GetParamExtents(minParam, maxParam)
    Dim arcLength As Double = 0
    Call evaluator.GetLengthAtParam(minParam, maxParam, arcLength)
    Return arcLength
End Function

Function MeasureEdgeChordLength(edge As Edge, tg As TransientGeometry) As Double
    If edge.StartVertex Is Nothing OrElse edge.StopVertex Is Nothing Then Return 0
    Dim p1 As Point = edge.StartVertex.Point
    Dim p2 As Point = edge.StopVertex.Point
    Dim dx As Double = p2.X - p1.X
    Dim dy As Double = p2.Y - p1.Y
    Dim dz As Double = p2.Z - p1.Z
    Return Math.Sqrt(dx * dx + dy * dy + dz * dz)
End Function

Function GetEdgeGeometryType(edge As Edge) As String
    Try
        Dim geom As Object = edge.Geometry
        If TypeOf geom Is Line Then Return "Line"
        If TypeOf geom Is LineSegment Then Return "LineSegment"
        If TypeOf geom Is Arc3d Then Return "Arc3d"
        If TypeOf geom Is Circle Then Return "Circle"
        If TypeOf geom Is BSplineCurve Then Return "BSpline"
        If TypeOf geom Is EllipticalArc Then Return "EllipticalArc"
        Return geom.GetType().Name
    Catch
        Return "Unknown"
    End Try
End Function

Function GetFaceGeometryType(face As Face) As String
    Try
        Dim geom As Object = face.Geometry
        If TypeOf geom Is Plane Then Return "Plane"
        If TypeOf geom Is Cylinder Then Return "Cylinder"
        If TypeOf geom Is Cone Then Return "Cone"
        If TypeOf geom Is Sphere Then Return "Sphere"
        If TypeOf geom Is Torus Then Return "Torus"
        If TypeOf geom Is BSplineSurface Then Return "BSpline"
        Return geom.GetType().Name
    Catch
        Return "Unknown"
    End Try
End Function

Function VertexKey(v As Vertex) As String
    If v Is Nothing Then Return ""
    Dim p As Point = v.Point
    Return Math.Round(p.X, 6) & "," & Math.Round(p.Y, 6) & "," & Math.Round(p.Z, 6)
End Function

Function FormatMm(valueCm As Double) As String
    Return (valueCm * 10).ToString("0.00")
End Function
