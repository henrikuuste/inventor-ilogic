' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' POC-1: Edge Arc Length Measurement
'
' Validates CurveEvaluator.GetLengthAtParam on foam sweep parts.
' Iterates all edges on all bodies, measures arc length, logs geometry type,
' and compares arc length vs. chord length.
'
' Run on: active part document (foam part with sweep feature)
' Output: iLogic log window with per-edge measurements
' ============================================================================

AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"

Imports Inventor
Imports System.Collections.Generic

Sub Main()
    UtilsLib.SetLogger(Logger)
    
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("POC-1: Ava detaili dokument (.ipt)")
        MessageBox.Show("Ava detaili dokument (.ipt)", "POC-1")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    Dim tg As TransientGeometry = app.TransientGeometry
    
    Logger.Info("POC-1: === Edge Arc Length Measurement ===")
    Logger.Info("POC-1: Part: " & partDoc.DisplayName)
    Logger.Info("POC-1: Bodies: " & compDef.SurfaceBodies.Count)
    
    Dim totalEdges As Integer = 0
    Dim measuredEdges As Integer = 0
    Dim failedEdges As Integer = 0
    
    For bodyIdx As Integer = 1 To compDef.SurfaceBodies.Count
        Dim body As SurfaceBody = compDef.SurfaceBodies.Item(bodyIdx)
        Dim isSolid As Boolean = False
        Try
            isSolid = body.Volume(0.01) > 0
        Catch
        End Try
        
        Logger.Info("POC-1: --- Body " & bodyIdx & ": '" & body.Name & "' (" & If(isSolid, "Solid", "Surface") & ") ---")
        Logger.Info("POC-1:     Faces: " & body.Faces.Count & ", Edges: " & body.Edges.Count & ", Vertices: " & body.Vertices.Count)
        
        For edgeIdx As Integer = 1 To body.Edges.Count
            totalEdges += 1
            Dim edge As Edge = body.Edges.Item(edgeIdx)
            
            Try
                Dim geomType As String = GetEdgeGeometryType(edge)
                Dim arcLen As Double = MeasureEdgeArcLength(edge)
                Dim chordLen As Double = MeasureEdgeChordLength(edge, tg)
                Dim ratio As Double = 0
                If chordLen > 0.0001 Then ratio = arcLen / chordLen
                
                Dim startPt As String = FormatVertex(edge.StartVertex)
                Dim stopPt As String = FormatVertex(edge.StopVertex)
                
                Logger.Info("POC-1:     Edge " & edgeIdx & ": " & geomType & _
                            " | Arc=" & FormatMm(arcLen) & "mm" & _
                            " | Chord=" & FormatMm(chordLen) & "mm" & _
                            " | Ratio=" & ratio.ToString("0.000") & _
                            " | " & startPt & " -> " & stopPt)
                
                measuredEdges += 1
            Catch ex As Exception
                failedEdges += 1
                Logger.Warn("POC-1:     Edge " & edgeIdx & ": FAILED - " & ex.Message)
            End Try
        Next
        
        Logger.Info("POC-1:     Face summary:")
        For faceIdx As Integer = 1 To body.Faces.Count
            Dim face As Face = body.Faces.Item(faceIdx)
            Dim faceGeom As String = GetFaceGeometryType(face)
            Dim faceArea As Double = 0
            Try
                faceArea = face.Evaluator.Area
            Catch
            End Try
            
            Dim edgeLoops As String = ""
            Try
                For Each el As EdgeLoop In face.EdgeLoops
                    If edgeLoops <> "" Then edgeLoops &= ", "
                    edgeLoops &= el.Edges.Count & " edges"
                    If el.IsOuterEdgeLoop Then edgeLoops &= " (outer)"
                Next
            Catch
            End Try
            
            Logger.Info("POC-1:     Face " & faceIdx & ": " & faceGeom & _
                        " | Area=" & FormatMm2(faceArea) & "mm² | Loops: " & edgeLoops)
        Next
    Next
    
    Logger.Info("POC-1: === Summary ===")
    Logger.Info("POC-1: Total edges: " & totalEdges & " | Measured: " & measuredEdges & " | Failed: " & failedEdges)
    Logger.Info("POC-1: === Done ===")
End Sub

' Measure arc length of an edge using CurveEvaluator
Function MeasureEdgeArcLength(edge As Edge) As Double
    Dim evaluator As CurveEvaluator = edge.Evaluator
    
    Dim minParam As Double = 0
    Dim maxParam As Double = 0
    Call evaluator.GetParamExtents(minParam, maxParam)
    
    Dim arcLength As Double = 0
    Call evaluator.GetLengthAtParam(minParam, maxParam, arcLength)
    
    Return arcLength
End Function

' Measure chord length (straight-line distance between edge endpoints)
Function MeasureEdgeChordLength(edge As Edge, tg As TransientGeometry) As Double
    If edge.StartVertex Is Nothing OrElse edge.StopVertex Is Nothing Then Return 0
    
    Dim p1 As Point = edge.StartVertex.Point
    Dim p2 As Point = edge.StopVertex.Point
    
    Dim dx As Double = p2.X - p1.X
    Dim dy As Double = p2.Y - p1.Y
    Dim dz As Double = p2.Z - p1.Z
    
    Return Math.Sqrt(dx * dx + dy * dy + dz * dz)
End Function

' Get the geometry type name of an edge
Function GetEdgeGeometryType(edge As Edge) As String
    Try
        Dim geom As Object = edge.Geometry
        Dim typeName As String = geom.GetType().Name
        
        ' Replace COM-generated names with readable ones
        If TypeOf geom Is Line Then Return "Line"
        If TypeOf geom Is LineSegment Then Return "LineSegment"
        If TypeOf geom Is Arc3d Then Return "Arc3d"
        If TypeOf geom Is Circle Then Return "Circle"
        If TypeOf geom Is BSplineCurve Then Return "BSpline"
        If TypeOf geom Is EllipticalArc Then Return "EllipticalArc"
        
        Return typeName
    Catch
        Return "Unknown"
    End Try
End Function

' Get the geometry type name of a face
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

' Format a vertex point as a readable string (in mm)
Function FormatVertex(v As Vertex) As String
    If v Is Nothing Then Return "(none)"
    Dim p As Point = v.Point
    Return "(" & FormatMm(p.X) & ", " & FormatMm(p.Y) & ", " & FormatMm(p.Z) & ")"
End Function

' Convert cm to mm and format
Function FormatMm(valueCm As Double) As String
    Return (valueCm * 10).ToString("0.00")
End Function

Function FormatMm2(valueCm2 As Double) As String
    Return (valueCm2 * 100).ToString("0.00")
End Function
