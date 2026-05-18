' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' POC-3: End Tangent and Cut Angle Computation with User Selection
'
' User selects longitudinal edges (the curved edges running along the sweep).
' For each selected edge, computes:
' 1. Tangent vectors at start and end using GetFirstDerivatives
' 2. Chord direction (straight line between endpoints)
' 3. Angle between tangent and chord at each end
'
' These angles determine the end cuts needed on a flat blank.
'
' Run on: active part document (foam part with sweep feature)
' Output: iLogic log window with tangent vectors and computed angles
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
        Logger.Error("POC-3: Ava detaili dokument (.ipt)")
        MessageBox.Show("Ava detaili dokument (.ipt)", "POC-3")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    Dim tg As TransientGeometry = app.TransientGeometry
    
    Logger.Info("POC-3: === End Cut Angle Computation ===")
    Logger.Info("POC-3: Part: " & partDoc.DisplayName)
    
    ' ========================================================================
    ' User selects longitudinal edges (loop until done)
    ' ========================================================================
    Dim selectedEdges As New List(Of Edge)
    
    Logger.Info("POC-3: Select longitudinal edges (the curved edges along the sweep)")
    Logger.Info("POC-3: Press ESC when done selecting")
    
    Dim edgeNum As Integer = 1
    Do
        Dim edge As Edge = Nothing
        Try
            Dim picked As Object = app.CommandManager.Pick( _
                SelectionFilterEnum.kPartEdgeFilter, _
                "Vali pikiserv #" & edgeNum & " (kumer serv piki pühkimist) - ESC lõpetamiseks")
            If picked IsNot Nothing AndAlso TypeOf picked Is Edge Then
                edge = CType(picked, Edge)
            End If
        Catch
            ' ESC pressed - done selecting
            Exit Do
        End Try
        
        If edge Is Nothing Then Exit Do
        
        selectedEdges.Add(edge)
        Dim geomType As String = GetEdgeGeometryType(edge)
        Dim arcLen As Double = MeasureEdgeArcLength(edge)
        Logger.Info("POC-3: Selected edge " & edgeNum & ": " & geomType & " | Arc=" & FormatMm(arcLen) & "mm")
        edgeNum += 1
    Loop
    
    If selectedEdges.Count = 0 Then
        Logger.Info("POC-3: No edges selected")
        MessageBox.Show("Ühtegi serva ei valitud.", "POC-3")
        Exit Sub
    End If
    
    Logger.Info("POC-3: Total edges selected: " & selectedEdges.Count)
    
    ' ========================================================================
    ' Analyze each selected edge for tangent and angles
    ' ========================================================================
    Logger.Info("POC-3: --- Per-Edge Tangent & Angle Analysis ---")
    
    Dim startAngles As New List(Of Double)
    Dim endAngles As New List(Of Double)
    
    For i As Integer = 0 To selectedEdges.Count - 1
        Dim edge As Edge = selectedEdges(i)
        Dim evaluator As CurveEvaluator = edge.Evaluator
        
        ' Get param range
        Dim minP As Double = 0, maxP As Double = 0
        evaluator.GetParamExtents(minP, maxP)
        
        ' Get arc length
        Dim arcLen As Double = 0
        evaluator.GetLengthAtParam(minP, maxP, arcLen)
        
        ' Get points at start and end
        Dim startPts(2) As Double
        Dim endPts(2) As Double
        Dim startParamArr() As Double = {minP}
        Dim endParamArr() As Double = {maxP}
        evaluator.GetPointAtParam(startParamArr, startPts)
        evaluator.GetPointAtParam(endParamArr, endPts)
        
        ' Chord direction (start → end)
        Dim cx As Double = endPts(0) - startPts(0)
        Dim cy As Double = endPts(1) - startPts(1)
        Dim cz As Double = endPts(2) - startPts(2)
        Dim cLen As Double = Math.Sqrt(cx * cx + cy * cy + cz * cz)
        If cLen > 0.0001 Then
            cx /= cLen : cy /= cLen : cz /= cLen
        End If
        
        ' Get tangent at start (first derivative at minP)
        Dim startTangent(2) As Double
        Dim startParams() As Double = {minP}
        Try
            evaluator.GetFirstDerivatives(startParams, startTangent)
        Catch ex As Exception
            Logger.Warn("POC-3:   Edge " & (i + 1) & " start tangent failed: " & ex.Message)
            Continue For
        End Try
        
        ' Normalize start tangent
        Dim stLen As Double = Math.Sqrt(startTangent(0) * startTangent(0) + _
                                        startTangent(1) * startTangent(1) + _
                                        startTangent(2) * startTangent(2))
        If stLen > 0.0001 Then
            startTangent(0) /= stLen
            startTangent(1) /= stLen
            startTangent(2) /= stLen
        End If
        
        ' Get tangent at end (first derivative at maxP)
        Dim endTangent(2) As Double
        Dim endParams() As Double = {maxP}
        Try
            evaluator.GetFirstDerivatives(endParams, endTangent)
        Catch ex As Exception
            Logger.Warn("POC-3:   Edge " & (i + 1) & " end tangent failed: " & ex.Message)
            Continue For
        End Try
        
        ' Normalize end tangent
        Dim etLen As Double = Math.Sqrt(endTangent(0) * endTangent(0) + _
                                        endTangent(1) * endTangent(1) + _
                                        endTangent(2) * endTangent(2))
        If etLen > 0.0001 Then
            endTangent(0) /= etLen
            endTangent(1) /= etLen
            endTangent(2) /= etLen
        End If
        
        ' Angle between start tangent and chord
        Dim dotStart As Double = startTangent(0) * cx + startTangent(1) * cy + startTangent(2) * cz
        dotStart = Math.Max(-1, Math.Min(1, dotStart))
        Dim angleStart As Double = Math.Acos(Math.Abs(dotStart)) * 180 / Math.PI
        
        ' Angle between end tangent and chord
        Dim dotEnd As Double = endTangent(0) * cx + endTangent(1) * cy + endTangent(2) * cz
        dotEnd = Math.Max(-1, Math.Min(1, dotEnd))
        Dim angleEnd As Double = Math.Acos(Math.Abs(dotEnd)) * 180 / Math.PI
        
        startAngles.Add(angleStart)
        endAngles.Add(angleEnd)
        
        Logger.Info("POC-3:   Edge " & (i + 1) & " (" & GetEdgeGeometryType(edge) & ", Arc=" & FormatMm(arcLen) & "mm):")
        Logger.Info("POC-3:     Start tangent: " & FormatArr(startTangent))
        Logger.Info("POC-3:     End tangent:   " & FormatArr(endTangent))
        Logger.Info("POC-3:     Chord dir:     (" & cx.ToString("0.000") & ", " & cy.ToString("0.000") & ", " & cz.ToString("0.000") & ")")
        Logger.Info("POC-3:     Angle start tangent-chord: " & angleStart.ToString("0.00") & "°")
        Logger.Info("POC-3:     Angle end tangent-chord:   " & angleEnd.ToString("0.00") & "°")
    Next
    
    ' ========================================================================
    ' Summary: consistency check across edges
    ' ========================================================================
    If startAngles.Count > 0 Then
        Logger.Info("POC-3: --- Angle Consistency ---")
        
        Dim avgStartAngle As Double = 0
        Dim avgEndAngle As Double = 0
        Dim maxStartDev As Double = 0
        Dim maxEndDev As Double = 0
        
        For Each a As Double In startAngles
            avgStartAngle += a
        Next
        avgStartAngle /= startAngles.Count
        
        For Each a As Double In endAngles
            avgEndAngle += a
        Next
        avgEndAngle /= endAngles.Count
        
        For Each a As Double In startAngles
            Dim dev As Double = Math.Abs(a - avgStartAngle)
            If dev > maxStartDev Then maxStartDev = dev
        Next
        
        For Each a As Double In endAngles
            Dim dev As Double = Math.Abs(a - avgEndAngle)
            If dev > maxEndDev Then maxEndDev = dev
        Next
        
        Logger.Info("POC-3:   Start angle: avg=" & avgStartAngle.ToString("0.00") & "° | max deviation=" & maxStartDev.ToString("0.00") & "°")
        Logger.Info("POC-3:   End angle:   avg=" & avgEndAngle.ToString("0.00") & "° | max deviation=" & maxEndDev.ToString("0.00") & "°")
        
        If maxStartDev < 1.0 AndAlso maxEndDev < 1.0 Then
            Logger.Info("POC-3:   RESULT: Angles are consistent across edges -> simple planar cuts should work")
        ElseIf maxStartDev < 5.0 AndAlso maxEndDev < 5.0 Then
            Logger.Info("POC-3:   RESULT: Angles vary slightly -> planar cuts are approximate but likely acceptable for foam")
        Else
            Logger.Info("POC-3:   RESULT: Significant angle variation -> may need non-planar (ruled surface) cuts")
        End If
        
        Logger.Info("POC-3: --- Proposed Flat Blank Parameters ---")
        Dim maxArcLen As Double = 0
        For Each edge As Edge In selectedEdges
            Try
                Dim aLen As Double = MeasureEdgeArcLength(edge)
                If aLen > maxArcLen Then maxArcLen = aLen
            Catch
            End Try
        Next
        Logger.Info("POC-3:   Blank length (max arc): " & FormatMm(maxArcLen) & "mm")
        Logger.Info("POC-3:   Start cut angle: " & avgStartAngle.ToString("0.0") & "°")
        Logger.Info("POC-3:   End cut angle: " & avgEndAngle.ToString("0.0") & "°")
    End If
    
    Logger.Info("POC-3: === Done ===")
End Sub

Function MeasureEdgeArcLength(edge As Edge) As Double
    Dim evaluator As CurveEvaluator = edge.Evaluator
    Dim minParam As Double = 0, maxParam As Double = 0
    Call evaluator.GetParamExtents(minParam, maxParam)
    Dim arcLength As Double = 0
    Call evaluator.GetLengthAtParam(minParam, maxParam, arcLength)
    Return arcLength
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

Function FormatArr(arr() As Double) As String
    Return "(" & arr(0).ToString("0.000") & ", " & arr(1).ToString("0.000") & ", " & arr(2).ToString("0.000") & ")"
End Function

Function FormatMm(valueCm As Double) As String
    Return (valueCm * 10).ToString("0.00")
End Function
