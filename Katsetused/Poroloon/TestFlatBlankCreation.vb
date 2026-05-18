' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' POC-4: Flat Blank Creation Test
'
' Tests the workflow for creating a developed flat blank from a curved foam part:
' 1. User picks profile face (end cross-section)
' 2. User picks back surface (defines flat plane)
' 3. User picks outer edge chain (measures max developed length)
' 4. User picks inner edge chain (measures min developed length)
' 5. Creates flat blank solid body with angled end cuts
' 6. Sets up DVRs for visibility control
'
' The flat blank is created as a separate solid body in the same part file.
' DVRs control which body is visible:
' - "Komponent": curved body visible, flat blank hidden (for assemblies)
' - "Pinnalaotus": flat blank visible, curved body hidden (for drawings/CAM)
'
' Run on: active part document (foam part with curved geometry)
' ============================================================================

AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"

Imports Inventor
Imports System.Collections.Generic
Imports System.Windows.Forms

Sub Main()
    Const DVR_KOMPONENT As String = "Komponent"
    Const DVR_PINNALAOTUS As String = "Pinnalaotus"
    Const FLAT_BODY_NAME As String = "Arendus"
    
    UtilsLib.SetLogger(Logger)
    
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("POC-4: Ava detaili dokument (.ipt)")
        MessageBox.Show("Ava detaili dokument (.ipt)", "POC-4")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    Dim tg As TransientGeometry = app.TransientGeometry
    
    Logger.Info("POC-4: === Flat Blank Creation Test ===")
    Logger.Info("POC-4: Part: " & partDoc.DisplayName)
    
    ' ========================================================================
    ' Step 1: User picks profile sketch (the cross-section to extrude)
    ' ========================================================================
    Logger.Info("POC-4: Step 1 - Select profile sketch")
    Dim profileSketch As PlanarSketch = Nothing
    Try
        Dim picked As Object = app.CommandManager.Pick( _
            SelectionFilterEnum.kSketchCurveFilter, _
            "Vali profiili eskiisi joon (sweep algne profiil) - ESC tühistamiseks")
        If picked IsNot Nothing Then
            ' Get the sketch from the picked curve
            If TypeOf picked Is SketchEntity Then
                profileSketch = CType(picked, SketchEntity).Parent
            End If
        End If
    Catch
        Logger.Info("POC-4: Cancelled")
        Exit Sub
    End Try
    
    If profileSketch Is Nothing Then
        Logger.Info("POC-4: No profile sketch selected")
        Exit Sub
    End If
    
    Logger.Info("POC-4: Profile sketch selected: " & profileSketch.Name)
    
    ' ========================================================================
    ' Step 2: User picks flat direction surface (determines extrude orientation)
    ' ========================================================================
    Logger.Info("POC-4: Step 2 - Select surface for flat direction")
    Dim flatDirFace As Face = Nothing
    Try
        Dim picked As Object = app.CommandManager.Pick( _
            SelectionFilterEnum.kPartFaceFilter, _
            "Vali pind, mis määrab lameda suuna (nt tagapind) - ESC tühistamiseks")
        If picked IsNot Nothing AndAlso TypeOf picked Is Face Then
            flatDirFace = CType(picked, Face)
        End If
    Catch
        Logger.Info("POC-4: Cancelled")
        Exit Sub
    End Try
    
    If flatDirFace Is Nothing Then
        Logger.Info("POC-4: No flat direction surface selected")
        Exit Sub
    End If
    
    Logger.Info("POC-4: Flat direction surface selected: " & GetFaceGeometryType(flatDirFace))
    
    ' ========================================================================
    ' Step 3: User picks outer edge chain (longest developed length)
    ' ========================================================================
    Logger.Info("POC-4: Step 3 - Select outer edge(s) - press ESC when done")
    Dim outerEdges As New List(Of Edge)
    Dim outerEdgeNum As Integer = 1
    Do
        Dim edge As Edge = Nothing
        Try
            Dim picked As Object = app.CommandManager.Pick( _
                SelectionFilterEnum.kPartEdgeFilter, _
                "Vali välisserv #" & outerEdgeNum & " (pikem arendusrada) - ESC lõpetamiseks")
            If picked IsNot Nothing AndAlso TypeOf picked Is Edge Then
                edge = CType(picked, Edge)
            End If
        Catch
            Exit Do
        End Try
        If edge Is Nothing Then Exit Do
        outerEdges.Add(edge)
        outerEdgeNum += 1
    Loop
    
    If outerEdges.Count = 0 Then
        Logger.Info("POC-4: No outer edges selected")
        Exit Sub
    End If
    
    Dim outerLength As Double = 0
    For Each e As Edge In outerEdges
        outerLength += MeasureEdgeArcLength(e)
    Next
    Logger.Info("POC-4: Outer edge chain: " & outerEdges.Count & " edges, total length = " & FormatMm(outerLength) & " mm")
    
    ' ========================================================================
    ' Step 4: User picks inner edge chain (shortest developed length)
    ' ========================================================================
    Logger.Info("POC-4: Step 4 - Select inner edge(s) - press ESC when done")
    Dim innerEdges As New List(Of Edge)
    Dim innerEdgeNum As Integer = 1
    Do
        Dim edge As Edge = Nothing
        Try
            Dim picked As Object = app.CommandManager.Pick( _
                SelectionFilterEnum.kPartEdgeFilter, _
                "Vali sisserv #" & innerEdgeNum & " (lühem arendusrada) - ESC lõpetamiseks")
            If picked IsNot Nothing AndAlso TypeOf picked Is Edge Then
                edge = CType(picked, Edge)
            End If
        Catch
            Exit Do
        End Try
        If edge Is Nothing Then Exit Do
        innerEdges.Add(edge)
        innerEdgeNum += 1
    Loop
    
    If innerEdges.Count = 0 Then
        Logger.Info("POC-4: No inner edges selected")
        Exit Sub
    End If
    
    Dim innerLength As Double = 0
    For Each e As Edge In innerEdges
        innerLength += MeasureEdgeArcLength(e)
    Next
    Logger.Info("POC-4: Inner edge chain: " & innerEdges.Count & " edges, total length = " & FormatMm(innerLength) & " mm")
    
    ' ========================================================================
    ' Step 5: User picks optional cut profile sketch (for curved end cuts)
    ' ========================================================================
    Logger.Info("POC-4: Step 5 - Select cut profile sketch (optional, ESC to skip for straight cuts)")
    Dim cutProfileSketch As PlanarSketch = Nothing
    Dim cutProfileBasePoint As SketchPoint = Nothing
    
    Try
        Dim picked As Object = app.CommandManager.Pick( _
            SelectionFilterEnum.kSketchCurveFilter, _
            "Vali lõikeprofiili eskiisi joon (kõvera lõike jaoks) - ESC sirge lõike jaoks")
        If picked IsNot Nothing Then
            If TypeOf picked Is SketchEntity Then
                cutProfileSketch = CType(picked, SketchEntity).Parent
            End If
        End If
    Catch
        ' User pressed ESC - use straight cut planes
        Logger.Info("POC-4: Using straight cut planes (no profile sketch)")
    End Try
    
    If cutProfileSketch IsNot Nothing Then
        Logger.Info("POC-4: Cut profile sketch selected: " & cutProfileSketch.Name)
        
        ' Ask user to select the base point (point that aligns with rotation axis)
        Logger.Info("POC-4: Step 5b - Select base point on cut profile (point on rotation axis)")
        Try
            Dim picked As Object = app.CommandManager.Pick( _
                SelectionFilterEnum.kSketchPointFilter, _
                "Vali baaspunkt lõikeprofiilil (punkt, mis asub pöördeteljel) - ESC tühistamiseks")
            If picked IsNot Nothing AndAlso TypeOf picked Is SketchPoint Then
                cutProfileBasePoint = CType(picked, SketchPoint)
                Logger.Info("POC-4: Base point selected")
            End If
        Catch
            Logger.Info("POC-4: No base point selected, will use sketch origin")
        End Try
    End If
    
    ' ========================================================================
    ' Calculate flat blank dimensions
    ' ========================================================================
    Dim lengthDiff As Double = outerLength - innerLength
    Dim overhangPerEnd As Double = lengthDiff / 2
    
    ' Calculate foam thickness from distance between inner and outer edges
    ' This is the radial distance between the two surfaces
    Dim foamThickness As Double = MeasureEdgeDistance(outerEdges(0), innerEdges(0))
    
    ' Get profile width from sketch (for extrude direction)
    Dim sketchExtents() As Double = GetSketchExtents(profileSketch)
    Dim profileWidth As Double = sketchExtents(0)
    
    Logger.Info("POC-4: --- Flat Blank Calculations ---")
    Logger.Info("POC-4:   Outer length: " & FormatMm(outerLength) & " mm")
    Logger.Info("POC-4:   Inner length: " & FormatMm(innerLength) & " mm")
    Logger.Info("POC-4:   Length diff: " & FormatMm(lengthDiff) & " mm")
    Logger.Info("POC-4:   Overhang per end: " & FormatMm(overhangPerEnd) & " mm")
    Logger.Info("POC-4:   Foam thickness: " & FormatMm(foamThickness) & " mm")
    Logger.Info("POC-4:   Profile width: " & FormatMm(profileWidth) & " mm")
    
    Dim cutAngle As Double = 0
    If foamThickness > 0.001 Then
        cutAngle = Math.Atan(overhangPerEnd / foamThickness) * 180 / Math.PI
        Logger.Info("POC-4:   End cut angle: " & cutAngle.ToString("0.0") & "° from perpendicular")
    End If
    
    ' ========================================================================
    ' Wrap all geometry creation + DVR setup in one transaction for easy undo
    ' ========================================================================
    Dim trans As Transaction = Nothing
    Try
        trans = app.TransactionManager.StartTransaction(partDoc, "POC-4: Loo arendus")
        
        ' ========================================================================
        ' Step 5: Create flat blank geometry
        ' ========================================================================
        Logger.Info("POC-4: Step 5 - Creating flat blank geometry...")
        
        Dim flatBody As SurfaceBody = Nothing
        
        ' Create the flat blank using sketch + extrude + cuts approach
        flatBody = CreateFlatBlankBody(partDoc, compDef, tg, _
                                       profileSketch, flatDirFace, _
                                       outerLength, innerLength, foamThickness, _
                                       cutProfileSketch, cutProfileBasePoint)
        
        If flatBody IsNot Nothing Then
            ' Rename the body
            Try
                flatBody.Name = FLAT_BODY_NAME
            Catch
            End Try
            
            Logger.Info("POC-4: Flat blank body created: '" & flatBody.Name & "'")
            
            ' ========================================================================
            ' Step 6: Set up DVRs
            ' ========================================================================
            Logger.Info("POC-4: Step 6 - Setting up Design View Representations...")
            SetupDVRs(partDoc, compDef, flatBody, DVR_KOMPONENT, DVR_PINNALAOTUS, FLAT_BODY_NAME)
        Else
            Logger.Warn("POC-4: Failed to create flat blank body")
        End If
        
        trans.End()
        Logger.Info("POC-4: === Done (use Ctrl+Z to undo) ===")
        
    Catch ex As Exception
        Logger.Error("POC-4: Error: " & ex.Message)
        If trans IsNot Nothing Then
            Try : trans.Abort() : Catch : End Try
        End If
    End Try
End Sub

' ============================================================================
' Create the flat blank solid body
' ============================================================================
Function CreateFlatBlankBody(partDoc As PartDocument, compDef As PartComponentDefinition, _
                             tg As TransientGeometry, _
                             profileSketch As PlanarSketch, flatDirFace As Face, _
                             outerLength As Double, innerLength As Double, _
                             foamThickness As Double, _
                             cutProfileSketch As PlanarSketch, cutProfileBasePoint As SketchPoint) As SurfaceBody
    
    ' Approach:
    ' 1. Create a work plane offset from the flat direction surface
    ' 2. Create a new sketch and copy/project the original profile curves
    ' 3. Extrude by the outer (longer) length
    ' 4. Cut the ends at angles based on length difference
    
    ' Calculate overhang (how much longer outer edge is on each end)
    Dim overhang As Double = (outerLength - innerLength) / 2
    
    Logger.Info("POC-4:   Foam thickness: " & FormatMm(foamThickness) & " mm")
    Logger.Info("POC-4:   Overhang per end: " & FormatMm(overhang) & " mm")
    
    ' Create a work plane offset from the part for the flat blank
    Dim offsetDist As Double = 20  ' 200mm offset to place the flat blank away from the original
    
    Try
        ' Create work plane parallel to the original sketch's plane, offset away
        Dim workPlanes As WorkPlanes = compDef.WorkPlanes
        Dim flatWorkPlane As WorkPlane = Nothing
        
        Try
            ' Use the profile sketch's plane as reference
            flatWorkPlane = workPlanes.AddByPlaneAndOffset(profileSketch.PlanarEntity, offsetDist)
            flatWorkPlane.Name = "Arenduse tasand"
            flatWorkPlane.Visible = False
            Logger.Info("POC-4:   Created work plane offset from profile sketch plane")
        Catch ex As Exception
            Logger.Warn("POC-4: Could not create work plane from sketch: " & ex.Message)
            ' Fall back to XY plane offset
            flatWorkPlane = workPlanes.AddByPlaneAndOffset(compDef.WorkPlanes.Item(3), offsetDist)
            flatWorkPlane.Name = "Arenduse tasand"
            flatWorkPlane.Visible = False
        End Try
        
        ' Create sketch on the work plane and project original profile geometry
        Dim newSketch As PlanarSketch = compDef.Sketches.Add(flatWorkPlane)
        newSketch.Name = "Arenduse profiil"
        
        ' Project all curves from the original profile sketch
        Dim projectedCount As Integer = 0
        For Each entity As SketchEntity In profileSketch.SketchEntities
            If entity.Construction Then Continue For
            
            Try
                ' Project the entity onto the new sketch
                newSketch.AddByProjectingEntity(entity)
                projectedCount += 1
            Catch
                ' Some entities may not be projectable, skip them
            End Try
        Next
        
        Logger.Info("POC-4:   Projected " & projectedCount & " entities from original sketch")
        
        ' If projection failed, try a different approach - copy the sketch manually
        If projectedCount = 0 Then
            Logger.Warn("POC-4: Projection failed, using simplified rectangle")
            ' Fall back to simple rectangle
            Dim halfWidth As Double = profileWidth / 2
            Dim halfThick As Double = profileThickness / 2
            
            Dim p1 As Point2d = tg.CreatePoint2d(-halfWidth, -halfThick)
            Dim p2 As Point2d = tg.CreatePoint2d(halfWidth, -halfThick)
            Dim p3 As Point2d = tg.CreatePoint2d(halfWidth, halfThick)
            Dim p4 As Point2d = tg.CreatePoint2d(-halfWidth, halfThick)
            
            Dim lines As SketchLines = newSketch.SketchLines
            lines.AddByTwoPoints(p1, p2)
            lines.AddByTwoPoints(p2, p3)
            lines.AddByTwoPoints(p3, p4)
            lines.AddByTwoPoints(p4, p1)
        End If
        
        ' Create profile for extrusion
        Dim profile As Profile = newSketch.Profiles.AddForSolid()
        
        ' Extrude by the OUTER (longer) length to create the initial solid
        Dim extrudeDef As ExtrudeDefinition = compDef.Features.ExtrudeFeatures.CreateExtrudeDefinition( _
            profile, PartFeatureOperationEnum.kNewBodyOperation)
        extrudeDef.SetDistanceExtent(outerLength, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection)
        
        Dim extrudeFeature As ExtrudeFeature = compDef.Features.ExtrudeFeatures.Add(extrudeDef)
        extrudeFeature.Name = "Arenduse extrude"
        
        Logger.Info("POC-4:   Created extrude feature (length = " & FormatMm(outerLength) & " mm)")
        
        ' Get the created body
        If extrudeFeature.SurfaceBodies.Count = 0 Then
            Logger.Error("POC-4: Extrude created no bodies")
            Return Nothing
        End If
        
        Dim flatBody As SurfaceBody = extrudeFeature.SurfaceBodies.Item(1)
        
        ' Now cut the ends at angles to create the trapezoid shape
        ' The cut angle is determined by overhang / foamThickness
        If overhang > 0.001 AndAlso foamThickness > 0.001 Then
            Logger.Info("POC-4:   Creating angled end cuts...")
            CutAngledEnds(compDef, tg, flatBody, flatDirFace, outerLength, innerLength, foamThickness, _
                          cutProfileSketch, cutProfileBasePoint)
        End If
        
        Return flatBody
        
    Catch ex As Exception
        Logger.Error("POC-4: Error in CreateFlatBlankBody: " & ex.Message)
    End Try
    
    Return Nothing
End Function

' ============================================================================
' Cut angled ends to create trapezoid shape using split features or profile sketches
' ============================================================================
Sub CutAngledEnds(compDef As PartComponentDefinition, tg As TransientGeometry, _
                  flatBody As SurfaceBody, flatDirFace As Face, _
                  outerLength As Double, innerLength As Double, _
                  foamThickness As Double, _
                  cutProfileSketch As PlanarSketch, cutProfileBasePoint As SketchPoint)
    
    ' Calculate cut parameters
    Dim overhang As Double = (outerLength - innerLength) / 2
    Dim cutAngleDeg As Double = Math.Atan(overhang / foamThickness) * 180 / Math.PI
    
    Logger.Info("POC-4:     Cut angle: " & cutAngleDeg.ToString("0.0") & "°")
    Logger.Info("POC-4:     Overhang: " & FormatMm(overhang) & " mm, Thickness: " & FormatMm(foamThickness) & " mm")
    
    Try
        Dim workPlanes As WorkPlanes = compDef.WorkPlanes
        Dim workAxes As WorkAxes = compDef.WorkAxes
        
        ' Find the two end faces of the extrude (perpendicular to length direction)
        Dim endFace1 As Face = Nothing
        Dim endFace2 As Face = Nothing
        Dim endFace1Normal As UnitVector = Nothing
        
        For Each face As Face In flatBody.Faces
            If Not TypeOf face.Geometry Is Plane Then Continue For
            
            Dim facePlane As Plane = CType(face.Geometry, Plane)
            Dim normal As UnitVector = facePlane.Normal
            
            If endFace1 Is Nothing Then
                endFace1 = face
                endFace1Normal = normal
            Else
                ' Check if this face has opposite normal (the other end)
                Dim dot As Double = normal.X * endFace1Normal.X + _
                                   normal.Y * endFace1Normal.Y + _
                                   normal.Z * endFace1Normal.Z
                If dot < -0.9 Then
                    endFace2 = face
                    Exit For
                End If
            End If
        Next
        
        If endFace1 Is Nothing OrElse endFace2 Is Nothing Then
            Logger.Warn("POC-4:     Could not find both end faces")
            Return
        End If
        
        Logger.Info("POC-4:     Found both end faces")
        
        ' Create a work plane parallel to the flat direction face, passing through the flat body
        ' This represents the "back" plane of our flat blank
        Dim backPlane As WorkPlane = workPlanes.AddByPlaneAndOffset(flatDirFace, 0)
        backPlane.Name = "Tagapinna tasand"
        backPlane.Visible = False
        
        ' Create work axes at the intersection of back plane and each end face
        ' These are the rotation axes for the angled cut profiles
        Dim axis1 As WorkAxis = workAxes.AddByTwoPlanes(backPlane, endFace1)
        axis1.Name = "Lõiketelg 1"
        axis1.Visible = True
        Logger.Info("POC-4:     Created rotation axis 1 (intersection of back plane and end face 1)")
        
        Dim axis2 As WorkAxis = workAxes.AddByTwoPlanes(backPlane, endFace2)
        axis2.Name = "Lõiketelg 2"
        axis2.Visible = True
        Logger.Info("POC-4:     Created rotation axis 2 (intersection of back plane and end face 2)")
        
        If cutProfileSketch IsNot Nothing Then
            ' ================================================================
            ' CURVED CUT: Copy the cut profile sketch to both ends
            ' ================================================================
            Logger.Info("POC-4:     Using cut profile sketch for curved cuts")
            
            ' Create sketches on the end faces and copy the cut profile
            CreateCutProfileAtEnd(compDef, tg, cutProfileSketch, cutProfileBasePoint, _
                                  endFace1, axis1, cutAngleDeg, "Lõikeprofiil 1")
            CreateCutProfileAtEnd(compDef, tg, cutProfileSketch, cutProfileBasePoint, _
                                  endFace2, axis2, -cutAngleDeg, "Lõikeprofiil 2")  ' Mirrored angle
            
            Logger.Info("POC-4:     Cut profile sketches created at both ends")
            Logger.Info("POC-4:     Use these sketches with Split Body or Sweep Cut manually")
        Else
            ' ================================================================
            ' STRAIGHT CUT: Create angled work planes
            ' ================================================================
            Logger.Info("POC-4:     Using straight cut planes")
            
            Dim splitPlane1 As WorkPlane = workPlanes.AddByLinePlaneAndAngle( _
                axis1, endFace1, cutAngleDeg & " deg")
            splitPlane1.Name = "Lõikepind 1"
            splitPlane1.Visible = True
            Logger.Info("POC-4:     Created split plane 1 at " & cutAngleDeg.ToString("0.0") & "°")
            
            Dim splitPlane2 As WorkPlane = workPlanes.AddByLinePlaneAndAngle( _
                axis2, endFace2, cutAngleDeg & " deg")
            splitPlane2.Name = "Lõikepind 2"
            splitPlane2.Visible = True
            Logger.Info("POC-4:     Created split plane 2 at " & cutAngleDeg.ToString("0.0") & "°")
        End If
        
        Logger.Info("POC-4:     Use Split Body feature manually to complete the cuts")
        
    Catch ex As Exception
        Logger.Warn("POC-4:     Could not create angled cuts: " & ex.Message)
    End Try
End Sub

' ============================================================================
' Create a rotated copy of the cut profile sketch at an end face
' ============================================================================
Sub CreateCutProfileAtEnd(compDef As PartComponentDefinition, tg As TransientGeometry, _
                          sourceSketch As PlanarSketch, basePoint As SketchPoint, _
                          endFace As Face, rotationAxis As WorkAxis, _
                          angleDeg As Double, sketchName As String)
    Try
        ' Create a new sketch on the end face
        Dim newSketch As PlanarSketch = compDef.Sketches.Add(endFace)
        newSketch.Name = sketchName
        
        ' Get the base point coordinates in the source sketch
        Dim baseX As Double = 0
        Dim baseY As Double = 0
        If basePoint IsNot Nothing Then
            baseX = basePoint.Geometry.X
            baseY = basePoint.Geometry.Y
        End If
        
        ' Copy sketch entities from source to new sketch
        ' We need to transform them to align base point with the rotation axis
        ' and rotate by the cut angle
        
        Dim angleRad As Double = angleDeg * Math.PI / 180
        Dim cosA As Double = Math.Cos(angleRad)
        Dim sinA As Double = Math.Sin(angleRad)
        
        ' For each entity in the source sketch, copy it with transformation
        For Each entity As SketchEntity In sourceSketch.SketchEntities
            If entity.Construction Then Continue For
            
            Try
                If TypeOf entity Is SketchLine Then
                    Dim srcLine As SketchLine = CType(entity, SketchLine)
                    
                    ' Get start and end points
                    Dim x1 As Double = srcLine.StartSketchPoint.Geometry.X - baseX
                    Dim y1 As Double = srcLine.StartSketchPoint.Geometry.Y - baseY
                    Dim x2 As Double = srcLine.EndSketchPoint.Geometry.X - baseX
                    Dim y2 As Double = srcLine.EndSketchPoint.Geometry.Y - baseY
                    
                    ' Rotate around origin (which is now the base point)
                    Dim rx1 As Double = x1 * cosA - y1 * sinA
                    Dim ry1 As Double = x1 * sinA + y1 * cosA
                    Dim rx2 As Double = x2 * cosA - y2 * sinA
                    Dim ry2 As Double = x2 * sinA + y2 * cosA
                    
                    ' Create line in new sketch
                    Dim p1 As Point2d = tg.CreatePoint2d(rx1, ry1)
                    Dim p2 As Point2d = tg.CreatePoint2d(rx2, ry2)
                    newSketch.SketchLines.AddByTwoPoints(p1, p2)
                    
                ElseIf TypeOf entity Is SketchArc Then
                    Dim srcArc As SketchArc = CType(entity, SketchArc)
                    
                    ' Get center, start, end points
                    Dim cx As Double = srcArc.CenterSketchPoint.Geometry.X - baseX
                    Dim cy As Double = srcArc.CenterSketchPoint.Geometry.Y - baseY
                    Dim sx As Double = srcArc.StartSketchPoint.Geometry.X - baseX
                    Dim sy As Double = srcArc.StartSketchPoint.Geometry.Y - baseY
                    Dim ex As Double = srcArc.EndSketchPoint.Geometry.X - baseX
                    Dim ey As Double = srcArc.EndSketchPoint.Geometry.Y - baseY
                    
                    ' Rotate
                    Dim rcx As Double = cx * cosA - cy * sinA
                    Dim rcy As Double = cx * sinA + cy * cosA
                    Dim rsx As Double = sx * cosA - sy * sinA
                    Dim rsy As Double = sx * sinA + sy * cosA
                    Dim rex As Double = ex * cosA - ey * sinA
                    Dim rey As Double = ex * sinA + ey * cosA
                    
                    ' Create arc in new sketch
                    Dim center As Point2d = tg.CreatePoint2d(rcx, rcy)
                    Dim startPt As Point2d = tg.CreatePoint2d(rsx, rsy)
                    Dim endPt As Point2d = tg.CreatePoint2d(rex, rey)
                    newSketch.SketchArcs.AddByCenterStartEndPoint(center, startPt, endPt)
                    
                ElseIf TypeOf entity Is SketchSpline Then
                    ' For splines, try to project the original
                    Try
                        newSketch.AddByProjectingEntity(entity)
                    Catch
                        Logger.Warn("POC-4:       Could not copy spline to cut profile")
                    End Try
                End If
            Catch
                ' Skip entities that can't be copied
            End Try
        Next
        
        Logger.Info("POC-4:     Created cut profile sketch: " & sketchName)
        
    Catch ex As Exception
        Logger.Warn("POC-4:     Failed to create cut profile: " & ex.Message)
    End Try
End Sub

' ============================================================================
' Set up Design View Representations
' ============================================================================
Sub SetupDVRs(partDoc As PartDocument, compDef As PartComponentDefinition, flatBody As SurfaceBody, _
              DVR_KOMPONENT As String, DVR_PINNALAOTUS As String, FLAT_BODY_NAME As String)
    Try
        Dim dvrs As DesignViewRepresentations = compDef.RepresentationsManager.DesignViewRepresentations
        
        ' Find or create "Komponent" DVR (curved body visible, flat hidden)
        Dim dvrKomponent As DesignViewRepresentation = Nothing
        Try
            dvrKomponent = dvrs.Item(DVR_KOMPONENT)
        Catch
            dvrKomponent = dvrs.Add(DVR_KOMPONENT)
            Logger.Info("POC-4:   Created DVR: " & DVR_KOMPONENT)
        End Try
        
        ' Find or create "Pinnalaotus" DVR (flat body visible, curved hidden)
        Dim dvrPinnalaotus As DesignViewRepresentation = Nothing
        Try
            dvrPinnalaotus = dvrs.Item(DVR_PINNALAOTUS)
        Catch
            dvrPinnalaotus = dvrs.Add(DVR_PINNALAOTUS)
            Logger.Info("POC-4:   Created DVR: " & DVR_PINNALAOTUS)
        End Try
        
        ' Activate Komponent DVR and hide the flat blank
        dvrKomponent.Activate()
        Try
            flatBody.Visible = False
        Catch
        End Try
        Logger.Info("POC-4:   Set " & DVR_KOMPONENT & " DVR: flat blank hidden")
        
        ' Activate Pinnalaotus DVR and show only the flat blank
        dvrPinnalaotus.Activate()
        Try
            ' Hide all bodies except the flat blank
            For Each body As SurfaceBody In compDef.SurfaceBodies
                If body.Name = FLAT_BODY_NAME Then
                    body.Visible = True
                Else
                    body.Visible = False
                End If
            Next
        Catch ex As Exception
            Logger.Warn("POC-4:   Could not set body visibility: " & ex.Message)
        End Try
        Logger.Info("POC-4:   Set " & DVR_PINNALAOTUS & " DVR: only flat blank visible")
        
        ' Return to default/Master DVR
        Try
            dvrs.Item("Master").Activate()
        Catch
            dvrs.Item(1).Activate()
        End Try
        
    Catch ex As Exception
        Logger.Error("POC-4: Error setting up DVRs: " & ex.Message)
    End Try
End Sub

' ============================================================================
' Helper functions
' ============================================================================

Function MeasureEdgeArcLength(edge As Edge) As Double
    Dim evaluator As CurveEvaluator = edge.Evaluator
    Dim minParam As Double = 0, maxParam As Double = 0
    Call evaluator.GetParamExtents(minParam, maxParam)
    Dim arcLength As Double = 0
    Call evaluator.GetLengthAtParam(minParam, maxParam, arcLength)
    Return arcLength
End Function

Function MeasureEdgeDistance(edge1 As Edge, edge2 As Edge) As Double
    ' Measure the distance between two edges by sampling their midpoints
    ' This gives the foam thickness (distance between inner and outer surfaces)
    
    Try
        ' Get midpoint of edge1
        Dim eval1 As CurveEvaluator = edge1.Evaluator
        Dim min1 As Double = 0, max1 As Double = 0
        Call eval1.GetParamExtents(min1, max1)
        Dim midParam1 As Double = (min1 + max1) / 2
        Dim pt1(2) As Double
        Call eval1.GetPointAtParam({midParam1}, pt1)
        
        ' Get midpoint of edge2
        Dim eval2 As CurveEvaluator = edge2.Evaluator
        Dim min2 As Double = 0, max2 As Double = 0
        Call eval2.GetParamExtents(min2, max2)
        Dim midParam2 As Double = (min2 + max2) / 2
        Dim pt2(2) As Double
        Call eval2.GetPointAtParam({midParam2}, pt2)
        
        ' Calculate distance
        Dim dx As Double = pt2(0) - pt1(0)
        Dim dy As Double = pt2(1) - pt1(1)
        Dim dz As Double = pt2(2) - pt1(2)
        
        Return Math.Sqrt(dx * dx + dy * dy + dz * dz)
    Catch ex As Exception
        Logger.Warn("POC-4: Could not measure edge distance: " & ex.Message)
        Return 0
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

Function GetSketchExtents(sketch As PlanarSketch) As Double()
    ' Calculate bounding extents from sketch geometry (2D, in sketch plane)
    ' Returns array of {larger, smaller} extent
    Dim minX As Double = Double.MaxValue
    Dim maxX As Double = Double.MinValue
    Dim minY As Double = Double.MaxValue
    Dim maxY As Double = Double.MinValue
    
    For Each curve As SketchEntity In sketch.SketchEntities
        If TypeOf curve Is SketchPoint Then
            Dim pt As SketchPoint = CType(curve, SketchPoint)
            Dim x As Double = pt.Geometry.X
            Dim y As Double = pt.Geometry.Y
            If x < minX Then minX = x
            If x > maxX Then maxX = x
            If y < minY Then minY = y
            If y > maxY Then maxY = y
        ElseIf TypeOf curve Is SketchLine Then
            Dim ln As SketchLine = CType(curve, SketchLine)
            For Each pt As SketchPoint In {ln.StartSketchPoint, ln.EndSketchPoint}
                Dim x As Double = pt.Geometry.X
                Dim y As Double = pt.Geometry.Y
                If x < minX Then minX = x
                If x > maxX Then maxX = x
                If y < minY Then minY = y
                If y > maxY Then maxY = y
            Next
        ElseIf TypeOf curve Is SketchArc Then
            Dim arc As SketchArc = CType(curve, SketchArc)
            For Each pt As SketchPoint In {arc.StartSketchPoint, arc.EndSketchPoint, arc.CenterSketchPoint}
                Dim x As Double = pt.Geometry.X
                Dim y As Double = pt.Geometry.Y
                If x < minX Then minX = x
                If x > maxX Then maxX = x
                If y < minY Then minY = y
                If y > maxY Then maxY = y
            Next
        End If
    Next
    
    Dim extX As Double = maxX - minX
    Dim extY As Double = maxY - minY
    
    If extX >= extY Then
        Return New Double() {extX, extY}
    Else
        Return New Double() {extY, extX}
    End If
End Function

Function GetFaceSmallestExtent(face As Face) As Double
    Dim extents() As Double = GetFaceExtents(face)
    Return extents(2)  ' Smallest
End Function

Function GetFaceExtents(face As Face) As Double()
    ' Calculate bounding extents from face vertices
    ' Returns array of {largest, middle, smallest}
    Dim minX As Double = Double.MaxValue
    Dim maxX As Double = Double.MinValue
    Dim minY As Double = Double.MaxValue
    Dim maxY As Double = Double.MinValue
    Dim minZ As Double = Double.MaxValue
    Dim maxZ As Double = Double.MinValue
    
    For Each v As Vertex In face.Vertices
        Dim pt As Point = v.Point
        If pt.X < minX Then minX = pt.X
        If pt.X > maxX Then maxX = pt.X
        If pt.Y < minY Then minY = pt.Y
        If pt.Y > maxY Then maxY = pt.Y
        If pt.Z < minZ Then minZ = pt.Z
        If pt.Z > maxZ Then maxZ = pt.Z
    Next
    
    Dim extX As Double = maxX - minX
    Dim extY As Double = maxY - minY
    Dim extZ As Double = maxZ - minZ
    
    ' Sort to get largest, middle, smallest
    Dim extents() As Double = {extX, extY, extZ}
    Array.Sort(extents)
    Array.Reverse(extents)  ' Now: largest, middle, smallest
    
    Return extents
End Function

Function GetLargestExtent(box As Box) As Double
    Dim x As Double = Math.Abs(box.MaxPoint.X - box.MinPoint.X)
    Dim y As Double = Math.Abs(box.MaxPoint.Y - box.MinPoint.Y)
    Dim z As Double = Math.Abs(box.MaxPoint.Z - box.MinPoint.Z)
    Return Math.Max(x, Math.Max(y, z))
End Function

Function FormatMm(valueCm As Double) As String
    Return (valueCm * 10).ToString("0.00")
End Function

Function FormatMm2(valueCm2 As Double) As String
    Return (valueCm2 * 100).ToString("0.00")
End Function
