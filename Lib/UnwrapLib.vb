' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' UnwrapLib - Pinnalaotus (Unwrap) feature detection and utilities
' 
' Provides functions to:
' - Detect UnwrapFeature in a part
' - Find associated ThickenFeature
' - Create ThickenFeature if missing
' - Get dimensions from thickened unwrap body
' - Design View Representations: Pinnalaotus (thickened flat — manufactured), Komponent (bent/original — hides unwrap+thicken outputs)
'
' Pinnalaotus workflow:
' 1. User creates solid body (bent/curved shape)
' 2. User creates Unwrap feature → surface body
' 3. User/script creates Thicken feature → flat solid body
' 4. Dimensions come from thickened body (thickness direction defaults to unwrap surface plane normal)
' 5. 1:1 CAM drawings use "Pinnalaotus" DVR (thickened manufactured solid); assemblies use "Komponent" DVR (bent part)
'
' Unwrap may exist on a normal part OR a sheet-metal subtype part. Detection uses
' UnwrapFeatures only — never infer Pinnalaotus from SubType alone.
'
' Usage: 
'   AddVbFile "Lib/UtilsLib.vb"
'   AddVbFile "Lib/UnwrapLib.vb"
'   UtilsLib.SetLogger(Logger) ' In Sub Main
'
' Dependencies: UtilsLib (for logging)
' ============================================================================

Imports Inventor
Imports System.Collections.Generic

Public Module UnwrapLib

    ''' <summary>Manufactured flat part: only the thickened unwrap solid is visible (1:1 CAM drawings).</summary>
    Public Const DVR_NAME_PINNALAOTUS As String = "Pinnalaotus"
    ''' <summary>Bent/original component: unwrap surface and thickened flat solid hidden (assemblies, design intent).</summary>
    Public Const DVR_NAME_KOMPONENT As String = "Komponent"
    ''' <summary>Legacy name for manufactured-solid-only DVR; treated like Pinnalaotus for CAM.</summary>
    Private Const LEGACY_MANUFACTURED_DVR_NAME As String = "BB_Pinnalaotus"
    
    ' Custom property for dimension source
    Public Const PROP_DIMENSION_SOURCE As String = "BB_DimensionSource"
    ''' <summary>Stored measurement solid body name for Uuenda rule (thickened unwrap body).</summary>
    Public Const PROP_PINNALAOTUS_BODY_NAME As String = "BB_PinnalaotusSolidBodyName"
    Public Const DIMENSION_SOURCE_PINNALAOTUS As String = "Pinnalaotus"
    Public Const DIMENSION_SOURCE_LEHTMETALL As String = "Lehtmetall"
    Public Const DIMENSION_SOURCE_NORMAL As String = "Normal"
    
    ' ============================================================================
    ' Detection Functions
    ' ============================================================================
    
    ''' <summary>
    ''' Check if part has any UnwrapFeature
    ''' </summary>
    Public Function HasUnwrapFeature(partDoc As PartDocument) As Boolean
        Try
            Dim unwraps As UnwrapFeatures = partDoc.ComponentDefinition.Features.UnwrapFeatures
            Return unwraps.Count > 0
        Catch
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Get the first UnwrapFeature (assumes single unwrap per part)
    ''' </summary>
    Public Function GetUnwrapFeature(partDoc As PartDocument) As UnwrapFeature
        Try
            Dim unwraps As UnwrapFeatures = partDoc.ComponentDefinition.Features.UnwrapFeatures
            If unwraps.Count > 0 Then
                Return unwraps.Item(1)
            End If
        Catch ex As Exception
            UtilsLib.LogWarn("UnwrapLib: Error getting UnwrapFeature: " & ex.Message)
        End Try
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Get all UnwrapFeatures in a part
    ''' </summary>
    Public Function GetUnwrapFeatures(partDoc As PartDocument) As List(Of UnwrapFeature)
        Dim result As New List(Of UnwrapFeature)
        Try
            Dim unwraps As UnwrapFeatures = partDoc.ComponentDefinition.Features.UnwrapFeatures
            For Each uw As UnwrapFeature In unwraps
                result.Add(uw)
            Next
        Catch
        End Try
        Return result
    End Function
    
    ''' <summary>
    ''' Get the surface body created by an UnwrapFeature
    ''' </summary>
    Public Function GetUnwrappedSurfaceBody(unwrapFeature As UnwrapFeature) As SurfaceBody
        Try
            Dim resultBodies As SurfaceBodies = unwrapFeature.SurfaceBodies
            If resultBodies.Count > 0 Then
                Return resultBodies.Item(1)
            End If
        Catch ex As Exception
            UtilsLib.LogWarn("UnwrapLib: Error getting unwrapped surface body: " & ex.Message)
        End Try
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Unwrap output is a flat surface; thickness for Pinnalaotus aligns with the plane normal.
    ''' Uses the first planar face found on the unwrap surface body (all should share the same normal direction).
    ''' </summary>
    Public Function TryGetUnwrapFlatSurfaceNormal(unwrapFeature As UnwrapFeature, _
                                                   ByRef nx As Double, ByRef ny As Double, ByRef nz As Double) As Boolean
        nx = 0 : ny = 0 : nz = 0
        Dim surf As SurfaceBody = GetUnwrappedSurfaceBody(unwrapFeature)
        If surf Is Nothing Then Return False
        Try
            For Each face As Face In surf.Faces
                Try
                    Dim geom As Object = face.Geometry
                    If TypeOf geom Is Plane Then
                        Dim pl As Plane = CType(geom, Plane)
                        Dim normal As UnitVector = pl.Normal
                        nx = normal.X
                        ny = normal.Y
                        nz = normal.Z
                        Dim len As Double = Math.Sqrt(nx * nx + ny * ny + nz * nz)
                        If len > 0.0001 Then
                            nx /= len : ny /= len : nz /= len
                        End If
                        If nx < -0.0001 OrElse (Math.Abs(nx) < 0.0001 AndAlso ny < -0.0001) OrElse _
                           (Math.Abs(nx) < 0.0001 AndAlso Math.Abs(ny) < 0.0001 AndAlso nz < -0.0001) Then
                            nx = -nx : ny = -ny : nz = -nz
                        End If
                        Return True
                    End If
                Catch
                End Try
            Next
        Catch ex As Exception
            UtilsLib.LogWarn("UnwrapLib: TryGetUnwrapFlatSurfaceNormal failed: " & ex.Message)
        End Try
        Return False
    End Function
    
    ' ============================================================================
    ' Thicken Feature Functions
    ' ============================================================================
    
    ''' <summary>
    ''' Find ThickenFeature that uses faces from the unwrapped surface body
    ''' </summary>
    Public Function GetThickenForUnwrap(partDoc As PartDocument, unwrapFeature As UnwrapFeature) As ThickenFeature
        Try
            Dim unwrapSurface As SurfaceBody = GetUnwrappedSurfaceBody(unwrapFeature)
            If unwrapSurface Is Nothing Then Return Nothing
            
            Dim thickens As ThickenFeatures = partDoc.ComponentDefinition.Features.ThickenFeatures
            
            For Each thicken As ThickenFeature In thickens
                If ThickenUsesUnwrapSurface(thicken, unwrapSurface) Then
                    UtilsLib.LogInfo("UnwrapLib: Found Thicken feature '" & thicken.Name & "' for unwrap")
                    Return thicken
                End If
            Next
            
            ' Heuristic: matching definition faces failed (API differences). Single thicken → assume unwrap chain.
            Try
                If thickens.Count = 1 Then
                    UtilsLib.LogInfo("UnwrapLib: Using sole Thicken feature '" & thickens.Item(1).Name & "' as unwrap thicken")
                    Return thickens.Item(1)
                End If
                If thickens.Count > 1 Then
                    Dim lastTh As ThickenFeature = thickens.Item(thickens.Count)
                    UtilsLib.LogWarn("UnwrapLib: Ambiguous Thicken count=" & thickens.Count & "; trying last feature '" & lastTh.Name & "'")
                    Return lastTh
                End If
            Catch
            End Try
        Catch ex As Exception
            UtilsLib.LogWarn("UnwrapLib: Error finding Thicken for unwrap: " & ex.Message)
        End Try
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Surface body used for Pinnalaotus extents: <see cref="ResolveManufacturedSolidBody"/> if set, else unwrap surface.
    ''' </summary>
    Public Function GetPinnalaotusMeasurementBody(partDoc As PartDocument) As SurfaceBody
        Try
            Dim resolved As SurfaceBody = ResolveManufacturedSolidBody(partDoc)
            If resolved IsNot Nothing Then Return resolved
            Dim unwrap As UnwrapFeature = GetUnwrapFeature(partDoc)
            If unwrap Is Nothing Then Return Nothing
            Return GetUnwrappedSurfaceBody(unwrap)
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Find a surface/solid body by display name (case-insensitive).
    ''' </summary>
    Public Function FindSurfaceBodyByName(compDef As PartComponentDefinition, bodyName As String) As SurfaceBody
        If compDef Is Nothing OrElse String.IsNullOrWhiteSpace(bodyName) Then Return Nothing
        Dim trimmed As String = bodyName.Trim()
        For Each body As SurfaceBody In compDef.SurfaceBodies
            Try
                If String.Equals(body.Name, trimmed, StringComparison.OrdinalIgnoreCase) Then Return body
            Catch
            End Try
        Next
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Stored manufactured-body hint (Thicken output, Extrude solid, etc.).
    ''' </summary>
    Public Function GetManufacturedSolidBodyNameProperty(partDoc As PartDocument) As String
        Try
            Dim propSet As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
            Dim v As Object = propSet.Item(PROP_PINNALAOTUS_BODY_NAME).Value
            If v Is Nothing Then Return ""
            Return CStr(v).Trim()
        Catch
            Return ""
        End Try
    End Function
    
    ''' <summary>
    ''' Persist manufactured solid body name for DVR resolution, dimensions, and CAM.
    ''' </summary>
    Public Sub SetManufacturedSolidBodyNameProperty(partDoc As PartDocument, bodyName As String)
        Try
            Dim propSet As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
            If String.IsNullOrWhiteSpace(bodyName) Then
                Try
                    propSet.Item(PROP_PINNALAOTUS_BODY_NAME).Value = ""
                Catch
                End Try
                Return
            End If
            Dim trimmed As String = bodyName.Trim()
            Try
                propSet.Item(PROP_PINNALAOTUS_BODY_NAME).Value = trimmed
            Catch
                propSet.Add(trimmed, PROP_PINNALAOTUS_BODY_NAME)
            End Try
        Catch ex As Exception
            UtilsLib.LogWarn("UnwrapLib: Could not set " & PROP_PINNALAOTUS_BODY_NAME & ": " & ex.Message)
        End Try
    End Sub
    
    ''' <summary>
    ''' Solid body from Unwrap+Thicken chain only (for UI default / autodetect).
    ''' </summary>
    Public Function TryGetThickenManufacturedSolidBody(partDoc As PartDocument) As SurfaceBody
        Dim unwrap As UnwrapFeature = GetUnwrapFeature(partDoc)
        If unwrap Is Nothing Then Return Nothing
        Dim thicken As ThickenFeature = GetThickenForUnwrap(partDoc, unwrap)
        If thicken Is Nothing Then Return Nothing
        Return GetThickenedSolidBody(thicken)
    End Function
    
    ''' <summary>
    ''' Manufactured flat solid for Pinnalaotus / Komponent DVRs: named body from <see cref="PROP_PINNALAOTUS_BODY_NAME"/> if valid, else Thicken output.
    ''' </summary>
    Public Function ResolveManufacturedSolidBody(partDoc As PartDocument) As SurfaceBody
        Try
            Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
            Dim propName As String = GetManufacturedSolidBodyNameProperty(partDoc)
            UtilsLib.LogInfo("UnwrapLib: BB_PinnalaotusSolidBodyName = '" & propName & "'")
            If Not String.IsNullOrWhiteSpace(propName) Then
                Dim fromProp As SurfaceBody = FindSurfaceBodyByName(compDef, propName)
                If fromProp IsNot Nothing Then
                    UtilsLib.LogInfo("UnwrapLib: Found body by property name: '" & fromProp.Name & "'")
                    Return fromProp
                Else
                    UtilsLib.LogWarn("UnwrapLib: Body '" & propName & "' not found in part")
                End If
            End If
            Return TryGetThickenManufacturedSolidBody(partDoc)
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Persist thickened body name so embedded dimension rule measures one solid only.
    ''' </summary>
    Public Sub StorePinnalaotusMeasurementBodyProperty(partDoc As PartDocument)
        Try
            Dim b As SurfaceBody = GetPinnalaotusMeasurementBody(partDoc)
            If b Is Nothing Then Return
            Dim propSet As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
            Try
                propSet.Item(PROP_PINNALAOTUS_BODY_NAME).Value = b.Name
            Catch
                propSet.Add(b.Name, PROP_PINNALAOTUS_BODY_NAME)
            End Try
        Catch ex As Exception
            UtilsLib.LogWarn("UnwrapLib: Could not store " & PROP_PINNALAOTUS_BODY_NAME & ": " & ex.Message)
        End Try
    End Sub
    
    ''' <summary>
    ''' Projected extent of one body along a direction (unit vector not required).
    ''' </summary>
    Public Function GetOrientedExtentForBody(body As SurfaceBody, dirX As Double, dirY As Double, dirZ As Double) As Double
        Dim minProj As Double = Double.MaxValue
        Dim maxProj As Double = Double.MinValue
        Try
            For Each vertex As Vertex In body.Vertices
                Dim pt As Point = vertex.Point
                Dim proj As Double = pt.X * dirX + pt.Y * dirY + pt.Z * dirZ
                If proj < minProj Then minProj = proj
                If proj > maxProj Then maxProj = proj
            Next
        Catch
        End Try
        If minProj = Double.MaxValue Then Return 0
        Return maxProj - minProj
    End Function
    
    ''' <summary>
    ''' Check if a ThickenFeature uses faces from the given surface body
    ''' </summary>
    Private Function ThickenUsesUnwrapSurface(thicken As ThickenFeature, unwrapSurface As SurfaceBody) As Boolean
        Try
            Dim collections As New List(Of Object)
            
            Try
                Dim fc As Object = thicken.ClientFaces
                If fc IsNot Nothing Then collections.Add(fc)
            Catch
            End Try
            Try
                Dim fc As Object = thicken.Faces
                If fc IsNot Nothing Then collections.Add(fc)
            Catch
            End Try
            
            Dim thickenDef As Object = Nothing
            Try
                thickenDef = thicken.Definition
            Catch
            End Try
            If thickenDef IsNot Nothing Then
                Try
                    Dim inputFaces As Object = thickenDef.FaceCollection
                    If inputFaces IsNot Nothing Then collections.Add(inputFaces)
                Catch
                End Try
                Try
                    Dim inputFaces As Object = thickenDef.Faces
                    If inputFaces IsNot Nothing Then collections.Add(inputFaces)
                Catch
                End Try
                Try
                    Dim inputFaces As Object = thickenDef.InputFaces
                    If inputFaces IsNot Nothing Then collections.Add(inputFaces)
                Catch
                End Try
            End If
            
            For Each inputFaces As Object In collections
                Try
                    For Each face As Face In inputFaces
                        If FaceBelongsToSurface(face, unwrapSurface) Then Return True
                    Next
                Catch
                End Try
            Next
        Catch
        End Try
        Return False
    End Function
    
    Private Function FaceBelongsToSurface(face As Face, unwrapSurface As SurfaceBody) As Boolean
        If face Is Nothing OrElse unwrapSurface Is Nothing Then Return False
        Try
            If face.Parent Is unwrapSurface Then Return True
        Catch
        End Try
        Try
            If CType(face.Parent, SurfaceBody).Name = unwrapSurface.Name Then Return True
        Catch
        End Try
        Return False
    End Function
    
    ''' <summary>
    ''' Get the solid body created by a ThickenFeature
    ''' </summary>
    Public Function GetThickenedSolidBody(thickenFeature As ThickenFeature) As SurfaceBody
        Try
            Dim resultBodies As SurfaceBodies = thickenFeature.SurfaceBodies
            If resultBodies.Count > 0 Then
                Return resultBodies.Item(1)
            End If
        Catch ex As Exception
            UtilsLib.LogWarn("UnwrapLib: Error getting thickened solid body: " & ex.Message)
        End Try
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Get thickness value from ThickenFeature (in cm, Inventor internal units)
    ''' </summary>
    Public Function GetThickenValue(thickenFeature As ThickenFeature) As Double
        Try
            Dim def As Object = thickenFeature.Definition
            If def Is Nothing Then Return 0
            Try
                Dim tp As Object = def.Thickness
                If tp IsNot Nothing Then Return Math.Abs(CDbl(tp.Value))
            Catch
            End Try
            Try
                Dim dp As Object = def.Distance
                If dp IsNot Nothing Then Return Math.Abs(CDbl(dp.Value))
            Catch
            End Try
            Try
                Dim od As Object = def.OffsetDistance
                If od IsNot Nothing Then Return Math.Abs(CDbl(od.Value))
            Catch
            End Try
        Catch
        End Try
        Return 0
    End Function
    
    ''' <summary>
    ''' Create a ThickenFeature from an unwrapped surface
    ''' </summary>
    ''' <param name="partDoc">Part document</param>
    ''' <param name="unwrapFeature">The UnwrapFeature to thicken</param>
    ''' <param name="thicknessCm">Thickness in cm (Inventor internal units)</param>
    ''' <returns>The created ThickenFeature, or Nothing on failure</returns>
    Public Function CreateThickenFromUnwrap(partDoc As PartDocument, _
                                            unwrapFeature As UnwrapFeature, _
                                            thicknessCm As Double) As ThickenFeature
        Try
            Dim unwrapSurface As SurfaceBody = GetUnwrappedSurfaceBody(unwrapFeature)
            If unwrapSurface Is Nothing Then
                UtilsLib.LogError("UnwrapLib: Cannot create Thicken - no unwrap surface body")
                Return Nothing
            End If
            
            ' Collect all faces from the unwrap surface
            Dim app As Inventor.Application = partDoc.Parent
            Dim faceCollection As FaceCollection = app.TransientObjects.CreateFaceCollection()
            
            For Each face As Face In unwrapSurface.Faces
                faceCollection.Add(face)
            Next
            
            If faceCollection.Count = 0 Then
                UtilsLib.LogError("UnwrapLib: Unwrap surface has no faces")
                Return Nothing
            End If
            
            UtilsLib.LogInfo("UnwrapLib: Creating Thicken with " & faceCollection.Count & " face(s), thickness=" & FormatNumber(thicknessCm * 10, 2) & " mm")
            
            Dim thickens As ThickenFeatures = partDoc.ComponentDefinition.Features.ThickenFeatures
            
            ' ThickenFeatures.Add(Faces, Distance, ExtentDirection, Operation) — no CreateDefinition in iLogic
            Dim distanceExpr As String = (thicknessCm * 10).ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) & " mm"
            Dim newThicken As ThickenFeature = thickens.Add( _
                faceCollection, _
                distanceExpr, _
                PartFeatureExtentDirectionEnum.kPositiveExtentDirection, _
                PartFeatureOperationEnum.kNewBodyOperation)
            
            ' Rename for clarity
            Try
                newThicken.Name = "Pinnalaotus_Thicken"
            Catch
            End Try
            
            UtilsLib.LogInfo("UnwrapLib: Created Thicken feature '" & newThicken.Name & "'")
            Return newThicken
            
        Catch ex As Exception
            UtilsLib.LogError("UnwrapLib: Failed to create Thicken: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ' ============================================================================
    ' Dimension Calculation
    ' ============================================================================
    
    ''' <summary>
    ''' Get dimensions from a Pinnalaotus part (thickened unwrap body). Uses oriented axes when provided,
    ''' otherwise sorts RangeBox extents (smallest = thickness, mid = width, largest = length).
    ''' </summary>
    Public Function GetPinnalaotusDimensions(partDoc As PartDocument, _
                                             ByRef thickness As Double, _
                                             ByRef width As Double, _
                                             ByRef length As Double, _
                                             Optional thicknessAxisStr As String = Nothing, _
                                             Optional widthAxisStr As String = Nothing, _
                                             Optional lengthAxisStr As String = Nothing) As Boolean
        thickness = 0
        width = 0
        length = 0
        
        Try
            Dim unwrap As UnwrapFeature = GetUnwrapFeature(partDoc)
            If unwrap Is Nothing Then Return False
            
            Dim thicken As ThickenFeature = GetThickenForUnwrap(partDoc, unwrap)
            If thicken Is Nothing Then Return False
            
            Dim thicknessCm As Double = GetThickenValue(thicken)
            
            ' Use stored body name property first, then fall back to thicken output
            Dim measBody As SurfaceBody = ResolveManufacturedSolidBody(partDoc)
            If measBody IsNot Nothing Then
                UtilsLib.LogInfo("UnwrapLib: Using resolved body '" & measBody.Name & "' for measurement")
                Try
                    Dim rb As Box = measBody.RangeBox
                    Dim rbX As Double = Math.Abs(rb.MaxPoint.X - rb.MinPoint.X) * 10
                    Dim rbY As Double = Math.Abs(rb.MaxPoint.Y - rb.MinPoint.Y) * 10
                    Dim rbZ As Double = Math.Abs(rb.MaxPoint.Z - rb.MinPoint.Z) * 10
                    UtilsLib.LogInfo("UnwrapLib: Body RangeBox - X:" & rbX.ToString("0.00") & " Y:" & rbY.ToString("0.00") & " Z:" & rbZ.ToString("0.00") & " mm")
                Catch
                End Try
            Else
                measBody = GetThickenedSolidBody(thicken)
                If measBody IsNot Nothing Then
                    UtilsLib.LogInfo("UnwrapLib: Using thicken output body '" & measBody.Name & "' for measurement")
                End If
            End If
            If measBody Is Nothing Then measBody = GetUnwrappedSurfaceBody(unwrap)
            If measBody Is Nothing Then Return False
            
            Dim tStr As String = If(thicknessAxisStr, "")
            Dim wStr As String = If(widthAxisStr, "")
            Dim lStr As String = If(lengthAxisStr, "")
            If tStr = "" Then tStr = ReadUserProp(partDoc, "BB_ThicknessAxis")
            If wStr = "" Then wStr = ReadUserProp(partDoc, "BB_WidthAxis")
            If lStr = "" Then lStr = ReadUserProp(partDoc, "BB_LengthAxis")
            
            Dim tx As Double = 0, ty As Double = 0, tz As Double = 0
            Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
            Dim lx As Double = 0, ly As Double = 0, lz As Double = 0
            Dim haveTW As Boolean = TryParseAxisVector(tStr, tx, ty, tz) AndAlso TryParseAxisVector(wStr, wx, wy, wz)
            
            If haveTW Then
                If Not TryParseAxisVector(lStr, lx, ly, lz) Then
                    lx = ty * wz - tz * wy
                    ly = tz * wx - tx * wz
                    lz = tx * wy - ty * wx
                End If
                thickness = GetOrientedExtentForBody(measBody, tx, ty, tz)
                width = GetOrientedExtentForBody(measBody, wx, wy, wz)
                length = GetOrientedExtentForBody(measBody, lx, ly, lz)
                If thicknessCm > 0.00001 Then thickness = thicknessCm
            Else
                ' Use unwrap surface normal as thickness direction for correct oriented measurement
                Dim nx As Double = 0, ny As Double = 0, nz As Double = 0
                If TryGetUnwrapFlatSurfaceNormal(unwrap, nx, ny, nz) Then
                    ' Compute perpendicular vectors for width/length
                    Dim perpWx As Double = 0, perpWy As Double = 0, perpWz As Double = 0
                    Dim perpLx As Double = 0, perpLy As Double = 0, perpLz As Double = 0
                    
                    ' Find a perpendicular vector (use cross product with a non-parallel vector)
                    If Math.Abs(nx) < 0.9 Then
                        perpWx = 0 : perpWy = -nz : perpWz = ny
                    Else
                        perpWx = -ny : perpWy = nx : perpWz = 0
                    End If
                    Dim wLen As Double = Math.Sqrt(perpWx * perpWx + perpWy * perpWy + perpWz * perpWz)
                    If wLen > 0.0001 Then
                        perpWx /= wLen : perpWy /= wLen : perpWz /= wLen
                    End If
                    
                    ' Third axis is cross product of thickness and width
                    perpLx = ny * perpWz - nz * perpWy
                    perpLy = nz * perpWx - nx * perpWz
                    perpLz = nx * perpWy - ny * perpWx
                    
                    ' Get oriented extents
                    Dim tExt As Double = GetOrientedExtentForBody(measBody, nx, ny, nz)
                    Dim wExt As Double = GetOrientedExtentForBody(measBody, perpWx, perpWy, perpWz)
                    Dim lExt As Double = GetOrientedExtentForBody(measBody, perpLx, perpLy, perpLz)
                    
                    UtilsLib.LogInfo("UnwrapLib: Normal=(" & nx.ToString("0.00") & "," & ny.ToString("0.00") & "," & nz.ToString("0.00") & ")" & _
                                     " perpW=(" & perpWx.ToString("0.00") & "," & perpWy.ToString("0.00") & "," & perpWz.ToString("0.00") & ")" & _
                                     " perpL=(" & perpLx.ToString("0.00") & "," & perpLy.ToString("0.00") & "," & perpLz.ToString("0.00") & ")")
                    UtilsLib.LogInfo("UnwrapLib: Raw extents - T:" & (tExt*10).ToString("0.00") & " W:" & (wExt*10).ToString("0.00") & " L:" & (lExt*10).ToString("0.00") & " mm")
                    
                    ' Use thicken value for thickness if available, otherwise use extent
                    If thicknessCm > 0.00001 Then
                        thickness = thicknessCm
                    Else
                        thickness = tExt
                    End If
                    
                    ' Assign width as smaller, length as larger
                    If wExt <= lExt Then
                        width = wExt
                        length = lExt
                    Else
                        width = lExt
                        length = wExt
                    End If
                    
                    UtilsLib.LogInfo("UnwrapLib: Using unwrap normal for orientation")
                Else
                    ' Fallback to axis-aligned bounding box if no unwrap normal found
                    Dim box As Box = measBody.RangeBox
                    Dim sx As Double = Math.Abs(box.MaxPoint.X - box.MinPoint.X)
                    Dim sy As Double = Math.Abs(box.MaxPoint.Y - box.MinPoint.Y)
                    Dim sz As Double = Math.Abs(box.MaxPoint.Z - box.MinPoint.Z)
                    Dim a As Double = 0, b As Double = 0, c As Double = 0
                    SortThreeExtents(sx, sy, sz, a, b, c)
                    If thicknessCm > 0.00001 Then
                        thickness = thicknessCm
                    Else
                        thickness = a
                    End If
                    width = b
                    length = c
                End If
            End If
            
            UtilsLib.LogInfo("UnwrapLib: Pinnalaotus dimensions - T:" & FormatNumber(thickness * 10, 2) & _
                             " W:" & FormatNumber(width * 10, 2) & " L:" & FormatNumber(length * 10, 2) & " mm")
            Return True
            
        Catch ex As Exception
            UtilsLib.LogWarn("UnwrapLib: Error getting Pinnalaotus dimensions: " & ex.Message)
        End Try
        
        Return False
    End Function
    
    Private Function ReadUserProp(partDoc As PartDocument, propName As String) As String
        Try
            Return CStr(partDoc.PropertySets.Item("Inventor User Defined Properties").Item(propName).Value)
        Catch
            Return ""
        End Try
    End Function
    
    Private Function TryParseAxisVector(axis As String, ByRef vx As Double, ByRef vy As Double, ByRef vz As Double) As Boolean
        vx = 0 : vy = 0 : vz = 0
        If String.IsNullOrEmpty(axis) OrElse Not axis.StartsWith("V:") Then Return False
        Try
            Dim parts() As String = axis.Substring(2).Split(","c)
            If parts.Length <> 3 Then Return False
            vx = Double.Parse(parts(0), System.Globalization.CultureInfo.InvariantCulture)
            vy = Double.Parse(parts(1), System.Globalization.CultureInfo.InvariantCulture)
            vz = Double.Parse(parts(2), System.Globalization.CultureInfo.InvariantCulture)
            Return True
        Catch
            Return False
        End Try
    End Function
    
    Private Sub SortThreeExtents(sx As Double, sy As Double, sz As Double, ByRef a As Double, ByRef b As Double, ByRef c As Double)
        Dim arr() As Double = {sx, sy, sz}
        System.Array.Sort(arr)
        a = arr(0)
        b = arr(1)
        c = arr(2)
    End Sub
    
    ''' <summary>
    ''' When Unwrap exists but Thicken is missing: preview W/L from unwrap surface RangeBox (cm).
    ''' Thickness is not available until Thicken exists.
    ''' </summary>
    Public Function TryGetUnwrapSurfacePreviewExtents(partDoc As PartDocument, _
                                                      ByRef width As Double, _
                                                      ByRef length As Double) As Boolean
        width = 0
        length = 0
        Try
            Dim unwrap As UnwrapFeature = GetUnwrapFeature(partDoc)
            If unwrap Is Nothing Then Return False
            Dim surf As SurfaceBody = GetUnwrappedSurfaceBody(unwrap)
            If surf Is Nothing Then Return False
            Dim box As Box = surf.RangeBox
            Dim sx As Double = Math.Abs(box.MaxPoint.X - box.MinPoint.X)
            Dim sy As Double = Math.Abs(box.MaxPoint.Y - box.MinPoint.Y)
            Dim sz As Double = Math.Abs(box.MaxPoint.Z - box.MinPoint.Z)
            Dim a As Double = 0, b As Double = 0, c As Double = 0
            SortThreeExtents(sx, sy, sz, a, b, c)
            width = b
            length = c
            Return True
        Catch
            Return False
        End Try
    End Function
    
    ' ============================================================================
    ' Design View Representations
    ' ============================================================================
    
    Private Function FindDesignViewByName(dvrs As DesignViewRepresentations, name As String) As DesignViewRepresentation
        For Each dvr As DesignViewRepresentation In dvrs
            If dvr.Name = name Then Return dvr
        Next
        Return Nothing
    End Function
    
    ''' <summary>Pinnalaotus or legacy BB_Pinnalaotus (both mean manufactured thickened solid for CAM).</summary>
    Private Function FindManufacturedPinnalaotusDvr(dvrs As DesignViewRepresentations) As DesignViewRepresentation
        Dim p As DesignViewRepresentation = FindDesignViewByName(dvrs, DVR_NAME_PINNALAOTUS)
        If p IsNot Nothing Then Return p
        Return FindDesignViewByName(dvrs, LEGACY_MANUFACTURED_DVR_NAME)
    End Function
    
    Private Sub UnlockAndActivateDvr(dvr As DesignViewRepresentation)
        Try
            dvr.Locked = False
        Catch
        End Try
        dvr.Activate()
    End Sub
    
    ''' <summary>Show only the thickened unwrap solid.</summary>
    Private Sub ApplyManufacturedPinnalaotusVisibility(compDef As PartComponentDefinition, thickenedBody As SurfaceBody)
        For Each body As SurfaceBody In compDef.SurfaceBodies
            Dim shouldShow As Boolean = False
            If thickenedBody IsNot Nothing Then
                If body Is thickenedBody OrElse body.Name = thickenedBody.Name Then shouldShow = True
            End If
            body.Visible = shouldShow
        Next
    End Sub
    
    ''' <summary>Hide unwrap surface and thickened flat solid; show remaining bodies (bent/original geometry).</summary>
    Private Sub ApplyKomponentBentVisibility(compDef As PartComponentDefinition, _
                                             unwrapSurface As SurfaceBody, _
                                             thickenedBody As SurfaceBody)
        For Each body As SurfaceBody In compDef.SurfaceBodies
            Dim hide As Boolean = False
            If unwrapSurface IsNot Nothing Then
                If body Is unwrapSurface OrElse body.Name = unwrapSurface.Name Then hide = True
            End If
            If thickenedBody IsNot Nothing Then
                If body Is thickenedBody OrElse body.Name = thickenedBody.Name Then hide = True
            End If
            body.Visible = Not hide
        Next
    End Sub
    
    ''' <summary>
    ''' Get or create <c>Pinnalaotus</c> DVR: **manufactured flat solid only** (Thicken or user-picked solid via <see cref="PROP_PINNALAOTUS_BODY_NAME"/>).
    ''' Recognizes legacy <c>BB_Pinnalaotus</c>. Refreshes body visibility when DVR already exists (fixes older wrong states).
    ''' Requires Unwrap and a resolved manufactured body unless <paramref name="manufacturedBody"/> is supplied.
    ''' </summary>
    Public Function GetOrCreatePinnalaotusDVR(partDoc As PartDocument, Optional manufacturedBody As SurfaceBody = Nothing) As DesignViewRepresentation
        Try
            Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
            Dim dvrs As DesignViewRepresentations = compDef.RepresentationsManager.DesignViewRepresentations
            
            Dim unwrap As UnwrapFeature = GetUnwrapFeature(partDoc)
            If unwrap Is Nothing Then
                UtilsLib.LogWarn("UnwrapLib: Cannot create """ & DVR_NAME_PINNALAOTUS & """ DVR — no UnwrapFeature")
                Return Nothing
            End If
            
            Dim solidBody As SurfaceBody = manufacturedBody
            If solidBody Is Nothing Then solidBody = ResolveManufacturedSolidBody(partDoc)
            If solidBody Is Nothing Then
                UtilsLib.LogWarn("UnwrapLib: Cannot create """ & DVR_NAME_PINNALAOTUS & """ DVR — no manufactured solid (Thicken or property " & PROP_PINNALAOTUS_BODY_NAME & ")")
                Return Nothing
            End If
            
            Dim target As DesignViewRepresentation = FindManufacturedPinnalaotusDvr(dvrs)
            Dim createdNew As Boolean = False
            
            If target Is Nothing Then
                UtilsLib.LogInfo("UnwrapLib: Creating DVR """ & DVR_NAME_PINNALAOTUS & """ (manufactured solid)")
                target = dvrs.Add(DVR_NAME_PINNALAOTUS)
                createdNew = True
            Else
                UtilsLib.LogInfo("UnwrapLib: Updating manufactured DVR """ & target.Name & """")
            End If
            
            UnlockAndActivateDvr(target)
            ApplyManufacturedPinnalaotusVisibility(compDef, solidBody)
            target.Locked = True
            ActivateDefaultOrMasterDesignView(partDoc)
            
            If createdNew Then
                UtilsLib.LogInfo("UnwrapLib: Created DVR """ & DVR_NAME_PINNALAOTUS & """")
            End If
            Return target
        Catch ex As Exception
            UtilsLib.LogError("UnwrapLib: Error with """ & DVR_NAME_PINNALAOTUS & """ DVR: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Get or create <c>Komponent</c> DVR: bent/original bodies visible; unwrap surface and manufactured flat solid hidden.
    ''' If the part has no Unwrap, all bodies are shown (single-state component).
    ''' Refreshes visibility when DVR already exists.
    ''' </summary>
    Public Function GetOrCreateKomponentDVR(partDoc As PartDocument, Optional manufacturedBody As SurfaceBody = Nothing) As DesignViewRepresentation
        Try
            Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
            Dim dvrs As DesignViewRepresentations = compDef.RepresentationsManager.DesignViewRepresentations
            
            Dim unwrap As UnwrapFeature = GetUnwrapFeature(partDoc)
            Dim unwrapSurface As SurfaceBody = Nothing
            Dim thickenedBody As SurfaceBody = Nothing
            If unwrap IsNot Nothing Then
                unwrapSurface = GetUnwrappedSurfaceBody(unwrap)
                thickenedBody = manufacturedBody
                If thickenedBody Is Nothing Then thickenedBody = ResolveManufacturedSolidBody(partDoc)
            End If
            
            Dim target As DesignViewRepresentation = FindDesignViewByName(dvrs, DVR_NAME_KOMPONENT)
            Dim createdNew As Boolean = False
            
            If target Is Nothing Then
                UtilsLib.LogInfo("UnwrapLib: Creating DVR """ & DVR_NAME_KOMPONENT & """ (bent component)")
                target = dvrs.Add(DVR_NAME_KOMPONENT)
                createdNew = True
            Else
                UtilsLib.LogInfo("UnwrapLib: Updating Komponent DVR """ & target.Name & """")
            End If
            
            UnlockAndActivateDvr(target)
            If unwrap Is Nothing Then
                For Each body As SurfaceBody In compDef.SurfaceBodies
                    body.Visible = True
                Next
            Else
                ApplyKomponentBentVisibility(compDef, unwrapSurface, thickenedBody)
            End If
            
            target.Locked = True
            ActivateDefaultOrMasterDesignView(partDoc)
            
            If createdNew Then
                UtilsLib.LogInfo("UnwrapLib: Created DVR """ & DVR_NAME_KOMPONENT & """")
            End If
            Return target
        Catch ex As Exception
            UtilsLib.LogError("UnwrapLib: Error with """ & DVR_NAME_KOMPONENT & """ DVR: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' True when Unwrap exists and a manufactured solid is available (Thicken and/or <see cref="PROP_PINNALAOTUS_BODY_NAME"/>).
    ''' </summary>
    Public Function HasCompletePinnalaotus(partDoc As PartDocument) As Boolean
        If GetUnwrapFeature(partDoc) Is Nothing Then Return False
        Return ResolveManufacturedSolidBody(partDoc) IsNot Nothing
    End Function
    
    ''' <summary>
    ''' Check if part has Unwrap but is missing Thicken
    ''' </summary>
    Public Function HasUnwrapWithoutThicken(partDoc As PartDocument) As Boolean
        Dim unwrap As UnwrapFeature = GetUnwrapFeature(partDoc)
        If unwrap Is Nothing Then Return False
        
        Dim thicken As ThickenFeature = GetThickenForUnwrap(partDoc, unwrap)
        Return thicken Is Nothing
    End Function
    
    ' ============================================================================
    ' Utility Functions
    ' ============================================================================
    
    ''' <summary>
    ''' Set dimension source property on part
    ''' </summary>
    Public Sub SetDimensionSourceProperty(partDoc As PartDocument, source As String)
        Try
            Dim propSet As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
            Try
                propSet.Item(PROP_DIMENSION_SOURCE).Value = source
            Catch
                propSet.Add(source, PROP_DIMENSION_SOURCE)
            End Try
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Get dimension source property from part
    ''' </summary>
    Public Function GetDimensionSourceProperty(partDoc As PartDocument) As String
        Try
            Dim propSet As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
            Return CStr(propSet.Item(PROP_DIMENSION_SOURCE).Value)
        Catch
            Return ""
        End Try
    End Function
    
    ''' <summary>
    ''' True when 1:1 CAM drawings should use the <c>Pinnalaotus</c> DVR (thickened manufactured solid).
    ''' False when user chose Normal or Lehtmetall dimensioning (full model / flat pattern semantics).
    ''' When <paramref name="force"/> is True, uses Pinnalaotus DVR whenever Unwrap and a manufactured solid exist (e.g. Detailid/Pinnalaotuse vaated.vb).
    ''' </summary>
    Public Function ShouldUsePinnalaotusDvrInDrawing(partDoc As PartDocument, Optional force As Boolean = False) As Boolean
        If Not HasCompletePinnalaotus(partDoc) Then Return False
        If force Then Return True
        Dim src As String = GetDimensionSourceProperty(partDoc)
        If src = DIMENSION_SOURCE_NORMAL OrElse src = DIMENSION_SOURCE_LEHTMETALL Then Return False
        Return True
    End Function
    
    ''' <summary>
    ''' Restore Default or Master design view after temporarily activating Pinnalaotus (manufactured) DVR for drawing creation.
    ''' </summary>
    Public Sub ActivateDefaultOrMasterDesignView(partDoc As PartDocument)
        Try
            Dim dvrs As DesignViewRepresentations = partDoc.ComponentDefinition.RepresentationsManager.DesignViewRepresentations
            Try
                dvrs.Item("Default").Activate()
                Return
            Catch
            End Try
            Try
                dvrs.Item("Master").Activate()
            Catch
            End Try
        Catch
        End Try
    End Sub
    
    Private Function FormatNumber(value As Double, decimals As Integer) As String
        Return value.ToString("F" & decimals.ToString(), System.Globalization.CultureInfo.InvariantCulture)
    End Function

End Module
