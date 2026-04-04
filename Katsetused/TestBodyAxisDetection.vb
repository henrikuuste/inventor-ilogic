' ============================================================================
' TestBodyAxisDetection - Test axis detection on individual solid bodies
' 
' Tests:
' - Can we iterate SurfaceBodies in a multi-body part?
' - Does axis detection logic work per-body (adapted from AutoDetectAxesFromGeometry)?
' - Are T/W/L values calculated correctly for each body?
'
' Usage: Open a multi-body part with 2+ solid bodies, then run this rule.
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    Logger.Info("TestBodyAxisDetection: Starting body axis detection tests...")
    
    ' Validate document type
    If doc Is Nothing Then
        Logger.Error("TestBodyAxisDetection: No active document.")
        MessageBox.Show("Aktiivne dokument puudub.", "TestBodyAxisDetection")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        Logger.Error("TestBodyAxisDetection: Active document is not a part file.")
        MessageBox.Show("See reegel töötab ainult detaili failidega (.ipt).", "TestBodyAxisDetection")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    
    ' Check for solid bodies
    Dim solidBodies As New System.Collections.Generic.List(Of SurfaceBody)
    For Each body As SurfaceBody In compDef.SurfaceBodies
        If body.IsSolid Then
            solidBodies.Add(body)
        End If
    Next
    
    If solidBodies.Count = 0 Then
        Logger.Error("TestBodyAxisDetection: No solid bodies found in part.")
        MessageBox.Show("Detailis ei ole tahkkehasid.", "TestBodyAxisDetection")
        Exit Sub
    End If
    
    Logger.Info("TestBodyAxisDetection: Found " & solidBodies.Count & " solid body(ies)")
    Logger.Info("TestBodyAxisDetection: ========================================")
    
    ' Process each body
    For i As Integer = 0 To solidBodies.Count - 1
        Dim body As SurfaceBody = solidBodies(i)
        Logger.Info("TestBodyAxisDetection: Body " & (i + 1) & ": '" & body.Name & "'")
        
        ' Detect axes for this body
        Dim thicknessAxis As String = ""
        Dim widthAxis As String = ""
        Dim lengthAxis As String = ""
        Dim thicknessVal As Double = 0
        Dim widthVal As Double = 0
        Dim lengthVal As Double = 0
        
        Dim success As Boolean = AutoDetectAxesForBody(body, thicknessAxis, widthAxis, lengthAxis, _
                                                        thicknessVal, widthVal, lengthVal)
        
        If success Then
            Logger.Info("TestBodyAxisDetection:   Thickness axis: " & thicknessAxis & " = " & FormatDimension(thicknessVal))
            Logger.Info("TestBodyAxisDetection:   Width axis:     " & widthAxis & " = " & FormatDimension(widthVal))
            Logger.Info("TestBodyAxisDetection:   Length axis:    " & lengthAxis & " = " & FormatDimension(lengthVal))
        Else
            Logger.Warn("TestBodyAxisDetection:   Could not detect axes for this body")
        End If
        
        Logger.Info("TestBodyAxisDetection: ----------------------------------------")
    Next
    
    Logger.Info("TestBodyAxisDetection: All body detection tests completed!")
    MessageBox.Show("Kehadetektsioon lõpetatud. Vaata iLogic logi tulemuste jaoks.", "TestBodyAxisDetection")
End Sub

Function FormatDimension(value As Double) As String
    ' Convert from cm to mm and format
    Return FormatNumber(value * 10, 2) & " mm"
End Function

' Adapted from BoundingBoxStockLib.AutoDetectAxesFromGeometry to work on a single SurfaceBody
Function AutoDetectAxesForBody(body As SurfaceBody, _
                               ByRef thicknessAxis As String, ByRef widthAxis As String, ByRef lengthAxis As String, _
                               ByRef thicknessVal As Double, ByRef widthVal As Double, ByRef lengthVal As Double) As Boolean
    
    ' Collect unique face normals and find the one with smallest extent
    Dim bestNormalX As Double = 0, bestNormalY As Double = 0, bestNormalZ As Double = 0
    Dim minExtent As Double = Double.MaxValue
    Dim foundNormal As Boolean = False
    
    ' Track normals we've already checked (to avoid duplicates like top/bottom faces)
    Dim checkedNormals As New System.Collections.Generic.List(Of String)
    
    Try
        For Each face As Face In body.Faces
            Dim nx As Double = 0, ny As Double = 0, nz As Double = 0
            If GetFaceNormal(face, nx, ny, nz) Then
                ' Normalize
                Dim len As Double = Math.Sqrt(nx * nx + ny * ny + nz * nz)
                If len > 0.0001 Then
                    nx /= len : ny /= len : nz /= len
                End If
                
                ' Make normal canonical (always point in "positive" direction)
                If nx < -0.0001 OrElse (Math.Abs(nx) < 0.0001 AndAlso ny < -0.0001) OrElse _
                   (Math.Abs(nx) < 0.0001 AndAlso Math.Abs(ny) < 0.0001 AndAlso nz < -0.0001) Then
                    nx = -nx : ny = -ny : nz = -nz
                End If
                
                ' Create key for this normal direction
                Dim normalKey As String = Math.Round(nx, 3).ToString() & "," & _
                                          Math.Round(ny, 3).ToString() & "," & _
                                          Math.Round(nz, 3).ToString()
                
                If checkedNormals.Contains(normalKey) Then Continue For
                checkedNormals.Add(normalKey)
                
                ' Calculate extent along this normal for this body only
                Dim extent As Double = GetOrientedExtentForBody(body, nx, ny, nz)
                
                If extent > 0 AndAlso extent < minExtent Then
                    minExtent = extent
                    bestNormalX = nx
                    bestNormalY = ny
                    bestNormalZ = nz
                    foundNormal = True
                End If
            End If
        Next
    Catch ex As Exception
        Logger.Warn("TestBodyAxisDetection: Exception in axis detection: " & ex.Message)
        Return False
    End Try
    
    If Not foundNormal Then Return False
    
    ' Convert to axis string
    Dim dotX As Double = Math.Abs(bestNormalX)
    Dim dotY As Double = Math.Abs(bestNormalY)
    Dim dotZ As Double = Math.Abs(bestNormalZ)
    
    If dotX > 0.9998 Then
        thicknessAxis = "X"
    ElseIf dotY > 0.9998 Then
        thicknessAxis = "Y"
    ElseIf dotZ > 0.9998 Then
        thicknessAxis = "Z"
    Else
        thicknessAxis = VectorToString(bestNormalX, bestNormalY, bestNormalZ)
    End If
    
    ' Calculate thickness value
    thicknessVal = minExtent
    
    ' Compute perpendicular vectors for width and length
    Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
    Dim lx As Double = 0, ly As Double = 0, lz As Double = 0
    ComputePerpendicularVectors(bestNormalX, bestNormalY, bestNormalZ, wx, wy, wz, lx, ly, lz)
    
    ' Measure extents
    Dim widthExtent As Double = GetOrientedExtentForBody(body, wx, wy, wz)
    Dim lengthExtent As Double = GetOrientedExtentForBody(body, lx, ly, lz)
    
    ' Assign width (smaller) and length (larger)
    If lengthExtent >= widthExtent Then
        widthVal = widthExtent
        lengthVal = lengthExtent
        widthAxis = VectorToString(wx, wy, wz)
        lengthAxis = VectorToString(lx, ly, lz)
    Else
        widthVal = lengthExtent
        lengthVal = widthExtent
        widthAxis = VectorToString(lx, ly, lz)
        lengthAxis = VectorToString(wx, wy, wz)
    End If
    
    ' Simplify axis strings if aligned to principal axes
    widthAxis = SimplifyAxis(widthAxis)
    lengthAxis = SimplifyAxis(lengthAxis)
    
    Return True
End Function

Function SimplifyAxis(axis As String) As String
    If Not axis.StartsWith("V:") Then Return axis
    
    Dim vx As Double = 0, vy As Double = 0, vz As Double = 0
    ParseVectorComponents(axis, vx, vy, vz)
    
    If Math.Abs(vx) > 0.9998 Then Return "X"
    If Math.Abs(vy) > 0.9998 Then Return "Y"
    If Math.Abs(vz) > 0.9998 Then Return "Z"
    
    Return axis
End Function

Function GetOrientedExtentForBody(body As SurfaceBody, dirX As Double, dirY As Double, dirZ As Double) As Double
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

Function GetFaceNormal(face As Face, ByRef nx As Double, ByRef ny As Double, ByRef nz As Double) As Boolean
    Try
        Dim geom As Object = face.Geometry
        If TypeOf geom Is Plane Then
            Dim plane As Plane = CType(geom, Plane)
            Dim normal As UnitVector = plane.Normal
            nx = normal.X
            ny = normal.Y
            nz = normal.Z
            Return True
        End If
    Catch
    End Try
    Return False
End Function

Sub ComputePerpendicularVectors(nx As Double, ny As Double, nz As Double, _
                                ByRef wx As Double, ByRef wy As Double, ByRef wz As Double, _
                                ByRef lx As Double, ByRef ly As Double, ByRef lz As Double)
    ' Find a vector not parallel to normal
    Dim refX As Double = 1, refY As Double = 0, refZ As Double = 0
    Dim dot As Double = nx * refX + ny * refY + nz * refZ
    If Math.Abs(dot) > 0.9 Then
        refX = 0 : refY = 1 : refZ = 0
    End If
    
    ' Cross product: w = n x ref
    wx = ny * refZ - nz * refY
    wy = nz * refX - nx * refZ
    wz = nx * refY - ny * refX
    
    ' Normalize w
    Dim wLen As Double = Math.Sqrt(wx * wx + wy * wy + wz * wz)
    If wLen > 0.0001 Then
        wx /= wLen : wy /= wLen : wz /= wLen
    End If
    
    ' Cross product: l = n x w
    lx = ny * wz - nz * wy
    ly = nz * wx - nx * wz
    lz = nx * wy - ny * wx
    
    ' Normalize l
    Dim lLen As Double = Math.Sqrt(lx * lx + ly * ly + lz * lz)
    If lLen > 0.0001 Then
        lx /= lLen : ly /= lLen : lz /= lLen
    End If
End Sub

Function VectorToString(vx As Double, vy As Double, vz As Double) As String
    Return "V:" & vx.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture) & "," & _
                  vy.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture) & "," & _
                  vz.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture)
End Function

Function ParseVectorComponents(axis As String, ByRef vx As Double, ByRef vy As Double, ByRef vz As Double) As Boolean
    If Not axis.StartsWith("V:") Then Return False
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
