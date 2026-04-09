' ============================================================================
' CenterPatternLib - Center-Based Occurrence Pattern Library
' 
' Creates parametric occurrence patterns that distribute instances evenly 
' across a span, with the seed assumed to be initially centered.
'
' Features:
' - Uniform or symmetric-from-center distribution
' - Include/exclude instances at span boundaries  
' - Constraint-based seed positioning (updates with geometry changes)
' - Parametric spacing via assembly parameters with formulas
'
' Depends on: UtilsLib.vb, GeoLib.vb, WorkFeatureLib.vb, PatternLib.vb
'
' Usage: 
'   AddVbFile "Lib/UtilsLib.vb"
'   AddVbFile "Lib/GeoLib.vb"
'   AddVbFile "Lib/WorkFeatureLib.vb"
'   AddVbFile "Lib/PatternLib.vb"
'   AddVbFile "Lib/CenterPatternLib.vb"
'
' Estonian naming convention for created features:
'   <Nimi>_Ulatus      - span distance parameter (mm)
'   <Nimi>_MaxVahe     - max spacing parameter (mm)
'   <Nimi>_Arv         - total instance count (unitless)
'   <Nimi>_Samm        - actual spacing (mm)
'   <Nimi>_Nihe        - offset from start (mm)
'   <Nimi>_KeskNihe    - center offset of original part (mm)
'   <Nimi>_AlgusTasand - start boundary work plane
'   <Nimi>_LõpuTasand  - end boundary work plane
'   <Nimi>_Telg        - pattern direction work axis
'   <Nimi>_Muster      - the occurrence pattern
'   <Nimi>_Asend       - positioning constraint
'
' ============================================================================

Option Strict Off
Imports Inventor

Public Module CenterPatternLib

    ' ============================================================================
    ' SECTION 1: Constants and Modes
    ' ============================================================================
    
    ' Distribution modes
    Public Const MODE_UNIFORM As String = "UNIFORM"
    Public Const MODE_SYMMETRIC As String = "SYMMETRIC"
    
    ' Attribute set name for storing pattern configuration
    Private Const ATTR_SET_NAME As String = "CenterPattern"

    ' ============================================================================
    ' SECTION 2: Formula Builders
    ' ============================================================================
    
    ''' <summary>
    ''' Build the count formula based on mode and ends option.
    ''' 
    ''' maxSpacing is the MAXIMUM allowed gap between supports.
    ''' We use ceiling() to ensure we have enough instances so all gaps ≤ maxSpacing.
    ''' 
    ''' For Uniform WITHOUT endpoints (N instances, N+1 gaps):
    '''   spacing = span/(N+1), need spacing ≤ maxSpacing
    '''   N ≥ ceiling(span/maxSpacing) - 1
    ''' 
    ''' For Uniform WITH endpoints (N instances, N-1 gaps):
    '''   spacing = span/(N-1), need spacing ≤ maxSpacing
    '''   N ≥ ceiling(span/maxSpacing) + 1
    ''' 
    ''' For Symmetric: same logic but applied to half span, ensuring odd count.
    ''' </summary>
    Public Function BuildCountFormula(spanParam As String, maxSpacingParam As String, _
                                       mode As String, includeEnds As Boolean) As String
        If mode = MODE_SYMMETRIC Then
            If includeEnds Then
                ' Symmetric with ends: 2 * ceil(span/2 / maxSpacing) + 1
                ' k gaps per half-span, k = ceil((span/2) / maxSpacing)
                ' Total = 2k + 1 (center + k on each side including endpoints)
                Return "2 * ceil(" & spanParam & " / 2 / " & maxSpacingParam & ") + 1"
            Else
                ' Symmetric without ends: 2 * max(0; ceil(span/2 / maxSpacing) - 1) + 1
                ' k+1 gaps per half-span (including boundary), k = ceil((span/2) / maxSpacing) - 1
                ' Total = 2k + 1 (center + k on each side)
                Return "2 * max(0; ceil(" & spanParam & " / 2 / " & maxSpacingParam & ") - 1) + 1"
            End If
        Else ' MODE_UNIFORM
            If includeEnds Then
                ' Uniform with ends: ceil(span / maxSpacing) + 1
                ' N-1 gaps between N instances, need N-1 ≥ span/maxSpacing
                Return "ceil(" & spanParam & " / " & maxSpacingParam & ") + 1"
            Else
                ' Uniform without ends: max(1; ceil(span / maxSpacing) - 1)
                ' N+1 gaps (including boundaries), need N+1 ≥ span/maxSpacing
                Return "max(1; ceil(" & spanParam & " / " & maxSpacingParam & ") - 1)"
            End If
        End If
    End Function
    
    ''' <summary>
    ''' Build the spacing formula based on mode and ends option.
    ''' </summary>
    Public Function BuildSpacingFormula(spanParam As String, countParam As String, _
                                         mode As String, includeEnds As Boolean) As String
        If includeEnds Then
            ' With ends: spacing = span / (count - 1)
            ' Guard against division by zero when count = 1
            Return spanParam & " / max(1; " & countParam & " - 1)"
        Else
            ' Without ends: spacing = span / (count + 1)
            Return spanParam & " / (" & countParam & " + 1)"
        End If
    End Function
    
    ''' <summary>
    ''' Build the offset formula (distance from start plane to first instance).
    ''' For symmetric mode without endpoints, calculates from center to maintain symmetry when span changes.
    ''' </summary>
    Public Function BuildOffsetFormula(spacingParam As String, includeEnds As Boolean, _
                                        mode As String, spanParam As String, countParam As String) As String
        If includeEnds Then
            ' With ends: first instance at start plane (offset = 0)
            Return "0 mm"
        Else
            If mode = "Symmetric" Then
                ' Symmetric without ends: first instance relative to center
                ' offset = span/2 - floor((count-1)/2) * spacing
                ' Since count is always odd (2k+1), floor((count-1)/2) = k
                ' This ensures when span changes, the pattern stays symmetric around span/2
                Return spanParam & " / 2 - floor((" & countParam & " - 1) / 2) * " & spacingParam
            Else
                ' Uniform without ends: first instance offset by one spacing
                Return spacingParam
            End If
        End If
    End Function
    
    ''' <summary>
    ''' Build the effective span formula.
    ''' effectiveSpan = span - startOffset - endOffset
    ''' </summary>
    Public Function BuildEffectiveSpanFormula(spanParam As String, _
                                               startOffsetParam As String, _
                                               endOffsetParam As String) As String
        Return spanParam & " - " & startOffsetParam & " - " & endOffsetParam
    End Function
    
    ''' <summary>
    ''' Build the count formula using effective span (span minus offsets).
    ''' </summary>
    Public Function BuildCountFormulaWithOffsets(effectiveSpanParam As String, maxSpacingParam As String, _
                                                  mode As String, includeEnds As Boolean) As String
        If mode = MODE_SYMMETRIC Then
            If includeEnds Then
                Return "2 * ceil(" & effectiveSpanParam & " / 2 / " & maxSpacingParam & ") + 1"
            Else
                Return "2 * max(0; ceil(" & effectiveSpanParam & " / 2 / " & maxSpacingParam & ") - 1) + 1"
            End If
        Else ' MODE_UNIFORM
            If includeEnds Then
                Return "ceil(" & effectiveSpanParam & " / " & maxSpacingParam & ") + 1"
            Else
                Return "max(1; ceil(" & effectiveSpanParam & " / " & maxSpacingParam & ") - 1)"
            End If
        End If
    End Function
    
    ''' <summary>
    ''' Build the spacing formula using effective span.
    ''' </summary>
    Public Function BuildSpacingFormulaWithOffsets(effectiveSpanParam As String, countParam As String, _
                                                    mode As String, includeEnds As Boolean) As String
        If includeEnds Then
            Return effectiveSpanParam & " / max(1; " & countParam & " - 1)"
        Else
            Return effectiveSpanParam & " / (" & countParam & " + 1)"
        End If
    End Function
    
    ''' <summary>
    ''' Build the first instance offset formula (from start plane to first instance).
    ''' This now includes the start offset.
    ''' </summary>
    Public Function BuildOffsetFormulaWithStartOffset(spacingParam As String, startOffsetParam As String, _
                                                       includeEnds As Boolean, mode As String, _
                                                       effectiveSpanParam As String, countParam As String) As String
        If includeEnds Then
            ' With ends: first instance at start plane + startOffset
            Return startOffsetParam
        Else
            If mode = "Symmetric" Then
                ' Symmetric without ends: startOffset + effectiveSpan/2 - k*spacing
                Return startOffsetParam & " + " & effectiveSpanParam & " / 2 - floor((" & countParam & " - 1) / 2) * " & spacingParam
            Else
                ' Uniform without ends: first instance at startOffset + spacing
                Return startOffsetParam & " + " & spacingParam
            End If
        End If
    End Function

    ' ============================================================================
    ' SECTION 3: Parameter Management
    ' ============================================================================
    
    ''' <summary>
    ''' Create or update a user parameter with a numeric value.
    ''' </summary>
    Public Function SetParameter(asmDoc As AssemblyDocument, paramName As String, _
                                  value As Double, Optional units As String = "mm") As Parameter
        Dim params As Parameters = asmDoc.ComponentDefinition.Parameters
        Dim expression As String = value.ToString(System.Globalization.CultureInfo.InvariantCulture) & " " & units
        
        Try
            Dim param As Parameter = params.Item(paramName)
            param.Expression = expression
            Return param
        Catch
            Try
                Return params.UserParameters.AddByExpression(paramName, expression, UnitsTypeEnum.kDefaultDisplayLengthUnits)
            Catch
                Return Nothing
            End Try
        End Try
    End Function
    
    ''' <summary>
    ''' Create or update a parameter with a formula expression.
    ''' </summary>
    Public Function SetParameterFormula(asmDoc As AssemblyDocument, paramName As String, _
                                         formula As String) As Parameter
        Dim params As Parameters = asmDoc.ComponentDefinition.Parameters
        
        Try
            Dim param As Parameter = params.Item(paramName)
            param.Expression = formula
            Return param
        Catch
            Try
                Return params.UserParameters.AddByExpression(paramName, formula, UnitsTypeEnum.kDefaultDisplayLengthUnits)
            Catch ex1 As Exception
                Try
                    Return params.UserParameters.AddByExpression(paramName, formula, UnitsTypeEnum.kUnitlessUnits)
                Catch
                    Return Nothing
                End Try
            End Try
        End Try
    End Function
    
    ''' <summary>
    ''' Create or update a UNITLESS parameter with a formula.
    ''' </summary>
    Public Function SetUnitlessFormula(asmDoc As AssemblyDocument, paramName As String, _
                                        formula As String) As Parameter
        Dim params As Parameters = asmDoc.ComponentDefinition.Parameters
        
        Try
            Dim param As Parameter = params.Item(paramName)
            param.Expression = formula
            Return param
        Catch
            Try
                Return params.UserParameters.AddByExpression(paramName, formula, UnitsTypeEnum.kUnitlessUnits)
            Catch
                Try
                    Return params.UserParameters.AddByExpression(paramName, formula, UnitsTypeEnum.kDefaultDisplayLengthUnits)
                Catch
                    Return Nothing
                End Try
            End Try
        End Try
    End Function
    
    ''' <summary>
    ''' Get parameter value in cm (internal units).
    ''' </summary>
    Public Function GetParameterValue(asmDoc As AssemblyDocument, paramName As String) As Double
        Try
            Return asmDoc.ComponentDefinition.Parameters.Item(paramName).Value
        Catch
            Return 0
        End Try
    End Function
    
    ''' <summary>
    ''' Get list of user parameter names (for dropdown population).
    ''' </summary>
    Public Function GetUserParameterNames(asmDoc As AssemblyDocument) As String()
        Dim names As New System.Collections.Generic.List(Of String)
        
        Try
            Dim allParams As Parameters = asmDoc.ComponentDefinition.Parameters
            
            For Each param As Parameter In allParams
                Try
                    If param.ParameterType = ParameterTypeEnum.kModelParameter Then
                        Continue For
                    End If
                    If Not names.Contains(param.Name) Then
                        names.Add(param.Name)
                    End If
                Catch
                End Try
            Next
        Catch
        End Try
        
        names.Sort()
        Return names.ToArray()
    End Function

    ' ============================================================================
    ' SECTION 4: Geometry-to-Principal Offset Calculation
    ' ============================================================================
    
    ''' <summary>
    ''' Calculate the offset from geometry center to principal plane origin.
    ''' 
    ''' This is an INTRINSIC property of the part - how far the principal plane
    ''' (which we constrain) is from the actual geometry center (which we care about).
    ''' 
    ''' The principal plane can be anywhere - it doesn't need to intersect the geometry.
    ''' This offset allows us to correctly position the geometry even though we
    ''' constrain the principal plane.
    ''' 
    ''' Returns: signed distance from geometry center to principal plane origin,
    '''          projected onto the axis direction. In cm (internal units).
    '''          
    ''' Usage: To place geometry center at position P from start plane,
    '''        use constraint offset = P + GeomToPrincipalOffset
    ''' </summary>
    Public Function CalculateGeomToPrincipalOffset(app As Inventor.Application, _
                                                    occ As ComponentOccurrence, _
                                                    axisDirection As UnitVector, _
                                                    Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing) As Double
        If occ Is Nothing OrElse axisDirection Is Nothing Then
            Return 0
        End If
        
        If logs IsNot Nothing Then
            logs.Add("CenterPatternLib: === GEOM-TO-PRINCIPAL OFFSET CALCULATION ===")
            logs.Add("CenterPatternLib: Occurrence: " & occ.Name)
            logs.Add("CenterPatternLib: Axis direction: (" & _
                     axisDirection.X.ToString("0.0000") & ", " & _
                     axisDirection.Y.ToString("0.0000") & ", " & _
                     axisDirection.Z.ToString("0.0000") & ")")
        End If
        
        Try
            ' Get geometry center from bounding box
            Dim rangeBox As Box = occ.RangeBox
            Dim geomCenterX As Double = (rangeBox.MinPoint.X + rangeBox.MaxPoint.X) / 2
            Dim geomCenterY As Double = (rangeBox.MinPoint.Y + rangeBox.MaxPoint.Y) / 2
            Dim geomCenterZ As Double = (rangeBox.MinPoint.Z + rangeBox.MaxPoint.Z) / 2
            
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: RangeBox min: (" & _
                         (rangeBox.MinPoint.X * 10).ToString("0.00") & ", " & _
                         (rangeBox.MinPoint.Y * 10).ToString("0.00") & ", " & _
                         (rangeBox.MinPoint.Z * 10).ToString("0.00") & ") mm")
                logs.Add("CenterPatternLib: RangeBox max: (" & _
                         (rangeBox.MaxPoint.X * 10).ToString("0.00") & ", " & _
                         (rangeBox.MaxPoint.Y * 10).ToString("0.00") & ", " & _
                         (rangeBox.MaxPoint.Z * 10).ToString("0.00") & ") mm")
                logs.Add("CenterPatternLib: Geometry center: (" & _
                         (geomCenterX * 10).ToString("0.00") & ", " & _
                         (geomCenterY * 10).ToString("0.00") & ", " & _
                         (geomCenterZ * 10).ToString("0.00") & ") mm")
            End If
            
            ' Find the principal plane of the occurrence (perpendicular to axis)
            Dim principalPlaneProxy As Object = GeoLib.FindPrincipalPlane(occ, axisDirection)
            If principalPlaneProxy Is Nothing Then
                If logs IsNot Nothing Then logs.Add("CenterPatternLib: ERROR - Could not find principal plane")
                Return 0
            End If
            
            Dim principalGeom As Plane = GeoLib.GetPlaneGeometry(principalPlaneProxy)
            If principalGeom Is Nothing Then
                If logs IsNot Nothing Then logs.Add("CenterPatternLib: ERROR - Could not get principal plane geometry")
                Return 0
            End If
            
            Dim planeOriginX As Double = principalGeom.RootPoint.X
            Dim planeOriginY As Double = principalGeom.RootPoint.Y
            Dim planeOriginZ As Double = principalGeom.RootPoint.Z
            
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: Principal plane origin: (" & _
                         (planeOriginX * 10).ToString("0.00") & ", " & _
                         (planeOriginY * 10).ToString("0.00") & ", " & _
                         (planeOriginZ * 10).ToString("0.00") & ") mm")
                logs.Add("CenterPatternLib: Principal plane normal: (" & _
                         principalGeom.Normal.X.ToString("0.0000") & ", " & _
                         principalGeom.Normal.Y.ToString("0.0000") & ", " & _
                         principalGeom.Normal.Z.ToString("0.0000") & ")")
            End If
            
            ' Vector from geometry center to principal plane origin
            Dim toPlaneX As Double = planeOriginX - geomCenterX
            Dim toPlaneY As Double = planeOriginY - geomCenterY
            Dim toPlaneZ As Double = planeOriginZ - geomCenterZ
            
            ' Project onto axis (signed distance)
            Dim geomToPrincipalOffset As Double = toPlaneX * axisDirection.X + _
                                                   toPlaneY * axisDirection.Y + _
                                                   toPlaneZ * axisDirection.Z
            
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: Vector from geom center to plane origin: (" & _
                         (toPlaneX * 10).ToString("0.00") & ", " & _
                         (toPlaneY * 10).ToString("0.00") & ", " & _
                         (toPlaneZ * 10).ToString("0.00") & ") mm")
                logs.Add("CenterPatternLib: Geom-to-principal offset (along axis): " & _
                         (geomToPrincipalOffset * 10).ToString("0.00") & " mm")
                logs.Add("CenterPatternLib: === END GEOM-TO-PRINCIPAL OFFSET CALCULATION ===")
            End If
            
            Return geomToPrincipalOffset
            
        Catch ex As Exception
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: ERROR calculating geom-to-principal offset: " & ex.Message)
            End If
            Return 0
        End Try
    End Function
    
    ''' <summary>
    ''' Legacy wrapper - calls CalculateGeomToPrincipalOffset for backward compatibility.
    ''' The spanCm and startPlane parameters are ignored as they're not needed for the new calculation.
    ''' </summary>
    Public Function CalculateCenterOffset(app As Inventor.Application, _
                                           occ As ComponentOccurrence, _
                                           startPlane As WorkPlane, _
                                           axisDirection As UnitVector, _
                                           spanCm As Double, _
                                           Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing) As Double
        ' The new calculation is independent of start plane and span
        Return CalculateGeomToPrincipalOffset(app, occ, axisDirection, logs)
    End Function

    ' ============================================================================
    ' SECTION 4B: Part Extent Calculation
    ' ============================================================================
    
    ' Alignment modes for start and end instances
    Public Const ALIGN_CENTER As String = "CENTER"
    Public Const ALIGN_INWARD As String = "INWARD"
    Public Const ALIGN_OUTWARD As String = "OUTWARD"
    
    ''' <summary>
    ''' Calculate the extent of a part along a given axis direction.
    ''' Projects all 8 corners of the occurrence's bounding box onto the axis.
    ''' 
    ''' Returns:
    ''' - minExtent: distance from principal plane to the "back" edge (negative = behind principal plane)
    ''' - maxExtent: distance from principal plane to the "front" edge (positive = in front of principal plane)
    ''' - halfExtent: half of total part thickness along axis = (maxExtent - minExtent) / 2
    ''' 
    ''' All values in internal units (cm).
    ''' </summary>
    Public Sub CalculatePartExtentAlongAxis(app As Inventor.Application, _
                                             occ As ComponentOccurrence, _
                                             axisDirection As UnitVector, _
                                             ByRef minExtent As Double, _
                                             ByRef maxExtent As Double, _
                                             ByRef halfExtent As Double, _
                                             Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing)
        minExtent = 0
        maxExtent = 0
        halfExtent = 0
        
        If occ Is Nothing OrElse axisDirection Is Nothing Then
            If logs IsNot Nothing Then logs.Add("CenterPatternLib: CalculatePartExtentAlongAxis - null input")
            Exit Sub
        End If
        
        If logs IsNot Nothing Then
            logs.Add("CenterPatternLib: === BEGIN PART EXTENT CALCULATION ===")
            logs.Add("CenterPatternLib: Occurrence: " & occ.Name)
            logs.Add("CenterPatternLib: Axis direction: (" & _
                     axisDirection.X.ToString("0.0000") & ", " & _
                     axisDirection.Y.ToString("0.0000") & ", " & _
                     axisDirection.Z.ToString("0.0000") & ")")
        End If
        
        Try
            ' Get the occurrence's range box
            Dim rangeBox As Box = occ.RangeBox
            Dim minPt As Point = rangeBox.MinPoint
            Dim maxPt As Point = rangeBox.MaxPoint
            
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: RangeBox min: (" & _
                         (minPt.X * 10).ToString("0.00") & ", " & _
                         (minPt.Y * 10).ToString("0.00") & ", " & _
                         (minPt.Z * 10).ToString("0.00") & ") mm")
                logs.Add("CenterPatternLib: RangeBox max: (" & _
                         (maxPt.X * 10).ToString("0.00") & ", " & _
                         (maxPt.Y * 10).ToString("0.00") & ", " & _
                         (maxPt.Z * 10).ToString("0.00") & ") mm")
                logs.Add("CenterPatternLib: RangeBox size: (" & _
                         ((maxPt.X - minPt.X) * 10).ToString("0.00") & " x " & _
                         ((maxPt.Y - minPt.Y) * 10).ToString("0.00") & " x " & _
                         ((maxPt.Z - minPt.Z) * 10).ToString("0.00") & ") mm")
            End If
            
            ' Get principal plane position as reference point
            Dim principalPlaneProxy As Object = GeoLib.FindPrincipalPlane(occ, axisDirection)
            Dim refPoint As Point = Nothing
            Dim refSource As String = ""
            
            If principalPlaneProxy IsNot Nothing Then
                Dim planeGeom As Plane = GeoLib.GetPlaneGeometry(principalPlaneProxy)
                If planeGeom IsNot Nothing Then
                    refPoint = planeGeom.RootPoint
                    refSource = "principal plane"
                    
                    If logs IsNot Nothing Then
                        Dim planeNormal As UnitVector = planeGeom.Normal
                        logs.Add("CenterPatternLib: Principal plane found")
                        logs.Add("CenterPatternLib:   Normal: (" & _
                                 planeNormal.X.ToString("0.0000") & ", " & _
                                 planeNormal.Y.ToString("0.0000") & ", " & _
                                 planeNormal.Z.ToString("0.0000") & ")")
                    End If
                End If
            Else
                If logs IsNot Nothing Then logs.Add("CenterPatternLib: Principal plane NOT found")
            End If
            
            ' Fallback to occurrence origin
            If refPoint Is Nothing Then
                refPoint = GeoLib.GetOccurrencePosition(app, occ)
                refSource = "occurrence origin"
            End If
            
            If refPoint Is Nothing Then
                If logs IsNot Nothing Then logs.Add("CenterPatternLib: ERROR - no reference point available")
                Exit Sub
            End If
            
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: Reference point source: " & refSource)
                logs.Add("CenterPatternLib: Reference point: (" & _
                         (refPoint.X * 10).ToString("0.00") & ", " & _
                         (refPoint.Y * 10).ToString("0.00") & ", " & _
                         (refPoint.Z * 10).ToString("0.00") & ") mm")
            End If
            
            ' Project reference point onto axis
            Dim refPosOnAxis As Double = refPoint.X * axisDirection.X + _
                                          refPoint.Y * axisDirection.Y + _
                                          refPoint.Z * axisDirection.Z
            
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: Reference position on axis: " & (refPosOnAxis * 10).ToString("0.00") & " mm")
            End If
            
            ' Generate all 8 corners of the bounding box
            Dim corners(7) As Point
            Dim tg As TransientGeometry = app.TransientGeometry
            
            corners(0) = tg.CreatePoint(minPt.X, minPt.Y, minPt.Z)
            corners(1) = tg.CreatePoint(maxPt.X, minPt.Y, minPt.Z)
            corners(2) = tg.CreatePoint(minPt.X, maxPt.Y, minPt.Z)
            corners(3) = tg.CreatePoint(maxPt.X, maxPt.Y, minPt.Z)
            corners(4) = tg.CreatePoint(minPt.X, minPt.Y, maxPt.Z)
            corners(5) = tg.CreatePoint(maxPt.X, minPt.Y, maxPt.Z)
            corners(6) = tg.CreatePoint(minPt.X, maxPt.Y, maxPt.Z)
            corners(7) = tg.CreatePoint(maxPt.X, maxPt.Y, maxPt.Z)
            
            ' Project each corner onto axis and find min/max relative to reference
            Dim minPos As Double = Double.MaxValue
            Dim maxPos As Double = Double.MinValue
            
            For i As Integer = 0 To 7
                Dim pos As Double = corners(i).X * axisDirection.X + _
                                    corners(i).Y * axisDirection.Y + _
                                    corners(i).Z * axisDirection.Z
                
                ' Position relative to reference point
                Dim relPos As Double = pos - refPosOnAxis
                
                If relPos < minPos Then minPos = relPos
                If relPos > maxPos Then maxPos = relPos
            Next
            
            minExtent = minPos
            maxExtent = maxPos
            halfExtent = (maxPos - minPos) / 2
            
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: Part extent along axis:")
                logs.Add("CenterPatternLib:   Min extent = " & (minExtent * 10).ToString("0.00") & " mm (relative to ref point)")
                logs.Add("CenterPatternLib:   Max extent = " & (maxExtent * 10).ToString("0.00") & " mm (relative to ref point)")
                logs.Add("CenterPatternLib:   Half extent = " & (halfExtent * 10).ToString("0.00") & " mm")
                logs.Add("CenterPatternLib:   Total thickness = " & ((maxExtent - minExtent) * 10).ToString("0.00") & " mm")
                logs.Add("CenterPatternLib: === END PART EXTENT CALCULATION ===")
            End If
            
        Catch ex As Exception
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: ERROR calculating part extent: " & ex.Message)
                logs.Add("CenterPatternLib: === END PART EXTENT CALCULATION (ERROR) ===")
            End If
        End Try
    End Sub
    
    ''' <summary>
    ''' Calculate alignment adjustment for offset parameters (AlgusNihe or LõpuNihe).
    ''' 
    ''' These adjustments are ADDED to the respective offset parameter to achieve alignment.
    ''' Adding to an offset parameter reduces the effective span from that side,
    ''' which shifts the first/last instance position inward.
    ''' 
    ''' Alignment modes:
    ''' - CENTER: geometry center at boundary (adjustment = 0)
    ''' - INWARD: inner edge at boundary → add +halfExtent (shift center inward)
    ''' - OUTWARD: outer edge at boundary → add -halfExtent (shift center outward)
    ''' 
    ''' Note: The logic is now the same for both start and end because the adjustment
    ''' is applied to the offset parameters, not the constraint. The isStartInstance
    ''' parameter is kept for API compatibility but is no longer used.
    ''' 
    ''' Returns adjustment in internal units (cm).
    ''' </summary>
    Public Function CalculateAlignmentAdjustment(alignMode As String, _
                                                  halfExtent As Double, _
                                                  isStartInstance As Boolean) As Double
        Select Case alignMode
            Case ALIGN_CENTER
                Return 0
                
            Case ALIGN_INWARD
                ' Inner edge at boundary → shift geometry center inward by halfExtent
                Return halfExtent
                
            Case ALIGN_OUTWARD
                ' Outer edge at boundary → shift geometry center outward by halfExtent
                Return -halfExtent
                
            Case Else
                Return 0
        End Select
    End Function
    
    ''' <summary>
    ''' Build alignment adjustment formula string for offset parameters.
    ''' Uses a stored half-extent parameter.
    ''' 
    ''' Note: The isStartInstance parameter is kept for API compatibility but is no longer used.
    ''' </summary>
    Public Function BuildAlignmentFormula(alignMode As String, _
                                           halfExtentParam As String, _
                                           isStartInstance As Boolean) As String
        Select Case alignMode
            Case ALIGN_CENTER
                Return "0 mm"
                
            Case ALIGN_INWARD
                Return halfExtentParam  ' +halfExtent
                
            Case ALIGN_OUTWARD
                Return "-" & halfExtentParam  ' -halfExtent
                
            Case Else
                Return "0 mm"
        End Select
    End Function

    ' ============================================================================
    ' SECTION 5: Constraint-Based Seed Positioning
    ' ============================================================================
    
    ''' <summary>
    ''' Create constraints for the perpendicular axes to lock the seed's position
    ''' in the directions not controlled by the pattern.
    ''' 
    ''' Creates two Flush constraints with 0 offset to assembly work planes
    ''' that match the original seed's perpendicular positions.
    ''' </summary>
    Public Sub CreatePerpendicularConstraints(app As Inventor.Application, _
                                               asmDoc As AssemblyDocument, _
                                               seedOcc As ComponentOccurrence, _
                                               axisDirection As UnitVector, _
                                               constraintBaseName As String, _
                                               Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing)
        If seedOcc Is Nothing OrElse axisDirection Is Nothing Then Exit Sub
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        Dim tg As TransientGeometry = app.TransientGeometry
        
        ' Get the principal plane index (perpendicular to axis)
        Dim principalIndex As Integer = GeoLib.GetPrincipalPlaneIndex(seedOcc, axisDirection)
        If principalIndex = 0 Then principalIndex = 1 ' Default to YZ
        
        If logs IsNot Nothing Then
            logs.Add("CenterPatternLib: Principal plane index = " & principalIndex.ToString())
        End If
        
        ' The other two planes need to be constrained
        ' Plane indices: 1=YZ, 2=XZ, 3=XY
        Dim perpIndices As New System.Collections.Generic.List(Of Integer)
        For i As Integer = 1 To 3
            If i <> principalIndex Then perpIndices.Add(i)
        Next
        
        Dim constraintNum As Integer = 1
        For Each planeIdx As Integer In perpIndices
            Dim planeName As String = ""
            Select Case planeIdx
                Case 1 : planeName = "YZ"
                Case 2 : planeName = "XZ"
                Case 3 : planeName = "XY"
            End Select
            
            Try
                ' Get proxy for seed's work plane
                Dim seedPlaneProxy As Object = WorkFeatureLib.CreateWorkPlaneProxy(seedOcc, planeIdx)
                If seedPlaneProxy Is Nothing Then
                    If logs IsNot Nothing Then logs.Add("CenterPatternLib: Could not get proxy for " & planeName & " plane")
                    Continue For
                End If
                
                ' Get the plane geometry to create a matching assembly work plane
                Dim seedPlaneGeom As Plane = GeoLib.GetPlaneGeometry(seedPlaneProxy)
                If seedPlaneGeom Is Nothing Then
                    If logs IsNot Nothing Then logs.Add("CenterPatternLib: Could not get geometry for " & planeName & " plane")
                    Continue For
                End If
                
                ' Create a fixed work plane in the assembly at this position
                Dim asmPlaneName As String = constraintBaseName & "_Tasand" & constraintNum.ToString()
                
                ' Delete existing if any
                Try
                    Dim existingPlane As WorkPlane = WorkFeatureLib.FindWorkPlaneByName(asmDef, asmPlaneName)
                    If existingPlane IsNot Nothing Then existingPlane.Delete()
                Catch
                End Try
                
                ' Create fixed work plane at seed's perpendicular position
                Dim origin As Point = seedPlaneGeom.RootPoint
                Dim normal As UnitVector = seedPlaneGeom.Normal
                
                ' Create X and Y axes for the work plane
                Dim refVec As Vector
                If Math.Abs(normal.Z) < 0.9 Then
                    refVec = tg.CreateVector(0, 0, 1)
                Else
                    refVec = tg.CreateVector(1, 0, 0)
                End If
                
                Dim xVec As Vector = refVec.CrossProduct(tg.CreateVector(normal.X, normal.Y, normal.Z))
                xVec.Normalize()
                Dim xAxis As UnitVector = tg.CreateUnitVector(xVec.X, xVec.Y, xVec.Z)
                
                Dim yVec As Vector = tg.CreateVector(normal.X, normal.Y, normal.Z).CrossProduct(xVec)
                yVec.Normalize()
                Dim yAxis As UnitVector = tg.CreateUnitVector(yVec.X, yVec.Y, yVec.Z)
                
                Dim asmPlane As WorkPlane = asmDef.WorkPlanes.AddFixed(origin, xAxis, yAxis)
                asmPlane.Name = asmPlaneName
                asmPlane.Visible = False
                
                ' Create Flush constraint with 0 offset
                Dim constraintName As String = constraintBaseName & "_Piirang" & constraintNum.ToString()
                
                ' Delete existing constraint if any
                Try
                    For Each c As AssemblyConstraint In asmDef.Constraints
                        If c.Name = constraintName Then
                            c.Delete()
                            Exit For
                        End If
                    Next
                Catch
                End Try
                
                ' Check normal alignment to determine Flush vs Mate
                Dim seedNormal As UnitVector = seedPlaneGeom.Normal
                Dim asmPlaneGeom As Plane = asmPlane.Plane
                Dim asmNormal As UnitVector = asmPlaneGeom.Normal
                
                Dim dot As Double = seedNormal.X * asmNormal.X + seedNormal.Y * asmNormal.Y + seedNormal.Z * asmNormal.Z
                
                If dot > 0 Then
                    ' Same direction - use Flush
                    Dim constraint As FlushConstraint = asmDef.Constraints.AddFlushConstraint(seedPlaneProxy, asmPlane, 0)
                    constraint.Name = constraintName
                Else
                    ' Opposite direction - use Mate
                    Dim constraint As MateConstraint = asmDef.Constraints.AddMateConstraint(seedPlaneProxy, asmPlane, 0)
                    constraint.Name = constraintName
                End If
                
                If logs IsNot Nothing Then
                    logs.Add("CenterPatternLib: Created perpendicular constraint for " & planeName & " plane")
                End If
                
                constraintNum += 1
            Catch ex As Exception
                If logs IsNot Nothing Then
                    logs.Add("CenterPatternLib: Failed to create perpendicular constraint for " & planeName & ": " & ex.Message)
                End If
            End Try
        Next
    End Sub
    
    ''' <summary>
    ''' Create perpendicular constraints between pattern seed and original occurrence.
    ''' This constrains the seed directly to the original's work planes, eliminating
    ''' the need for fixed assembly work planes. Moving the original will move the pattern.
    ''' </summary>
    Public Sub CreatePerpendicularConstraintsToOriginal(app As Inventor.Application, _
                                                         asmDoc As AssemblyDocument, _
                                                         seedOcc As ComponentOccurrence, _
                                                         originalOcc As ComponentOccurrence, _
                                                         axisDirection As UnitVector, _
                                                         constraintBaseName As String, _
                                                         Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing)
        If seedOcc Is Nothing OrElse originalOcc Is Nothing OrElse axisDirection Is Nothing Then Exit Sub
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' Get the principal plane index (perpendicular to axis)
        Dim principalIndex As Integer = GeoLib.GetPrincipalPlaneIndex(seedOcc, axisDirection)
        If principalIndex = 0 Then principalIndex = 1 ' Default to YZ
        
        If logs IsNot Nothing Then
            logs.Add("CenterPatternLib: Principal plane index = " & principalIndex.ToString())
        End If
        
        ' The other two planes need to be constrained to original's corresponding planes
        ' Plane indices: 1=YZ, 2=XZ, 3=XY
        Dim perpIndices As New System.Collections.Generic.List(Of Integer)
        For i As Integer = 1 To 3
            If i <> principalIndex Then perpIndices.Add(i)
        Next
        
        Dim constraintNum As Integer = 1
        For Each planeIdx As Integer In perpIndices
            Dim planeName As String = ""
            Select Case planeIdx
                Case 1 : planeName = "YZ"
                Case 2 : planeName = "XZ"
                Case 3 : planeName = "XY"
            End Select
            
            Try
                ' Get proxy for seed's work plane
                Dim seedPlaneProxy As Object = WorkFeatureLib.CreateWorkPlaneProxy(seedOcc, planeIdx)
                If seedPlaneProxy Is Nothing Then
                    If logs IsNot Nothing Then logs.Add("CenterPatternLib: Could not get seed proxy for " & planeName & " plane")
                    Continue For
                End If
                
                ' Get proxy for original's work plane (same plane index)
                Dim origPlaneProxy As Object = WorkFeatureLib.CreateWorkPlaneProxy(originalOcc, planeIdx)
                If origPlaneProxy Is Nothing Then
                    If logs IsNot Nothing Then logs.Add("CenterPatternLib: Could not get original proxy for " & planeName & " plane")
                    Continue For
                End If
                
                ' Create Flush constraint with 0 offset between seed and original planes
                Dim constraintName As String = constraintBaseName & "_Piirang" & constraintNum.ToString()
                
                ' Delete existing constraint if any
                Try
                    For Each c As AssemblyConstraint In asmDef.Constraints
                        If c.Name = constraintName Then
                            c.Delete()
                            Exit For
                        End If
                    Next
                Catch
                End Try
                
                ' Check normal alignment to determine Flush vs Mate
                Dim seedPlaneGeom As Plane = GeoLib.GetPlaneGeometry(seedPlaneProxy)
                Dim origPlaneGeom As Plane = GeoLib.GetPlaneGeometry(origPlaneProxy)
                
                If seedPlaneGeom Is Nothing OrElse origPlaneGeom Is Nothing Then
                    If logs IsNot Nothing Then logs.Add("CenterPatternLib: Could not get plane geometry for " & planeName)
                    Continue For
                End If
                
                Dim seedNormal As UnitVector = seedPlaneGeom.Normal
                Dim origNormal As UnitVector = origPlaneGeom.Normal
                Dim dot As Double = seedNormal.X * origNormal.X + seedNormal.Y * origNormal.Y + seedNormal.Z * origNormal.Z
                
                If dot > 0 Then
                    ' Same direction - use Flush
                    Dim constraint As FlushConstraint = asmDef.Constraints.AddFlushConstraint(seedPlaneProxy, origPlaneProxy, 0)
                    constraint.Name = constraintName
                Else
                    ' Opposite direction - use Mate
                    Dim constraint As MateConstraint = asmDef.Constraints.AddMateConstraint(seedPlaneProxy, origPlaneProxy, 0)
                    constraint.Name = constraintName
                End If
                
                If logs IsNot Nothing Then
                    logs.Add("CenterPatternLib: Created perpendicular constraint to original for " & planeName & " plane")
                End If
                
                constraintNum += 1
            Catch ex As Exception
                If logs IsNot Nothing Then
                    logs.Add("CenterPatternLib: Failed to create perpendicular constraint for " & planeName & ": " & ex.Message)
                End If
            End Try
        Next
    End Sub
    
    ''' <summary>
    ''' Create a constraint between the seed's principal plane and the start work plane,
    ''' with a parametric offset expression.
    ''' 
    ''' The constraint type and offset sign are determined by analyzing the geometry:
    ''' - We need the offset to move the seed FROM its current position TO the first instance position
    ''' - The constraint offset direction depends on normal directions
    ''' 
    ''' Total offset = Nihe (first instance offset from start) + KeskNihe (center offset)
    ''' </summary>
    Public Function CreateSeedConstraint(asmDoc As AssemblyDocument, _
                                          seedOcc As ComponentOccurrence, _
                                          startPlane As WorkPlane, _
                                          axisDirection As UnitVector, _
                                          offsetExpression As String, _
                                          constraintName As String, _
                                          Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing) As Object
        If seedOcc Is Nothing OrElse startPlane Is Nothing Then Return Nothing
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' Delete existing constraint if any
        Try
            For Each c As AssemblyConstraint In asmDef.Constraints
                If c.Name = constraintName Then
                    c.Delete()
                    Exit For
                End If
            Next
        Catch
        End Try
        
        ' Find the seed's principal plane (perpendicular to axis)
        Dim principalPlaneProxy As Object = GeoLib.FindPrincipalPlane(seedOcc, axisDirection)
        If principalPlaneProxy Is Nothing Then
            ' Fallback to XY plane (index 3)
            principalPlaneProxy = WorkFeatureLib.CreateWorkPlaneProxy(seedOcc, 3)
        End If
        If principalPlaneProxy Is Nothing Then
            If logs IsNot Nothing Then logs.Add("CenterPatternLib: ERROR - Could not find principal plane")
            Return Nothing
        End If
        
        ' Get geometries of both planes
        Dim principalGeom As Plane = GeoLib.GetPlaneGeometry(principalPlaneProxy)
        Dim startPlaneGeom As Plane = GeoLib.GetPlaneGeometry(startPlane)
        
        If principalGeom Is Nothing OrElse startPlaneGeom Is Nothing Then
            If logs IsNot Nothing Then logs.Add("CenterPatternLib: ERROR - Could not get plane geometries")
            Return Nothing
        End If
        
        Dim principalNormal As UnitVector = principalGeom.Normal
        Dim startNormal As UnitVector = startPlaneGeom.Normal
        Dim principalRoot As Point = principalGeom.RootPoint
        Dim startRoot As Point = startPlaneGeom.RootPoint
        
        ' Calculate dot products
        Dim principalDotAxis As Double = principalNormal.X * axisDirection.X + _
                                          principalNormal.Y * axisDirection.Y + _
                                          principalNormal.Z * axisDirection.Z
        Dim startDotAxis As Double = startNormal.X * axisDirection.X + _
                                      startNormal.Y * axisDirection.Y + _
                                      startNormal.Z * axisDirection.Z
        Dim normalsDotProduct As Double = principalNormal.X * startNormal.X + _
                                           principalNormal.Y * startNormal.Y + _
                                           principalNormal.Z * startNormal.Z
        
        If logs IsNot Nothing Then
            logs.Add("CenterPatternLib: === CONSTRAINT CREATION ===")
            logs.Add("CenterPatternLib: Start plane root = (" & _
                     (startRoot.X * 10).ToString("0.00") & ", " & _
                     (startRoot.Y * 10).ToString("0.00") & ", " & _
                     (startRoot.Z * 10).ToString("0.00") & ") mm")
            logs.Add("CenterPatternLib: Start plane normal = (" & _
                     startNormal.X.ToString("0.000") & ", " & _
                     startNormal.Y.ToString("0.000") & ", " & _
                     startNormal.Z.ToString("0.000") & ")")
            logs.Add("CenterPatternLib: Principal plane root = (" & _
                     (principalRoot.X * 10).ToString("0.00") & ", " & _
                     (principalRoot.Y * 10).ToString("0.00") & ", " & _
                     (principalRoot.Z * 10).ToString("0.00") & ") mm")
            logs.Add("CenterPatternLib: Principal plane normal = (" & _
                     principalNormal.X.ToString("0.000") & ", " & _
                     principalNormal.Y.ToString("0.000") & ", " & _
                     principalNormal.Z.ToString("0.000") & ")")
            logs.Add("CenterPatternLib: Axis direction = (" & _
                     axisDirection.X.ToString("0.000") & ", " & _
                     axisDirection.Y.ToString("0.000") & ", " & _
                     axisDirection.Z.ToString("0.000") & ")")
            logs.Add("CenterPatternLib: Principal·Axis = " & principalDotAxis.ToString("0.000"))
            logs.Add("CenterPatternLib: Start·Axis = " & startDotAxis.ToString("0.000"))
            logs.Add("CenterPatternLib: Principal·Start (normals) = " & normalsDotProduct.ToString("0.000"))
        End If
        
        ' Determine constraint type based on whether normals face same or opposite directions
        ' Normals dot > 0: same direction → Flush
        ' Normals dot < 0: opposite directions → Mate
        Dim useMate As Boolean = (normalsDotProduct < 0)
        
        ' Determine if we need to negate the offset
        ' Our offset formula measures distance along AXIS direction (positive = toward end)
        ' 
        ' For BOTH Flush and Mate constraints (from testing):
        '   - Positive offset moves seed OPPOSITE to principal plane's normal direction
        '   - If principal normal = +axis: positive offset moves seed OPPOSITE to axis (toward start) → NEGATE
        '   - If principal normal = -axis: positive offset moves seed SAME as axis (toward end) → DON'T NEGATE
        '   
        ' Therefore: negate when principal normal aligns with axis (principalDotAxis > 0)
        Dim needNegateOffset As Boolean = (principalDotAxis > 0)
        
        If logs IsNot Nothing Then
            logs.Add("CenterPatternLib: Using " & If(useMate, "MATE", "FLUSH") & " constraint")
            logs.Add("CenterPatternLib: Need to negate offset: " & needNegateOffset.ToString())
            logs.Add("CenterPatternLib: Offset expression: " & offsetExpression)
        End If
        
        ' Build final offset expression
        Dim finalOffsetExpr As String = offsetExpression
        If needNegateOffset Then
            finalOffsetExpr = "-(" & offsetExpression & ")"
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: Final offset expression: " & finalOffsetExpr)
            End If
        End If
        
        ' Create the appropriate constraint type
        If useMate Then
            ' Mate: planes face each other, offset pushes them apart
            Try
                Dim constraint As MateConstraint = asmDef.Constraints.AddMateConstraint( _
                    principalPlaneProxy, startPlane, 0)
                constraint.Name = constraintName
                
                Try
                    constraint.Offset.Expression = finalOffsetExpr
                    If logs IsNot Nothing Then logs.Add("CenterPatternLib: Mate constraint created with offset: " & finalOffsetExpr)
                Catch ex As Exception
                    If logs IsNot Nothing Then logs.Add("CenterPatternLib: WARNING - Could not set offset expression: " & ex.Message)
                End Try
                
                If logs IsNot Nothing Then logs.Add("CenterPatternLib: === END CONSTRAINT CREATION ===")
                Return constraint
            Catch ex As Exception
                If logs IsNot Nothing Then logs.Add("CenterPatternLib: Mate constraint failed: " & ex.Message)
            End Try
        Else
            ' Flush: planes face same direction, offset separates them
            Try
                Dim constraint As FlushConstraint = asmDef.Constraints.AddFlushConstraint( _
                    principalPlaneProxy, startPlane, 0)
                constraint.Name = constraintName
                
                Try
                    constraint.Offset.Expression = finalOffsetExpr
                    If logs IsNot Nothing Then logs.Add("CenterPatternLib: Flush constraint created with offset: " & finalOffsetExpr)
                Catch ex As Exception
                    If logs IsNot Nothing Then logs.Add("CenterPatternLib: WARNING - Could not set offset expression: " & ex.Message)
                End Try
                
                If logs IsNot Nothing Then logs.Add("CenterPatternLib: === END CONSTRAINT CREATION ===")
                Return constraint
            Catch ex As Exception
                If logs IsNot Nothing Then logs.Add("CenterPatternLib: Flush constraint failed: " & ex.Message)
            End Try
        End If
        
        ' If preferred type failed, try the other type
        If logs IsNot Nothing Then logs.Add("CenterPatternLib: Trying alternate constraint type...")
        
        If useMate Then
            Try
                Dim constraint As FlushConstraint = asmDef.Constraints.AddFlushConstraint( _
                    principalPlaneProxy, startPlane, 0)
                constraint.Name = constraintName
                Try
                    constraint.Offset.Expression = finalOffsetExpr
                Catch
                End Try
                If logs IsNot Nothing Then logs.Add("CenterPatternLib: === END CONSTRAINT CREATION ===")
                Return constraint
            Catch
            End Try
        Else
            Try
                Dim constraint As MateConstraint = asmDef.Constraints.AddMateConstraint( _
                    principalPlaneProxy, startPlane, 0)
                constraint.Name = constraintName
                Try
                    constraint.Offset.Expression = finalOffsetExpr
                Catch
                End Try
                If logs IsNot Nothing Then logs.Add("CenterPatternLib: === END CONSTRAINT CREATION ===")
                Return constraint
            Catch
            End Try
        End If
        
        If logs IsNot Nothing Then logs.Add("CenterPatternLib: ERROR - Both constraint types failed")
        If logs IsNot Nothing Then logs.Add("CenterPatternLib: === END CONSTRAINT CREATION ===")
        Return Nothing
    End Function

    ' ============================================================================
    ' SECTION 6: Main Pattern Creation
    ' ============================================================================
    
    ''' <summary>
    ''' Create a center-based occurrence pattern.
    ''' 
    ''' The seed is assumed to be initially positioned at the center of the span.
    ''' A constraint is created to position the seed at the first instance position,
    ''' and a rectangular pattern creates additional instances.
    ''' 
    ''' Parameters:
    '''   app - Inventor.Application
    '''   asmDoc - Assembly document
    '''   iLogicAuto - iLogicVb.Automation object (for update handler registration)
    '''   seedOcc - The occurrence to pattern
    '''   startGeometry - Face/plane defining start boundary
    '''   endGeometry - Face/plane defining end boundary
    '''   maxSpacingMm - Maximum spacing between instances (mm)
    '''   mode - MODE_UNIFORM or MODE_SYMMETRIC
    '''   includeEnds - Whether to include instances at boundaries
    '''   baseName - Base name for created features/parameters
    '''   logs - Optional list for logging
    '''   
    ''' Returns True if pattern created successfully.
    ''' </summary>
    Public Function CreateCenterPattern(app As Inventor.Application, _
                                         asmDoc As AssemblyDocument, _
                                         iLogicAuto As Object, _
                                         seedOcc As ComponentOccurrence, _
                                         startGeometry As Object, _
                                         endGeometry As Object, _
                                         maxSpacingMm As Double, _
                                         mode As String, _
                                         includeEnds As Boolean, _
                                         baseName As String, _
                                         Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing) As Boolean
        
        If logs Is Nothing Then logs = New System.Collections.Generic.List(Of String)
        
        ' Validate inputs
        If seedOcc Is Nothing Then
            logs.Add("CenterPatternLib: ERROR - No seed occurrence")
            Return False
        End If
        
        ' Handle parameter names that start with digits
        Dim paramPrefix As String = baseName
        If baseName.Length > 0 AndAlso Char.IsDigit(baseName(0)) Then
            paramPrefix = "M_" & baseName
        End If
        
        logs.Add("CenterPatternLib: Creating pattern '" & baseName & "'")
        logs.Add("CenterPatternLib: Mode = " & mode & ", IncludeEnds = " & includeEnds.ToString())
        
        ' Define parameter and feature names
        Dim spanParamName As String = paramPrefix & "_Ulatus"
        Dim maxSpacingParamName As String = paramPrefix & "_MaxVahe"
        Dim countParamName As String = paramPrefix & "_Arv"
        Dim spacingParamName As String = paramPrefix & "_Samm"
        Dim offsetParamName As String = paramPrefix & "_Nihe"
        Dim centerOffsetParamName As String = paramPrefix & "_KeskNihe"
        Dim startPlaneName As String = baseName & "_AlgusTasand"
        Dim endPlaneName As String = baseName & "_LõpuTasand"
        Dim axisName As String = baseName & "_Telg"
        Dim patternName As String = baseName & "_Muster"
        Dim constraintName As String = baseName & "_Asend"
        
        ' 1. Create associative work planes for start and end
        logs.Add("CenterPatternLib: Creating work planes...")
        Dim startPlane As WorkPlane = WorkFeatureLib.GetOrCreateWorkPlane(app, asmDoc, startGeometry, startPlaneName)
        Dim endPlane As WorkPlane = WorkFeatureLib.GetOrCreateWorkPlane(app, asmDoc, endGeometry, endPlaneName)
        
        If startPlane Is Nothing OrElse endPlane Is Nothing Then
            logs.Add("CenterPatternLib: ERROR - Failed to create work planes")
            Return False
        End If
        
        ' Log plane positions
        Dim startPlaneGeom As Plane = GeoLib.GetPlaneGeometry(startPlane)
        Dim endPlaneGeom As Plane = GeoLib.GetPlaneGeometry(endPlane)
        If startPlaneGeom IsNot Nothing Then
            logs.Add("CenterPatternLib: Start plane at (" & _
                     (startPlaneGeom.RootPoint.X * 10).ToString("0.00") & ", " & _
                     (startPlaneGeom.RootPoint.Y * 10).ToString("0.00") & ", " & _
                     (startPlaneGeom.RootPoint.Z * 10).ToString("0.00") & ") mm")
        End If
        
        ' 2. Create work axis between planes
        logs.Add("CenterPatternLib: Creating direction axis...")
        Dim dirAxis As WorkAxis = WorkFeatureLib.GetOrCreateWorkAxis(app, asmDoc, startPlane, endPlane, axisName)
        
        If dirAxis Is Nothing Then
            logs.Add("CenterPatternLib: ERROR - Failed to create work axis")
            Return False
        End If
        
        Dim axisDirection As UnitVector = GeoLib.GetAxisDirection(dirAxis)
        If axisDirection IsNot Nothing Then
            logs.Add("CenterPatternLib: Axis direction = (" & _
                     axisDirection.X.ToString("0.000") & ", " & _
                     axisDirection.Y.ToString("0.000") & ", " & _
                     axisDirection.Z.ToString("0.000") & ")")
        End If
        
        ' 3. Measure span (fixed value at creation time)
        Dim spanCm As Double = GeoLib.MeasurePlaneDistance(startPlane, endPlane)
        Dim spanMm As Double = spanCm * 10
        logs.Add("CenterPatternLib: Span = " & spanMm.ToString("0.00") & " mm")
        
        ' 4. Calculate center offset (before copying seed)
        logs.Add("CenterPatternLib: Calculating center offset...")
        Dim centerOffsetCm As Double = CalculateCenterOffset(app, seedOcc, startPlane, axisDirection, spanCm, logs)
        Dim centerOffsetMm As Double = centerOffsetCm * 10
        logs.Add("CenterPatternLib: Center offset result = " & centerOffsetMm.ToString("0.00") & " mm")
        
        ' 5. Copy seed and suppress original
        logs.Add("CenterPatternLib: Copying seed occurrence...")
        Dim patternSeed As ComponentOccurrence = PatternLib.CopyAndSuppressSeed(app, asmDoc, seedOcc, baseName)
        
        If patternSeed Is Nothing Then
            logs.Add("CenterPatternLib: ERROR - Failed to copy seed")
            Return False
        End If
        logs.Add("CenterPatternLib: Seed copied, original suppressed")
        
        ' 6. Create parameters
        logs.Add("CenterPatternLib: Creating parameters...")
        
        ' Span (fixed value)
        SetParameter(asmDoc, spanParamName, spanMm, "mm")
        logs.Add("CenterPatternLib: " & spanParamName & " = " & spanMm.ToString("0.00") & " mm")
        
        ' Max spacing
        SetParameter(asmDoc, maxSpacingParamName, maxSpacingMm, "mm")
        logs.Add("CenterPatternLib: " & maxSpacingParamName & " = " & maxSpacingMm.ToString("0.00") & " mm")
        
        ' Count (unitless formula)
        Dim countFormula As String = BuildCountFormula(spanParamName, maxSpacingParamName, mode, includeEnds)
        SetUnitlessFormula(asmDoc, countParamName, countFormula)
        logs.Add("CenterPatternLib: " & countParamName & " = " & countFormula)
        
        ' Spacing (formula)
        Dim spacingFormula As String = BuildSpacingFormula(spanParamName, countParamName, mode, includeEnds)
        SetParameterFormula(asmDoc, spacingParamName, spacingFormula)
        logs.Add("CenterPatternLib: " & spacingParamName & " = " & spacingFormula)
        
        ' Offset (formula)
        Dim offsetFormula As String = BuildOffsetFormula(spacingParamName, includeEnds, mode, spanParamName, countParamName)
        SetParameterFormula(asmDoc, offsetParamName, offsetFormula)
        logs.Add("CenterPatternLib: " & offsetParamName & " = " & offsetFormula)
        
        ' Center offset (fixed value)
        SetParameter(asmDoc, centerOffsetParamName, centerOffsetMm, "mm")
        logs.Add("CenterPatternLib: " & centerOffsetParamName & " = " & centerOffsetMm.ToString("0.00") & " mm")
        
        ' Log actual values
        Try
            Dim actualCount As Double = GetParameterValue(asmDoc, countParamName)
            Dim actualSpacing As Double = GetParameterValue(asmDoc, spacingParamName) * 10
            Dim actualOffset As Double = GetParameterValue(asmDoc, offsetParamName) * 10
            logs.Add("CenterPatternLib: Actual values - Count=" & CInt(actualCount).ToString() & _
                     ", Spacing=" & actualSpacing.ToString("0.00") & "mm" & _
                     ", Offset=" & actualOffset.ToString("0.00") & "mm")
        Catch
        End Try
        
        ' Register update handler to automatically update span when geometry changes
        If iLogicAuto IsNot Nothing Then
            RegisterSpanUpdateHandler(asmDoc, iLogicAuto, baseName, startPlaneName, endPlaneName, axisName, spanParamName, logs)
        End If
        
        ' 7. Create perpendicular constraints FIRST (to lock Y and Z before moving along axis)
        logs.Add("CenterPatternLib: Creating perpendicular constraints...")
        CreatePerpendicularConstraints(app, asmDoc, patternSeed, axisDirection, baseName, logs)
        
        ' 8. Create seed positioning constraint along pattern axis
        logs.Add("CenterPatternLib: Creating seed constraint...")
        
        ' Offset expression: Nihe + KeskNihe
        Dim constraintOffsetExpr As String = offsetParamName & " + " & centerOffsetParamName
        logs.Add("CenterPatternLib: Constraint offset = " & constraintOffsetExpr)
        
        Dim seedConstraint As Object = CreateSeedConstraint( _
            asmDoc, patternSeed, startPlane, axisDirection, constraintOffsetExpr, constraintName, logs)
        
        If seedConstraint Is Nothing Then
            logs.Add("CenterPatternLib: WARNING - Seed constraint creation failed")
        Else
            logs.Add("CenterPatternLib: Seed constraint created")
        End If
        
        ' 9. Create rectangular pattern (always create even if count=1, so it updates when count changes)
        Dim totalCount As Double = GetParameterValue(asmDoc, countParamName)
        logs.Add("CenterPatternLib: Total count = " & CInt(totalCount).ToString())
        
        logs.Add("CenterPatternLib: Creating pattern...")
        
        Dim pattern As RectangularOccurrencePattern = PatternLib.CreateRectangularPatternFromOccurrence( _
            app, asmDoc, patternSeed, dirAxis, countParamName, spacingParamName, patternName)
        
        If pattern Is Nothing Then
            logs.Add("CenterPatternLib: WARNING - Pattern creation failed")
        Else
            logs.Add("CenterPatternLib: Pattern created successfully")
            
            ' Log final positions
            Dim patternOccs As System.Collections.Generic.List(Of ComponentOccurrence) = _
                PatternLib.GetPatternOccurrences(pattern)
            logs.Add("CenterPatternLib: Pattern has " & patternOccs.Count.ToString() & " element(s)")
        End If
        
        ' 10. Store configuration for re-runs
        StorePatternConfig(patternSeed, baseName, startPlaneName, endPlaneName, _
                          maxSpacingMm.ToString(), mode, includeEnds)
        
        logs.Add("CenterPatternLib: Done")
        Return True
    End Function
    
    ''' <summary>
    ''' Extended version of CreateCenterPattern with additional features:
    ''' - Start/end offsets to reduce effective span
    ''' - Start/end alignment modes (center, inward, outward)
    ''' - Zero instance handling (suppress pattern when span too small)
    ''' - Explicit axis specification for non-parallel boundaries
    ''' 
    ''' Parameters:
    '''   app - Inventor.Application
    '''   asmDoc - Assembly document
    '''   iLogicAuto - iLogicVb.Automation object
    '''   seedOcc - The occurrence to pattern
    '''   startGeometry - Geometry defining start boundary (face/plane/point/vertex)
    '''   endGeometry - Geometry defining end boundary
    '''   explicitAxis - Optional explicit axis (work axis, edge) for non-parallel boundaries
    '''   maxSpacingMm - Maximum spacing between instances (mm)
    '''   mode - MODE_UNIFORM or MODE_SYMMETRIC
    '''   includeEnds - Whether to include instances at boundaries
    '''   baseName - Base name for created features/parameters
    '''   startOffsetMm - Offset from start plane (reduces effective span)
    '''   endOffsetMm - Offset from end plane (reduces effective span)
    '''   startAlignment - ALIGN_CENTER, ALIGN_INWARD, or ALIGN_OUTWARD
    '''   endAlignment - ALIGN_CENTER, ALIGN_INWARD, or ALIGN_OUTWARD
    '''   allowZeroInstances - Allow suppressing pattern when span <= maxSpacing
    '''   logs - Optional list for logging
    ''' </summary>
    Public Function CreateCenterPatternEx(app As Inventor.Application, _
                                           asmDoc As AssemblyDocument, _
                                           iLogicAuto As Object, _
                                           seedOcc As ComponentOccurrence, _
                                           startGeometry As Object, _
                                           endGeometry As Object, _
                                           explicitAxis As Object, _
                                           maxSpacingInput As String, _
                                           mode As String, _
                                           includeEnds As Boolean, _
                                           baseName As String, _
                                           startOffsetMm As Double, _
                                           endOffsetMm As Double, _
                                           startAlignment As String, _
                                           endAlignment As String, _
                                           allowZeroInstances As Boolean, _
                                           Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing) As Boolean
        
        If logs Is Nothing Then logs = New System.Collections.Generic.List(Of String)
        
        ' Validate inputs
        If seedOcc Is Nothing Then
            logs.Add("CenterPatternLib: ERROR - No seed occurrence")
            Return False
        End If
        
        ' Handle parameter names that start with digits
        Dim paramPrefix As String = baseName
        If baseName.Length > 0 AndAlso Char.IsDigit(baseName(0)) Then
            paramPrefix = "M_" & baseName
        End If
        
        logs.Add("CenterPatternLib: Creating extended pattern '" & baseName & "'")
        logs.Add("CenterPatternLib: Mode = " & mode & ", IncludeEnds = " & includeEnds.ToString())
        logs.Add("CenterPatternLib: StartOffset = " & startOffsetMm.ToString("0.00") & " mm, EndOffset = " & endOffsetMm.ToString("0.00") & " mm")
        logs.Add("CenterPatternLib: StartAlign = " & startAlignment & ", EndAlign = " & endAlignment)
        logs.Add("CenterPatternLib: AllowZeroInstances = " & allowZeroInstances.ToString())
        
        ' Define parameter and feature names
        Dim spanParamName As String = paramPrefix & "_Ulatus"
        Dim effectiveSpanParamName As String = paramPrefix & "_TegelikUlatus"
        Dim maxSpacingParamName As String = paramPrefix & "_MaxVahe"
        Dim countParamName As String = paramPrefix & "_Arv"
        Dim spacingParamName As String = paramPrefix & "_Samm"
        Dim offsetParamName As String = paramPrefix & "_Nihe"
        Dim centerOffsetParamName As String = paramPrefix & "_KeskNihe"
        Dim startOffsetParamName As String = paramPrefix & "_AlgusNihe"
        Dim endOffsetParamName As String = paramPrefix & "_LõpuNihe"
        Dim halfExtentParamName As String = paramPrefix & "_PoolUlatus"
        Dim startPlaneName As String = baseName & "_AlgusTasand"
        Dim endPlaneName As String = baseName & "_LõpuTasand"
        Dim axisName As String = baseName & "_Telg"
        Dim patternName As String = baseName & "_Muster"
        Dim constraintName As String = baseName & "_Asend"
        
        ' 1. Determine axis direction first (needed for work plane creation with non-planar geometry)
        logs.Add("CenterPatternLib: Determining axis...")
        Dim dirAxis As WorkAxis = Nothing
        Dim axisDirection As UnitVector = Nothing
        
        ' Try to determine axis from boundaries or explicit axis
        dirAxis = WorkFeatureLib.DetermineAxisFromBoundaries(app, asmDoc, startGeometry, endGeometry, explicitAxis, axisName)
        
        If dirAxis IsNot Nothing Then
            axisDirection = GeoLib.GetAxisDirection(dirAxis)
        ElseIf explicitAxis IsNot Nothing Then
            axisDirection = GeoLib.GetAxisDirection(explicitAxis)
        End If
        
        ' 2. Create associative work planes for start and end
        logs.Add("CenterPatternLib: Creating work planes...")
        Dim startPlane As WorkPlane = Nothing
        Dim endPlane As WorkPlane = Nothing
        
        If axisDirection IsNot Nothing Then
            startPlane = WorkFeatureLib.GetOrCreateWorkPlaneEx(app, asmDoc, startGeometry, axisDirection, startPlaneName)
            endPlane = WorkFeatureLib.GetOrCreateWorkPlaneEx(app, asmDoc, endGeometry, axisDirection, endPlaneName)
        Else
            startPlane = WorkFeatureLib.GetOrCreateWorkPlane(app, asmDoc, startGeometry, startPlaneName)
            endPlane = WorkFeatureLib.GetOrCreateWorkPlane(app, asmDoc, endGeometry, endPlaneName)
        End If
        
        If startPlane Is Nothing OrElse endPlane Is Nothing Then
            logs.Add("CenterPatternLib: ERROR - Failed to create work planes")
            Return False
        End If
        
        ' Log plane positions
        Dim startPlaneGeom As Plane = GeoLib.GetPlaneGeometry(startPlane)
        If startPlaneGeom IsNot Nothing Then
            logs.Add("CenterPatternLib: Start plane at (" & _
                     (startPlaneGeom.RootPoint.X * 10).ToString("0.00") & ", " & _
                     (startPlaneGeom.RootPoint.Y * 10).ToString("0.00") & ", " & _
                     (startPlaneGeom.RootPoint.Z * 10).ToString("0.00") & ") mm")
        End If
        
        ' 3. Create work axis if not already created
        If dirAxis Is Nothing Then
            logs.Add("CenterPatternLib: Creating direction axis...")
            dirAxis = WorkFeatureLib.GetOrCreateWorkAxis(app, asmDoc, startPlane, endPlane, axisName)
        End If
        
        If dirAxis Is Nothing Then
            logs.Add("CenterPatternLib: ERROR - Failed to create work axis")
            Return False
        End If
        
        axisDirection = GeoLib.GetAxisDirection(dirAxis)
        If axisDirection IsNot Nothing Then
            logs.Add("CenterPatternLib: Axis direction = (" & _
                     axisDirection.X.ToString("0.000") & ", " & _
                     axisDirection.Y.ToString("0.000") & ", " & _
                     axisDirection.Z.ToString("0.000") & ")")
        End If
        
        ' 4. Measure span
        Dim spanCm As Double = GeoLib.MeasurePlaneDistance(startPlane, endPlane)
        Dim spanMm As Double = spanCm * 10
        logs.Add("CenterPatternLib: Span = " & spanMm.ToString("0.00") & " mm")
        
        ' 5. Calculate part extent along axis for alignment
        Dim minExtent As Double = 0
        Dim maxExtent As Double = 0
        Dim halfExtent As Double = 0
        CalculatePartExtentAlongAxis(app, seedOcc, axisDirection, minExtent, maxExtent, halfExtent, logs)
        Dim halfExtentMm As Double = halfExtent * 10
        
        ' 6. Calculate center offset (before copying seed)
        logs.Add("CenterPatternLib: Calculating center offset...")
        Dim centerOffsetCm As Double = CalculateCenterOffset(app, seedOcc, startPlane, axisDirection, spanCm, logs)
        Dim centerOffsetMm As Double = centerOffsetCm * 10
        logs.Add("CenterPatternLib: Center offset result = " & centerOffsetMm.ToString("0.00") & " mm")
        
        ' 7. Copy seed and hide original (BOM=Reference, Visible=False)
        logs.Add("CenterPatternLib: Copying seed occurrence...")
        Dim patternSeed As ComponentOccurrence = PatternLib.CopyAndHideSeed(app, asmDoc, seedOcc, baseName)
        
        If patternSeed Is Nothing Then
            logs.Add("CenterPatternLib: ERROR - Failed to copy seed")
            Return False
        End If
        logs.Add("CenterPatternLib: Seed copied, original hidden (BOM=Reference)")
        
        ' 7a. Move original to Template folder IMMEDIATELY (before creating more features)
        ' This must happen before creating pattern, otherwise original ends up below pattern in browser
        logs.Add("CenterPatternLib: Moving original to Template folder...")
        MoveOriginalToTemplateFolder(asmDoc, seedOcc, logs)
        
        ' 8. Create parameters
        logs.Add("CenterPatternLib: Creating parameters...")
        
        ' Span (fixed value - updated by update handler)
        SetParameter(asmDoc, spanParamName, spanMm, "mm")
        logs.Add("CenterPatternLib: " & spanParamName & " = " & spanMm.ToString("0.00") & " mm")
        
        ' Calculate alignment adjustments (in mm)
        Dim startAlignAdjustMm As Double = CalculateAlignmentAdjustment(startAlignment, halfExtent, True) * 10
        Dim endAlignAdjustMm As Double = CalculateAlignmentAdjustment(endAlignment, halfExtent, False) * 10
        
        logs.Add("CenterPatternLib: === ALIGNMENT ADJUSTMENTS ===")
        logs.Add("CenterPatternLib: Start alignment: " & startAlignment & " → adjustment = " & startAlignAdjustMm.ToString("0.00") & " mm")
        logs.Add("CenterPatternLib: End alignment: " & endAlignment & " → adjustment = " & endAlignAdjustMm.ToString("0.00") & " mm")
        
        ' Start and end offsets (include alignment adjustments)
        ' The alignment adjustment is ADDED to the user offset to achieve the desired positioning
        Dim effectiveStartOffsetMm As Double = startOffsetMm + startAlignAdjustMm
        Dim effectiveEndOffsetMm As Double = endOffsetMm + endAlignAdjustMm
        
        SetParameter(asmDoc, startOffsetParamName, effectiveStartOffsetMm, "mm")
        SetParameter(asmDoc, endOffsetParamName, effectiveEndOffsetMm, "mm")
        logs.Add("CenterPatternLib: " & startOffsetParamName & " = " & effectiveStartOffsetMm.ToString("0.00") & " mm (user: " & startOffsetMm.ToString("0.00") & " + align: " & startAlignAdjustMm.ToString("0.00") & ")")
        logs.Add("CenterPatternLib: " & endOffsetParamName & " = " & effectiveEndOffsetMm.ToString("0.00") & " mm (user: " & endOffsetMm.ToString("0.00") & " + align: " & endAlignAdjustMm.ToString("0.00") & ")")
        
        ' Effective span (formula)
        Dim effectiveSpanFormula As String = BuildEffectiveSpanFormula(spanParamName, startOffsetParamName, endOffsetParamName)
        SetParameterFormula(asmDoc, effectiveSpanParamName, effectiveSpanFormula)
        logs.Add("CenterPatternLib: " & effectiveSpanParamName & " = " & effectiveSpanFormula)
        
        ' Max spacing - can be number or formula/parameter reference
        Dim maxSpacingMm As Double = 0
        Dim numValue As Double
        If Double.TryParse(maxSpacingInput, System.Globalization.NumberStyles.Any, _
                           System.Globalization.CultureInfo.InvariantCulture, numValue) Then
            ' Pure number - set as fixed value with units
            maxSpacingMm = numValue
            SetParameter(asmDoc, maxSpacingParamName, maxSpacingMm, "mm")
            logs.Add("CenterPatternLib: " & maxSpacingParamName & " = " & maxSpacingMm.ToString("0.00") & " mm (fixed)")
        Else
            ' Formula/expression - let Inventor evaluate it
            SetParameterFormula(asmDoc, maxSpacingParamName, maxSpacingInput)
            maxSpacingMm = GetParameterValue(asmDoc, maxSpacingParamName) * 10 ' Get evaluated value (cm to mm)
            logs.Add("CenterPatternLib: " & maxSpacingParamName & " = " & maxSpacingInput & " (formula, evaluates to " & maxSpacingMm.ToString("0.00") & " mm)")
        End If
        
        If maxSpacingMm <= 0 Then
            logs.Add("CenterPatternLib: ERROR - Max spacing must be positive")
            Return False
        End If
        
        ' Half extent (for alignment calculations)
        SetParameter(asmDoc, halfExtentParamName, halfExtentMm, "mm")
        logs.Add("CenterPatternLib: " & halfExtentParamName & " = " & halfExtentMm.ToString("0.00") & " mm")
        
        ' Count (unitless formula using effective span)
        Dim countFormula As String = BuildCountFormulaWithOffsets(effectiveSpanParamName, maxSpacingParamName, mode, includeEnds)
        SetUnitlessFormula(asmDoc, countParamName, countFormula)
        logs.Add("CenterPatternLib: " & countParamName & " = " & countFormula)
        
        ' Spacing (formula using effective span)
        Dim spacingFormula As String = BuildSpacingFormulaWithOffsets(effectiveSpanParamName, countParamName, mode, includeEnds)
        SetParameterFormula(asmDoc, spacingParamName, spacingFormula)
        logs.Add("CenterPatternLib: " & spacingParamName & " = " & spacingFormula)
        
        ' Offset from start (includes start offset)
        Dim offsetFormula As String = BuildOffsetFormulaWithStartOffset(spacingParamName, startOffsetParamName, _
                                                                         includeEnds, mode, effectiveSpanParamName, countParamName)
        SetParameterFormula(asmDoc, offsetParamName, offsetFormula)
        logs.Add("CenterPatternLib: " & offsetParamName & " = " & offsetFormula)
        
        ' Center offset (fixed value)
        SetParameter(asmDoc, centerOffsetParamName, centerOffsetMm, "mm")
        logs.Add("CenterPatternLib: " & centerOffsetParamName & " = " & centerOffsetMm.ToString("0.00") & " mm")
        
        ' Log actual values
        Try
            Dim actualCount As Double = GetParameterValue(asmDoc, countParamName)
            Dim actualSpacing As Double = GetParameterValue(asmDoc, spacingParamName) * 10
            Dim actualOffset As Double = GetParameterValue(asmDoc, offsetParamName) * 10
            Dim actualEffSpan As Double = GetParameterValue(asmDoc, effectiveSpanParamName) * 10
            logs.Add("CenterPatternLib: Actual values - Count=" & CInt(actualCount).ToString() & _
                     ", EffSpan=" & actualEffSpan.ToString("0.00") & "mm" & _
                     ", Spacing=" & actualSpacing.ToString("0.00") & "mm" & _
                     ", Offset=" & actualOffset.ToString("0.00") & "mm")
        Catch
        End Try
        
        ' Register extended update handler
        If iLogicAuto IsNot Nothing Then
            RegisterSpanUpdateHandlerEx(asmDoc, iLogicAuto, baseName, startPlaneName, endPlaneName, axisName, _
                                         spanParamName, effectiveSpanParamName, maxSpacingParamName, patternName, _
                                         allowZeroInstances, logs)
        End If
        
        ' 9. Create perpendicular constraints to original occurrence
        ' This constrains the seed to the original's work planes (not fixed assembly planes)
        ' Moving the original will move the entire pattern
        logs.Add("CenterPatternLib: Creating perpendicular constraints to original...")
        CreatePerpendicularConstraintsToOriginal(app, asmDoc, patternSeed, seedOcc, axisDirection, baseName, logs)
        
        ' 10. Create seed positioning constraint along pattern axis
        ' The constraint offset is simple: Nihe (first instance position) + KeskNihe (geom-to-principal offset)
        ' Alignment adjustments are already baked into AlgusNihe/LõpuNihe → Nihe
        logs.Add("CenterPatternLib: Creating seed constraint...")
        
        Dim constraintOffsetExpr As String = offsetParamName & " + " & centerOffsetParamName
        logs.Add("CenterPatternLib: Constraint offset expression: " & constraintOffsetExpr)
        
        Dim seedConstraint As Object = CreateSeedConstraint( _
            asmDoc, patternSeed, startPlane, axisDirection, constraintOffsetExpr, constraintName, logs)
        
        If seedConstraint Is Nothing Then
            logs.Add("CenterPatternLib: WARNING - Seed constraint creation failed")
        Else
            logs.Add("CenterPatternLib: Seed constraint created")
        End If
        
        ' 11. Create rectangular pattern
        Dim totalCount As Double = GetParameterValue(asmDoc, countParamName)
        logs.Add("CenterPatternLib: Total count = " & CInt(totalCount).ToString())
        
        logs.Add("CenterPatternLib: Creating pattern...")
        
        Dim pattern As RectangularOccurrencePattern = PatternLib.CreateRectangularPatternFromOccurrence( _
            app, asmDoc, patternSeed, dirAxis, countParamName, spacingParamName, patternName)
        
        If pattern Is Nothing Then
            logs.Add("CenterPatternLib: WARNING - Pattern creation failed")
        Else
            logs.Add("CenterPatternLib: Pattern created successfully")
            
            ' Check if should be suppressed for zero instance case
            ' Use effective offset values (which include alignment adjustments)
            Dim actualEffSpanMm As Double = spanMm - effectiveStartOffsetMm - effectiveEndOffsetMm
            If allowZeroInstances AndAlso actualEffSpanMm <= maxSpacingMm Then
                Try
                    pattern.Suppress()
                    logs.Add("CenterPatternLib: Pattern suppressed (effective span " & _
                             actualEffSpanMm.ToString("0.00") & " mm <= maxSpacing " & maxSpacingMm.ToString("0.00") & " mm)")
                Catch
                End Try
            End If
        End If
        
        ' 12. Store extended configuration for re-runs
        ' Store the ORIGINAL user-provided values (maxSpacingInput as formula/value, offsets without alignment)
        ' Also store original occurrence name for restore/delete operations
        StorePatternConfigEx(patternSeed, baseName, startPlaneName, endPlaneName, axisName, _
                             maxSpacingInput, mode, includeEnds, _
                             startOffsetMm.ToString(), endOffsetMm.ToString(), _
                             startAlignment, endAlignment, allowZeroInstances.ToString(), _
                             seedOcc.Name)
        
        ' 13. Move helpers to Abivahendid folder
        logs.Add("CenterPatternLib: Moving helpers to Abivahendid folder...")
        MoveHelpersToFolder(asmDoc, baseName, logs)
        
        logs.Add("CenterPatternLib: Done")
        Return True
    End Function

    ' ============================================================================
    ' SECTION 7: Span Update Handler
    ' ============================================================================
    
    ''' <summary>
    ''' Registers an update handler with DocumentUpdateLib to automatically
    ''' update the span parameter when geometry changes.
    ''' </summary>
    Public Sub RegisterSpanUpdateHandler(doc As Document, _
                                          iLogicAuto As Object, _
                                          baseName As String, _
                                          startPlaneName As String, _
                                          endPlaneName As String, _
                                          axisName As String, _
                                          spanParamName As String, _
                                          Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing)
        If doc Is Nothing OrElse iLogicAuto Is Nothing Then
            If logs IsNot Nothing Then logs.Add("CenterPatternLib: Cannot register update handler - missing doc or iLogicAuto")
            Exit Sub
        End If
        
        ' Build the update code that will run when parameters change
        ' Uses the work axis direction to measure the distance between planes
        Dim updateCode() As String = { _
            "' Update span for pattern: " & baseName, _
            "Dim asmDoc As AssemblyDocument = CType(ThisDoc.Document, AssemblyDocument)", _
            "Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition", _
            "Dim startWP As WorkPlane = Nothing", _
            "Dim endWP As WorkPlane = Nothing", _
            "Dim dirAxis As WorkAxis = Nothing", _
            "For Each wp As WorkPlane In asmDef.WorkPlanes", _
            "    If wp.Name = """ & startPlaneName & """ Then startWP = wp", _
            "    If wp.Name = """ & endPlaneName & """ Then endWP = wp", _
            "Next", _
            "For Each wa As WorkAxis In asmDef.WorkAxes", _
            "    If wa.Name = """ & axisName & """ Then dirAxis = wa", _
            "Next", _
            "If startWP IsNot Nothing AndAlso endWP IsNot Nothing AndAlso dirAxis IsNot Nothing Then", _
            "    Dim p1 As Plane = startWP.Plane", _
            "    Dim p2 As Plane = endWP.Plane", _
            "    Dim axisDir As UnitVector = dirAxis.Line.Direction", _
            "    ' Measure distance along axis direction", _
            "    Dim dx As Double = p2.RootPoint.X - p1.RootPoint.X", _
            "    Dim dy As Double = p2.RootPoint.Y - p1.RootPoint.Y", _
            "    Dim dz As Double = p2.RootPoint.Z - p1.RootPoint.Z", _
            "    Dim dist As Double = Math.Abs(dx * axisDir.X + dy * axisDir.Y + dz * axisDir.Z) * 10", _
            "    Dim params As Parameters = asmDoc.ComponentDefinition.Parameters", _
            "    Try", _
            "        Dim p As Parameter = params.Item(""" & spanParamName & """)", _
            "        Dim currentVal As Double = p.Value * 10", _
            "        If Math.Abs(currentVal - dist) > 0.01 Then", _
            "            p.Expression = dist.ToString(System.Globalization.CultureInfo.InvariantCulture) & "" mm""", _
            "        End If", _
            "    Catch", _
            "    End Try", _
            "End If" _
        }
        
        ' Triggers: model parameter change causes geometry to move, which should update span
        Dim triggers() As DocumentUpdateLib.UpdateTrigger = { _
            DocumentUpdateLib.UpdateTrigger.ModelParameterChange _
        }
        
        Try
            Dim success As Boolean = DocumentUpdateLib.RegisterUpdateHandler( _
                doc, iLogicAuto, "CenterPattern_" & baseName, updateCode, triggers)
            
            If logs IsNot Nothing Then
                If success Then
                    logs.Add("CenterPatternLib: Registered span update handler for '" & baseName & "'")
                Else
                    logs.Add("CenterPatternLib: WARNING - Failed to register span update handler")
                End If
            End If
        Catch ex As Exception
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: ERROR registering update handler: " & ex.Message)
            End If
        End Try
    End Sub
    
    ''' <summary>
    ''' Extended update handler that also handles zero-instance case.
    ''' Self-contained VB code without library dependencies for portability.
    ''' </summary>
    Public Sub RegisterSpanUpdateHandlerEx(doc As Document, _
                                            iLogicAuto As Object, _
                                            baseName As String, _
                                            startPlaneName As String, _
                                            endPlaneName As String, _
                                            axisName As String, _
                                            spanParamName As String, _
                                            effectiveSpanParamName As String, _
                                            maxSpacingParamName As String, _
                                            patternName As String, _
                                            allowZeroInstances As Boolean, _
                                            Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing)
        If doc Is Nothing OrElse iLogicAuto Is Nothing Then
            If logs IsNot Nothing Then logs.Add("CenterPatternLib: Cannot register update handler - missing doc or iLogicAuto")
            Exit Sub
        End If
        
        ' Build self-contained update code
        ' This code does NOT use any external library functions for portability
        Dim codeLines As New System.Collections.Generic.List(Of String)
        
        codeLines.Add("' Update span and handle zero-instance case for pattern: " & baseName)
        codeLines.Add("Dim asmDoc As AssemblyDocument = CType(ThisDoc.Document, AssemblyDocument)")
        codeLines.Add("Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition")
        codeLines.Add("")
        codeLines.Add("' Find work features")
        codeLines.Add("Dim startWP As WorkPlane = Nothing")
        codeLines.Add("Dim endWP As WorkPlane = Nothing")
        codeLines.Add("Dim dirAxis As WorkAxis = Nothing")
        codeLines.Add("For Each wp As WorkPlane In asmDef.WorkPlanes")
        codeLines.Add("    If wp.Name = """ & startPlaneName & """ Then startWP = wp")
        codeLines.Add("    If wp.Name = """ & endPlaneName & """ Then endWP = wp")
        codeLines.Add("Next")
        codeLines.Add("For Each wa As WorkAxis In asmDef.WorkAxes")
        codeLines.Add("    If wa.Name = """ & axisName & """ Then dirAxis = wa")
        codeLines.Add("Next")
        codeLines.Add("")
        codeLines.Add("' Measure span if we have all work features")
        codeLines.Add("If startWP IsNot Nothing AndAlso endWP IsNot Nothing AndAlso dirAxis IsNot Nothing Then")
        codeLines.Add("    Dim p1 As Plane = startWP.Plane")
        codeLines.Add("    Dim p2 As Plane = endWP.Plane")
        codeLines.Add("    Dim axisDir As UnitVector = dirAxis.Line.Direction")
        codeLines.Add("    ")
        codeLines.Add("    ' Measure distance along axis direction")
        codeLines.Add("    Dim dx As Double = p2.RootPoint.X - p1.RootPoint.X")
        codeLines.Add("    Dim dy As Double = p2.RootPoint.Y - p1.RootPoint.Y")
        codeLines.Add("    Dim dz As Double = p2.RootPoint.Z - p1.RootPoint.Z")
        codeLines.Add("    Dim dist As Double = Math.Abs(dx * axisDir.X + dy * axisDir.Y + dz * axisDir.Z) * 10")
        codeLines.Add("    ")
        codeLines.Add("    ' Update span parameter if changed")
        codeLines.Add("    Dim params As Parameters = asmDoc.ComponentDefinition.Parameters")
        codeLines.Add("    Try")
        codeLines.Add("        Dim spanP As Parameter = params.Item(""" & spanParamName & """)")
        codeLines.Add("        Dim currentSpan As Double = spanP.Value * 10")
        codeLines.Add("        If Math.Abs(currentSpan - dist) > 0.01 Then")
        codeLines.Add("            spanP.Expression = dist.ToString(System.Globalization.CultureInfo.InvariantCulture) & "" mm""")
        codeLines.Add("        End If")
        codeLines.Add("    Catch")
        codeLines.Add("    End Try")
        
        ' Add zero-instance handling if enabled
        If allowZeroInstances Then
            codeLines.Add("    ")
            codeLines.Add("    ' Handle zero-instance case - suppress/unsuppress pattern")
            codeLines.Add("    Try")
            codeLines.Add("        Dim effSpanP As Parameter = params.Item(""" & effectiveSpanParamName & """)")
            codeLines.Add("        Dim maxSpP As Parameter = params.Item(""" & maxSpacingParamName & """)")
            codeLines.Add("        Dim effSpan As Double = effSpanP.Value * 10")
            codeLines.Add("        Dim maxSp As Double = maxSpP.Value * 10")
            codeLines.Add("        ")
            codeLines.Add("        ' Find the pattern")
            codeLines.Add("        For Each pat As OccurrencePattern In asmDef.OccurrencePatterns")
            codeLines.Add("            If pat.Name = """ & patternName & """ Then")
            codeLines.Add("                If effSpan <= maxSp Then")
            codeLines.Add("                    If Not pat.Suppressed Then pat.Suppress()")
            codeLines.Add("                Else")
            codeLines.Add("                    If pat.Suppressed Then pat.Unsuppress()")
            codeLines.Add("                End If")
            codeLines.Add("                Exit For")
            codeLines.Add("            End If")
            codeLines.Add("        Next")
            codeLines.Add("    Catch")
            codeLines.Add("    End Try")
        End If
        
        codeLines.Add("End If")
        
        ' Triggers
        Dim triggers() As DocumentUpdateLib.UpdateTrigger = { _
            DocumentUpdateLib.UpdateTrigger.ModelParameterChange _
        }
        
        Try
            Dim success As Boolean = DocumentUpdateLib.RegisterUpdateHandler( _
                doc, iLogicAuto, "CenterPattern_" & baseName, codeLines.ToArray(), triggers)
            
            If logs IsNot Nothing Then
                If success Then
                    logs.Add("CenterPatternLib: Registered extended update handler for '" & baseName & "'")
                Else
                    logs.Add("CenterPatternLib: WARNING - Failed to register extended update handler")
                End If
            End If
        Catch ex As Exception
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: ERROR registering extended update handler: " & ex.Message)
            End If
        End Try
    End Sub

    ' ============================================================================
    ' SECTION 7B: Update Mode (Re-run on Existing Pattern)
    ' ============================================================================
    
    ''' <summary>
    ''' Detect if a pattern already exists for the given occurrence.
    ''' Checks both the occurrence and its original (if suppressed).
    ''' Returns the seed occurrence that has configuration, or Nothing if not found.
    ''' </summary>
    Public Function DetectExistingPattern(asmDoc As AssemblyDocument, _
                                           occ As ComponentOccurrence) As ComponentOccurrence
        If asmDoc Is Nothing OrElse occ Is Nothing Then Return Nothing
        
        ' First check if this occurrence itself has pattern config
        If HasPatternConfig(occ) Then
            Return occ
        End If
        
        ' Check if occurrence name suggests it's a copy (contains "_Koopia")
        ' and look for related pattern seed
        Dim baseName As String = ""
        Dim occName As String = occ.Name
        
        ' Extract base name from occurrence name
        Dim colonPos As Integer = occName.LastIndexOf(":")
        If colonPos > 0 Then
            baseName = occName.Substring(0, colonPos)
        Else
            baseName = occName
        End If
        
        ' Check all occurrences for pattern config with matching base name
        For Each checkOcc As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
            If HasPatternConfig(checkOcc) Then
                Dim configBase As String = ""
                Dim dummy1, dummy2, dummy3, dummy4, dummy5, dummy6 As String
                Dim dummy7 As Boolean
                If LoadPatternConfig(checkOcc, configBase, dummy1, dummy2, dummy3, dummy4, dummy7) Then
                    ' Check if this pattern's base name matches
                    If baseName.Contains(configBase) OrElse configBase.Contains(baseName) Then
                        Return checkOcc
                    End If
                End If
            End If
        Next
        
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Update an existing pattern by re-measuring span and updating parameters.
    ''' Does not recreate work features or constraints, just updates values.
    ''' </summary>
    Public Function UpdateCenterPattern(app As Inventor.Application, _
                                         asmDoc As AssemblyDocument, _
                                         patternSeed As ComponentOccurrence, _
                                         Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing) As Boolean
        If logs Is Nothing Then logs = New System.Collections.Generic.List(Of String)
        
        If patternSeed Is Nothing Then
            logs.Add("CenterPatternLib: ERROR - No pattern seed for update")
            Return False
        End If
        
        ' Load configuration
        Dim baseName As String = ""
        Dim startPlaneName As String = ""
        Dim endPlaneName As String = ""
        Dim axisName As String = ""
        Dim maxSpacing As String = ""
        Dim mode As String = ""
        Dim includeEnds As Boolean = False
        Dim startOffset As String = ""
        Dim endOffset As String = ""
        Dim startAlignment As String = ""
        Dim endAlignment As String = ""
        Dim allowZeroInstances As Boolean = False
        
        If Not LoadPatternConfigEx(patternSeed, baseName, startPlaneName, endPlaneName, axisName, _
                                    maxSpacing, mode, includeEnds, startOffset, endOffset, _
                                    startAlignment, endAlignment, allowZeroInstances) Then
            ' Try basic config
            If Not LoadPatternConfig(patternSeed, baseName, startPlaneName, endPlaneName, _
                                      maxSpacing, mode, includeEnds) Then
                logs.Add("CenterPatternLib: ERROR - Could not load pattern configuration")
                Return False
            End If
        End If
        
        logs.Add("CenterPatternLib: Updating pattern '" & baseName & "'")
        
        ' Handle parameter names that start with digits
        Dim paramPrefix As String = baseName
        If baseName.Length > 0 AndAlso Char.IsDigit(baseName(0)) Then
            paramPrefix = "M_" & baseName
        End If
        
        ' Find work features
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        Dim startPlane As WorkPlane = WorkFeatureLib.FindWorkPlaneByName(asmDef, startPlaneName)
        Dim endPlane As WorkPlane = WorkFeatureLib.FindWorkPlaneByName(asmDef, endPlaneName)
        Dim dirAxis As WorkAxis = Nothing
        
        If Not String.IsNullOrEmpty(axisName) Then
            dirAxis = WorkFeatureLib.FindWorkAxisByName(asmDef, axisName)
        End If
        
        If startPlane Is Nothing OrElse endPlane Is Nothing Then
            logs.Add("CenterPatternLib: ERROR - Work planes not found")
            Return False
        End If
        
        ' Re-measure span
        Dim spanCm As Double = GeoLib.MeasurePlaneDistance(startPlane, endPlane)
        Dim spanMm As Double = spanCm * 10
        logs.Add("CenterPatternLib: New span = " & spanMm.ToString("0.00") & " mm")
        
        ' Update span parameter
        Dim spanParamName As String = paramPrefix & "_Ulatus"
        Try
            SetParameter(asmDoc, spanParamName, spanMm, "mm")
            logs.Add("CenterPatternLib: Updated " & spanParamName)
        Catch ex As Exception
            logs.Add("CenterPatternLib: WARNING - Could not update span: " & ex.Message)
        End Try
        
        ' Recalculate center offset if axis is available
        If dirAxis IsNot Nothing Then
            Dim axisDirection As UnitVector = GeoLib.GetAxisDirection(dirAxis)
            If axisDirection IsNot Nothing Then
                Dim centerOffsetCm As Double = CalculateCenterOffset(app, patternSeed, startPlane, axisDirection, spanCm, logs)
                Dim centerOffsetMm As Double = centerOffsetCm * 10
                
                Dim centerOffsetParamName As String = paramPrefix & "_KeskNihe"
                Try
                    SetParameter(asmDoc, centerOffsetParamName, centerOffsetMm, "mm")
                    logs.Add("CenterPatternLib: Updated " & centerOffsetParamName & " = " & centerOffsetMm.ToString("0.00") & " mm")
                Catch ex As Exception
                    logs.Add("CenterPatternLib: WARNING - Could not update center offset: " & ex.Message)
                End Try
                
                ' Recalculate part extent for alignment
                Dim minExtent As Double = 0
                Dim maxExtent As Double = 0
                Dim halfExtent As Double = 0
                CalculatePartExtentAlongAxis(app, patternSeed, axisDirection, minExtent, maxExtent, halfExtent, logs)
                
                Dim halfExtentParamName As String = paramPrefix & "_PoolUlatus"
                Try
                    SetParameter(asmDoc, halfExtentParamName, halfExtent * 10, "mm")
                    logs.Add("CenterPatternLib: Updated " & halfExtentParamName)
                Catch
                End Try
            End If
        End If
        
        logs.Add("CenterPatternLib: Pattern update complete")
        Return True
    End Function

    ' ============================================================================
    ' SECTION 8: Configuration Storage (Attributes)
    ' ============================================================================
    
    ''' <summary>
    ''' Store pattern configuration in attributes on the seed occurrence.
    ''' </summary>
    Public Sub StorePatternConfig(occ As ComponentOccurrence, _
                                   baseName As String, _
                                   startPlaneName As String, _
                                   endPlaneName As String, _
                                   maxSpacing As String, _
                                   mode As String, _
                                   includeEnds As Boolean)
        If occ Is Nothing Then Exit Sub
        
        Try
            ' Get or create attribute set
            Dim attrSet As AttributeSet
            If occ.AttributeSets.NameIsUsed(ATTR_SET_NAME) Then
                attrSet = occ.AttributeSets.Item(ATTR_SET_NAME)
            Else
                attrSet = occ.AttributeSets.Add(ATTR_SET_NAME)
            End If
            
            ' Store values
            SetAttribute(attrSet, "BaseName", baseName)
            SetAttribute(attrSet, "StartPlane", startPlaneName)
            SetAttribute(attrSet, "EndPlane", endPlaneName)
            SetAttribute(attrSet, "MaxSpacing", maxSpacing)
            SetAttribute(attrSet, "Mode", mode)
            SetAttribute(attrSet, "IncludeEnds", includeEnds.ToString())
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Load pattern configuration from attributes on an occurrence.
    ''' </summary>
    Public Function LoadPatternConfig(occ As ComponentOccurrence, _
                                       ByRef baseName As String, _
                                       ByRef startPlaneName As String, _
                                       ByRef endPlaneName As String, _
                                       ByRef maxSpacing As String, _
                                       ByRef mode As String, _
                                       ByRef includeEnds As Boolean) As Boolean
        baseName = ""
        startPlaneName = ""
        endPlaneName = ""
        maxSpacing = ""
        mode = MODE_UNIFORM
        includeEnds = False
        
        If occ Is Nothing Then Return False
        
        Try
            If Not occ.AttributeSets.NameIsUsed(ATTR_SET_NAME) Then Return False
            
            Dim attrSet As AttributeSet = occ.AttributeSets.Item(ATTR_SET_NAME)
            
            baseName = GetAttribute(attrSet, "BaseName")
            startPlaneName = GetAttribute(attrSet, "StartPlane")
            endPlaneName = GetAttribute(attrSet, "EndPlane")
            maxSpacing = GetAttribute(attrSet, "MaxSpacing")
            mode = GetAttribute(attrSet, "Mode")
            
            Dim includeEndsStr As String = GetAttribute(attrSet, "IncludeEnds")
            includeEnds = includeEndsStr.ToLower() = "true"
            
            Return baseName <> ""
        Catch
            Return False
        End Try
    End Function
    
    Private Sub SetAttribute(attrSet As AttributeSet, name As String, value As String)
        Try
            If attrSet.NameIsUsed(name) Then
                attrSet.Item(name).Value = value
            Else
                attrSet.Add(name, ValueTypeEnum.kStringType, value)
            End If
        Catch
        End Try
    End Sub
    
    Private Function GetAttribute(attrSet As AttributeSet, name As String) As String
        Try
            If attrSet.NameIsUsed(name) Then
                Return CStr(attrSet.Item(name).Value)
            End If
        Catch
        End Try
        Return ""
    End Function
    
    ''' <summary>
    ''' Store extended pattern configuration in attributes on the seed occurrence.
    ''' </summary>
    Public Sub StorePatternConfigEx(occ As ComponentOccurrence, _
                                     baseName As String, _
                                     startPlaneName As String, _
                                     endPlaneName As String, _
                                     axisName As String, _
                                     maxSpacing As String, _
                                     mode As String, _
                                     includeEnds As Boolean, _
                                     startOffset As String, _
                                     endOffset As String, _
                                     startAlignment As String, _
                                     endAlignment As String, _
                                     allowZeroInstances As String, _
                                     Optional originalOccName As String = "")
        If occ Is Nothing Then Exit Sub
        
        Try
            ' Get or create attribute set
            Dim attrSet As AttributeSet
            If occ.AttributeSets.NameIsUsed(ATTR_SET_NAME) Then
                attrSet = occ.AttributeSets.Item(ATTR_SET_NAME)
            Else
                attrSet = occ.AttributeSets.Add(ATTR_SET_NAME)
            End If
            
            ' Store basic values
            SetAttribute(attrSet, "BaseName", baseName)
            SetAttribute(attrSet, "StartPlane", startPlaneName)
            SetAttribute(attrSet, "EndPlane", endPlaneName)
            SetAttribute(attrSet, "AxisName", axisName)
            SetAttribute(attrSet, "MaxSpacing", maxSpacing)
            SetAttribute(attrSet, "Mode", mode)
            SetAttribute(attrSet, "IncludeEnds", includeEnds.ToString())
            
            ' Store extended values
            SetAttribute(attrSet, "StartOffset", startOffset)
            SetAttribute(attrSet, "EndOffset", endOffset)
            SetAttribute(attrSet, "StartAlignment", startAlignment)
            SetAttribute(attrSet, "EndAlignment", endAlignment)
            SetAttribute(attrSet, "AllowZeroInstances", allowZeroInstances)
            
            ' Store original occurrence name (for restore/delete operations)
            If Not String.IsNullOrEmpty(originalOccName) Then
                SetAttribute(attrSet, "OriginalOccName", originalOccName)
            End If
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Load extended pattern configuration from attributes on an occurrence.
    ''' </summary>
    Public Function LoadPatternConfigEx(occ As ComponentOccurrence, _
                                         ByRef baseName As String, _
                                         ByRef startPlaneName As String, _
                                         ByRef endPlaneName As String, _
                                         ByRef axisName As String, _
                                         ByRef maxSpacing As String, _
                                         ByRef mode As String, _
                                         ByRef includeEnds As Boolean, _
                                         ByRef startOffset As String, _
                                         ByRef endOffset As String, _
                                         ByRef startAlignment As String, _
                                         ByRef endAlignment As String, _
                                         ByRef allowZeroInstances As Boolean) As Boolean
        ' Initialize defaults
        baseName = ""
        startPlaneName = ""
        endPlaneName = ""
        axisName = ""
        maxSpacing = ""
        mode = MODE_UNIFORM
        includeEnds = False
        startOffset = "0"
        endOffset = "0"
        startAlignment = ALIGN_CENTER
        endAlignment = ALIGN_CENTER
        allowZeroInstances = False
        
        If occ Is Nothing Then Return False
        
        Try
            If Not occ.AttributeSets.NameIsUsed(ATTR_SET_NAME) Then Return False
            
            Dim attrSet As AttributeSet = occ.AttributeSets.Item(ATTR_SET_NAME)
            
            ' Load basic values
            baseName = GetAttribute(attrSet, "BaseName")
            startPlaneName = GetAttribute(attrSet, "StartPlane")
            endPlaneName = GetAttribute(attrSet, "EndPlane")
            axisName = GetAttribute(attrSet, "AxisName")
            maxSpacing = GetAttribute(attrSet, "MaxSpacing")
            mode = GetAttribute(attrSet, "Mode")
            
            Dim includeEndsStr As String = GetAttribute(attrSet, "IncludeEnds")
            includeEnds = includeEndsStr.ToLower() = "true"
            
            ' Load extended values
            startOffset = GetAttribute(attrSet, "StartOffset")
            If String.IsNullOrEmpty(startOffset) Then startOffset = "0"
            
            endOffset = GetAttribute(attrSet, "EndOffset")
            If String.IsNullOrEmpty(endOffset) Then endOffset = "0"
            
            startAlignment = GetAttribute(attrSet, "StartAlignment")
            If String.IsNullOrEmpty(startAlignment) Then startAlignment = ALIGN_CENTER
            
            endAlignment = GetAttribute(attrSet, "EndAlignment")
            If String.IsNullOrEmpty(endAlignment) Then endAlignment = ALIGN_CENTER
            
            Dim azStr As String = GetAttribute(attrSet, "AllowZeroInstances")
            allowZeroInstances = azStr.ToLower() = "true"
            
            Return baseName <> ""
        Catch
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Get the original occurrence name from pattern configuration.
    ''' Returns empty string if not found.
    ''' </summary>
    Public Function GetOriginalOccurrenceName(occ As ComponentOccurrence) As String
        If occ Is Nothing Then Return ""
        
        Try
            If Not occ.AttributeSets.NameIsUsed(ATTR_SET_NAME) Then Return ""
            
            Dim attrSet As AttributeSet = occ.AttributeSets.Item(ATTR_SET_NAME)
            Return GetAttribute(attrSet, "OriginalOccName")
        Catch
            Return ""
        End Try
    End Function
    
    ''' <summary>
    ''' Find the original occurrence by name from pattern config.
    ''' </summary>
    Public Function FindOriginalOccurrence(asmDoc As AssemblyDocument, patternSeed As ComponentOccurrence) As ComponentOccurrence
        If asmDoc Is Nothing OrElse patternSeed Is Nothing Then Return Nothing
        
        Dim origName As String = GetOriginalOccurrenceName(patternSeed)
        If String.IsNullOrEmpty(origName) Then Return Nothing
        
        Try
            Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
            For Each occ As ComponentOccurrence In asmDef.Occurrences
                If occ.Name = origName Then Return occ
            Next
        Catch
        End Try
        
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Check if an occurrence has pattern configuration stored.
    ''' </summary>
    Public Function HasPatternConfig(occ As ComponentOccurrence) As Boolean
        If occ Is Nothing Then Return False
        
        Try
            If Not occ.AttributeSets.NameIsUsed(ATTR_SET_NAME) Then Return False
            
            Dim attrSet As AttributeSet = occ.AttributeSets.Item(ATTR_SET_NAME)
            Dim baseName As String = GetAttribute(attrSet, "BaseName")
            Return Not String.IsNullOrEmpty(baseName)
        Catch
            Return False
        End Try
    End Function

    ' ============================================================================
    ' SECTION 10: Browser Organization
    ' ============================================================================
    
    Private Const HELPERS_FOLDER As String = "Abivahendid"
    Private Const TEMPLATE_FOLDER As String = "Template"
    
    ''' <summary>
    ''' Move original occurrence to Template folder.
    ''' This must be called BEFORE creating the pattern so the original stays at top level
    ''' in the browser tree and can be moved to a folder.
    ''' </summary>
    Public Sub MoveOriginalToTemplateFolder(asmDoc As AssemblyDocument, _
                                             originalOcc As ComponentOccurrence, _
                                             Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing)
        If asmDoc Is Nothing OrElse originalOcc Is Nothing Then Exit Sub
        
        Try
            Dim oPane As BrowserPane = asmDoc.BrowserPanes.Item("Model")
            If oPane Is Nothing Then Exit Sub
            
            ' Create/get Template folder
            Dim templateFolder As BrowserFolder = UtilsLib.GetOrCreateFolder(oPane, TEMPLATE_FOLDER)
            
            ' Move original occurrence to Template folder
            MoveObjectToFolder(oPane, templateFolder, originalOcc, logs)
            
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: Original moved to Template folder")
            End If
            
        Catch ex As Exception
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: WARNING - Could not move original to Template folder: " & ex.Message)
            End If
        End Try
    End Sub
    
    ''' <summary>
    ''' Move work features (planes, axis) to Abivahendid folder.
    ''' Called after creating work features but can be before or after pattern creation.
    ''' </summary>
    Public Sub MoveHelpersToFolder(asmDoc As AssemblyDocument, _
                                    baseName As String, _
                                    Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing)
        If asmDoc Is Nothing Then Exit Sub
        
        Try
            Dim oPane As BrowserPane = asmDoc.BrowserPanes.Item("Model")
            If oPane Is Nothing Then Exit Sub
            
            ' Create/get Abivahendid folder
            Dim helpersFolder As BrowserFolder = UtilsLib.GetOrCreateFolder(oPane, HELPERS_FOLDER)
            
            Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
            
            ' Move work features to Abivahendid folder
            Dim workFeatureNames As String() = { _
                baseName & "_AlgusTasand", _
                baseName & "_LõpuTasand", _
                baseName & "_Telg" _
            }
            
            For Each wfName As String In workFeatureNames
                Try
                    ' Try work planes
                    Dim wp As WorkPlane = WorkFeatureLib.FindWorkPlaneByName(asmDef, wfName)
                    If wp IsNot Nothing Then
                        MoveObjectToFolder(oPane, helpersFolder, wp, logs)
                    End If
                Catch
                End Try
                
                Try
                    ' Try work axes
                    Dim wa As WorkAxis = WorkFeatureLib.FindWorkAxisByName(asmDef, wfName)
                    If wa IsNot Nothing Then
                        MoveObjectToFolder(oPane, helpersFolder, wa, logs)
                    End If
                Catch
                End Try
            Next
            
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: Helpers moved to Abivahendid folder")
            End If
            
        Catch ex As Exception
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: WARNING - Could not move helpers to folder: " & ex.Message)
            End If
        End Try
    End Sub
    
    ''' <summary>
    ''' Organize browser items: move work features to "Abivahendid" folder,
    ''' move original occurrence to "Template" folder.
    ''' NOTE: This is the original combined function - prefer using the split functions
    ''' MoveOriginalToTemplateFolder and MoveHelpersToFolder for proper ordering.
    ''' </summary>
    Public Sub OrganizeBrowserItems(asmDoc As AssemblyDocument, _
                                     baseName As String, _
                                     originalOcc As ComponentOccurrence, _
                                     Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing)
        MoveOriginalToTemplateFolder(asmDoc, originalOcc, logs)
        MoveHelpersToFolder(asmDoc, baseName, logs)
    End Sub
    
    ''' <summary>
    ''' Move an object to a browser folder.
    ''' </summary>
    Private Sub MoveObjectToFolder(oPane As BrowserPane, folder As BrowserFolder, obj As Object, _
                                   Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing)
        If oPane Is Nothing OrElse folder Is Nothing OrElse obj Is Nothing Then Exit Sub
        
        Try
            Dim node As BrowserNode = oPane.GetBrowserNodeFromObject(obj)
            If node IsNot Nothing Then
                ' Get the movable parent node
                Dim movableNode As BrowserNode = UtilsLib.GetMovableParentNode(oPane, node)
                If movableNode IsNot Nothing Then
                    folder.Add(movableNode)
                End If
            End If
        Catch ex As Exception
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: Could not move object to folder: " & ex.Message)
            End If
        End Try
    End Sub
    
    ''' <summary>
    ''' Remove an object from its browser folder (move back to top level).
    ''' </summary>
    Public Sub RemoveFromBrowserFolder(asmDoc As AssemblyDocument, obj As Object, _
                                        Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing)
        If asmDoc Is Nothing OrElse obj Is Nothing Then Exit Sub
        
        Try
            Dim oPane As BrowserPane = asmDoc.BrowserPanes.Item("Model")
            If oPane Is Nothing Then Exit Sub
            
            Dim node As BrowserNode = oPane.GetBrowserNodeFromObject(obj)
            If node IsNot Nothing Then
                Dim movableNode As BrowserNode = UtilsLib.GetMovableParentNode(oPane, node)
                If movableNode IsNot Nothing AndAlso movableNode.Parent IsNot Nothing Then
                    ' Check if parent is a BrowserFolder
                    If movableNode.Parent.NativeObject IsNot Nothing Then
                        If TypeName(movableNode.Parent.NativeObject) = "BrowserFolder" Then
                            ' Remove from folder by adding to top node
                            ' Note: Inventor doesn't have a direct "remove from folder" API
                            ' The item will stay where it is, but we can try to move it back
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            If logs IsNot Nothing Then
                logs.Add("CenterPatternLib: Could not remove object from folder: " & ex.Message)
            End If
        End Try
    End Sub

    ' ============================================================================
    ' SECTION 11: Pattern Rebuild and Delete/Cleanup
    ' ============================================================================
    
    ''' <summary>
    ''' Rebuild an existing pattern with new settings.
    ''' Deletes the existing pattern and creates a new one.
    ''' Work features are handled individually - only replaced if geometry changed.
    ''' </summary>
    Public Function RebuildCenterPattern(app As Inventor.Application, _
                                          asmDoc As AssemblyDocument, _
                                          iLogicAuto As Object, _
                                          existingSeed As ComponentOccurrence, _
                                          startGeometry As Object, _
                                          endGeometry As Object, _
                                          explicitAxis As Object, _
                                          maxSpacingInput As String, _
                                          mode As String, _
                                          includeEnds As Boolean, _
                                          baseName As String, _
                                          startOffsetMm As Double, _
                                          endOffsetMm As Double, _
                                          startAlignment As String, _
                                          endAlignment As String, _
                                          allowZeroInstances As Boolean, _
                                          Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing) As Boolean
        If logs Is Nothing Then logs = New System.Collections.Generic.List(Of String)
        
        logs.Add("CenterPatternLib: === REBUILDING PATTERN ===")
        
        ' Find original occurrence before deleting
        Dim originalOcc As ComponentOccurrence = FindOriginalOccurrence(asmDoc, existingSeed)
        If originalOcc Is Nothing Then
            logs.Add("CenterPatternLib: ERROR - Could not find original occurrence for rebuild")
            Return False
        End If
        
        logs.Add("CenterPatternLib: Found original occurrence: " & originalOcc.Name)
        
        ' Load existing config to get work feature names
        Dim cfgBaseName As String = ""
        Dim startPlaneName As String = ""
        Dim endPlaneName As String = ""
        Dim axisName As String = ""
        Dim cfgMaxSpacing As String = ""
        Dim cfgMode As String = ""
        Dim cfgIncludeEnds As Boolean = False
        Dim cfgStartOffset As String = ""
        Dim cfgEndOffset As String = ""
        Dim cfgStartAlign As String = ""
        Dim cfgEndAlign As String = ""
        Dim cfgAllowZero As Boolean = False
        
        LoadPatternConfigEx(existingSeed, cfgBaseName, startPlaneName, endPlaneName, axisName, _
                            cfgMaxSpacing, cfgMode, cfgIncludeEnds, cfgStartOffset, cfgEndOffset, _
                            cfgStartAlign, cfgEndAlign, cfgAllowZero)
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' Check each work feature individually - delete if geometry changed
        Dim deleteStartPlane As Boolean = Not IsGeometrySameAsWorkFeature(startGeometry, startPlaneName, asmDef)
        Dim deleteEndPlane As Boolean = Not IsGeometrySameAsWorkFeature(endGeometry, endPlaneName, asmDef)
        Dim deleteAxis As Boolean = Not IsGeometrySameAsWorkFeature(explicitAxis, axisName, asmDef)
        
        If deleteStartPlane Then
            logs.Add("CenterPatternLib: Start geometry changed - will recreate start plane")
        End If
        If deleteEndPlane Then
            logs.Add("CenterPatternLib: End geometry changed - will recreate end plane")
        End If
        If deleteAxis Then
            logs.Add("CenterPatternLib: Axis changed - will recreate axis")
        End If
        
        ' Delete existing pattern (keep all work features for now)
        logs.Add("CenterPatternLib: Deleting existing pattern...")
        If Not DeleteCenterPattern(asmDoc, iLogicAuto, existingSeed, keepWorkFeatures:=True, logs:=logs) Then
            logs.Add("CenterPatternLib: WARNING - Pattern delete had issues, continuing with create")
        End If
        
        ' Now delete the work features that need to be recreated
        If deleteStartPlane AndAlso Not String.IsNullOrEmpty(startPlaneName) Then
            DeleteWorkFeatureByName(asmDef, startPlaneName, logs)
        End If
        If deleteEndPlane AndAlso Not String.IsNullOrEmpty(endPlaneName) Then
            DeleteWorkFeatureByName(asmDef, endPlaneName, logs)
        End If
        If deleteAxis AndAlso Not String.IsNullOrEmpty(axisName) Then
            DeleteWorkFeatureByName(asmDef, axisName, logs)
        End If
        
        ' Create new pattern with the original occurrence as seed
        logs.Add("CenterPatternLib: Creating new pattern with updated settings...")
        Return CreateCenterPatternEx(app, asmDoc, iLogicAuto, originalOcc, _
                                      startGeometry, endGeometry, explicitAxis, _
                                      maxSpacingInput, mode, includeEnds, baseName, _
                                      startOffsetMm, endOffsetMm, _
                                      startAlignment, endAlignment, _
                                      allowZeroInstances, logs)
    End Function
    
    ''' <summary>
    ''' Check if the selected geometry is the same as an existing work feature.
    ''' Returns True if geometry IS the work feature (so it should be kept).
    ''' Returns False if geometry is different (so work feature should be recreated).
    ''' </summary>
    Private Function IsGeometrySameAsWorkFeature(geometry As Object, workFeatureName As String, _
                                                  asmDef As AssemblyComponentDefinition) As Boolean
        If geometry Is Nothing Then Return False
        If String.IsNullOrEmpty(workFeatureName) Then Return False
        
        ' If geometry is a WorkPlane, check if it's the same as the named one
        If TypeOf geometry Is WorkPlane Then
            Dim geoPlane As WorkPlane = CType(geometry, WorkPlane)
            Dim existingPlane As WorkPlane = WorkFeatureLib.FindWorkPlaneByName(asmDef, workFeatureName)
            If existingPlane IsNot Nothing Then
                ' Compare by name (most reliable way)
                Return geoPlane.Name = existingPlane.Name
            End If
        End If
        
        ' If geometry is a WorkAxis, check if it's the same as the named one
        If TypeOf geometry Is WorkAxis Then
            Dim geoAxis As WorkAxis = CType(geometry, WorkAxis)
            Dim existingAxis As WorkAxis = WorkFeatureLib.FindWorkAxisByName(asmDef, workFeatureName)
            If existingAxis IsNot Nothing Then
                Return geoAxis.Name = existingAxis.Name
            End If
        End If
        
        ' Geometry is something else (face, point, etc.) - not the same as work feature
        Return False
    End Function
    
    ''' <summary>
    ''' Delete a center pattern and restore the state before the pattern was created.
    ''' This removes:
    ''' - The occurrence pattern
    ''' - All constraints (perpendicular + positioning)
    ''' - Work features (start/end planes, axis) - unless keepWorkFeatures is True
    ''' - All parameters (_Ulatus, _MaxVahe, etc.)
    ''' - The update handler rule
    ''' - The seed copy occurrence
    ''' And restores:
    ''' - Original occurrence (visible, BOM=Default)
    ''' 
    ''' Set keepWorkFeatures=True when rebuilding to reuse existing work planes/axis.
    ''' </summary>
    Public Function DeleteCenterPattern(asmDoc As AssemblyDocument, _
                                         iLogicAuto As Object, _
                                         patternSeed As ComponentOccurrence, _
                                         Optional keepWorkFeatures As Boolean = False, _
                                         Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing) As Boolean
        If logs Is Nothing Then logs = New System.Collections.Generic.List(Of String)
        
        If asmDoc Is Nothing OrElse patternSeed Is Nothing Then
            logs.Add("CenterPatternLib: ERROR - Invalid parameters for delete")
            Return False
        End If
        
        ' Load configuration
        Dim baseName As String = ""
        Dim startPlaneName As String = ""
        Dim endPlaneName As String = ""
        Dim axisName As String = ""
        Dim maxSpacing As String = ""
        Dim mode As String = ""
        Dim includeEnds As Boolean = False
        Dim startOffset As String = ""
        Dim endOffset As String = ""
        Dim startAlignment As String = ""
        Dim endAlignment As String = ""
        Dim allowZeroInstances As Boolean = False
        
        If Not LoadPatternConfigEx(patternSeed, baseName, startPlaneName, endPlaneName, axisName, _
                                    maxSpacing, mode, includeEnds, startOffset, endOffset, _
                                    startAlignment, endAlignment, allowZeroInstances) Then
            logs.Add("CenterPatternLib: ERROR - Could not load pattern configuration")
            Return False
        End If
        
        logs.Add("CenterPatternLib: Deleting pattern '" & baseName & "'")
        
        ' Handle parameter prefix (same as in creation)
        Dim paramPrefix As String = baseName
        If baseName.Length > 0 AndAlso Char.IsDigit(baseName(0)) Then
            paramPrefix = "M_" & baseName
        End If
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' 1. Find and store reference to original occurrence BEFORE deleting seed
        Dim originalOcc As ComponentOccurrence = FindOriginalOccurrence(asmDoc, patternSeed)
        
        ' 2. Delete occurrence pattern
        Try
            Dim patternName As String = baseName & "_Muster"
            PatternLib.DeletePatternByName(asmDef, patternName)
            logs.Add("CenterPatternLib: Deleted pattern '" & patternName & "'")
        Catch ex As Exception
            logs.Add("CenterPatternLib: WARNING - Could not delete pattern: " & ex.Message)
        End Try
        
        ' 3. Delete constraints (perpendicular + positioning)
        DeleteConstraintsWithPrefix(asmDef, baseName & "_", logs)
        
        ' 4. Delete work features (unless rebuilding)
        If Not keepWorkFeatures Then
            DeleteWorkFeatureByName(asmDef, startPlaneName, logs)
            DeleteWorkFeatureByName(asmDef, endPlaneName, logs)
            If Not String.IsNullOrEmpty(axisName) Then
                DeleteWorkFeatureByName(asmDef, axisName, logs)
            End If
        Else
            logs.Add("CenterPatternLib: Keeping work features for rebuild")
        End If
        
        ' 5. Delete parameters
        DeleteParametersWithPrefix(asmDoc, paramPrefix & "_", logs)
        
        ' 6. Delete iLogic update handler rule
        If iLogicAuto IsNot Nothing Then
            Dim ruleName As String = baseName & " Uuenda"
            Try
                Dim existingRule As Object = Nothing
                Try
                    existingRule = iLogicAuto.GetRule(asmDoc, ruleName)
                Catch
                End Try
                
                If existingRule IsNot Nothing Then
                    iLogicAuto.DeleteRule(asmDoc, ruleName)
                    logs.Add("CenterPatternLib: Deleted rule '" & ruleName & "'")
                End If
            Catch ex As Exception
                logs.Add("CenterPatternLib: WARNING - Could not delete rule: " & ex.Message)
            End Try
        End If
        
        ' 7. Restore original occurrence
        If originalOcc IsNot Nothing Then
            PatternLib.RestoreHiddenSeed(originalOcc)
            logs.Add("CenterPatternLib: Restored original occurrence '" & originalOcc.Name & "'")
        Else
            logs.Add("CenterPatternLib: WARNING - Could not find original occurrence to restore")
        End If
        
        ' 8. Remove config attributes from seed
        Try
            If patternSeed.AttributeSets.NameIsUsed(ATTR_SET_NAME) Then
                patternSeed.AttributeSets.Item(ATTR_SET_NAME).Delete()
                logs.Add("CenterPatternLib: Removed config attributes")
            End If
        Catch
        End Try
        
        ' 9. Delete seed copy occurrence
        Try
            patternSeed.Delete()
            logs.Add("CenterPatternLib: Deleted seed copy")
        Catch ex As Exception
            logs.Add("CenterPatternLib: WARNING - Could not delete seed copy: " & ex.Message)
        End Try
        
        logs.Add("CenterPatternLib: Pattern deleted successfully")
        Return True
    End Function
    
    ''' <summary>
    ''' Delete all assembly constraints that start with the given prefix.
    ''' </summary>
    Private Sub DeleteConstraintsWithPrefix(asmDef As AssemblyComponentDefinition, prefix As String, _
                                            Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing)
        If asmDef Is Nothing OrElse String.IsNullOrEmpty(prefix) Then Exit Sub
        
        Dim toDelete As New System.Collections.Generic.List(Of AssemblyConstraint)
        
        Try
            For Each c As AssemblyConstraint In asmDef.Constraints
                If c.Name.StartsWith(prefix) Then
                    toDelete.Add(c)
                End If
            Next
            
            For Each c As AssemblyConstraint In toDelete
                Try
                    Dim cName As String = c.Name
                    c.Delete()
                    If logs IsNot Nothing Then logs.Add("CenterPatternLib: Deleted constraint '" & cName & "'")
                Catch
                End Try
            Next
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Delete a work feature (plane or axis) by name.
    ''' </summary>
    Private Sub DeleteWorkFeatureByName(asmDef As AssemblyComponentDefinition, featureName As String, _
                                        Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing)
        If asmDef Is Nothing OrElse String.IsNullOrEmpty(featureName) Then Exit Sub
        
        ' Try as work plane
        Try
            Dim wp As WorkPlane = WorkFeatureLib.FindWorkPlaneByName(asmDef, featureName)
            If wp IsNot Nothing Then
                wp.Delete()
                If logs IsNot Nothing Then logs.Add("CenterPatternLib: Deleted work plane '" & featureName & "'")
                Exit Sub
            End If
        Catch
        End Try
        
        ' Try as work axis
        Try
            Dim wa As WorkAxis = WorkFeatureLib.FindWorkAxisByName(asmDef, featureName)
            If wa IsNot Nothing Then
                wa.Delete()
                If logs IsNot Nothing Then logs.Add("CenterPatternLib: Deleted work axis '" & featureName & "'")
                Exit Sub
            End If
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Delete all parameters that start with the given prefix.
    ''' </summary>
    Private Sub DeleteParametersWithPrefix(asmDoc As AssemblyDocument, prefix As String, _
                                            Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing)
        If asmDoc Is Nothing OrElse String.IsNullOrEmpty(prefix) Then Exit Sub
        
        Dim params As Parameters = asmDoc.ComponentDefinition.Parameters
        Dim toDelete As New System.Collections.Generic.List(Of String)
        
        Try
            For Each p As Parameter In params.UserParameters
                If p.Name.StartsWith(prefix) Then
                    toDelete.Add(p.Name)
                End If
            Next
            
            For Each pName As String In toDelete
                Try
                    Dim p As Parameter = params.Item(pName)
                    p.Delete()
                    If logs IsNot Nothing Then logs.Add("CenterPatternLib: Deleted parameter '" & pName & "'")
                Catch
                End Try
            Next
        Catch
        End Try
    End Sub

End Module
