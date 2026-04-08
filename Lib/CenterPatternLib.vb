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
    ' SECTION 4: Center Offset Calculation
    ' ============================================================================
    
    ''' <summary>
    ''' Calculate the center offset - how far the part's principal plane is from
    ''' the center of the span when measured along the axis.
    ''' 
    ''' This captures the initial position of the part so that when we constrain
    ''' it to the start plane with an offset, it ends up at the correct position.
    ''' 
    ''' centerOffset = (distance from start plane to part's principal plane) - (span/2)
    ''' 
    ''' Returns offset in cm (internal units).
    ''' </summary>
    Public Function CalculateCenterOffset(app As Inventor.Application, _
                                           occ As ComponentOccurrence, _
                                           startPlane As WorkPlane, _
                                           axisDirection As UnitVector, _
                                           spanCm As Double, _
                                           Optional ByRef logs As System.Collections.Generic.List(Of String) = Nothing) As Double
        If occ Is Nothing OrElse startPlane Is Nothing OrElse axisDirection Is Nothing Then
            Return 0
        End If
        
        ' Get start plane geometry
        Dim startPlaneGeom As Plane = GeoLib.GetPlaneGeometry(startPlane)
        If startPlaneGeom Is Nothing Then Return 0
        
        Dim startPt As Point = startPlaneGeom.RootPoint
        Dim startNormal As UnitVector = startPlaneGeom.Normal
        
        If logs IsNot Nothing Then
            logs.Add("CenterPatternLib: === OFFSET CALCULATION ===")
            logs.Add("CenterPatternLib: Start plane root = (" & _
                     (startPt.X * 10).ToString("0.00") & ", " & _
                     (startPt.Y * 10).ToString("0.00") & ", " & _
                     (startPt.Z * 10).ToString("0.00") & ") mm")
            logs.Add("CenterPatternLib: Start plane normal = (" & _
                     startNormal.X.ToString("0.000") & ", " & _
                     startNormal.Y.ToString("0.000") & ", " & _
                     startNormal.Z.ToString("0.000") & ")")
            logs.Add("CenterPatternLib: Axis direction = (" & _
                     axisDirection.X.ToString("0.000") & ", " & _
                     axisDirection.Y.ToString("0.000") & ", " & _
                     axisDirection.Z.ToString("0.000") & ")")
        End If
        
        ' Find the principal plane of the occurrence (perpendicular to axis)
        Dim principalPlaneProxy As Object = GeoLib.FindPrincipalPlane(occ, axisDirection)
        If principalPlaneProxy Is Nothing Then Return 0
        
        Dim principalGeom As Plane = GeoLib.GetPlaneGeometry(principalPlaneProxy)
        If principalGeom Is Nothing Then Return 0
        
        Dim partPlanePt As Point = principalGeom.RootPoint
        Dim principalNormal As UnitVector = principalGeom.Normal
        
        If logs IsNot Nothing Then
            logs.Add("CenterPatternLib: Principal plane root = (" & _
                     (partPlanePt.X * 10).ToString("0.00") & ", " & _
                     (partPlanePt.Y * 10).ToString("0.00") & ", " & _
                     (partPlanePt.Z * 10).ToString("0.00") & ") mm")
            logs.Add("CenterPatternLib: Principal plane normal = (" & _
                     principalNormal.X.ToString("0.000") & ", " & _
                     principalNormal.Y.ToString("0.000") & ", " & _
                     principalNormal.Z.ToString("0.000") & ")")
        End If
        
        ' Calculate distance from start plane to part plane along axis
        Dim toPartX As Double = partPlanePt.X - startPt.X
        Dim toPartY As Double = partPlanePt.Y - startPt.Y
        Dim toPartZ As Double = partPlanePt.Z - startPt.Z
        
        Dim distFromStart As Double = toPartX * axisDirection.X + _
                                       toPartY * axisDirection.Y + _
                                       toPartZ * axisDirection.Z
        
        ' Dot products for understanding directions
        Dim startNormalDotAxis As Double = startNormal.X * axisDirection.X + _
                                            startNormal.Y * axisDirection.Y + _
                                            startNormal.Z * axisDirection.Z
        Dim principalNormalDotAxis As Double = principalNormal.X * axisDirection.X + _
                                                principalNormal.Y * axisDirection.Y + _
                                                principalNormal.Z * axisDirection.Z
        
        If logs IsNot Nothing Then
            logs.Add("CenterPatternLib: Distance from start to principal (along axis) = " & _
                     (distFromStart * 10).ToString("0.00") & " mm")
            logs.Add("CenterPatternLib: Start normal · axis = " & startNormalDotAxis.ToString("0.000"))
            logs.Add("CenterPatternLib: Principal normal · axis = " & principalNormalDotAxis.ToString("0.000"))
            logs.Add("CenterPatternLib: Span/2 = " & (spanCm * 10 / 2).ToString("0.00") & " mm")
        End If
        
        ' Center offset = actual distance - half span
        ' If part is at center (spanCm/2 from start), centerOffset = 0
        ' If part is closer to start, centerOffset < 0
        ' If part is closer to end, centerOffset > 0
        Dim centerOffset As Double = distFromStart - (spanCm / 2)
        
        If logs IsNot Nothing Then
            logs.Add("CenterPatternLib: Center offset = " & (centerOffset * 10).ToString("0.00") & " mm")
            logs.Add("CenterPatternLib: === END OFFSET CALCULATION ===")
        End If
        
        Return centerOffset
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

End Module
