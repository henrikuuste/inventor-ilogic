' Lehtmetall - Convert solid part to sheet metal
' Converts the active part to sheet metal, measures thickness from geometry,
' exports Thickness as iProperty, sets Width/Length custom properties
' with sheet metal expressions, and creates flat pattern.

AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/CustomPropertiesLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = ThisDoc.Document
    
    ' Enable immediate logging
    UtilsLib.SetLogger(Logger)
    
    UtilsLib.LogInfo("Lehtmetall: Starting conversion...")
    
    ' Validate document type
    If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        UtilsLib.LogError("Lehtmetall: This rule can only be run on a part document.")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    
    ' Check if solid body exists
    If compDef.SurfaceBodies.Count = 0 Then
        UtilsLib.LogError("Lehtmetall: Part has no solid body. Ensure the part contains geometry before converting to sheet metal.")
        Exit Sub
    End If
    
    ' Check if solid body exists (not just surfaces)
    Dim hasSolidBody As Boolean = False
    For Each body As SurfaceBody In compDef.SurfaceBodies
        If body.IsSolid Then
            hasSolidBody = True
            Exit For
        End If
    Next
    
    If Not hasSolidBody Then
        UtilsLib.LogError("Lehtmetall: Part has no solid body. Only surfaces found, which are not suitable for sheet metal conversion.")
        Exit Sub
    End If
    
    ' Check if already sheet metal
    Const SHEET_METAL_GUID As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    If partDoc.SubType = SHEET_METAL_GUID Then
        UtilsLib.LogInfo("Lehtmetall: Part is already sheet metal. Validating properties and flat pattern...")
        ValidateAndRepairExistingSheetMetal(app, partDoc)
        Exit Sub
    End If
    
    ' Pick A-side face BEFORE conversion (to measure thickness)
    Dim aSideFace As Face = PickASideFace(app)
    If aSideFace Is Nothing Then
        Exit Sub
    End If
    
    ' Measure thickness along normal
    Dim thickness As Double = MeasureThicknessAlongNormal(app, aSideFace)
    UtilsLib.LogInfo("Lehtmetall: Measured thickness = " & FormatNumber(thickness * 10, 2) & " mm")
    
    If thickness <= 0 Then
        UtilsLib.LogError("Lehtmetall: Could not measure thickness. Check that the selected face has an opposite face.")
        Exit Sub
    End If
    
    ' Convert to sheet metal
    UtilsLib.LogInfo("Lehtmetall: Converting to sheet metal...")
    partDoc.SubType = SHEET_METAL_GUID
    partDoc.Update()
    
    ' Get sheet metal component definition
    Dim smCompDef As SheetMetalComponentDefinition = partDoc.ComponentDefinition
    
    ' Set active sheet metal style to Default_mm
    SetSheetMetalStyle(smCompDef, "Default_mm")
    UtilsLib.LogInfo("Lehtmetall: Set style to Default_mm")
    
    ' Set measured thickness
    SetMeasuredThickness(smCompDef, thickness)
    
    ' Export Thickness parameter as iProperty (in mm)
    ExportThicknessAsProperty(smCompDef)
    UtilsLib.LogInfo("Lehtmetall: Exported Thickness as iProperty")
    
    ' Set Width and Length custom properties as numeric values
    SetSheetMetalProperties(partDoc)
    UtilsLib.LogInfo("Lehtmetall: Set Width and Length properties")
    
    ' Create flat pattern using the already selected A-side face
    CreateFlatPattern(smCompDef, aSideFace)
    
    partDoc.Update()
    
    UtilsLib.LogInfo("Lehtmetall: Conversion complete! Thickness: " & FormatNumber(thickness * 10, 2) & " mm")
End Sub

Function PickASideFace(app As Inventor.Application) As Face
    Dim aSideFace As Face = Nothing
    
    Try
        aSideFace = app.CommandManager.Pick(SelectionFilterEnum.kPartFacePlanarFilter, _
            "Vali A-külje pind (ülemine pind) - ESC tühistamiseks")
    Catch
        UtilsLib.LogWarn("Lehtmetall: Face selection cancelled.")
        Return Nothing
    End Try
    
    If aSideFace Is Nothing Then
        UtilsLib.LogError("Lehtmetall: No face selected.")
        Return Nothing
    End If
    
    Return aSideFace
End Function

Function MeasureThicknessAlongNormal(app As Inventor.Application, aSideFace As Face) As Double
    ' Get face normal at center point
    Dim evaluator As SurfaceEvaluator = aSideFace.Evaluator
    Dim paramRange As Box2d = evaluator.ParamRangeRect
    Dim centerU As Double = (paramRange.MinPoint.X + paramRange.MaxPoint.X) / 2
    Dim centerV As Double = (paramRange.MinPoint.Y + paramRange.MaxPoint.Y) / 2
    
    ' GetNormal takes params array and returns normals array
    Dim paramsArr() As Double = {centerU, centerV}
    Dim normalArr() As Double = {}
    Call evaluator.GetNormal(paramsArr, normalArr)
    
    ' Find opposite face - search all faces for parallel face with max distance
    Dim body As SurfaceBody = aSideFace.Parent
    Dim maxDistance As Double = 0
    
    For Each face As Face In body.Faces
        If face Is aSideFace Then Continue For
        If face.SurfaceType <> SurfaceTypeEnum.kPlaneSurface Then Continue For
        
        ' Check if parallel (normals are opposite)
        Dim otherEval As SurfaceEvaluator = face.Evaluator
        Dim otherRange As Box2d = otherEval.ParamRangeRect
        Dim otherCenterU As Double = (otherRange.MinPoint.X + otherRange.MaxPoint.X) / 2
        Dim otherCenterV As Double = (otherRange.MinPoint.Y + otherRange.MaxPoint.Y) / 2
        
        Dim otherParamsArr() As Double = {otherCenterU, otherCenterV}
        Dim otherNormalArr() As Double = {}
        Call otherEval.GetNormal(otherParamsArr, otherNormalArr)
        
        ' Check if anti-parallel (dot product ~ -1)
        Dim dot As Double = normalArr(0) * otherNormalArr(0) + _
                            normalArr(1) * otherNormalArr(1) + _
                            normalArr(2) * otherNormalArr(2)
        
        If Math.Abs(dot + 1) < 0.01 Then
            ' Measure distance between faces
            Dim dist As Double = app.MeasureTools.GetMinimumDistance(aSideFace, face)
            If dist > maxDistance Then
                maxDistance = dist
            End If
        End If
    Next
    
    Return maxDistance
End Function

Sub SetSheetMetalStyle(smCompDef As SheetMetalComponentDefinition, styleName As String)
    Try
        Dim mmStyle As SheetMetalStyle = smCompDef.SheetMetalStyles.Item(styleName)
        mmStyle.Activate()
    Catch
        ' Style not found, continue with default
    End Try
End Sub

Sub SetMeasuredThickness(smCompDef As SheetMetalComponentDefinition, thickness As Double)
    Try
        smCompDef.UseSheetMetalStyleThickness = False
        smCompDef.Thickness.Value = thickness
    Catch ex As Exception
        UtilsLib.LogWarn("Lehtmetall: Could not set thickness. " & ex.Message)
    End Try
End Sub

Sub ExportThicknessAsProperty(smCompDef As SheetMetalComponentDefinition)
    Try
        CustomPropertiesLib.EnsureSheetMetalThicknessExport(smCompDef)
    Catch ex As Exception
        UtilsLib.LogWarn("Lehtmetall: Could not export Thickness as iProperty. " & ex.Message)
    End Try
End Sub

Sub SetSheetMetalProperties(partDoc As PartDocument)
    Try
        CustomPropertiesLib.ValidateAndFixDimensionProperties(partDoc)
    Catch ex As Exception
        UtilsLib.LogWarn("Lehtmetall: Could not set Width/Length properties. " & ex.Message)
    End Try
End Sub

Sub SetOrAddProperty(propSet As PropertySet, propName As String, propValue As String)
    Try
        propSet.Item(propName).Value = propValue
    Catch
        Try
            propSet.Add(propValue, propName)
        Catch
        End Try
    End Try
End Sub

Sub CreateFlatPattern(smCompDef As SheetMetalComponentDefinition, aSideFace As Face)
    Try
        smCompDef.ASideFace = aSideFace
        smCompDef.Unfold()
        If smCompDef.HasFlatPattern Then
            smCompDef.FlatPattern.ExitEdit()
        End If
        UtilsLib.LogInfo("Lehtmetall: Flat pattern created")
    Catch ex As Exception
        UtilsLib.LogError("Lehtmetall: Could not create flat pattern: " & ex.Message)
    End Try
End Sub

Sub ValidateAndRepairExistingSheetMetal(app As Inventor.Application, partDoc As PartDocument)
    Dim smCompDef As SheetMetalComponentDefinition = CType(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
    Dim fixes As Integer = 0

    SetSheetMetalStyle(smCompDef, "Default_mm")

    Dim aSideFace As Face = Nothing
    Try
        aSideFace = smCompDef.ASideFace
    Catch
    End Try

    If aSideFace IsNot Nothing Then
        Dim measuredThickness As Double = MeasureThicknessAlongNormal(app, aSideFace)
        If measuredThickness > 0 Then
            Dim currentThickness As Double = smCompDef.Thickness.Value
            If Math.Abs(currentThickness - measuredThickness) > 0.001 Then
                SetMeasuredThickness(smCompDef, measuredThickness)
                fixes += 1
                UtilsLib.LogInfo("Lehtmetall: Fixed sheet metal thickness to " & FormatNumber(measuredThickness * 10, 3) & " mm")
            End If
        Else
            UtilsLib.LogWarn("Lehtmetall: Could not measure thickness from the current A-side face.")
        End If
    Else
        UtilsLib.LogWarn("Lehtmetall: A-side face is not set. Thickness verification skipped.")
    End If

    ExportThicknessAsProperty(smCompDef)
    If CustomPropertiesLib.ValidateAndFixDimensionProperties(partDoc) Then
        fixes += 1
        UtilsLib.LogInfo("Lehtmetall: Repaired dimension custom properties.")
    End If

    If Not smCompDef.HasFlatPattern Then
        If aSideFace Is Nothing Then
            UtilsLib.LogInfo("Lehtmetall: Flat pattern missing and A-side not set. Please select A-side face.")
            aSideFace = PickASideFace(app)
        End If

        If aSideFace IsNot Nothing Then
            CreateFlatPattern(smCompDef, aSideFace)
            SetSheetMetalProperties(partDoc)
            fixes += 1
        Else
            UtilsLib.LogWarn("Lehtmetall: Flat pattern was not created because A-side face was not selected.")
        End If
    End If

    partDoc.Update()
    If fixes = 0 Then
        UtilsLib.LogInfo("Lehtmetall: Validation complete. No fixes were needed.")
    Else
        UtilsLib.LogInfo("Lehtmetall: Validation complete. Applied " & fixes & " fix(es).")
    End If
End Sub

