' ============================================================================
' DimensionUpdateLib - Register Self-Contained Dimension Update Handler
'
' This library registers a dimension update handler in the "Uuenda" rule via
' DocumentUpdateLib. The generated update code is SELF-CONTAINED and does not
' depend on any external library files - it can run on any computer.
'
' Usage: AddVbFile "Lib/DimensionUpdateLib.vb"
'        AddVbFile "Lib/DocumentUpdateLib.vb"
'
' Example:
'   DimensionUpdateLib.RegisterDimensionHandler(partDoc, iLogicVb.Automation, "Z", "X", "Y")
'   ' For sheet metal (axes are empty, uses flat pattern formulas):
'   DimensionUpdateLib.RegisterDimensionHandler(partDoc, iLogicVb.Automation, "", "", "")
' ============================================================================

Imports Inventor

Public Module DimensionUpdateLib

    ' ============================================================================
    ' CONSTANTS
    ' ============================================================================
    
    Public Const HANDLER_UID As String = "Dimensions"
    Private Const LEGACY_RULE_NAME As String = "Uuenda mõõdud"
    Private Const SHEET_METAL_GUID As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    
    ' ============================================================================
    ' LOGGING
    ' ============================================================================
    
    Private m_Logger As Object = Nothing
    
    Public Sub SetLogger(logger As Object)
        m_Logger = logger
    End Sub
    
    Private Sub LogInfo(message As String)
        If m_Logger IsNot Nothing Then
            Try
                m_Logger.Info("DimensionUpdateLib: " & message)
            Catch
            End Try
        End If
    End Sub
    
    Private Sub LogError(message As String)
        If m_Logger IsNot Nothing Then
            Try
                m_Logger.Error("DimensionUpdateLib: " & message)
            Catch
            End Try
        End If
    End Sub
    
    ' ============================================================================
    ' PUBLIC API
    ' ============================================================================
    
    ''' <summary>
    ''' Registers a self-contained dimension update handler in the "Uuenda" rule.
    ''' Removes legacy "Uuenda mõõdud" rule if present.
    ''' </summary>
    ''' <param name="doc">Part document</param>
    ''' <param name="iLogicAuto">iLogicVb.Automation object</param>
    ''' <param name="thicknessAxis">Thickness axis (X/Y/Z or V:x,y,z format), empty for sheet metal</param>
    ''' <param name="widthAxis">Width axis (X/Y/Z or V:x,y,z format), empty for sheet metal</param>
    ''' <param name="lengthAxis">Length axis (X/Y/Z or V:x,y,z format), empty for sheet metal</param>
    ''' <returns>True if successful</returns>
    Public Function RegisterDimensionHandler(ByVal doc As Document, ByVal iLogicAuto As Object, _
                                             ByVal thicknessAxis As String, ByVal widthAxis As String, _
                                             ByVal lengthAxis As String) As Boolean
        LogInfo("RegisterDimensionHandler called - T:" & thicknessAxis & " W:" & widthAxis & " L:" & lengthAxis)
        
        If doc Is Nothing OrElse iLogicAuto Is Nothing Then
            LogError("RegisterDimensionHandler: doc or iLogicAuto is Nothing")
            Return False
        End If
        
        If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
            LogError("RegisterDimensionHandler: Not a part document")
            Return False
        End If
        
        Dim partDoc As PartDocument = CType(doc, PartDocument)
        LogInfo("Processing: " & partDoc.DisplayName)
        
        ' Remove legacy rule if present
        RemoveLegacyRule(partDoc, iLogicAuto)
        
        ' Store axis configuration in custom properties (only for normal parts)
        If Not String.IsNullOrEmpty(thicknessAxis) Then
            SetCustomProperty(partDoc, "BB_ThicknessAxis", thicknessAxis)
            SetCustomProperty(partDoc, "BB_WidthAxis", widthAxis)
            LogInfo("Stored axis config: T=" & thicknessAxis & " W=" & widthAxis)
        End If
        
        ' Build self-contained update code
        Dim codeLines() As String = BuildDimensionUpdateCode()
        LogInfo("Built dimension update code: " & codeLines.Length & " lines")
        
        ' Register with DocumentUpdateLib
        Dim triggers() As DocumentUpdateLib.UpdateTrigger = { _
            DocumentUpdateLib.UpdateTrigger.PartGeometryChange, _
            DocumentUpdateLib.UpdateTrigger.UserParameterChange, _
            DocumentUpdateLib.UpdateTrigger.ModelParameterChange, _
            DocumentUpdateLib.UpdateTrigger.BeforeVaultCheckIn _
        }
        
        Dim result As Boolean = DocumentUpdateLib.RegisterUpdateHandler(doc, iLogicAuto, HANDLER_UID, codeLines, triggers)
        LogInfo("RegisterUpdateHandler result: " & result.ToString())
        
        ' Run the update rule immediately to set initial values
        If result Then
            Try
                iLogicAuto.RunRule(doc, DocumentUpdateLib.RULE_NAME)
                LogInfo("Ran " & DocumentUpdateLib.RULE_NAME & " rule")
            Catch ex As Exception
                LogError("Failed to run " & DocumentUpdateLib.RULE_NAME & ": " & ex.Message)
            End Try
        End If
        
        Return result
    End Function
    
    ''' <summary>
    ''' Removes the dimension handler from the "Uuenda" rule.
    ''' </summary>
    Public Function RemoveDimensionHandler(ByVal doc As Document, ByVal iLogicAuto As Object) As Boolean
        If doc Is Nothing OrElse iLogicAuto Is Nothing Then
            Return False
        End If
        
        Return DocumentUpdateLib.RemoveUpdateHandler(doc, iLogicAuto, HANDLER_UID)
    End Function
    
    ' ============================================================================
    ' PRIVATE HELPERS
    ' ============================================================================
    
    ''' <summary>
    ''' Removes the legacy "Uuenda mõõdud" rule if present.
    ''' </summary>
    Private Sub RemoveLegacyRule(ByVal doc As Document, ByVal iLogicAuto As Object)
        Try
            Dim legacyRule As Object = iLogicAuto.GetRule(doc, LEGACY_RULE_NAME)
            If legacyRule IsNot Nothing Then
                iLogicAuto.DeleteRule(doc, LEGACY_RULE_NAME)
            End If
        Catch
            ' Rule doesn't exist, nothing to remove
        End Try
    End Sub
    
    ''' <summary>
    ''' Sets a custom property value.
    ''' </summary>
    Private Sub SetCustomProperty(ByVal doc As Document, ByVal propName As String, ByVal propValue As String)
        Try
            Dim propSet As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
            Try
                propSet.Item(propName).Value = propValue
            Catch
                propSet.Add(propValue, propName)
            End Try
        Catch
        End Try
    End Sub
    
    ''' <summary>
    ''' Builds self-contained VB code lines for dimension updates.
    ''' This code has NO external dependencies and works on any computer.
    ''' All logic is inline (no Sub/Function definitions) because DocumentUpdateLib
    ''' places code inside Sub Main().
    ''' Creates numeric custom properties (Thickness, Width, Length) with values in mm.
    ''' </summary>
    Private Function BuildDimensionUpdateCode() As String()
        Dim lines As New System.Collections.Generic.List(Of String)
        
        ' All code must be inline - no Sub/Function definitions allowed
        ' because DocumentUpdateLib places this inside Sub Main()
        
        lines.Add("Dim doc As Document = ThisDoc.Document")
        lines.Add("If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then")
        lines.Add("    Dim partDoc As PartDocument = CType(doc, PartDocument)")
        lines.Add("    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition")
        lines.Add("    ")
        lines.Add("    Const SHEET_METAL_GUID As String = ""{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}""")
        lines.Add("    ")
        lines.Add("    Dim isSheetMetal As Boolean = False")
        lines.Add("    Try")
        lines.Add("        isSheetMetal = (partDoc.SubType = SHEET_METAL_GUID)")
        lines.Add("    Catch")
        lines.Add("    End Try")
        lines.Add("    ")
        lines.Add("    Dim thicknessVal As Double = 0")
        lines.Add("    Dim widthVal As Double = 0")
        lines.Add("    Dim lengthVal As Double = 0")
        lines.Add("    ")
        lines.Add("    If isSheetMetal Then")
        lines.Add("        ' Sheet metal: get dimensions from flat pattern via API")
        lines.Add("        Try")
        lines.Add("            Dim smCompDef As SheetMetalComponentDefinition = CType(compDef, SheetMetalComponentDefinition)")
        lines.Add("            ")
        lines.Add("            ' Get thickness from sheet metal definition (in cm)")
        lines.Add("            thicknessVal = smCompDef.Thickness.Value")
        lines.Add("            ")
        lines.Add("            ' Get width/length from flat pattern RangeBox (in cm)")
        lines.Add("            If smCompDef.HasFlatPattern Then")
        lines.Add("                Dim fpBox As Box = smCompDef.FlatPattern.RangeBox")
        lines.Add("                Dim fpX As Double = Math.Abs(fpBox.MaxPoint.X - fpBox.MinPoint.X)")
        lines.Add("                Dim fpY As Double = Math.Abs(fpBox.MaxPoint.Y - fpBox.MinPoint.Y)")
        lines.Add("                ' Width is smaller, Length is larger")
        lines.Add("                If fpX <= fpY Then")
        lines.Add("                    widthVal = fpX")
        lines.Add("                    lengthVal = fpY")
        lines.Add("                Else")
        lines.Add("                    widthVal = fpY")
        lines.Add("                    lengthVal = fpX")
        lines.Add("                End If")
        lines.Add("            End If")
        lines.Add("        Catch")
        lines.Add("        End Try")
        lines.Add("    Else")
        lines.Add("        ' Normal part: calculate from bounding box")
        lines.Add("        Dim thicknessAxis As String = ""Z""")
        lines.Add("        Dim widthAxis As String = ""X""")
        lines.Add("        Try")
        lines.Add("            thicknessAxis = CStr(partDoc.PropertySets.Item(""Inventor User Defined Properties"").Item(""BB_ThicknessAxis"").Value)")
        lines.Add("        Catch")
        lines.Add("        End Try")
        lines.Add("        Try")
        lines.Add("            widthAxis = CStr(partDoc.PropertySets.Item(""Inventor User Defined Properties"").Item(""BB_WidthAxis"").Value)")
        lines.Add("        Catch")
        lines.Add("        End Try")
        lines.Add("        ")
        lines.Add("        If thicknessAxis.StartsWith(""V:"") Then")
        lines.Add("            ' Oriented bounding box calculation")
        lines.Add("            Dim tx As Double = 0, ty As Double = 0, tz As Double = 0")
        lines.Add("            Dim wx As Double = 0, wy As Double = 0, wz As Double = 0")
        lines.Add("            Dim lx As Double = 0, ly As Double = 0, lz As Double = 0")
        lines.Add("            ")
        lines.Add("            ' Parse thickness vector")
        lines.Add("            Try")
        lines.Add("                Dim tParts() As String = thicknessAxis.Substring(2).Split("",""c)")
        lines.Add("                If tParts.Length = 3 Then")
        lines.Add("                    tx = Double.Parse(tParts(0), System.Globalization.CultureInfo.InvariantCulture)")
        lines.Add("                    ty = Double.Parse(tParts(1), System.Globalization.CultureInfo.InvariantCulture)")
        lines.Add("                    tz = Double.Parse(tParts(2), System.Globalization.CultureInfo.InvariantCulture)")
        lines.Add("                End If")
        lines.Add("            Catch")
        lines.Add("            End Try")
        lines.Add("            ")
        lines.Add("            If widthAxis.StartsWith(""V:"") Then")
        lines.Add("                ' Parse width vector")
        lines.Add("                Try")
        lines.Add("                    Dim wParts() As String = widthAxis.Substring(2).Split("",""c)")
        lines.Add("                    If wParts.Length = 3 Then")
        lines.Add("                        wx = Double.Parse(wParts(0), System.Globalization.CultureInfo.InvariantCulture)")
        lines.Add("                        wy = Double.Parse(wParts(1), System.Globalization.CultureInfo.InvariantCulture)")
        lines.Add("                        wz = Double.Parse(wParts(2), System.Globalization.CultureInfo.InvariantCulture)")
        lines.Add("                    End If")
        lines.Add("                Catch")
        lines.Add("                End Try")
        lines.Add("                ' Length = cross(thickness, width)")
        lines.Add("                lx = ty * wz - tz * wy")
        lines.Add("                ly = tz * wx - tx * wz")
        lines.Add("                lz = tx * wy - ty * wx")
        lines.Add("            Else")
        lines.Add("                ' Compute perpendicular vectors")
        lines.Add("                Dim refX As Double = 0, refY As Double = 0, refZ As Double = 0")
        lines.Add("                If Math.Abs(tx) <= Math.Abs(ty) AndAlso Math.Abs(tx) <= Math.Abs(tz) Then")
        lines.Add("                    refX = 1 : refY = 0 : refZ = 0")
        lines.Add("                ElseIf Math.Abs(ty) <= Math.Abs(tz) Then")
        lines.Add("                    refX = 0 : refY = 1 : refZ = 0")
        lines.Add("                Else")
        lines.Add("                    refX = 0 : refY = 0 : refZ = 1")
        lines.Add("                End If")
        lines.Add("                wx = ty * refZ - tz * refY")
        lines.Add("                wy = tz * refX - tx * refZ")
        lines.Add("                wz = tx * refY - ty * refX")
        lines.Add("                Dim wLen As Double = Math.Sqrt(wx * wx + wy * wy + wz * wz)")
        lines.Add("                If wLen > 0.0001 Then")
        lines.Add("                    wx = wx / wLen : wy = wy / wLen : wz = wz / wLen")
        lines.Add("                End If")
        lines.Add("                lx = ty * wz - tz * wy")
        lines.Add("                ly = tz * wx - tx * wz")
        lines.Add("                lz = tx * wy - ty * wx")
        lines.Add("                Dim lLen As Double = Math.Sqrt(lx * lx + ly * ly + lz * lz)")
        lines.Add("                If lLen > 0.0001 Then")
        lines.Add("                    lx = lx / lLen : ly = ly / lLen : lz = lz / lLen")
        lines.Add("                End If")
        lines.Add("            End If")
        lines.Add("            ")
        lines.Add("            ' Calculate oriented extents")
        lines.Add("            Dim minT As Double = Double.MaxValue, maxT As Double = Double.MinValue")
        lines.Add("            Dim minW As Double = Double.MaxValue, maxW As Double = Double.MinValue")
        lines.Add("            Dim minL As Double = Double.MaxValue, maxL As Double = Double.MinValue")
        lines.Add("            Try")
        lines.Add("                For Each body As SurfaceBody In partDoc.ComponentDefinition.SurfaceBodies")
        lines.Add("                    For Each vertex As Vertex In body.Vertices")
        lines.Add("                        Dim pt As Point = vertex.Point")
        lines.Add("                        Dim projT As Double = pt.X * tx + pt.Y * ty + pt.Z * tz")
        lines.Add("                        Dim projW As Double = pt.X * wx + pt.Y * wy + pt.Z * wz")
        lines.Add("                        Dim projL As Double = pt.X * lx + pt.Y * ly + pt.Z * lz")
        lines.Add("                        If projT < minT Then minT = projT")
        lines.Add("                        If projT > maxT Then maxT = projT")
        lines.Add("                        If projW < minW Then minW = projW")
        lines.Add("                        If projW > maxW Then maxW = projW")
        lines.Add("                        If projL < minL Then minL = projL")
        lines.Add("                        If projL > maxL Then maxL = projL")
        lines.Add("                    Next")
        lines.Add("                Next")
        lines.Add("            Catch")
        lines.Add("            End Try")
        lines.Add("            If minT < Double.MaxValue Then thicknessVal = maxT - minT")
        lines.Add("            If minW < Double.MaxValue Then widthVal = maxW - minW")
        lines.Add("            If minL < Double.MaxValue Then lengthVal = maxL - minL")
        lines.Add("        Else")
        lines.Add("            ' Standard axis-aligned bounding box")
        lines.Add("            Dim rangebox As Box = partDoc.ComponentDefinition.RangeBox")
        lines.Add("            Dim xSize As Double = rangebox.MaxPoint.X - rangebox.MinPoint.X")
        lines.Add("            Dim ySize As Double = rangebox.MaxPoint.Y - rangebox.MinPoint.Y")
        lines.Add("            Dim zSize As Double = rangebox.MaxPoint.Z - rangebox.MinPoint.Z")
        lines.Add("            ")
        lines.Add("            ' Determine length axis")
        lines.Add("            Dim lengthAxis As String = ""X""")
        lines.Add("            If (thicknessAxis = ""X"" AndAlso widthAxis = ""Y"") OrElse (thicknessAxis = ""Y"" AndAlso widthAxis = ""X"") Then")
        lines.Add("                lengthAxis = ""Z""")
        lines.Add("            ElseIf (thicknessAxis = ""X"" AndAlso widthAxis = ""Z"") OrElse (thicknessAxis = ""Z"" AndAlso widthAxis = ""X"") Then")
        lines.Add("                lengthAxis = ""Y""")
        lines.Add("            End If")
        lines.Add("            ")
        lines.Add("            ' Get axis values")
        lines.Add("            If thicknessAxis = ""X"" Then thicknessVal = xSize")
        lines.Add("            If thicknessAxis = ""Y"" Then thicknessVal = ySize")
        lines.Add("            If thicknessAxis = ""Z"" Then thicknessVal = zSize")
        lines.Add("            If widthAxis = ""X"" Then widthVal = xSize")
        lines.Add("            If widthAxis = ""Y"" Then widthVal = ySize")
        lines.Add("            If widthAxis = ""Z"" Then widthVal = zSize")
        lines.Add("            If lengthAxis = ""X"" Then lengthVal = xSize")
        lines.Add("            If lengthAxis = ""Y"" Then lengthVal = ySize")
        lines.Add("            If lengthAxis = ""Z"" Then lengthVal = zSize")
        lines.Add("        End If")
        lines.Add("    End If")
        lines.Add("    ")
        lines.Add("    ' Check for manual overrides (applies to both sheet metal and normal parts)")
        lines.Add("    Try")
        lines.Add("        Dim ovT As Parameter = compDef.Parameters.Item(""ThicknessOverride"")")
        lines.Add("        If ovT IsNot Nothing Then thicknessVal = ovT.Value")
        lines.Add("    Catch")
        lines.Add("    End Try")
        lines.Add("    Try")
        lines.Add("        Dim ovW As Parameter = compDef.Parameters.Item(""WidthOverride"")")
        lines.Add("        If ovW IsNot Nothing Then widthVal = ovW.Value")
        lines.Add("    Catch")
        lines.Add("    End Try")
        lines.Add("    Try")
        lines.Add("        Dim ovL As Parameter = compDef.Parameters.Item(""LengthOverride"")")
        lines.Add("        If ovL IsNot Nothing Then lengthVal = ovL.Value")
        lines.Add("    Catch")
        lines.Add("    End Try")
        lines.Add("    ")
        lines.Add("    ' Delete any existing User Parameters (cleanup from old approach)")
        lines.Add("    Dim userParams As UserParameters = compDef.Parameters.UserParameters")
        lines.Add("    Try")
        lines.Add("        Dim oldParam As Parameter = userParams.Item(""Thickness"")")
        lines.Add("        If oldParam IsNot Nothing Then oldParam.Delete()")
        lines.Add("    Catch")
        lines.Add("    End Try")
        lines.Add("    Try")
        lines.Add("        Dim oldParam As Parameter = userParams.Item(""Width"")")
        lines.Add("        If oldParam IsNot Nothing Then oldParam.Delete()")
        lines.Add("    Catch")
        lines.Add("    End Try")
        lines.Add("    Try")
        lines.Add("        Dim oldParam As Parameter = userParams.Item(""Length"")")
        lines.Add("        If oldParam IsNot Nothing Then oldParam.Delete()")
        lines.Add("    Catch")
        lines.Add("    End Try")
        lines.Add("    ")
        lines.Add("    ' Delete any existing custom properties (to recreate fresh)")
        lines.Add("    Dim customProps As PropertySet = partDoc.PropertySets.Item(""Inventor User Defined Properties"")")
        lines.Add("    Try")
        lines.Add("        Dim oldT As [Property] = customProps.Item(""Thickness"")")
        lines.Add("        If oldT IsNot Nothing Then")
        lines.Add("            Logger.Info(""[DimensionUpdate] Deleting old Thickness property, value="" & oldT.Value.ToString())")
        lines.Add("            oldT.Delete()")
        lines.Add("        End If")
        lines.Add("    Catch")
        lines.Add("    End Try")
        lines.Add("    Try")
        lines.Add("        Dim oldW As [Property] = customProps.Item(""Width"")")
        lines.Add("        If oldW IsNot Nothing Then")
        lines.Add("            Logger.Info(""[DimensionUpdate] Deleting old Width property, value="" & oldW.Value.ToString())")
        lines.Add("            oldW.Delete()")
        lines.Add("        End If")
        lines.Add("    Catch")
        lines.Add("    End Try")
        lines.Add("    Try")
        lines.Add("        Dim oldL As [Property] = customProps.Item(""Length"")")
        lines.Add("        If oldL IsNot Nothing Then")
        lines.Add("            Logger.Info(""[DimensionUpdate] Deleting old Length property, value="" & oldL.Value.ToString())")
        lines.Add("            oldL.Delete()")
        lines.Add("        End If")
        lines.Add("    Catch")
        lines.Add("    End Try")
        lines.Add("    ")
        lines.Add("    ' Convert cm to mm and round to 1 decimal place")
        lines.Add("    Dim thicknessMm As Double = Math.Round(thicknessVal * 10.0, 1)")
        lines.Add("    Dim widthMm As Double = Math.Round(widthVal * 10.0, 1)")
        lines.Add("    Dim lengthMm As Double = Math.Round(lengthVal * 10.0, 1)")
        lines.Add("    ")
        lines.Add("    ' Debug logging")
        lines.Add("    Logger.Info(""[DimensionUpdate] isSheetMetal="" & isSheetMetal.ToString())")
        lines.Add("    Logger.Info(""[DimensionUpdate] Raw values (cm): T="" & thicknessVal.ToString() & "" W="" & widthVal.ToString() & "" L="" & lengthVal.ToString())")
        lines.Add("    Logger.Info(""[DimensionUpdate] Converted (mm): T="" & thicknessMm.ToString() & "" W="" & widthMm.ToString() & "" L="" & lengthMm.ToString())")
        lines.Add("    ")
        lines.Add("    ' Set Thickness property (numeric value in mm)")
        lines.Add("    Try")
        lines.Add("        customProps.Add(thicknessMm, ""Thickness"")")
        lines.Add("        Logger.Info(""[DimensionUpdate] Set Thickness: "" & thicknessMm.ToString() & "" mm"")")
        lines.Add("    Catch ex As Exception")
        lines.Add("        Logger.Error(""[DimensionUpdate] Failed to set Thickness: "" & ex.Message)")
        lines.Add("    End Try")
        lines.Add("    ")
        lines.Add("    ' Set Width property (numeric value in mm)")
        lines.Add("    Try")
        lines.Add("        customProps.Add(widthMm, ""Width"")")
        lines.Add("        Logger.Info(""[DimensionUpdate] Set Width: "" & widthMm.ToString() & "" mm"")")
        lines.Add("    Catch ex As Exception")
        lines.Add("        Logger.Error(""[DimensionUpdate] Failed to set Width: "" & ex.Message)")
        lines.Add("    End Try")
        lines.Add("    ")
        lines.Add("    ' Set Length property (numeric value in mm)")
        lines.Add("    Try")
        lines.Add("        customProps.Add(lengthMm, ""Length"")")
        lines.Add("        Logger.Info(""[DimensionUpdate] Set Length: "" & lengthMm.ToString() & "" mm"")")
        lines.Add("    Catch ex As Exception")
        lines.Add("        Logger.Error(""[DimensionUpdate] Failed to set Length: "" & ex.Message)")
        lines.Add("    End Try")
        lines.Add("End If")
        
        Return lines.ToArray()
    End Function

End Module
