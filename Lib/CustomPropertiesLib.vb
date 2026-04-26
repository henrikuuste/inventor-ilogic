' Copyright (c) 2026 Henri Kuuste
Imports Inventor
Imports System.Globalization

Public Module CustomPropertiesLib

    Public Const PROP_THICKNESS As String = "Thickness"
    Public Const PROP_WIDTH As String = "Width"
    Public Const PROP_LENGTH As String = "Length"
    Public Const SHEET_METAL_GUID As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"
    Public Const SHEET_METAL_WIDTH_FORMULA As String = "=<Sheet Metal Width>"
    Public Const SHEET_METAL_LENGTH_FORMULA As String = "=<Sheet Metal Length>"

    Public Function ToMm(valueCm As Double) As Double
        Return valueCm * 10.0
    End Function

    Public Function CheckPropertyValue(partDoc As PartDocument, propName As String, expectedValue As String) As Boolean
        Try
            Dim propSet As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
            Dim prop As [Property] = propSet.Item(propName)
            Return String.Equals(CStr(prop.Value), expectedValue, StringComparison.OrdinalIgnoreCase)
        Catch
            Return False
        End Try
    End Function

    Public Function FormatDimensionText(valueCm As Double) As String
        Dim roundedMm As Double = Math.Round(ToMm(valueCm), 0, MidpointRounding.AwayFromZero)
        Return roundedMm.ToString("0", CultureInfo.InvariantCulture)
    End Function

    Public Function IsSheetMetalPart(partDoc As PartDocument) As Boolean
        Try
            Return String.Equals(partDoc.SubType, SHEET_METAL_GUID, StringComparison.OrdinalIgnoreCase)
        Catch
            Return False
        End Try
    End Function

    Public Sub SetTextProperty(partDoc As PartDocument, propName As String, value As String)
        Try
            Dim propSet As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
            SetOrReplaceTextProperty(propSet, propName, value)
        Catch
        End Try
    End Sub

    Public Sub SetNormalPartDimensionProperties(partDoc As PartDocument, thicknessCm As Double, widthCm As Double, lengthCm As Double)
        SetTextProperty(partDoc, PROP_THICKNESS, FormatDimensionText(thicknessCm))
        SetTextProperty(partDoc, PROP_WIDTH, FormatDimensionText(widthCm))
        SetTextProperty(partDoc, PROP_LENGTH, FormatDimensionText(lengthCm))
    End Sub

    Public Sub SetSheetMetalDimensionProperties(partDoc As PartDocument)
        SetTextProperty(partDoc, PROP_WIDTH, SHEET_METAL_WIDTH_FORMULA)
        SetTextProperty(partDoc, PROP_LENGTH, SHEET_METAL_LENGTH_FORMULA)
    End Sub

    Public Sub EnsureSheetMetalThicknessExport(smCompDef As SheetMetalComponentDefinition)
        Try
            Dim thicknessParam As Parameter = smCompDef.Thickness
            thicknessParam.ExposedAsProperty = True
            thicknessParam.CustomPropertyFormat.PropertyType = CustomPropertyTypeEnum.kNumberPropertyType
            thicknessParam.CustomPropertyFormat.ShowUnitsString = False
            thicknessParam.CustomPropertyFormat.Units = "mm"
        Catch
        End Try
    End Sub

    ' Centralized validation/fix entry point used across rules.
    ' For normal parts, provide thickness/width/length values in cm when available.
    Public Function ValidateAndFixDimensionProperties(partDoc As PartDocument, _
                                                      Optional thicknessCm As Double = Double.NaN, _
                                                      Optional widthCm As Double = Double.NaN, _
                                                      Optional lengthCm As Double = Double.NaN) As Boolean
        If partDoc Is Nothing Then Return False

        Dim changed As Boolean = False

        If IsSheetMetalPart(partDoc) Then
            Try
                Dim smCompDef As SheetMetalComponentDefinition = CType(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
                EnsureSheetMetalThicknessExport(smCompDef)
            Catch
            End Try

            If Not CheckPropertyValue(partDoc, PROP_WIDTH, SHEET_METAL_WIDTH_FORMULA) Then
                SetTextProperty(partDoc, PROP_WIDTH, SHEET_METAL_WIDTH_FORMULA)
                changed = True
            End If
            If Not CheckPropertyValue(partDoc, PROP_LENGTH, SHEET_METAL_LENGTH_FORMULA) Then
                SetTextProperty(partDoc, PROP_LENGTH, SHEET_METAL_LENGTH_FORMULA)
                changed = True
            End If

            Return changed
        End If

        If Not Double.IsNaN(thicknessCm) AndAlso Not Double.IsNaN(widthCm) AndAlso Not Double.IsNaN(lengthCm) Then
            SetNormalPartDimensionProperties(partDoc, thicknessCm, widthCm, lengthCm)
            Return True
        End If

        If NormalizeExistingDimensionProperty(partDoc, PROP_THICKNESS) Then changed = True
        If NormalizeExistingDimensionProperty(partDoc, PROP_WIDTH) Then changed = True
        If NormalizeExistingDimensionProperty(partDoc, PROP_LENGTH) Then changed = True

        Return changed
    End Function

    Private Sub SetOrAddProperty(propSet As PropertySet, propName As String, propValue As Object)
        Try
            propSet.Item(propName).Value = propValue
        Catch
            Try
                propSet.Add(propValue, propName)
            Catch
            End Try
        End Try
    End Sub

    Private Sub SetOrReplaceTextProperty(propSet As PropertySet, propName As String, propValue As String)
        Try
            Dim existingProp As [Property] = propSet.Item(propName)
            If TypeOf existingProp.Value Is String Then
                existingProp.Value = propValue
                Return
            End If

            existingProp.Delete()
            propSet.Add(propValue, propName)
        Catch
            Try
                propSet.Add(propValue, propName)
            Catch
            End Try
        End Try
    End Sub

    Private Function NormalizeExistingDimensionProperty(partDoc As PartDocument, propName As String) As Boolean
        Try
            Dim propSet As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
            Dim prop As [Property] = propSet.Item(propName)
            Dim currentText As String = CStr(prop.Value)
            Dim normalized As String = NormalizeDimensionText(currentText)
            If String.IsNullOrEmpty(normalized) Then Return False

            If Not String.Equals(currentText, normalized, StringComparison.Ordinal) Then
                SetTextProperty(partDoc, propName, normalized)
                Return True
            End If
        Catch
        End Try
        Return False
    End Function

    Private Function NormalizeDimensionText(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""

        Dim normalized As String = value.Trim()
        normalized = normalized.Replace(" mm", "")
        normalized = normalized.Replace("MM", "")
        normalized = normalized.Replace("Mm", "")
        normalized = normalized.Replace("mM", "")
        normalized = normalized.Replace("mm", "")
        normalized = normalized.Trim()
        normalized = normalized.Replace(",", ".")

        Dim parsedMm As Double = 0
        If Double.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, parsedMm) Then
            Dim rounded As Double = Math.Round(parsedMm, 0, MidpointRounding.AwayFromZero)
            Return rounded.ToString("0", CultureInfo.InvariantCulture)
        End If

        Return ""
    End Function

End Module
