' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' MaterialRoutingLib
'
' Central material routing rules for part classification/output defaults.
' Keep this map explicit and editable so users can tune it to company needs.
' ============================================================================

Imports System.Collections.Generic

Public Module MaterialRoutingLib

    Public Const ROUTE_AUTO As String = "AUTO"
    Public Const ROUTE_FRAME As String = "KARKASS"
    Public Const ROUTE_PADDING As String = "POROLOON"
    Public Const ROUTE_CUSTOM As String = "CUSTOM"

    ' ------------------------------------------------------------------------
    ' POROLOON_PATTERNS
    ' Edit these regex patterns to adjust which materials are treated as padding.
    ' ------------------------------------------------------------------------
    Public Function PaddingMaterialPatterns() As List(Of String)
        Return New List(Of String) From {
            "HR\d{5}$",        ' e.g. HR35
            "HS\d{5}$",        ' e.g. HSXX
            "RP\d{5}$",        ' e.g. RPXX
            "RG\d{5}$",        ' e.g. RGXX
            "RX\d{5}$",        ' e.g. RXXX
            "LIMI.*",
            "Dryfeel.*",   ' e.g. Dryfeel Soft
            "^ST\d{4}$"    ' e.g. ST1234
        }
    End Function

    Public Function IsPaddingMaterial(materialName As String) As Boolean
        Return UtilsLib.MaterialMatchesPatterns(materialName, PaddingMaterialPatterns())
    End Function

    Public Function GetPartOutputKind(materialName As String) As String
        If IsPaddingMaterial(materialName) Then
            Return ROUTE_PADDING
        End If
        Return ROUTE_FRAME
    End Function

    Public Function GetDefaultDetailFolder(elementRoot As String, kind As String) As String
        If String.IsNullOrEmpty(elementRoot) Then Return ""

        If String.Equals(kind, ROUTE_PADDING, StringComparison.OrdinalIgnoreCase) Then
            Return System.IO.Path.Combine(elementRoot, BaseElementLayoutLib.SEG_PADDING, BaseElementLayoutLib.SEG_PARTS)
        End If

        Return System.IO.Path.Combine(elementRoot, BaseElementLayoutLib.SEG_FRAME, BaseElementLayoutLib.SEG_PARTS)
    End Function

End Module
