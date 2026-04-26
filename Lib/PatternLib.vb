' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' PatternLib - Generic Pattern Operations
' 
' Functions for creating, managing, and querying occurrence patterns.
' Supports rectangular patterns with parametric count and spacing.
'
' Depends on: WorkFeatureLib.vb
'
' Usage: 
'   AddVbFile "Lib/UtilsLib.vb"
'   AddVbFile "Lib/GeoLib.vb"
'   AddVbFile "Lib/WorkFeatureLib.vb"
'   AddVbFile "Lib/PatternLib.vb"
'
' ============================================================================

Option Strict Off
Imports Inventor

Public Module PatternLib

    ' ============================================================================
    ' SECTION 1: Find Patterns
    ' ============================================================================
    
    ''' <summary>
    ''' Find a rectangular occurrence pattern by name.
    ''' </summary>
    Public Function FindPatternByName(asmDef As AssemblyComponentDefinition, _
                                       name As String) As RectangularOccurrencePattern
        If asmDef Is Nothing OrElse String.IsNullOrEmpty(name) Then Return Nothing
        
        Try
            For Each pattern As OccurrencePattern In asmDef.OccurrencePatterns
                If pattern.Name = name Then
                    If TypeOf pattern Is RectangularOccurrencePattern Then
                        Return CType(pattern, RectangularOccurrencePattern)
                    End If
                End If
            Next
        Catch
        End Try
        Return Nothing
    End Function
    
    ''' <summary>
    ''' Check if a pattern with the given name exists.
    ''' </summary>
    Public Function PatternExists(asmDef As AssemblyComponentDefinition, _
                                   name As String) As Boolean
        Return FindPatternByName(asmDef, name) IsNot Nothing
    End Function

    ' ============================================================================
    ' SECTION 2: Delete Patterns
    ' ============================================================================
    
    ''' <summary>
    ''' Delete an occurrence pattern by name if it exists.
    ''' Returns True if deleted, False if not found or failed.
    ''' </summary>
    Public Function DeletePatternByName(asmDef As AssemblyComponentDefinition, _
                                         name As String) As Boolean
        If asmDef Is Nothing OrElse String.IsNullOrEmpty(name) Then Return False
        
        Try
            For Each pattern As OccurrencePattern In asmDef.OccurrencePatterns
                If pattern.Name = name Then
                    pattern.Delete()
                    Return True
                End If
            Next
        Catch
        End Try
        Return False
    End Function

    ' ============================================================================
    ' SECTION 3: Create Rectangular Patterns
    ' ============================================================================
    
    ''' <summary>
    ''' Create a rectangular occurrence pattern using parameter names for count and spacing.
    ''' The count parameter must be unitless.
    ''' 
    ''' seedOccs - ObjectCollection containing the seed occurrence(s)
    ''' directionAxis - WorkAxis defining pattern direction
    ''' countParamName - Name of parameter for total column count (unitless)
    ''' spacingParamName - Name of parameter for spacing between columns
    ''' patternName - Name for the pattern
    ''' </summary>
    Public Function CreateRectangularPattern(app As Inventor.Application, _
                                              asmDoc As AssemblyDocument, _
                                              seedOccs As ObjectCollection, _
                                              directionAxis As WorkAxis, _
                                              countParamName As String, _
                                              spacingParamName As String, _
                                              patternName As String) As RectangularOccurrencePattern
        If seedOccs Is Nothing OrElse seedOccs.Count = 0 Then Return Nothing
        If directionAxis Is Nothing Then Return Nothing
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' Delete existing pattern if it exists
        DeletePatternByName(asmDef, patternName)
        
        Try
            ' Create the rectangular pattern
            ' Parameters:
            '   ParentComponents - ObjectCollection of seed occurrences
            '   ColumnAxis - WorkAxis for direction
            '   ColumnDirection - True for natural axis direction
            '   ColumnOffset - Spacing expression (parameter name or value)
            '   ColumnCount - Count expression (parameter name, must be unitless)
            Dim pattern As RectangularOccurrencePattern = _
                asmDef.OccurrencePatterns.AddRectangularPattern( _
                    seedOccs, _
                    directionAxis, _
                    True, _
                    spacingParamName, _
                    countParamName)
            
            pattern.Name = patternName
            Return pattern
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Create a rectangular pattern from a single seed occurrence.
    ''' </summary>
    Public Function CreateRectangularPatternFromOccurrence(app As Inventor.Application, _
                                                            asmDoc As AssemblyDocument, _
                                                            seedOcc As ComponentOccurrence, _
                                                            directionAxis As WorkAxis, _
                                                            countParamName As String, _
                                                            spacingParamName As String, _
                                                            patternName As String) As RectangularOccurrencePattern
        If seedOcc Is Nothing Then Return Nothing
        
        ' Create object collection with seed
        Dim occs As ObjectCollection = app.TransientObjects.CreateObjectCollection()
        occs.Add(seedOcc)
        
        Return CreateRectangularPattern(app, asmDoc, occs, directionAxis, _
                                        countParamName, spacingParamName, patternName)
    End Function

    ' ============================================================================
    ' SECTION 4: Pattern Occurrences
    ' ============================================================================
    
    ''' <summary>
    ''' Get all occurrences in a pattern (including seed and pattern elements).
    ''' </summary>
    Public Function GetPatternOccurrences(pattern As RectangularOccurrencePattern) As System.Collections.Generic.List(Of ComponentOccurrence)
        Dim result As New System.Collections.Generic.List(Of ComponentOccurrence)
        If pattern Is Nothing Then Return result
        
        Try
            ' Get occurrences from pattern elements
            For i As Integer = 1 To pattern.OccurrencePatternElements.Count
                Dim elem As OccurrencePatternElement = pattern.OccurrencePatternElements.Item(i)
                For Each occ As ComponentOccurrence In elem.Occurrences
                    result.Add(occ)
                Next
            Next
        Catch
        End Try
        
        Return result
    End Function
    
    ''' <summary>
    ''' Get the count of pattern elements (not including the seed).
    ''' </summary>
    Public Function GetPatternElementCount(pattern As RectangularOccurrencePattern) As Integer
        If pattern Is Nothing Then Return 0
        
        Try
            Return pattern.OccurrencePatternElements.Count
        Catch
            Return 0
        End Try
    End Function

    ' ============================================================================
    ' SECTION 5: Suppress/Unsuppress Patterns
    ' ============================================================================
    
    ''' <summary>
    ''' Suppress a pattern (hide all pattern instances).
    ''' </summary>
    Public Function SuppressPattern(pattern As RectangularOccurrencePattern) As Boolean
        If pattern Is Nothing Then Return False
        
        Try
            If Not pattern.Suppressed Then
                pattern.Suppress()
            End If
            Return True
        Catch
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Unsuppress a pattern (show all pattern instances).
    ''' </summary>
    Public Function UnsuppressPattern(pattern As RectangularOccurrencePattern) As Boolean
        If pattern Is Nothing Then Return False
        
        Try
            If pattern.Suppressed Then
                pattern.Unsuppress()
            End If
            Return True
        Catch
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Suppress a pattern by name.
    ''' </summary>
    Public Function SuppressPatternByName(asmDef As AssemblyComponentDefinition, _
                                           name As String) As Boolean
        Dim pattern As RectangularOccurrencePattern = FindPatternByName(asmDef, name)
        Return SuppressPattern(pattern)
    End Function

    ' ============================================================================
    ' SECTION 6: Copy and Position Seed
    ' ============================================================================
    
    ''' <summary>
    ''' Copy an occurrence and suppress the original.
    ''' The copy is placed at the same position as the original.
    ''' Returns the new copy which will be used as the pattern seed.
    ''' </summary>
    Public Function CopyAndSuppressSeed(app As Inventor.Application, _
                                         asmDoc As AssemblyDocument, _
                                         seedOcc As ComponentOccurrence, _
                                         copyBaseName As String) As ComponentOccurrence
        If seedOcc Is Nothing Then Return Nothing
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' Get the document being referenced
        Dim partDoc As Document = seedOcc.Definition.Document
        
        ' Place a new instance at the SAME position as the original
        Dim matrix As Matrix = seedOcc.Transformation
        
        Try
            Dim newOcc As ComponentOccurrence = asmDef.Occurrences.Add(partDoc.FullDocumentName, matrix)
            
            ' Unground so it can be constrained later
            If newOcc.Grounded Then
                newOcc.Grounded = False
            End If
            
            ' Rename the new occurrence
            Try
                newOcc.Name = copyBaseName & ":1"
            Catch
            End Try
            
            ' Suppress the original
            If Not seedOcc.Suppressed Then
                seedOcc.Suppress()
            End If
            
            Return newOcc
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Copy an occurrence and hide the original (set as Reference in BOM).
    ''' The original is NOT suppressed, so it can still be moved/used for constraints.
    ''' The copy is placed at the same position as the original.
    ''' Returns the new copy which will be used as the pattern seed.
    ''' </summary>
    Public Function CopyAndHideSeed(app As Inventor.Application, _
                                     asmDoc As AssemblyDocument, _
                                     seedOcc As ComponentOccurrence, _
                                     copyBaseName As String) As ComponentOccurrence
        If seedOcc Is Nothing Then Return Nothing
        
        Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
        
        ' Get the document being referenced
        Dim partDoc As Document = seedOcc.Definition.Document
        
        ' Place a new instance at the SAME position as the original
        Dim matrix As Matrix = seedOcc.Transformation
        
        Try
            Dim newOcc As ComponentOccurrence = asmDef.Occurrences.Add(partDoc.FullDocumentName, matrix)
            
            ' Unground so it can be constrained later
            If newOcc.Grounded Then
                newOcc.Grounded = False
            End If
            
            ' Rename the new occurrence
            Try
                newOcc.Name = copyBaseName & ":1"
            Catch
            End Try
            
            ' Don't suppress - instead set as Reference (excludes from BOM) and hide
            Try
                seedOcc.BOMStructure = BOMStructureEnum.kReferenceBOMStructure
            Catch
            End Try
            
            Try
                seedOcc.Visible = False
            Catch
            End Try
            
            Return newOcc
        Catch
            Return Nothing
        End Try
    End Function
    
    ''' <summary>
    ''' Restore a hidden seed occurrence to normal state.
    ''' Sets BOM structure back to Default and makes visible.
    ''' </summary>
    Public Sub RestoreHiddenSeed(seedOcc As ComponentOccurrence)
        If seedOcc Is Nothing Then Exit Sub
        
        Try
            seedOcc.BOMStructure = BOMStructureEnum.kDefaultBOMStructure
        Catch
        End Try
        
        Try
            seedOcc.Visible = True
        Catch
        End Try
        
        Try
            If seedOcc.Suppressed Then
                seedOcc.Unsuppress()
            End If
        Catch
        End Try
    End Sub

End Module
