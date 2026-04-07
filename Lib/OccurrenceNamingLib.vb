' ============================================================================
' OccurrenceNamingLib - Descriptive Occurrence Naming for Assemblies
' 
' Provides functions to rename assembly occurrences using a descriptive pattern:
' "<Description> (<Part Number>):<instance>"
'
' This makes occurrences easier to identify in the browser compared to the
' default Vault-assigned part numbers.
'
' Usage: 
'   AddVbFile "Lib/UtilsLib.vb"
'   AddVbFile "Lib/OccurrenceNamingLib.vb"
'   UtilsLib.SetLogger(Logger) ' In Sub Main
'
' Dependencies: UtilsLib (for logging)
' ============================================================================

Imports Inventor

Public Module OccurrenceNamingLib

    ' ============================================================================
    ' SECTION 1: Property Accessors
    ' ============================================================================

    ''' <summary>
    ''' Get the Description property from a document.
    ''' Works for both PartDocument and AssemblyDocument.
    ''' </summary>
    Public Function GetDocumentDescription(doc As Document) As String
        Try
            Dim designProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
            Dim value As Object = designProps.Item("Description").Value
            If value IsNot Nothing Then
                Return CStr(value).Trim()
            End If
        Catch
        End Try
        Return ""
    End Function

    ''' <summary>
    ''' Get the Part Number property from a document.
    ''' Works for both PartDocument and AssemblyDocument.
    ''' </summary>
    Public Function GetDocumentPartNumber(doc As Document) As String
        Try
            Dim designProps As PropertySet = doc.PropertySets.Item("Design Tracking Properties")
            Dim value As Object = designProps.Item("Part Number").Value
            If value IsNot Nothing Then
                Return CStr(value).Trim()
            End If
        Catch
        End Try
        Return ""
    End Function

    ' ============================================================================
    ' SECTION 2: Name Building
    ' ============================================================================

    ''' <summary>
    ''' Build the occurrence base name from description and part number.
    ''' Returns "Description (PartNumber)" if both present.
    ''' Returns just description or part number if one is empty.
    ''' Returns "" if both are empty.
    ''' </summary>
    Public Function BuildOccurrenceBaseName(description As String, partNumber As String) As String
        Dim desc As String = If(description IsNot Nothing, description.Trim(), "")
        Dim pn As String = If(partNumber IsNot Nothing, partNumber.Trim(), "")
        
        If desc <> "" AndAlso pn <> "" Then
            Return desc & " (" & pn & ")"
        ElseIf desc <> "" Then
            Return desc
        ElseIf pn <> "" Then
            Return pn
        Else
            Return ""
        End If
    End Function

    ''' <summary>
    ''' Get the instance number from an occurrence name.
    ''' E.g., "PartName:1" returns "1", "Description (123456):3" returns "3".
    ''' Returns "1" if no colon found.
    ''' </summary>
    Public Function GetInstanceNumber(occName As String) As String
        If String.IsNullOrEmpty(occName) Then Return "1"
        
        Dim colonPos As Integer = occName.LastIndexOf(":")
        If colonPos >= 0 AndAlso colonPos < occName.Length - 1 Then
            Return occName.Substring(colonPos + 1)
        End If
        Return "1"
    End Function

    ''' <summary>
    ''' Get the base name from an occurrence name (everything before the last colon).
    ''' E.g., "PartName:1" returns "PartName".
    ''' </summary>
    Public Function GetOccurrenceBaseName(occName As String) As String
        If String.IsNullOrEmpty(occName) Then Return ""
        
        Dim colonPos As Integer = occName.LastIndexOf(":")
        If colonPos > 0 Then
            Return occName.Substring(0, colonPos)
        End If
        Return occName
    End Function

    ' ============================================================================
    ' SECTION 3: Occurrence Renaming
    ' ============================================================================

    ''' <summary>
    ''' Rename a single occurrence using the descriptive naming pattern.
    ''' Reads Description and Part Number from the occurrence's referenced document.
    ''' Returns True if renamed, False if skipped or failed.
    ''' </summary>
    Public Function RenameOccurrence(occ As ComponentOccurrence) As Boolean
        Try
            ' Get the referenced document
            Dim refDoc As Document = Nothing
            Try
                refDoc = CType(occ.Definition.Document, Document)
            Catch
                ' Virtual component or inaccessible document
                Return False
            End Try
            
            If refDoc Is Nothing Then Return False
            
            ' Get properties
            Dim description As String = GetDocumentDescription(refDoc)
            Dim partNumber As String = GetDocumentPartNumber(refDoc)
            
            ' Build new base name
            Dim newBaseName As String = BuildOccurrenceBaseName(description, partNumber)
            If newBaseName = "" Then
                ' Both properties empty, skip
                Return False
            End If
            
            ' Preserve instance number
            Dim instanceNum As String = GetInstanceNumber(occ.Name)
            Dim newName As String = newBaseName & ":" & instanceNum
            
            ' Only rename if different
            If occ.Name <> newName Then
                Dim oldName As String = occ.Name
                occ.Name = newName
                UtilsLib.LogInfo("OccurrenceNamingLib: Renamed '" & oldName & "' -> '" & newName & "'")
                Return True
            End If
            
            Return False
        Catch ex As Exception
            UtilsLib.LogWarn("OccurrenceNamingLib: Error renaming occurrence: " & ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Rename all top-level occurrences in an assembly.
    ''' Returns the count of occurrences that were renamed.
    ''' </summary>
    Public Function RenameAllOccurrences(asmDoc As AssemblyDocument) As Integer
        Dim renamedCount As Integer = 0
        
        Try
            For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                If RenameOccurrence(occ) Then
                    renamedCount += 1
                End If
            Next
        Catch ex As Exception
            UtilsLib.LogWarn("OccurrenceNamingLib: Error iterating occurrences: " & ex.Message)
        End Try
        
        Return renamedCount
    End Function

    ''' <summary>
    ''' Rename only the selected occurrences in an assembly.
    ''' Returns the count of occurrences that were renamed.
    ''' </summary>
    Public Function RenameSelectedOccurrences(asmDoc As AssemblyDocument) As Integer
        Dim renamedCount As Integer = 0
        
        Try
            For Each obj As Object In asmDoc.SelectSet
                If TypeOf obj Is ComponentOccurrence Then
                    Dim occ As ComponentOccurrence = CType(obj, ComponentOccurrence)
                    If RenameOccurrence(occ) Then
                        renamedCount += 1
                    End If
                End If
            Next
        Catch ex As Exception
            UtilsLib.LogWarn("OccurrenceNamingLib: Error iterating selection: " & ex.Message)
        End Try
        
        Return renamedCount
    End Function

    ''' <summary>
    ''' Get the count of ComponentOccurrence objects in the SelectSet.
    ''' </summary>
    Public Function GetSelectedOccurrenceCount(asmDoc As AssemblyDocument) As Integer
        Dim count As Integer = 0
        Try
            For Each obj As Object In asmDoc.SelectSet
                If TypeOf obj Is ComponentOccurrence Then
                    count += 1
                End If
            Next
        Catch
        End Try
        Return count
    End Function

End Module
