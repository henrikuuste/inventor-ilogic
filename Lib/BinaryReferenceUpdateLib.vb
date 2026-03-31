' ============================================================================
' BinaryReferenceUpdateLib - Update derived part references via binary editing
'
' This library provides functions to update file references in Inventor
' documents by directly editing the binary file content. This is necessary
' because the Inventor API does not provide a method to update derived part
' base references while preserving derive settings.
'
' IMPORTANT: Files must be CLOSED before calling these functions.
'
' Usage: AddVbFile "Lib/BinaryReferenceUpdateLib.vb"
' ============================================================================

Imports System.IO
Imports System.Text
Imports System.Collections.Generic

Public Module BinaryReferenceUpdateLib

    ''' <summary>
    ''' Update all file references in a document using binary replacement.
    ''' The document must be CLOSED before calling this function.
    ''' </summary>
    ''' <param name="filePath">Path to the file to update (must be closed)</param>
    ''' <param name="pathMap">Dictionary mapping old paths to new paths</param>
    ''' <param name="logMessages">Output log messages</param>
    ''' <returns>True if successful, False if errors occurred</returns>
    Public Function UpdateFileReferencesBinary(filePath As String, _
                                                pathMap As Dictionary(Of String, String), _
                                                ByRef logMessages As List(Of String)) As Boolean
        If logMessages Is Nothing Then logMessages = New List(Of String)
        
        ' Validate file exists
        If Not System.IO.File.Exists(filePath) Then
            logMessages.Add("  ERROR: File not found: " & filePath)
            Return False
        End If
        
        ' Read file bytes
        Dim fileBytes As Byte() = Nothing
        Try
            fileBytes = System.IO.File.ReadAllBytes(filePath)
        Catch ex As Exception
            logMessages.Add("  ERROR reading file: " & ex.Message)
            Return False
        End Try
        
        ' Create backup
        Dim backupPath As String = filePath & ".backup"
        Try
            System.IO.File.Copy(filePath, backupPath, True)
        Catch ex As Exception
            logMessages.Add("  WARNING: Could not create backup: " & ex.Message)
        End Try
        
        ' Perform replacements
        Dim totalReplacements As Integer = 0
        Dim modified As Boolean = False
        
        For Each kvp As KeyValuePair(Of String, String) In pathMap
            Dim oldPath As String = kvp.Key
            Dim newPath As String = kvp.Value
            
            ' Skip if same
            If oldPath.Equals(newPath, StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If
            
            ' Check if old path exists in file (Unicode search)
            Dim oldBytes As Byte() = Encoding.Unicode.GetBytes(oldPath)
            Dim pos As Integer = IndexOfBytes(fileBytes, oldBytes, 0)
            
            If pos < 0 Then
                ' Path not found in this file
                Continue For
            End If
            
            ' Handle length difference
            Dim newBytes As Byte() = Nothing
            
            If oldPath.Length = newPath.Length Then
                ' Same length - direct replacement
                newBytes = Encoding.Unicode.GetBytes(newPath)
            ElseIf newPath.Length < oldPath.Length Then
                ' New is shorter - pad with null characters
                Dim paddedNew As String = newPath & New String(ChrW(0), oldPath.Length - newPath.Length)
                newBytes = Encoding.Unicode.GetBytes(paddedNew)
            Else
                ' New is longer - this is a problem
                logMessages.Add("  ERROR: New path is longer than old path!")
                logMessages.Add("    Old (" & oldPath.Length & "): " & oldPath)
                logMessages.Add("    New (" & newPath.Length & "): " & newPath)
                logMessages.Add("    Path length difference: " & (newPath.Length - oldPath.Length))
                Return False
            End If
            
            ' Replace all occurrences
            Do While pos >= 0
                Array.Copy(newBytes, 0, fileBytes, pos, newBytes.Length)
                totalReplacements += 1
                pos = IndexOfBytes(fileBytes, oldBytes, pos + oldBytes.Length)
            Loop
            
            modified = True
            logMessages.Add("  Replaced: " & System.IO.Path.GetFileName(oldPath) & _
                          " -> " & System.IO.Path.GetFileName(newPath))
        Next
        
        ' Write modified file
        If modified Then
            Try
                System.IO.File.WriteAllBytes(filePath, fileBytes)
                logMessages.Add("  Updated " & totalReplacements & " reference(s)")
            Catch ex As Exception
                logMessages.Add("  ERROR writing file: " & ex.Message)
                ' Try to restore backup
                Try
                    System.IO.File.Copy(backupPath, filePath, True)
                    logMessages.Add("  Restored from backup")
                Catch
                End Try
                Return False
            End Try
        End If
        
        Return True
    End Function

    ''' <summary>
    ''' Calculate a variant folder path that matches the master folder path length.
    ''' </summary>
    ''' <param name="masterFolder">The master folder path (e.g., "C:\Project\Master")</param>
    ''' <param name="releaseRoot">The release root folder (e.g., "C:\Project\Release")</param>
    ''' <param name="variantName">The desired variant name</param>
    ''' <returns>A variant folder path that matches the master folder length</returns>
    Public Function CalculateMatchingVariantPath(masterFolder As String, _
                                                  releaseRoot As String, _
                                                  variantName As String) As String
        ' Normalize paths
        masterFolder = System.IO.Path.GetFullPath(masterFolder).TrimEnd(System.IO.Path.DirectorySeparatorChar)
        releaseRoot = System.IO.Path.GetFullPath(releaseRoot).TrimEnd(System.IO.Path.DirectorySeparatorChar)
        
        Dim masterLength As Integer = masterFolder.Length
        Dim releaseRootLength As Integer = releaseRoot.Length
        
        ' We need: releaseRoot + separator + variantFolderName = masterLength
        ' So: variantFolderName.Length = masterLength - releaseRootLength - 1
        Dim requiredVariantLength As Integer = masterLength - releaseRootLength - 1
        
        If requiredVariantLength <= 0 Then
            ' Release root is already longer than or equal to master
            ' Use minimum folder name
            Return System.IO.Path.Combine(releaseRoot, "v")
        End If
        
        ' Create variant folder name of required length
        Dim variantFolderName As String = ""
        
        If variantName.Length = requiredVariantLength Then
            ' Exact match
            variantFolderName = variantName
        ElseIf variantName.Length < requiredVariantLength Then
            ' Pad with underscores
            variantFolderName = variantName & New String("_"c, requiredVariantLength - variantName.Length)
        Else
            ' Truncate
            variantFolderName = variantName.Substring(0, requiredVariantLength)
        End If
        
        Return System.IO.Path.Combine(releaseRoot, variantFolderName)
    End Function

    ''' <summary>
    ''' Build a path map that ensures all new paths match the length of old paths.
    ''' </summary>
    ''' <param name="sourceFiles">List of source file paths</param>
    ''' <param name="masterRoot">Master folder root</param>
    ''' <param name="targetRoot">Target folder root (should be length-matched)</param>
    ''' <param name="variantName">Variant name for renaming main assembly</param>
    ''' <param name="mainAsmPath">Path to main assembly</param>
    ''' <returns>Dictionary mapping old paths to new paths (same length)</returns>
    Public Function BuildLengthMatchedCopyMap(sourceFiles As List(Of String), _
                                               masterRoot As String, _
                                               targetRoot As String, _
                                               variantName As String, _
                                               mainAsmPath As String) As Dictionary(Of String, String)
        Dim copyMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        
        ' Normalize roots
        masterRoot = System.IO.Path.GetFullPath(masterRoot).TrimEnd(System.IO.Path.DirectorySeparatorChar)
        targetRoot = System.IO.Path.GetFullPath(targetRoot).TrimEnd(System.IO.Path.DirectorySeparatorChar)
        
        ' Check if roots are same length
        If masterRoot.Length <> targetRoot.Length Then
            ' This will cause problems - log warning
            ' The caller should use CalculateMatchingVariantPath first
        End If
        
        Dim mainAsmName As String = System.IO.Path.GetFileNameWithoutExtension(mainAsmPath)
        
        For Each sourcePath As String In sourceFiles
            Dim normalizedSource As String = System.IO.Path.GetFullPath(sourcePath)
            
            ' Check if inside master root
            If Not normalizedSource.StartsWith(masterRoot, StringComparison.OrdinalIgnoreCase) Then
                ' External file - don't copy
                Continue For
            End If
            
            ' Get relative path from master root
            Dim relativePath As String = normalizedSource.Substring(masterRoot.Length)
            If relativePath.StartsWith(System.IO.Path.DirectorySeparatorChar) Then
                relativePath = relativePath.Substring(1)
            End If
            
            ' Build target path
            Dim targetPath As String = System.IO.Path.Combine(targetRoot, relativePath)
            
            ' Handle renaming of main assembly and related files
            Dim fileNameNoExt As String = System.IO.Path.GetFileNameWithoutExtension(sourcePath)
            Dim ext As String = System.IO.Path.GetExtension(sourcePath)
            
            If fileNameNoExt.Equals(mainAsmName, StringComparison.OrdinalIgnoreCase) Then
                ' Rename to variant name - but must keep same length!
                Dim newFileName As String = variantName
                
                If newFileName.Length <> fileNameNoExt.Length Then
                    ' Adjust variant name to match length
                    If newFileName.Length < fileNameNoExt.Length Then
                        newFileName = newFileName & New String("_"c, fileNameNoExt.Length - newFileName.Length)
                    Else
                        newFileName = newFileName.Substring(0, fileNameNoExt.Length)
                    End If
                End If
                
                Dim targetDir As String = System.IO.Path.GetDirectoryName(targetPath)
                targetPath = System.IO.Path.Combine(targetDir, newFileName & ext)
            End If
            
            ' Verify lengths match
            If normalizedSource.Length <> targetPath.Length Then
                ' Path length mismatch - this will cause issues
                ' For now, continue but this should be flagged
            End If
            
            copyMap(normalizedSource) = targetPath
        Next
        
        Return copyMap
    End Function

    ''' <summary>
    ''' Get all file references from a document without opening it fully.
    ''' Uses binary search to find referenced file paths.
    ''' </summary>
    ''' <param name="filePath">Path to the file</param>
    ''' <returns>List of referenced file paths found in the binary</returns>
    Public Function GetReferencesFromBinary(filePath As String) As List(Of String)
        Dim references As New List(Of String)
        
        If Not System.IO.File.Exists(filePath) Then
            Return references
        End If
        
        Try
            Dim fileBytes As Byte() = System.IO.File.ReadAllBytes(filePath)
            Dim fileText As String = Encoding.Unicode.GetString(fileBytes)
            
            ' Search for .ipt and .iam references
            Dim patterns As String() = {".ipt", ".iam"}
            
            For Each pattern As String In patterns
                Dim idx As Integer = 0
                Do
                    idx = fileText.IndexOf(pattern, idx, StringComparison.OrdinalIgnoreCase)
                    If idx < 0 Then Exit Do
                    
                    ' Extract the full path by looking backwards
                    Dim pathStr As String = ExtractPathBackwards(fileText, idx + pattern.Length)
                    
                    If Not String.IsNullOrEmpty(pathStr) AndAlso Not references.Contains(pathStr) Then
                        ' Validate it looks like a real path
                        If pathStr.Contains(":\") OrElse pathStr.StartsWith("\\") Then
                            references.Add(pathStr)
                        End If
                    End If
                    
                    idx += 1
                Loop
            Next
        Catch
        End Try
        
        Return references
    End Function

    ''' <summary>
    ''' Extract a file path by looking backwards from the extension.
    ''' </summary>
    Private Function ExtractPathBackwards(text As String, endIdx As Integer) As String
        Dim startIdx As Integer = endIdx
        
        ' Look back for drive letter or UNC path start
        For i As Integer = endIdx - 1 To Math.Max(0, endIdx - 500) Step -1
            Dim c As Char = text(i)
            
            ' Check for drive letter pattern (X:)
            If c = ":"c AndAlso i > 0 AndAlso Char.IsLetter(text(i - 1)) Then
                startIdx = i - 1
                Exit For
            End If
            
            ' Check for UNC path start (\\)
            If c = "\"c AndAlso i > 0 AndAlso text(i - 1) = "\"c Then
                startIdx = i - 1
                Exit For
            End If
            
            ' Stop at null or control characters
            If Asc(c) < 32 Then
                startIdx = i + 1
                Exit For
            End If
        Next
        
        If startIdx < endIdx Then
            Dim path As String = text.Substring(startIdx, endIdx - startIdx)
            ' Clean up null characters and spaces
            path = path.Replace(ChrW(0), "").Trim()
            Return path
        End If
        
        Return ""
    End Function

    ''' <summary>
    ''' Find byte sequence in array.
    ''' </summary>
    Private Function IndexOfBytes(source As Byte(), pattern As Byte(), startIndex As Integer) As Integer
        If source Is Nothing OrElse pattern Is Nothing Then Return -1
        If pattern.Length > source.Length - startIndex Then Return -1
        
        For i As Integer = startIndex To source.Length - pattern.Length
            Dim found As Boolean = True
            For j As Integer = 0 To pattern.Length - 1
                If source(i + j) <> pattern(j) Then
                    found = False
                    Exit For
                End If
            Next
            If found Then Return i
        Next
        
        Return -1
    End Function

End Module

