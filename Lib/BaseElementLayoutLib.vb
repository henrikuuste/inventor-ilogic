' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' BaseElementLayoutLib
'
' Shared helpers for base element path detection and folder creation.
' Terminology updated 2026-05-12: module -> element per UBIQUITOUS_LANGUAGE.md
' ============================================================================

Imports Inventor

Public Module BaseElementLayoutLib

    Public Const SEG_BASE_ELEMENTS As String = "Aluselemendid"
    Public Const SEG_SKETCH As String = "Eskiis"
    Public Const SEG_FRAME As String = "Karkass"
    Public Const SEG_PADDING As String = "Poroloon"
    Public Const SEG_PARTS As String = "Detailid"
    Public Const SEG_DRAWINGS As String = "Joonised"

    ' Detects <project>\Aluselemendid\<elementName> from any subfolder under the element root.
    ' Anchors detection to projectRoot/projectName to avoid matching unrelated paths.
    Public Function DetectElementRootFromMasterPath(masterFullPath As String, _
                                                    projectRoot As String, _
                                                    projectName As String) As String
        If String.IsNullOrEmpty(masterFullPath) Then Return ""

        Try
            Dim normalizedDocPath As String = masterFullPath.Replace("/", "\")
            Dim searchPath As String = normalizedDocPath
            If System.IO.File.Exists(normalizedDocPath) Then
                searchPath = System.IO.Path.GetDirectoryName(normalizedDocPath)
            End If

            Dim normalizedProjectRoot As String = projectRoot.Replace("/", "\")
            If Not String.IsNullOrEmpty(normalizedProjectRoot) AndAlso _
               Not normalizedProjectRoot.EndsWith("\") Then
                normalizedProjectRoot &= "\"
            End If

            ' 1) Preferred: bounded inside project root
            If Not String.IsNullOrEmpty(normalizedProjectRoot) AndAlso _
               searchPath.StartsWith(normalizedProjectRoot, StringComparison.OrdinalIgnoreCase) Then
                Dim relative As String = searchPath.Substring(normalizedProjectRoot.Length)
                Dim relSegments() As String = relative.Split("\"c)
                For i As Integer = 0 To relSegments.Length - 2
                    If relSegments(i).Equals(SEG_BASE_ELEMENTS, StringComparison.OrdinalIgnoreCase) Then
                        Dim elementName As String = relSegments(i + 1)
                        If Not String.IsNullOrEmpty(elementName) Then
                            Return System.IO.Path.Combine(normalizedProjectRoot.TrimEnd("\"c), SEG_BASE_ELEMENTS, elementName)
                        End If
                    End If
                Next
            End If

            ' 2) Fallback: anchor by project name in full path
            If Not String.IsNullOrEmpty(projectName) Then
                Dim segments() As String = searchPath.Split("\"c)
                For i As Integer = 0 To segments.Length - 2
                    If segments(i).Equals(SEG_BASE_ELEMENTS, StringComparison.OrdinalIgnoreCase) Then
                        Dim elementName As String = segments(i + 1)
                        If String.IsNullOrEmpty(elementName) Then Continue For

                        Dim candidateRoot As String = String.Join("\", segments, 0, i + 2)
                        If candidateRoot.IndexOf("\" & projectName & "\", StringComparison.OrdinalIgnoreCase) >= 0 OrElse _
                           candidateRoot.EndsWith("\" & projectName, StringComparison.OrdinalIgnoreCase) Then
                            Return candidateRoot
                        End If
                    End If
                Next
            End If
        Catch
        End Try

        Return ""
    End Function

    Public Function GetElementName(elementRoot As String) As String
        If String.IsNullOrEmpty(elementRoot) Then Return ""
        Try
            Return System.IO.Path.GetFileName(elementRoot.TrimEnd("\"c))
        Catch
            Return ""
        End Try
    End Function

    Public Function EnumerateExpectedFolders(projectPath As String, elementName As String) As System.Collections.Generic.List(Of String)
        Dim folders As New System.Collections.Generic.List(Of String)
        If String.IsNullOrEmpty(projectPath) OrElse String.IsNullOrEmpty(elementName) Then Return folders

        Dim elementRoot As String = System.IO.Path.Combine(projectPath, SEG_BASE_ELEMENTS, elementName)
        folders.Add(System.IO.Path.Combine(elementRoot, SEG_SKETCH))
        folders.Add(System.IO.Path.Combine(elementRoot, SEG_FRAME, SEG_PARTS))
        folders.Add(System.IO.Path.Combine(elementRoot, SEG_FRAME, SEG_DRAWINGS))
        folders.Add(System.IO.Path.Combine(elementRoot, SEG_PADDING, SEG_PARTS))
        folders.Add(System.IO.Path.Combine(elementRoot, SEG_PADDING, SEG_DRAWINGS))
        Return folders
    End Function

    ' Creates standard base-element folders locally and in Vault (if connected).
    ' Returns element root path: <projectPath>\Aluselemendid\<elementName>
    Public Function EnsureBaseElementLayout(projectPath As String, _
                                            elementName As String, _
                                            vaultConn As Object, _
                                            workspaceRoot As String) As String
        If String.IsNullOrEmpty(projectPath) OrElse String.IsNullOrEmpty(elementName) Then Return ""

        Dim elementRoot As String = System.IO.Path.Combine(projectPath, SEG_BASE_ELEMENTS, elementName)
        Dim folders As System.Collections.Generic.List(Of String) = EnumerateExpectedFolders(projectPath, elementName)

        For Each folderPath As String In folders
            Try
                If Not System.IO.Directory.Exists(folderPath) Then
                    System.IO.Directory.CreateDirectory(folderPath)
                    UtilsLib.LogInfo("BaseElementLayoutLib: Created local folder: " & folderPath)
                End If

                If vaultConn IsNot Nothing AndAlso Not String.IsNullOrEmpty(workspaceRoot) Then
                    Dim vaultPath As String = VaultNumberingLib.ConvertLocalPathToVaultPath(folderPath, workspaceRoot)
                    If Not String.IsNullOrEmpty(vaultPath) Then
                        Dim vaultFolder As Object = VaultNumberingLib.EnsureVaultFolderRecursive(vaultConn, vaultPath)
                        If vaultFolder IsNot Nothing Then
                            UtilsLib.LogInfo("BaseElementLayoutLib: Vault folder ready: " & vaultPath)
                        Else
                            UtilsLib.LogWarn("BaseElementLayoutLib: Could not create Vault folder: " & vaultPath)
                        End If
                    End If
                End If
            Catch ex As Exception
                UtilsLib.LogWarn("BaseElementLayoutLib: Failed to create folder '" & folderPath & "': " & ex.Message)
            End Try
        Next

        Return elementRoot
    End Function

End Module
