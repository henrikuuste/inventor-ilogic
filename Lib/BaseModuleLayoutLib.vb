' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' BaseModuleLayoutLib
'
' Shared helpers for base module path detection and folder creation.
' NOTE: when terminology/folder naming is refactored (module -> element),
' update the constants in this file and callers can remain mostly unchanged.
' ============================================================================

Imports Inventor

Public Module BaseModuleLayoutLib

    Public Const SEG_BASE_MODULES As String = "Alusmoodulid"
    Public Const SEG_SKETCH As String = "Eskiis"
    Public Const SEG_FRAME As String = "Karkass"
    Public Const SEG_PADDING As String = "Poroloon"
    Public Const SEG_PARTS As String = "Detailid"
    Public Const SEG_DRAWINGS As String = "Joonised"

    ' Detects <project>\Alusmoodulid\<moduleName> from any subfolder under the module root.
    ' Anchors detection to projectRoot/projectName to avoid matching unrelated paths.
    Public Function DetectModuleRootFromMasterPath(masterFullPath As String, _
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
                    If relSegments(i).Equals(SEG_BASE_MODULES, StringComparison.OrdinalIgnoreCase) Then
                        Dim moduleName As String = relSegments(i + 1)
                        If Not String.IsNullOrEmpty(moduleName) Then
                            Return System.IO.Path.Combine(normalizedProjectRoot.TrimEnd("\"c), SEG_BASE_MODULES, moduleName)
                        End If
                    End If
                Next
            End If

            ' 2) Fallback: anchor by project name in full path
            If Not String.IsNullOrEmpty(projectName) Then
                Dim segments() As String = searchPath.Split("\"c)
                For i As Integer = 0 To segments.Length - 2
                    If segments(i).Equals(SEG_BASE_MODULES, StringComparison.OrdinalIgnoreCase) Then
                        Dim moduleName As String = segments(i + 1)
                        If String.IsNullOrEmpty(moduleName) Then Continue For

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

    Public Function GetModuleName(moduleRoot As String) As String
        If String.IsNullOrEmpty(moduleRoot) Then Return ""
        Try
            Return System.IO.Path.GetFileName(moduleRoot.TrimEnd("\"c))
        Catch
            Return ""
        End Try
    End Function

    Public Function EnumerateExpectedFolders(projectPath As String, moduleName As String) As System.Collections.Generic.List(Of String)
        Dim folders As New System.Collections.Generic.List(Of String)
        If String.IsNullOrEmpty(projectPath) OrElse String.IsNullOrEmpty(moduleName) Then Return folders

        Dim moduleRoot As String = System.IO.Path.Combine(projectPath, SEG_BASE_MODULES, moduleName)
        folders.Add(System.IO.Path.Combine(moduleRoot, SEG_SKETCH))
        folders.Add(System.IO.Path.Combine(moduleRoot, SEG_FRAME, SEG_PARTS))
        folders.Add(System.IO.Path.Combine(moduleRoot, SEG_FRAME, SEG_DRAWINGS))
        folders.Add(System.IO.Path.Combine(moduleRoot, SEG_PADDING, SEG_PARTS))
        folders.Add(System.IO.Path.Combine(moduleRoot, SEG_PADDING, SEG_DRAWINGS))
        Return folders
    End Function

    ' Creates standard base-module folders locally and in Vault (if connected).
    ' Returns module root path: <projectPath>\Alusmoodulid\<moduleName>
    Public Function EnsureBaseModuleLayout(projectPath As String, _
                                           moduleName As String, _
                                           vaultConn As Object, _
                                           workspaceRoot As String) As String
        If String.IsNullOrEmpty(projectPath) OrElse String.IsNullOrEmpty(moduleName) Then Return ""

        Dim moduleRoot As String = System.IO.Path.Combine(projectPath, SEG_BASE_MODULES, moduleName)
        Dim folders As System.Collections.Generic.List(Of String) = EnumerateExpectedFolders(projectPath, moduleName)

        For Each folderPath As String In folders
            Try
                If Not System.IO.Directory.Exists(folderPath) Then
                    System.IO.Directory.CreateDirectory(folderPath)
                    UtilsLib.LogInfo("BaseModuleLayoutLib: Created local folder: " & folderPath)
                End If

                If vaultConn IsNot Nothing AndAlso Not String.IsNullOrEmpty(workspaceRoot) Then
                    Dim vaultPath As String = VaultNumberingLib.ConvertLocalPathToVaultPath(folderPath, workspaceRoot)
                    If Not String.IsNullOrEmpty(vaultPath) Then
                        Dim vaultFolder As Object = VaultNumberingLib.EnsureVaultFolderRecursive(vaultConn, vaultPath)
                        If vaultFolder IsNot Nothing Then
                            UtilsLib.LogInfo("BaseModuleLayoutLib: Vault folder ready: " & vaultPath)
                        Else
                            UtilsLib.LogWarn("BaseModuleLayoutLib: Could not create Vault folder: " & vaultPath)
                        End If
                    End If
                End If
            Catch ex As Exception
                UtilsLib.LogWarn("BaseModuleLayoutLib: Failed to create folder '" & folderPath & "': " & ex.Message)
            End Try
        Next

        Return moduleRoot
    End Function

End Module
