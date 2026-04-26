' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' FileSearchLib - Generic file search utilities with depth-first traversal
' 
' Provides optimized file search that:
' - Starts from document's folder (most likely location)
' - Searches current folder first, then children
' - Goes up one level, searches siblings
' - Stops at Vault root + 2 levels
'
' This approach finds nearby files quickly without scanning the entire workspace.
' ============================================================================

Imports System.Collections.Generic

Public Module FileSearchLib
    
    ' Get the depth of a path (number of directory levels)
    Public Function GetPathDepth(path As String) As Integer
        If String.IsNullOrEmpty(path) Then Return 0
        Return path.Split(System.IO.Path.DirectorySeparatorChar).Length
    End Function
    
    ' Calculate minimum search depth (vaultRoot + 2 levels)
    ' Search will not go above this depth
    Public Function GetMinSearchDepth(vaultRoot As String) As Integer
        If String.IsNullOrEmpty(vaultRoot) Then Return 2
        Return GetPathDepth(vaultRoot) + 2
    End Function
    
    ' Check if a folder path should be skipped (OldVersions, etc.)
    Public Function ShouldSkipFolder(folderPath As String) As Boolean
        If String.IsNullOrEmpty(folderPath) Then Return True
        Return folderPath.IndexOf("\OldVersions", StringComparison.OrdinalIgnoreCase) >= 0
    End Function
    
    ' Search for exact filename using depth-first folder order
    ' Fast - no file opening required, just checks if file exists
    ' Returns the full path of the first matching file, or empty string if not found
    Public Function FindFileByName(fileName As String, _
                                   startPath As String, _
                                   vaultRoot As String) As String
        If String.IsNullOrEmpty(fileName) OrElse String.IsNullOrEmpty(startPath) Then
            Return ""
        End If
        
        Try
            Dim minDepth As Integer = GetMinSearchDepth(vaultRoot)
            Dim currentPath As String = startPath
            Dim searchedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            
            ' Search upward from startPath until we reach minimum depth
            While Not String.IsNullOrEmpty(currentPath) AndAlso GetPathDepth(currentPath) >= minDepth
                ' Search this folder and its children for the file
                Dim result As String = SearchFolderForFile(fileName, currentPath, searchedPaths)
                If Not String.IsNullOrEmpty(result) Then
                    Return result
                End If
                
                ' Go up one level
                Dim parentPath As String = System.IO.Path.GetDirectoryName(currentPath)
                If String.IsNullOrEmpty(parentPath) OrElse parentPath = currentPath Then
                    Exit While
                End If
                currentPath = parentPath
            End While
            
        Catch ex As Exception
            ' Silently handle errors - return empty string
        End Try
        
        Return ""
    End Function
    
    ' Search a folder and its subfolders for a specific filename
    Private Function SearchFolderForFile(fileName As String, _
                                         folderPath As String, _
                                         searchedPaths As HashSet(Of String)) As String
        ' Skip if already searched or doesn't exist
        If searchedPaths.Contains(folderPath) Then Return ""
        If Not System.IO.Directory.Exists(folderPath) Then Return ""
        
        searchedPaths.Add(folderPath)
        
        ' Skip OldVersions folders
        If ShouldSkipFolder(folderPath) Then Return ""
        
        Try
            ' Check if file exists in current folder
            Dim filePath As String = System.IO.Path.Combine(folderPath, fileName)
            If System.IO.File.Exists(filePath) Then
                Return filePath
            End If
            
            ' Search subfolders
            Dim subDirs() As String = System.IO.Directory.GetDirectories(folderPath)
            For Each subDir As String In subDirs
                Dim result As String = SearchFolderForFile(fileName, subDir, searchedPaths)
                If Not String.IsNullOrEmpty(result) Then
                    Return result
                End If
            Next
            
        Catch ex As Exception
            ' Ignore access errors for individual folders
        End Try
        
        Return ""
    End Function
    
    ' Search files matching pattern with custom checker object
    ' Searches: current folder -> children -> parent -> siblings -> grandparent...
    ' Stops at vaultRoot + 2 levels
    ' fileChecker: Object with CheckFile(filePath As String) As Boolean method (late-bound)
    ' Returns first matching file path, or empty string if not found
    Public Function SearchFilesWithChecker(startPath As String, _
                                           vaultRoot As String, _
                                           filePattern As String, _
                                           fileChecker As Object) As String
        If String.IsNullOrEmpty(startPath) OrElse String.IsNullOrEmpty(filePattern) Then
            Return ""
        End If
        
        If fileChecker Is Nothing Then
            Return ""
        End If
        
        Try
            Dim minDepth As Integer = GetMinSearchDepth(vaultRoot)
            Dim currentPath As String = startPath
            Dim searchedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            
            ' Search upward from startPath until we reach minimum depth
            While Not String.IsNullOrEmpty(currentPath) AndAlso GetPathDepth(currentPath) >= minDepth
                ' Search this folder and its children
                Dim result As String = SearchFolderWithChecker(filePattern, currentPath, searchedPaths, fileChecker)
                If Not String.IsNullOrEmpty(result) Then
                    Return result
                End If
                
                ' Go up one level
                Dim parentPath As String = System.IO.Path.GetDirectoryName(currentPath)
                If String.IsNullOrEmpty(parentPath) OrElse parentPath = currentPath Then
                    Exit While
                End If
                currentPath = parentPath
            End While
            
        Catch ex As Exception
            ' Silently handle errors - return empty string
        End Try
        
        Return ""
    End Function
    
    ' Search a folder and its subfolders for files matching pattern, using checker
    Private Function SearchFolderWithChecker(filePattern As String, _
                                             folderPath As String, _
                                             searchedPaths As HashSet(Of String), _
                                             fileChecker As Object) As String
        ' Skip if already searched or doesn't exist
        If searchedPaths.Contains(folderPath) Then Return ""
        If Not System.IO.Directory.Exists(folderPath) Then Return ""
        
        searchedPaths.Add(folderPath)
        
        ' Skip OldVersions folders
        If ShouldSkipFolder(folderPath) Then Return ""
        
        Try
            ' Search files in current folder first
            Dim files() As String = System.IO.Directory.GetFiles(folderPath, filePattern, System.IO.SearchOption.TopDirectoryOnly)
            
            For Each filePath As String In files
                ' Skip OldVersions files (shouldn't happen but double-check)
                If filePath.IndexOf("\OldVersions\", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Continue For
                End If
                
                ' Call checker's CheckFile method via late binding
                Try
                    Dim isMatch As Boolean = CBool(CallByName(fileChecker, "CheckFile", Microsoft.VisualBasic.CallType.Method, filePath))
                    If isMatch Then
                        Return filePath
                    End If
                Catch
                    ' Skip file if checker fails
                End Try
            Next
            
            ' Search subfolders
            Dim subDirs() As String = System.IO.Directory.GetDirectories(folderPath)
            For Each subDir As String In subDirs
                Dim result As String = SearchFolderWithChecker(filePattern, subDir, searchedPaths, fileChecker)
                If Not String.IsNullOrEmpty(result) Then
                    Return result
                End If
            Next
            
        Catch ex As Exception
            ' Ignore access errors for individual folders
        End Try
        
        Return ""
    End Function
    
End Module
