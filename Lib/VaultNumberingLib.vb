' ============================================================================
' VaultNumberingLib - Vault WebServices API wrapper for number generation
' 
' Provides functions to:
' - Check Vault connection status
' - Enumerate available numbering schemes
' - Generate file numbers from a specific scheme
'
' Usage: 
'   In calling script (BEFORE AddVbFile):
'     AddReference "Autodesk.Connectivity.WebServices"
'     AddReference "Autodesk.DataManagement.Client.Framework.Vault"
'     AddReference "Autodesk.DataManagement.Client.Framework.Vault.Forms"
'     AddReference "Connectivity.InventorAddin.EdmAddin"
'     AddVbFile "Lib/VaultNumberingLib.vb"
'
' Note: Logger is not available in library modules.
'       Pass a List(Of String) to collect log messages.
' ============================================================================

Public Module VaultNumberingLib

    ' Get the current Vault connection, or Nothing if not logged in
    Public Function GetVaultConnection() As Object
        Try
            Return Connectivity.InventorAddin.EdmAddin.EdmSecurity.Instance.VaultConnection()
        Catch
            Return Nothing
        End Try
    End Function
    
    ' Check if user is logged into Vault
    Public Function IsVaultConnected() As Boolean
        Return GetVaultConnection() IsNot Nothing
    End Function
    
    ' Get connection info for logging
    Public Function GetConnectionInfo(conn As Object) As String
        If conn Is Nothing Then Return "Not connected"
        Try
            Return "Server: " & conn.Server & ", Vault: " & conn.Vault & ", User: " & conn.UserName
        Catch
            Return "Connected (details unavailable)"
        End Try
    End Function
    
    ' Get available numbering schemes for files
    Public Function GetNumberingSchemes(conn As Object, _
                                        logs As System.Collections.Generic.List(Of String)) As Object()
        If conn Is Nothing Then
            logs.Add("VaultNumberingLib: No Vault connection")
            Return Nothing
        End If
        
        Try
            Dim schemes As Object() = conn.WebServiceManager.NumberingService.GetNumberingSchemes("FILE", Nothing)
            logs.Add("VaultNumberingLib: Found " & schemes.Length & " numbering scheme(s)")
            Return schemes
        Catch ex As Exception
            logs.Add("VaultNumberingLib: Error getting schemes: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ' Get scheme names as a list (for dropdown)
    Public Function GetSchemeNames(conn As Object, _
                                   logs As System.Collections.Generic.List(Of String)) As System.Collections.Generic.List(Of String)
        Dim names As New System.Collections.Generic.List(Of String)
        Dim schemes As Object() = GetNumberingSchemes(conn, logs)
        
        If schemes IsNot Nothing Then
            For Each scheme As Object In schemes
                names.Add(scheme.Name)
            Next
        End If
        
        Return names
    End Function
    
    ' Find a scheme by name
    Public Function FindSchemeByName(conn As Object, _
                                     schemeName As String, _
                                     logs As System.Collections.Generic.List(Of String)) As Object
        Dim schemes As Object() = GetNumberingSchemes(conn, logs)
        
        If schemes Is Nothing Then Return Nothing
        
        Dim searchName As String = schemeName.Trim()
        logs.Add("VaultNumberingLib: Looking for scheme '" & searchName & "' (len=" & searchName.Length & ")")
        
        For Each scheme As Object In schemes
            Dim schName As String = CStr(scheme.Name).Trim()
            logs.Add("VaultNumberingLib:   Comparing with '" & schName & "' (len=" & schName.Length & ")")
            If schName.Equals(searchName, StringComparison.OrdinalIgnoreCase) Then
                logs.Add("VaultNumberingLib: Found matching scheme")
                Return scheme
            End If
        Next
        
        logs.Add("VaultNumberingLib: Scheme '" & searchName & "' not found")
        Return Nothing
    End Function
    
    ' Generate a file number from a specific scheme
    Public Function GenerateFileNumber(conn As Object, _
                                       scheme As Object, _
                                       logs As System.Collections.Generic.List(Of String)) As String
        If conn Is Nothing Then
            logs.Add("VaultNumberingLib: No Vault connection")
            Return ""
        End If
        
        If scheme Is Nothing Then
            logs.Add("VaultNumberingLib: No scheme specified")
            Return ""
        End If
        
        Try
            Dim numGenArgs() As String = {""}
            Dim number As String = conn.WebServiceManager.DocumentService.GenerateFileNumber(scheme.SchmID, numGenArgs)
            logs.Add("VaultNumberingLib: Generated number: " & number)
            Return number
        Catch ex As Exception
            logs.Add("VaultNumberingLib: Error generating number: " & ex.Message)
            Return ""
        End Try
    End Function
    
    ' Generate a file number by scheme name (convenience function)
    Public Function GenerateFileNumberByName(conn As Object, _
                                             schemeName As String, _
                                             logs As System.Collections.Generic.List(Of String)) As String
        Dim scheme As Object = FindSchemeByName(conn, schemeName, logs)
        If scheme Is Nothing Then Return ""
        Return GenerateFileNumber(conn, scheme, logs)
    End Function
    
    ' Generate multiple file numbers at once
    Public Function GenerateFileNumbers(conn As Object, _
                                        scheme As Object, _
                                        count As Integer, _
                                        logs As System.Collections.Generic.List(Of String)) As System.Collections.Generic.List(Of String)
        Dim numbers As New System.Collections.Generic.List(Of String)
        
        For i As Integer = 1 To count
            Dim num As String = GenerateFileNumber(conn, scheme, logs)
            If String.IsNullOrEmpty(num) Then
                logs.Add("VaultNumberingLib: Failed to generate number " & i & " of " & count)
                Exit For
            End If
            numbers.Add(num)
        Next
        
        Return numbers
    End Function

End Module
