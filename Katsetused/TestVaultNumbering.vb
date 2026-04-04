' ============================================================================
' TestVaultNumbering - Test Vault WebServices API for number generation
' 
' Tests:
' - Can we get VaultConnection?
' - Can we enumerate numbering schemes?
' - Can we generate a number from a specific scheme?
' - What happens when not logged in?
'
' Usage: Run while logged into Vault to test full functionality.
'        Run while logged out to test fallback behavior.
' ============================================================================

AddReference "Autodesk.Connectivity.WebServices"
AddReference "Autodesk.DataManagement.Client.Framework.Vault"
AddReference "Autodesk.DataManagement.Client.Framework.Vault.Forms"
AddReference "Connectivity.InventorAddin.EdmAddin"

Imports ACW = Autodesk.Connectivity.WebServices
Imports VDF = Autodesk.DataManagement.Client.Framework
Imports edm = Connectivity.InventorAddin.EdmAddin

Sub Main()
    Logger.Info("TestVaultNumbering: Starting Vault API tests...")
    
    ' Test 1: Get Vault connection
    Logger.Info("TestVaultNumbering: Test 1 - Getting Vault connection...")
    Dim conn As VDF.Vault.Currency.Connections.Connection = Nothing
    
    Try
        conn = edm.EdmSecurity.Instance.VaultConnection()
    Catch ex As Exception
        Logger.Error("TestVaultNumbering: Exception getting Vault connection: " & ex.Message)
        MessageBox.Show("Vault ühenduse viga: " & ex.Message, "TestVaultNumbering")
        Exit Sub
    End Try
    
    If conn Is Nothing Then
        Logger.Warn("TestVaultNumbering: No Vault connection available.")
        Logger.Warn("TestVaultNumbering: User is not logged into Vault.")
        MessageBox.Show("Vault ühendus puudub. Palun logi Vault'i sisse ja proovi uuesti.", "TestVaultNumbering")
        Exit Sub
    End If
    
    Logger.Info("TestVaultNumbering: Vault connection established successfully!")
    Logger.Info("TestVaultNumbering: Server: " & conn.Server)
    Logger.Info("TestVaultNumbering: Vault: " & conn.Vault)
    Logger.Info("TestVaultNumbering: User: " & conn.UserName)
    
    ' Test 2: Enumerate numbering schemes
    Logger.Info("TestVaultNumbering: Test 2 - Enumerating numbering schemes...")
    Dim schemes As ACW.NumSchm() = Nothing
    
    Try
        ' Get file numbering schemes using NumberingService
        ' EntityClassId for files is "FILE" (string)
        schemes = conn.WebServiceManager.NumberingService.GetNumberingSchemes("FILE", Nothing)
    Catch ex As Exception
        Logger.Error("TestVaultNumbering: Exception getting numbering schemes: " & ex.Message)
        Exit Sub
    End Try
    
    If schemes Is Nothing OrElse schemes.Length = 0 Then
        Logger.Warn("TestVaultNumbering: No numbering schemes found.")
        Exit Sub
    End If
    
    Logger.Info("TestVaultNumbering: Found " & schemes.Length & " numbering scheme(s):")
    For Each scheme As ACW.NumSchm In schemes
        Logger.Info("TestVaultNumbering:   - Name: '" & scheme.Name & "' (ID: " & scheme.SchmID & ")")
    Next
    
    ' Test 3: Find specific scheme (try common names)
    Logger.Info("TestVaultNumbering: Test 3 - Looking for specific schemes...")
    Dim schemeNames() As String = {"Softcom numbriskeem", "Softcom numbrisüsteem", "Default", "Sequential"}
    Dim targetScheme As ACW.NumSchm = Nothing
    
    For Each schemeName As String In schemeNames
        For Each scheme As ACW.NumSchm In schemes
            If scheme.Name.Equals(schemeName, StringComparison.OrdinalIgnoreCase) Then
                targetScheme = scheme
                Logger.Info("TestVaultNumbering: Found scheme '" & schemeName & "'")
                Exit For
            End If
        Next
        If targetScheme IsNot Nothing Then Exit For
    Next
    
    If targetScheme Is Nothing Then
        ' Use first available scheme
        targetScheme = schemes(0)
        Logger.Info("TestVaultNumbering: Using first available scheme: '" & targetScheme.Name & "'")
    End If
    
    ' Test 4: Generate a number
    Logger.Info("TestVaultNumbering: Test 4 - Generating a number from scheme '" & targetScheme.Name & "'...")
    Dim generatedNumber As String = ""
    
    Try
        Dim numGenArgs() As String = {""}
        generatedNumber = conn.WebServiceManager.DocumentService.GenerateFileNumber(targetScheme.SchmID, numGenArgs)
    Catch ex As Exception
        Logger.Error("TestVaultNumbering: Exception generating number: " & ex.Message)
        Exit Sub
    End Try
    
    If String.IsNullOrEmpty(generatedNumber) Then
        Logger.Warn("TestVaultNumbering: Generated number is empty.")
    Else
        Logger.Info("TestVaultNumbering: Generated number: '" & generatedNumber & "'")
    End If
    
    ' Summary
    Logger.Info("TestVaultNumbering: ========================================")
    Logger.Info("TestVaultNumbering: TEST SUMMARY")
    Logger.Info("TestVaultNumbering: ========================================")
    Logger.Info("TestVaultNumbering: Vault connection: OK")
    Logger.Info("TestVaultNumbering: Numbering schemes found: " & schemes.Length)
    Logger.Info("TestVaultNumbering: Generated number: " & generatedNumber)
    Logger.Info("TestVaultNumbering: ========================================")
    Logger.Info("TestVaultNumbering: All tests completed successfully!")
    
    MessageBox.Show("Vault numbrite genereerimine töötab!" & vbCrLf & vbCrLf & _
                    "Genereeritud number: " & generatedNumber, "TestVaultNumbering")
End Sub
