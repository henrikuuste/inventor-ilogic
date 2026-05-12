' Copyright (c) 2026 Henri Kuuste
' Test11_UserInfo.vb
' PURPOSE: Explore what user data is available through Inventor API
' 
' TESTS:
' 1. Inventor Application user info
' 2. Autodesk account / cloud user
' 3. Vault user (if logged in)
' 4. Windows/system user info
' 5. Document author/creator info
' 6. Any other user-related properties
'
' RUN: Open any document (or none), then run this rule

AddVbFile "Lib/StringsLib.vb"
AddVbFile "Lib/UtilsLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    
    Logger.Info("=== Test11_UserInfo: Starting ===")
    Logger.Info("Purpose: List all available user data through Inventor API")
    Logger.Info("")
    
    ' === 1. INVENTOR APPLICATION PROPERTIES ===
    Logger.Info("--- 1. Inventor Application Properties ---")
    
    Try
        Logger.Info("  UserName: " & app.UserName)
    Catch ex As Exception
        Logger.Error("  UserName: FAILED - " & ex.Message)
    End Try
    
    Try
        ' Try getting user initials if available
        Dim initials As String = GetPropertySafe(app, "UserInitials")
        Logger.Info("  UserInitials: " & initials)
    Catch ex As Exception
        Logger.Error("  UserInitials: FAILED - " & ex.Message)
    End Try
    
    ' === 2. GENERAL OPTIONS (user settings) ===
    Logger.Info("")
    Logger.Info("--- 2. GeneralOptions (User Settings) ---")
    
    Try
        Dim genOpts As GeneralOptions = app.GeneralOptions
        
        Try : Logger.Info("  UserName: " & genOpts.UserName) : Catch ex As Exception : Logger.Warn("  UserName: " & ex.Message) : End Try
        Try : Logger.Info("  UserType: " & genOpts.UserType.ToString()) : Catch ex As Exception : Logger.Warn("  UserType: " & ex.Message) : End Try
        Try : Logger.Info("  StartupProject: " & genOpts.StartupProject) : Catch ex As Exception : Logger.Warn("  StartupProject: " & ex.Message) : End Try
        Try : Logger.Info("  TemplatesPath: " & genOpts.TemplatesPath) : Catch ex As Exception : Logger.Warn("  TemplatesPath: " & ex.Message) : End Try
        Try : Logger.Info("  DesignDataPath: " & genOpts.DesignDataPath) : Catch ex As Exception : Logger.Warn("  DesignDataPath: " & ex.Message) : End Try
        Try : Logger.Info("  DefaultVBAProject: " & genOpts.DefaultVBAProjectFullFileName) : Catch ex As Exception : Logger.Warn("  DefaultVBAProject: " & ex.Message) : End Try
    Catch ex As Exception
        Logger.Error("  GeneralOptions: FAILED - " & ex.Message)
    End Try
    
    ' === 3. APPLICATION IDENTITY ===
    Logger.Info("")
    Logger.Info("--- 3. Application Identity ---")
    
    Try
        Logger.Info("  ProductVersion: " & app.SoftwareVersion.ProductVersion)
    Catch ex As Exception : Logger.Warn("  ProductVersion: " & ex.Message) : End Try
    
    Try
        Logger.Info("  DisplayVersion: " & app.SoftwareVersion.DisplayVersion)
    Catch ex As Exception : Logger.Warn("  DisplayVersion: " & ex.Message) : End Try
    
    Try
        Logger.Info("  ReleaseType: " & app.SoftwareVersion.ReleaseType.ToString())
    Catch ex As Exception : Logger.Warn("  ReleaseType: " & ex.Message) : End Try
    
    Try
        Logger.Info("  ServicePack: " & app.SoftwareVersion.ServicePack.ToString())
    Catch ex As Exception : Logger.Warn("  ServicePack: " & ex.Message) : End Try
    
    Try
        Logger.Info("  Locale: " & app.Locale.ToString())
    Catch ex As Exception : Logger.Warn("  Locale: " & ex.Message) : End Try
    
    Try
        Logger.Info("  Language: " & app.LanguageName)
    Catch ex As Exception : Logger.Warn("  Language: " & ex.Message) : End Try
    
    ' === 4. AUTODESK ACCOUNT / CLOUD USER ===
    Logger.Info("")
    Logger.Info("--- 4. Autodesk Account / Cloud User ---")
    
    ' Try WebServicesManager for cloud login info
    Try
        Dim wsm As Object = Nothing
        Try
            wsm = app.WebServicesManager
            Logger.Info("  WebServicesManager: ACCESSIBLE")
        Catch
            Logger.Warn("  WebServicesManager: NOT ACCESSIBLE")
        End Try
        
        If wsm IsNot Nothing Then
            ' Try to get user info from WebServicesManager
            Try
                Dim userId As String = CallByName(wsm, "UserId", Microsoft.VisualBasic.CallType.Get, Nothing)
                Logger.Info("  UserId: " & userId)
            Catch ex As Exception
                Logger.Warn("  UserId: " & ex.Message)
            End Try
            
            Try
                Dim userName As String = CallByName(wsm, "UserName", Microsoft.VisualBasic.CallType.Get, Nothing)
                Logger.Info("  WSM.UserName: " & userName)
            Catch ex As Exception
                Logger.Warn("  WSM.UserName: " & ex.Message)
            End Try
            
            Try
                Dim isOnline As Boolean = CallByName(wsm, "IsOnline", Microsoft.VisualBasic.CallType.Get, Nothing)
                Logger.Info("  IsOnline: " & isOnline.ToString())
            Catch ex As Exception
                Logger.Warn("  IsOnline: " & ex.Message)
            End Try
            
            Try
                Dim connState As Object = CallByName(wsm, "ConnectionState", Microsoft.VisualBasic.CallType.Get, Nothing)
                Logger.Info("  ConnectionState: " & connState.ToString())
            Catch ex As Exception
                Logger.Warn("  ConnectionState: " & ex.Message)
            End Try
        End If
    Catch ex As Exception
        Logger.Error("  WebServicesManager block FAILED: " & ex.Message)
    End Try
    
    ' Try ApplicationAddIns for Autodesk account info
    Try
        Logger.Info("  Checking ApplicationAddIns for account info...")
        For Each addin As ApplicationAddIn In app.ApplicationAddIns
            Try
                If addin.DisplayName.ToLower().Contains("autodesk") OrElse _
                   addin.DisplayName.ToLower().Contains("account") OrElse _
                   addin.DisplayName.ToLower().Contains("cloud") Then
                    Logger.Info("    AddIn: " & addin.DisplayName & " (Active: " & addin.Activated.ToString() & ")")
                End If
            Catch : End Try
        Next
    Catch ex As Exception
        Logger.Warn("  ApplicationAddIns: " & ex.Message)
    End Try
    
    ' === 5. VAULT USER (if Vault add-in is active) ===
    Logger.Info("")
    Logger.Info("--- 5. Vault User ---")
    
    Try
        Dim vaultAddin As ApplicationAddIn = Nothing
        
        ' Find Vault add-in
        For Each addin As ApplicationAddIn In app.ApplicationAddIns
            If addin.DisplayName.ToLower().Contains("vault") Then
                Logger.Info("  Found: " & addin.DisplayName)
                If addin.Activated Then
                    vaultAddin = addin
                    Logger.Info("  Status: ACTIVATED")
                Else
                    Logger.Info("  Status: Not activated")
                End If
            End If
        Next
        
        If vaultAddin IsNot Nothing Then
            ' Try to get Vault connection info through automation
            Try
                Dim vaultAuto As Object = vaultAddin.Automation
                If vaultAuto IsNot Nothing Then
                    Logger.Info("  Vault.Automation: ACCESSIBLE")
                    
                    ' Try common property names for user info
                    TryGetVaultProperty(vaultAuto, "UserName")
                    TryGetVaultProperty(vaultAuto, "CurrentUser")
                    TryGetVaultProperty(vaultAuto, "User")
                    TryGetVaultProperty(vaultAuto, "LoggedInUser")
                    TryGetVaultProperty(vaultAuto, "LoginName")
                    TryGetVaultProperty(vaultAuto, "Server")
                    TryGetVaultProperty(vaultAuto, "VaultName")
                    TryGetVaultProperty(vaultAuto, "Database")
                    TryGetVaultProperty(vaultAuto, "IsLoggedIn")
                    TryGetVaultProperty(vaultAuto, "ConnectionState")
                    
                    ' Try getting Connection object
                    Try
                        Dim conn As Object = CallByName(vaultAuto, "Connection", Microsoft.VisualBasic.CallType.Get, Nothing)
                        If conn IsNot Nothing Then
                            Logger.Info("  Vault.Connection: ACCESSIBLE")
                            TryGetVaultProperty(conn, "UserName")
                            TryGetVaultProperty(conn, "Server")
                            TryGetVaultProperty(conn, "Vault")
                            TryGetVaultProperty(conn, "AuthenticationType")
                        End If
                    Catch : End Try
                Else
                    Logger.Warn("  Vault.Automation: NULL (not logged in?)")
                End If
            Catch ex As Exception
                Logger.Warn("  Vault.Automation: " & ex.Message)
            End Try
        Else
            Logger.Info("  No active Vault add-in found")
        End If
    Catch ex As Exception
        Logger.Error("  Vault section FAILED: " & ex.Message)
    End Try
    
    ' === 6. WINDOWS/SYSTEM USER ===
    Logger.Info("")
    Logger.Info("--- 6. Windows/System User ---")
    
    Try
        Logger.Info("  Environment.UserName: " & System.Environment.UserName)
    Catch ex As Exception : Logger.Warn("  Environment.UserName: " & ex.Message) : End Try
    
    Try
        Logger.Info("  Environment.UserDomainName: " & System.Environment.UserDomainName)
    Catch ex As Exception : Logger.Warn("  Environment.UserDomainName: " & ex.Message) : End Try
    
    Try
        Logger.Info("  Environment.MachineName: " & System.Environment.MachineName)
    Catch ex As Exception : Logger.Warn("  Environment.MachineName: " & ex.Message) : End Try
    
    Try
        Dim userProfilePath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile)
        Logger.Info("  UserProfilePath: " & userProfilePath)
    Catch ex As Exception : Logger.Warn("  UserProfilePath: " & ex.Message) : End Try
    
    ' === 7. DOCUMENT AUTHOR INFO (if document is open) ===
    Logger.Info("")
    Logger.Info("--- 7. Document Author/Creator Info ---")
    
    If app.ActiveDocument IsNot Nothing Then
        Dim doc As Document = app.ActiveDocument
        Logger.Info("  Active document: " & doc.DisplayName)
        
        ' Property sets
        Try
            Dim summaryInfo As PropertySet = doc.PropertySets.Item("Inventor Summary Information")
            
            For Each prop As Inventor.Property In summaryInfo
                Try
                    If prop.Name.ToLower().Contains("author") OrElse _
                       prop.Name.ToLower().Contains("creator") OrElse _
                       prop.Name.ToLower().Contains("user") OrElse _
                       prop.Name.ToLower().Contains("manager") Then
                        Dim val As String = If(prop.Value IsNot Nothing, prop.Value.ToString(), "<empty>")
                        Logger.Info("  Summary." & prop.Name & ": " & val)
                    End If
                Catch : End Try
            Next
        Catch ex As Exception
            Logger.Warn("  Summary Information: " & ex.Message)
        End Try
        
        Try
            Dim docSummary As PropertySet = doc.PropertySets.Item("Inventor Document Summary Information")
            
            For Each prop As Inventor.Property In docSummary
                Try
                    If prop.Name.ToLower().Contains("author") OrElse _
                       prop.Name.ToLower().Contains("creator") OrElse _
                       prop.Name.ToLower().Contains("user") OrElse _
                       prop.Name.ToLower().Contains("manager") OrElse _
                       prop.Name.ToLower().Contains("company") Then
                        Dim val As String = If(prop.Value IsNot Nothing, prop.Value.ToString(), "<empty>")
                        Logger.Info("  DocSummary." & prop.Name & ": " & val)
                    End If
                Catch : End Try
            Next
        Catch ex As Exception
            Logger.Warn("  Document Summary: " & ex.Message)
        End Try
        
        ' List ALL properties for reference
        Logger.Info("")
        Logger.Info("  --- All Summary Properties ---")
        Try
            For Each propSet As PropertySet In doc.PropertySets
                For Each prop As Inventor.Property In propSet
                    Try
                        Dim val As String = ""
                        If prop.Value IsNot Nothing Then
                            val = prop.Value.ToString()
                            If val.Length > 50 Then val = val.Substring(0, 50) & "..."
                        Else
                            val = "<null>"
                        End If
                        Logger.Info("    [" & propSet.Name & "] " & prop.Name & " = " & val)
                    Catch : End Try
                Next
            Next
        Catch ex As Exception
            Logger.Warn("  PropertySets enumeration: " & ex.Message)
        End Try
        
        ' File info
        Logger.Info("")
        Logger.Info("  --- File System Info ---")
        Try
            Dim fileInfo As New System.IO.FileInfo(doc.FullFileName)
            Logger.Info("  File.CreationTime: " & fileInfo.CreationTime.ToString())
            Logger.Info("  File.LastWriteTime: " & fileInfo.LastWriteTime.ToString())
            Logger.Info("  File.LastAccessTime: " & fileInfo.LastAccessTime.ToString())
        Catch ex As Exception
            Logger.Warn("  FileInfo: " & ex.Message)
        End Try
    Else
        Logger.Info("  No document open - skipping document-specific info")
    End If
    
    ' === 8. AUTODESK IDENTITY MANAGER (config files) ===
    Logger.Info("")
    Logger.Info("--- 8. Autodesk Identity Manager Config Files ---")
    
    Try
        Dim localAppData As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData)
        Logger.Info("  LocalAppData: " & localAppData)
        
        ' Check Web Services folder (Autodesk Identity Manager stores login info here)
        Dim webServicesPath As String = System.IO.Path.Combine(localAppData, "Autodesk", "Web Services")
        Logger.Info("  Checking: " & webServicesPath)
        
        If System.IO.Directory.Exists(webServicesPath) Then
            Logger.Info("  Web Services folder EXISTS")
            
            ' List files in Web Services folder
            For Each filePath As String In System.IO.Directory.GetFiles(webServicesPath, "*.*", System.IO.SearchOption.AllDirectories)
                Dim fileName As String = System.IO.Path.GetFileName(filePath)
                Dim fileSize As Long = New System.IO.FileInfo(filePath).Length
                Logger.Info("    File: " & fileName & " (" & fileSize & " bytes)")
                
                ' Try to read JSON/XML config files that might contain user info
                If fileName.EndsWith(".json") OrElse fileName.EndsWith(".xml") OrElse fileName.EndsWith(".config") Then
                    Try
                        Dim content As String = System.IO.File.ReadAllText(filePath)
                        ' Look for email patterns (Autodesk ID is typically email)
                        If content.Contains("@") OrElse content.ToLower().Contains("user") OrElse content.ToLower().Contains("email") Then
                            ' Truncate for display
                            If content.Length > 500 Then content = content.Substring(0, 500) & "..."
                            Logger.Info("      Content preview: " & content.Replace(vbCrLf, " ").Replace(vbLf, " "))
                        End If
                    Catch : End Try
                End If
            Next
        Else
            Logger.Info("  Web Services folder NOT FOUND")
        End If
        
        ' Check Autodesk Identity folder
        Dim identityPath As String = System.IO.Path.Combine(localAppData, "Autodesk", "Identity")
        If System.IO.Directory.Exists(identityPath) Then
            Logger.Info("  Identity folder EXISTS: " & identityPath)
            For Each filePath As String In System.IO.Directory.GetFiles(identityPath, "*.*", System.IO.SearchOption.AllDirectories)
                Logger.Info("    File: " & System.IO.Path.GetFileName(filePath))
            Next
        End If
        
        ' Check AdskLicensing folder (modern licensing)
        Dim licensingPath As String = System.IO.Path.Combine(localAppData, "Autodesk", "AdskLicensing")
        If System.IO.Directory.Exists(licensingPath) Then
            Logger.Info("  AdskLicensing folder EXISTS: " & licensingPath)
            For Each filePath As String In System.IO.Directory.GetFiles(licensingPath, "*.*", System.IO.SearchOption.TopDirectoryOnly)
                Logger.Info("    File: " & System.IO.Path.GetFileName(filePath))
            Next
        End If
        
        ' Check common Autodesk subfolders
        Dim autodeskPath As String = System.IO.Path.Combine(localAppData, "Autodesk")
        If System.IO.Directory.Exists(autodeskPath) Then
            Logger.Info("  Autodesk subfolders:")
            For Each dirPath As String In System.IO.Directory.GetDirectories(autodeskPath)
                Dim dirName As String = System.IO.Path.GetFileName(dirPath)
                Logger.Info("    " & dirName)
            Next
        End If
        
    Catch ex As Exception
        Logger.Error("  Config files section FAILED: " & ex.Message)
    End Try
    
    ' === 9. LICENSING INFO ===
    Logger.Info("")
    Logger.Info("--- 9. Licensing Information ---")
    
    Try
        Logger.Info("  SoftwareVersion.IsPrerelease: " & app.SoftwareVersion.IsPrerelease.ToString())
    Catch ex As Exception : Logger.Warn("  IsPrerelease: " & ex.Message) : End Try
    
    Try
        ' Try getting license info through reflection
        Dim licInfo As Object = Nothing
        Try
            licInfo = CallByName(app, "LicenseInfo", Microsoft.VisualBasic.CallType.Get, Nothing)
            If licInfo IsNot Nothing Then
                Logger.Info("  LicenseInfo: ACCESSIBLE")
                TryGetObjectProperty(licInfo, "LicenseType")
                TryGetObjectProperty(licInfo, "ProductKey")
                TryGetObjectProperty(licInfo, "SerialNumber")
                TryGetObjectProperty(licInfo, "UserName")
            End If
        Catch ex As Exception
            Logger.Warn("  LicenseInfo: " & ex.Message)
        End Try
    Catch ex As Exception
        Logger.Warn("  License section: " & ex.Message)
    End Try
    
    ' === 10. ENVIRONMENT VARIABLES (related to Autodesk) ===
    Logger.Info("")
    Logger.Info("--- 10. Relevant Environment Variables ---")
    
    Dim envVars() As String = {
        "ADSK_3DSMAX_x64_2024",
        "AUTODESK_ADLM_THINCLIENT_ENV",
        "INVENTOR_LOCALE",
        "INVENTOR_ADDINS",
        "USERNAME",
        "USERDOMAIN",
        "COMPUTERNAME"
    }
    
    For Each envName As String In envVars
        Try
            Dim val As String = System.Environment.GetEnvironmentVariable(envName)
            If Not String.IsNullOrEmpty(val) Then
                Logger.Info("  " & envName & " = " & val)
            End If
        Catch : End Try
    Next
    
    ' Search for any Autodesk-related env vars
    Try
        Dim envDict As System.Collections.IDictionary = System.Environment.GetEnvironmentVariables()
        For Each key As Object In envDict.Keys
            Dim keyStr As String = key.ToString()
            If keyStr.ToUpper().Contains("AUTODESK") OrElse _
               keyStr.ToUpper().Contains("INVENTOR") OrElse _
               keyStr.ToUpper().Contains("VAULT") OrElse _
               keyStr.ToUpper().Contains("ADSK") Then
                Logger.Info("  " & keyStr & " = " & envDict(key).ToString())
            End If
        Next
    Catch ex As Exception
        Logger.Warn("  Environment variables scan: " & ex.Message)
    End Try
    
    Logger.Info("")
    Logger.Info("=== Test11_UserInfo: Complete ===")
End Sub

' Helper: Try to get a property value safely
Function GetPropertySafe(obj As Object, propName As String) As String
    Try
        Dim val As Object = CallByName(obj, propName, Microsoft.VisualBasic.CallType.Get, Nothing)
        Return If(val IsNot Nothing, val.ToString(), "<null>")
    Catch
        Return "<not available>"
    End Try
End Function

' Helper: Try to get and log a property from Vault automation object
Sub TryGetVaultProperty(obj As Object, propName As String)
    Try
        Dim val As Object = CallByName(obj, propName, Microsoft.VisualBasic.CallType.Get, Nothing)
        Logger.Info("    " & propName & ": " & If(val IsNot Nothing, val.ToString(), "<null>"))
    Catch ex As Exception
        ' Silently skip - property doesn't exist
    End Try
End Sub

' Helper: Try to get and log a property from any object
Sub TryGetObjectProperty(obj As Object, propName As String)
    Try
        Dim val As Object = CallByName(obj, propName, Microsoft.VisualBasic.CallType.Get, Nothing)
        Logger.Info("    " & propName & ": " & If(val IsNot Nothing, val.ToString(), "<null>"))
    Catch ex As Exception
        ' Silently skip
    End Try
End Sub
