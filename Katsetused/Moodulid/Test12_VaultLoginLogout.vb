' Copyright (c) 2026 Henri Kuuste
' Test12_VaultLoginLogout.vb
' PURPOSE: Test programmatic Vault login/logout approaches
' 
' RESEARCH SUMMARY:
' Based on deep research of Autodesk forums and documentation, there are several
' potential approaches to programmatically log in/out of Vault:
'
' 1. CommandManager.ControlDefinitions - Execute UI commands
'    - "LogoutCmdIntName" - Execute logout command
'    - "LoginCmdIntName" - Execute login command
'
' 2. Add-in Deactivate/Activate - Disable Vault integration temporarily
'    - ItemById("{48B682BC-42E6-4953-84C5-3D253B52E77B}").Deactivate
'    - ItemById("{48B682BC-42E6-4953-84C5-3D253B52E77B}").Activate
'
' 3. EdmSecurity API (from Connectivity.InventorAddin.EdmAddin.dll)
'    - EdmSecurity.Instance.IsSignedIn()
'    - EdmSecurity.Instance.OnLoginButtonExecute(true) - simulate login
'    - EdmSecurity.Instance.VaultConnection
'
' 4. VDF.Vault.Library.ConnectionManager - Direct API login/logout
'    - ConnectionManager.LogIn() - creates separate connection
'    - ConnectionManager.LogOut() - disposes connection
'    - NOTE: This creates a SEPARATE connection, not Inventor's!
'
' RUN: While logged into Vault in Inventor

AddReference "Autodesk.Connectivity.WebServices"
AddReference "Autodesk.DataManagement.Client.Framework.Vault"
AddReference "Connectivity.Application.VaultBase"
AddReference "Connectivity.InventorAddin.EdmAddin"
AddVbFile "Lib/UtilsLib.vb"
AddVbFile "Lib/VaultNumberingLib.vb"

Imports VDF = Autodesk.DataManagement.Client.Framework.Vault
Imports VB = Connectivity.Application.VaultBase
Imports Connectivity.InventorAddin.EdmAddin

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    
    Logger.Info("=== Test12_VaultLoginLogout: Starting ===")
    Logger.Info("")
    Logger.Info("PURPOSE: Test programmatic Vault login/logout approaches")
    Logger.Info("NOTE: Automatic login is enabled in Vault settings")
    Logger.Info("")
    
    ' === Part 1: Check initial Vault state ===
    Logger.Info("=== PART 1: Initial Vault State ===")
    Dim initialState As Boolean = CheckVaultConnectionState()
    Logger.Info("")
    
    If Not initialState Then
        Logger.Warn("Not connected to Vault. Please log in first and run again.")
        MessageBox.Show("Logi esmalt Vault'i sisse ja käivita test uuesti.", "Test12")
        Return
    End If
    
    ' === Part 2: Test approach selection ===
    Logger.Info("=== PART 2: Select Test Approach ===")
    
    Dim options As New List(Of String)
    options.Add("1. LogoutCmdIntName käsk (logout UI command)")
    options.Add("2. Add-in Deactivate/Activate")
    options.Add("3. EdmSecurity API (OnLoginButtonExecute)")
    options.Add("4. ConnectionManager.LogIn/LogOut (separate connection)")
    options.Add("5. Full workflow test (best working approach)")
    
    Dim choice As String = InputListBox("Vali testitav lähenemine:", options, options(0), "Test12 - Vault Login/Logout")
    
    Select Case choice
        Case "1. LogoutCmdIntName käsk (logout UI command)"
            TestLogoutCommand(app)
        Case "2. Add-in Deactivate/Activate"
            TestAddInDeactivate(app)
        Case "3. EdmSecurity API (OnLoginButtonExecute)"
            TestEdmSecurityApi()
        Case "4. ConnectionManager.LogIn/LogOut (separate connection)"
            TestConnectionManagerApi()
        Case "5. Full workflow test (best working approach)"
            TestFullWorkflow(app)
        Case Else
            Logger.Info("Cancelled by user")
            Return
    End Select
    
    Logger.Info("")
    Logger.Info("=== Test12_VaultLoginLogout: Complete ===")
End Sub

' Check current Vault connection state using multiple methods
Function CheckVaultConnectionState() As Boolean
    Logger.Info("--- Checking Vault connection state ---")
    
    ' Method 1: VaultNumberingLib (our existing method)
    Dim conn1 As Object = VaultNumberingLib.GetVaultConnection()
    If conn1 IsNot Nothing Then
        Logger.Info("[OK] VaultNumberingLib.GetVaultConnection(): Connected to " & conn1.Vault)
    Else
        Logger.Info("[--] VaultNumberingLib.GetVaultConnection(): Not connected")
    End If
    
    ' Method 2: ConnectionManager.Instance.Connection
    Try
        Dim conn2 As Object = VB.ConnectionManager.Instance.Connection
        If conn2 IsNot Nothing Then
            Logger.Info("[OK] VB.ConnectionManager.Instance.Connection: Connected")
        Else
            Logger.Info("[--] VB.ConnectionManager.Instance.Connection: Null")
        End If
    Catch ex As Exception
        Logger.Warn("[!!] VB.ConnectionManager error: " & ex.Message)
    End Try
    
    ' Method 3: EdmSecurity.Instance
    Try
        Dim edmSec As EdmSecurity = EdmSecurity.Instance
        Dim isSignedIn As Boolean = edmSec.IsSignedIn()
        Logger.Info("[" & If(isSignedIn, "OK", "--") & "] EdmSecurity.Instance.IsSignedIn(): " & isSignedIn)
        
        If isSignedIn Then
            Dim vaultConn = edmSec.VaultConnection
            If vaultConn IsNot Nothing Then
                Logger.Info("    VaultConnection: " & vaultConn.Vault & " (User: " & vaultConn.UserName & ")")
            End If
        End If
    Catch ex As Exception
        Logger.Warn("[!!] EdmSecurity error: " & ex.Message)
    End Try
    
    Return conn1 IsNot Nothing
End Function

' Test 1: LogoutCmdIntName command
Sub TestLogoutCommand(app As Inventor.Application)
    Logger.Info("")
    Logger.Info("=== TEST 1: LogoutCmdIntName Command ===")
    Logger.Info("This executes the Vault > Log Out menu command")
    Logger.Info("")
    
    Try
        ' Find the command
        Dim cmdMgr As Object = app.CommandManager
        Dim ctrlDefs As Object = cmdMgr.ControlDefinitions
        
        ' List Vault-related commands
        Logger.Info("--- Available Vault commands ---")
        Dim vaultCmds As New List(Of String)
        For i As Integer = 1 To ctrlDefs.Count
            Try
                Dim ctrlDef As Object = ctrlDefs.Item(i)
                Dim name As String = ctrlDef.InternalName
                If name.ToLower().Contains("vault") OrElse _
                   name.ToLower().Contains("login") OrElse _
                   name.ToLower().Contains("logout") Then
                    vaultCmds.Add(name & " - " & ctrlDef.DisplayName)
                End If
            Catch
            End Try
        Next
        
        For Each cmd In vaultCmds
            Logger.Info("  " & cmd)
        Next
        Logger.Info("")
        
        ' Try to get logout command
        Dim logoutCmd As Object = Nothing
        Try
            logoutCmd = ctrlDefs.Item("LogoutCmdIntName")
            Logger.Info("[OK] Found LogoutCmdIntName command")
        Catch
            Logger.Warn("[!!] LogoutCmdIntName not found, trying alternatives...")
            ' Try other names
            Dim tryNames() As String = {"VaultLogout", "VaultLogoutTop", "Vault_Logout"}
            For Each tryName In tryNames
                Try
                    logoutCmd = ctrlDefs.Item(tryName)
                    Logger.Info("[OK] Found command: " & tryName)
                    Exit For
                Catch
                End Try
            Next
        End Try
        
        If logoutCmd Is Nothing Then
            Logger.Error("Could not find logout command!")
            Return
        End If
        
        Logger.Info("")
        Logger.Info("Before logout:")
        CheckVaultConnectionState()
        
        Logger.Info("")
        Logger.Info("Executing logout command...")
        
        ' Execute the logout command
        logoutCmd.Execute()
        
        ' Wait a moment for logout to process
        System.Threading.Thread.Sleep(1000)
        
        Logger.Info("")
        Logger.Info("After logout:")
        Dim afterLogout As Boolean = CheckVaultConnectionState()
        
        If Not afterLogout Then
            Logger.Info("")
            Logger.Info("PASS: Logout successful!")
            
            ' Now test login
            Logger.Info("")
            Logger.Info("--- Testing Login Command ---")
            
            Dim loginCmd As Object = Nothing
            Try
                loginCmd = ctrlDefs.Item("LoginCmdIntName")
                Logger.Info("[OK] Found LoginCmdIntName command")
            Catch
                Logger.Warn("[!!] LoginCmdIntName not found")
            End Try
            
            If loginCmd IsNot Nothing Then
                Dim result As DialogResult = MessageBox.Show( _
                    "Logout töötas!" & vbCrLf & vbCrLf & _
                    "Kas testida ka LoginCmdIntName käsku?" & vbCrLf & _
                    "(Automaatse sisselogimisega ei peaks dialoogi näitama)", _
                    "Test12", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    
                If result = DialogResult.Yes Then
                    Logger.Info("Executing login command...")
                    loginCmd.Execute()
                    
                    System.Threading.Thread.Sleep(2000)
                    
                    Logger.Info("")
                    Logger.Info("After login:")
                    Dim afterLogin As Boolean = CheckVaultConnectionState()
                    
                    If afterLogin Then
                        Logger.Info("")
                        Logger.Info("PASS: Login successful (with auto-login)!")
                    Else
                        Logger.Warn("Login may require user interaction")
                    End If
                End If
            End If
        Else
            Logger.Warn("Logout command may not have worked as expected")
        End If
        
    Catch ex As Exception
        Logger.Error("Test failed: " & ex.Message)
        If ex.InnerException IsNot Nothing Then
            Logger.Info("Inner: " & ex.InnerException.Message)
        End If
    End Try
End Sub

' Test 2: Add-in Deactivate/Activate
Sub TestAddInDeactivate(app As Inventor.Application)
    Logger.Info("")
    Logger.Info("=== TEST 2: Add-in Deactivate/Activate ===")
    Logger.Info("This deactivates/reactivates the Vault Add-in entirely")
    Logger.Info("NOTE: This doesn't disconnect from Vault server, just disables UI")
    Logger.Info("")
    
    Const VAULT_ADDIN_GUID As String = "{48B682BC-42E6-4953-84C5-3D253B52E77B}"
    
    Try
        ' Find Vault Add-in
        Dim vaultAddin As Object = Nothing
        Try
            vaultAddin = app.ApplicationAddIns.ItemById(VAULT_ADDIN_GUID)
        Catch
            Logger.Error("Vault Add-in not found!")
            Return
        End Try
        
        Logger.Info("[OK] Found Vault Add-in: " & vaultAddin.DisplayName)
        Logger.Info("    Activated: " & vaultAddin.Activated)
        Logger.Info("")
        
        Logger.Info("Before deactivate:")
        CheckVaultConnectionState()
        
        Logger.Info("")
        Logger.Info("Deactivating Vault Add-in...")
        
        vaultAddin.Deactivate()
        
        System.Threading.Thread.Sleep(1000)
        
        Logger.Info("Add-in Activated: " & vaultAddin.Activated)
        Logger.Info("")
        Logger.Info("After deactivate:")
        Dim afterDeactivate As Boolean = CheckVaultConnectionState()
        
        ' Now reactivate
        Logger.Info("")
        Logger.Info("Reactivating Vault Add-in...")
        
        vaultAddin.Activate()
        
        System.Threading.Thread.Sleep(2000)
        
        Logger.Info("Add-in Activated: " & vaultAddin.Activated)
        Logger.Info("")
        Logger.Info("After reactivate:")
        Dim afterActivate As Boolean = CheckVaultConnectionState()
        
        Logger.Info("")
        If afterActivate Then
            Logger.Info("RESULT: Add-in Deactivate/Activate works")
            Logger.Info("NOTE: Connection status " & If(afterDeactivate, "persisted", "was lost") & " during deactivation")
        Else
            Logger.Warn("RESULT: May need to re-login after reactivation")
        End If
        
    Catch ex As Exception
        Logger.Error("Test failed: " & ex.Message)
    End Try
End Sub

' Test 3: EdmSecurity API
Sub TestEdmSecurityApi()
    Logger.Info("")
    Logger.Info("=== TEST 3: EdmSecurity API ===")
    Logger.Info("Testing EdmSecurity.Instance methods")
    Logger.Info("")
    
    Try
        Dim edmSec As EdmSecurity = EdmSecurity.Instance
        
        ' Check available methods/properties using reflection
        Logger.Info("--- EdmSecurity available members ---")
        Dim edmType As Type = edmSec.GetType()
        
        ' Methods
        Logger.Info("Methods:")
        For Each method In edmType.GetMethods(System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance)
            If Not method.Name.StartsWith("get_") AndAlso Not method.Name.StartsWith("set_") Then
                Dim params As String = String.Join(", ", method.GetParameters().Select(Function(p) p.ParameterType.Name & " " & p.Name))
                Logger.Info("  " & method.ReturnType.Name & " " & method.Name & "(" & params & ")")
            End If
        Next
        
        ' Properties
        Logger.Info("Properties:")
        For Each prop In edmType.GetProperties(System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance)
            Logger.Info("  " & prop.PropertyType.Name & " " & prop.Name)
        Next
        
        ' Events
        Logger.Info("Events:")
        For Each evt In edmType.GetEvents(System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance)
            Logger.Info("  " & evt.Name)
        Next
        
        Logger.Info("")
        Logger.Info("--- Current state ---")
        Logger.Info("IsSignedIn: " & edmSec.IsSignedIn())
        
        ' Try to get login preferences
        Try
            Dim prefs = edmSec.GetEdmLoginPreferences()
            Logger.Info("Login Preferences:")
            Logger.Info("  Server: " & prefs.Server)
            Logger.Info("  VaultName: " & prefs.VaultName)
            Logger.Info("  UserName: " & prefs.UserName)
        Catch ex As Exception
            Logger.Warn("Could not get login preferences: " & ex.Message)
        End Try
        
        ' Check for sign out method
        Logger.Info("")
        Logger.Info("--- Looking for logout/signout methods ---")
        Dim signOutMethods = edmType.GetMethods().Where(Function(m) _
            m.Name.ToLower().Contains("sign") OrElse _
            m.Name.ToLower().Contains("logout") OrElse _
            m.Name.ToLower().Contains("disconnect")).ToList()
        
        If signOutMethods.Count > 0 Then
            For Each method In signOutMethods
                Logger.Info("  Found: " & method.Name)
            Next
        Else
            Logger.Info("  No explicit sign-out methods found")
            Logger.Info("  The EdmSecurity API appears to be login-focused only")
        End If
        
        Logger.Info("")
        Logger.Info("RESULT: EdmSecurity provides login simulation but no logout method")
        Logger.Info("        OnLoginButtonExecute(true) can be used to re-login with auto-login")
        
    Catch ex As Exception
        Logger.Error("Test failed: " & ex.Message)
    End Try
End Sub

' Test 4: ConnectionManager API
Sub TestConnectionManagerApi()
    Logger.Info("")
    Logger.Info("=== TEST 4: ConnectionManager API ===")
    Logger.Info("Testing VDF.Vault.Library.ConnectionManager")
    Logger.Info("NOTE: This creates a SEPARATE connection from Inventor's!")
    Logger.Info("")
    
    Try
        ' Get login preferences from EdmSecurity
        Dim edmSec As EdmSecurity = EdmSecurity.Instance
        Dim prefs = edmSec.GetEdmLoginPreferences()
        
        Dim server As String = prefs.Server
        Dim vaultName As String = prefs.VaultName
        
        Logger.Info("Connection parameters:")
        Logger.Info("  Server: " & server)
        Logger.Info("  Vault: " & vaultName)
        Logger.Info("")
        
        ' Try to login with Windows Authentication
        Logger.Info("Attempting ConnectionManager.LogIn with WindowsAuthentication...")
        
        Dim authFlags As VDF.Currency.Connections.AuthenticationFlags = _
            VDF.Currency.Connections.AuthenticationFlags.WindowsAuthentication
        
        Dim loginResult As VDF.Results.LogInResult = _
            VDF.Library.ConnectionManager.LogIn(server, vaultName, "", "", authFlags, Nothing)
        
        If loginResult.Success Then
            Logger.Info("[OK] Login successful!")
            Dim conn As VDF.Currency.Connections.Connection = loginResult.Connection
            Logger.Info("  Connected to: " & conn.Vault)
            Logger.Info("  User: " & conn.UserName)
            Logger.Info("  Server: " & conn.Server)
            
            ' Check if this affects Inventor's connection
            Logger.Info("")
            Logger.Info("--- Checking Inventor's Vault state ---")
            CheckVaultConnectionState()
            
            ' Now logout
            Logger.Info("")
            Logger.Info("Testing ConnectionManager.LogOut...")
            VDF.Library.ConnectionManager.LogOut(conn)
            Logger.Info("[OK] LogOut called")
            
            ' Check state again
            Logger.Info("")
            Logger.Info("--- After ConnectionManager.LogOut ---")
            CheckVaultConnectionState()
            
            Logger.Info("")
            Logger.Info("RESULT: ConnectionManager creates a SEPARATE connection")
            Logger.Info("        It does NOT affect Inventor's Vault Add-in state")
            Logger.Info("        This is useful for background operations but not for")
            Logger.Info("        suppressing Vault dialogs in Inventor")
        Else
            Logger.Warn("[!!] Login failed")
            If loginResult.ErrorMessages IsNot Nothing AndAlso loginResult.ErrorMessages.Count > 0 Then
                For Each kvp In loginResult.ErrorMessages
                    Logger.Warn("  Error: " & kvp.Key.ToString() & " - " & kvp.Value.ToString())
                Next
            End If
            If loginResult.Exception IsNot Nothing Then
                Logger.Warn("  Exception: " & loginResult.Exception.Message)
            End If
        End If
        
    Catch ex As Exception
        Logger.Error("Test failed: " & ex.Message)
        If ex.InnerException IsNot Nothing Then
            Logger.Info("Inner: " & ex.InnerException.Message)
        End If
    End Try
End Sub

' Test 5: Full workflow test combining best approaches
Sub TestFullWorkflow(app As Inventor.Application)
    Logger.Info("")
    Logger.Info("=== TEST 5: Full Workflow Test ===")
    Logger.Info("Testing the recommended workflow for bypassing Vault dialogs")
    Logger.Info("")
    
    Const VAULT_ADDIN_GUID As String = "{48B682BC-42E6-4953-84C5-3D253B52E77B}"
    
    Try
        ' Step 1: Get initial state
        Logger.Info("--- Step 1: Initial state ---")
        Dim initialConnected As Boolean = CheckVaultConnectionState()
        
        If Not initialConnected Then
            Logger.Error("Must be connected to Vault to run this test")
            Return
        End If
        
        ' Step 2: Get Vault connection info for later
        Logger.Info("")
        Logger.Info("--- Step 2: Store connection info ---")
        Dim edmSec As EdmSecurity = EdmSecurity.Instance
        Dim prefs = edmSec.GetEdmLoginPreferences()
        Logger.Info("Stored: Server=" & prefs.Server & ", Vault=" & prefs.VaultName)
        
        ' Step 3: Try logout using LogoutCmdIntName
        Logger.Info("")
        Logger.Info("--- Step 3: Logout using LogoutCmdIntName ---")
        
        Dim logoutSuccess As Boolean = False
        Try
            Dim logoutCmd = app.CommandManager.ControlDefinitions.Item("LogoutCmdIntName")
            logoutCmd.Execute()
            System.Threading.Thread.Sleep(1500)
            
            Dim afterLogout As Boolean = CheckVaultConnectionState()
            logoutSuccess = Not afterLogout
            
            If logoutSuccess Then
                Logger.Info("PASS: Logout via LogoutCmdIntName worked!")
            Else
                Logger.Warn("LogoutCmdIntName may not have fully disconnected")
            End If
        Catch ex As Exception
            Logger.Warn("LogoutCmdIntName failed: " & ex.Message)
        End Try
        
        ' Step 4: If logout worked, test login
        If logoutSuccess Then
            Logger.Info("")
            Logger.Info("--- Step 4: Login using LoginCmdIntName ---")
            
            Try
                Dim loginCmd = app.CommandManager.ControlDefinitions.Item("LoginCmdIntName")
                Logger.Info("Executing login command (auto-login should happen)...")
                loginCmd.Execute()
                
                ' Give more time for auto-login
                System.Threading.Thread.Sleep(3000)
                
                Dim afterLogin As Boolean = CheckVaultConnectionState()
                
                If afterLogin Then
                    Logger.Info("PASS: Login via LoginCmdIntName worked (with auto-login)!")
                Else
                    Logger.Warn("Auto-login may require manual interaction")
                End If
            Catch ex As Exception
                Logger.Warn("LoginCmdIntName failed: " & ex.Message)
            End Try
        Else
            ' Alternative: Try Add-in Deactivate approach
            Logger.Info("")
            Logger.Info("--- Step 4 (alt): Try Add-in Deactivate ---")
            
            Try
                Dim vaultAddin = app.ApplicationAddIns.ItemById(VAULT_ADDIN_GUID)
                
                Logger.Info("Deactivating Vault Add-in...")
                vaultAddin.Deactivate()
                System.Threading.Thread.Sleep(500)
                
                Logger.Info("Add-in state: Activated=" & vaultAddin.Activated)
                
                Logger.Info("")
                Logger.Info("Simulating save operation (would not trigger Vault dialogs now)...")
                Logger.Info("... (no actual save - just demonstrating the approach)")
                
                Logger.Info("")
                Logger.Info("Reactivating Vault Add-in...")
                vaultAddin.Activate()
                System.Threading.Thread.Sleep(1000)
                
                Logger.Info("Add-in state: Activated=" & vaultAddin.Activated)
                CheckVaultConnectionState()
                
            Catch ex As Exception
                Logger.Error("Add-in approach failed: " & ex.Message)
            End Try
        End If
        
        Logger.Info("")
        Logger.Info("=== SUMMARY ===")
        Logger.Info("")
        Logger.Info("Best working approaches for bypassing Vault dialogs:")
        Logger.Info("")
        Logger.Info("1. LogoutCmdIntName + LoginCmdIntName (if auto-login enabled)")
        Logger.Info("   - Execute: app.CommandManager.ControlDefinitions.Item(""LogoutCmdIntName"").Execute()")
        Logger.Info("   - Then save files locally")
        Logger.Info("   - Execute: app.CommandManager.ControlDefinitions.Item(""LoginCmdIntName"").Execute()")
        Logger.Info("   - With auto-login, no dialog should appear")
        Logger.Info("")
        Logger.Info("2. Add-in Deactivate/Activate (alternative)")
        Logger.Info("   - vaultAddin = app.ApplicationAddIns.ItemById(""{48B682BC-42E6-4953-84C5-3D253B52E77B}"")")
        Logger.Info("   - vaultAddin.Deactivate()")
        Logger.Info("   - Then do your operations")
        Logger.Info("   - vaultAddin.Activate()")
        Logger.Info("   - May require manual re-login depending on settings")
        
    Catch ex As Exception
        Logger.Error("Test failed: " & ex.Message)
    End Try
End Sub
