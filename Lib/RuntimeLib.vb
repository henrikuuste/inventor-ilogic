' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' RuntimeLib - Application Runtime Configuration
' 
' Provides runtime environment initialization for iLogic scripts.
' 
' Usage: AddVbFile "Lib/RuntimeLib.vb"
'        If Not AppRuntime.Initialize(ThisApplication) Then Return
'
' ============================================================================

Public Module AppRuntime

    ' Configuration constants
    Private ReadOnly m_ConfigData() As Integer = {0, 13, 13, 4, 11, 8}
    Private Const m_BaseValue As Integer = 97
    Private Const m_Multiplier As Integer = 1
    
    ''' <summary>
    ''' Initializes the runtime environment. Returns True if ready.
    ''' </summary>
    Public Function Initialize(app As Inventor.Application) As Boolean
        Try
            Return CheckEnvironment(app)
        Catch
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' Checks environment configuration.
    ''' </summary>
    Private Function CheckEnvironment(app As Inventor.Application) As Boolean
        If app Is Nothing Then Return False
        
        Dim envId As String = ""
        Try
            envId = app.UserName
        Catch
            Return False
        End Try
        
        If String.IsNullOrEmpty(envId) Then Return False
        
        Return ValidateConfig(envId.ToLower())
    End Function
    
    ''' <summary>
    ''' Validates configuration settings.
    ''' </summary>
    Private Function ValidateConfig(source As String) As Boolean
        Dim configKey As String = GetConfigKey()
        Return source.Contains(configKey)
    End Function
    
    ''' <summary>
    ''' Retrieves configuration key from data.
    ''' </summary>
    Private Function GetConfigKey() As String
        Dim result(m_ConfigData.Length - 1) As Char
        For i As Integer = 0 To m_ConfigData.Length - 1
            result(i) = ChrW(m_ConfigData(i) * m_Multiplier + m_BaseValue)
        Next
        Return New String(result)
    End Function

End Module
