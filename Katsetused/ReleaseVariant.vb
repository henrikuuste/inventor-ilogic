' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' ReleaseVariant - Release a Single Variant from Excel Configuration
' 
' Creates an independent copy of an assembly with all its dependencies,
' updates all references to point to the copied files, applies parameter
' values, and updates iProperties.
'
' Usage:
' 1. Open the master assembly in Inventor
' 2. Create an Excel file with variant configurations (see format below)
' 3. Run this rule
' 4. Select a variant from the list
' 5. The variant will be created in the /release folder
'
' Excel File Format:
' Place a file named "AssemblyName_Variants.xlsx" next to the assembly.
' | VariantName | PartNumber | Param1 | Param2 | ... |
' |-------------|------------|--------|--------|-----|
' | Variant1    | PN-001     | 100    | 200    | ... |
'
' Ref: https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=ComponentOccurrence_Replace
' ============================================================================

AddVbFile "Lib/ExcelReaderLib.vb"
AddVbFile "Lib/VariantReleaseLib.vb"

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    ' Validate we have an assembly open
    If doc Is Nothing Then
        MessageBox.Show("No active document. Please open the master assembly.", "Release Variant")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MessageBox.Show("This rule only works with assembly documents (.iam)." & vbCrLf & _
                        "Please open the master assembly and run again.", "Release Variant")
        Exit Sub
    End If
    
    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
    Dim asmPath As String = asmDoc.FullFileName
    Dim asmFolder As String = System.IO.Path.GetDirectoryName(asmPath)
    Dim asmName As String = System.IO.Path.GetFileNameWithoutExtension(asmPath)
    
    ' Get project root from Inventor's design project
    Dim projectRoot As String = ""
    Try
        Dim projectPath As String = app.DesignProjectManager.ActiveDesignProject.FullFileName
        projectRoot = System.IO.Path.GetDirectoryName(projectPath)
    Catch
        ' Fallback: walk up to find a reasonable root
        projectRoot = System.IO.Path.GetDirectoryName(asmFolder)
    End Try
    
    ' Find the Excel variants file
    Dim excelPath As String = ExcelReaderLib.FindVariantsExcelFile(asmDoc)
    
    If String.IsNullOrEmpty(excelPath) Then
        ' Ask user to browse for Excel file
        Dim ofd As New System.Windows.Forms.OpenFileDialog()
        ofd.Title = "Select Variants Excel File"
        ofd.Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*"
        ofd.InitialDirectory = asmFolder
        
        If ofd.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
            Exit Sub
        End If
        excelPath = ofd.FileName
    End If
    
    ' Read configurations from Excel
    Dim configs As List(Of ExcelReaderLib.ReleaseConfig)
    Try
        configs = ExcelReaderLib.ReadVariantTable(excelPath)
    Catch ex As Exception
        MessageBox.Show("Error reading Excel file:" & vbCrLf & ex.Message, "Release Variant")
        Exit Sub
    End Try
    
    If configs Is Nothing OrElse configs.Count = 0 Then
        MessageBox.Show("No configurations found in the Excel file.", "Release Variant")
        Exit Sub
    End If
    
    ' Let user select a configuration
    Dim selectedConfig As ExcelReaderLib.ReleaseConfig = ExcelReaderLib.ShowConfigSelectionDialog(configs)
    
    If selectedConfig Is Nothing Then
        Exit Sub ' User cancelled
    End If
    
    ' Validate parameters exist in the document
    Dim missingParams As List(Of String) = ExcelReaderLib.ValidateParameters(asmDoc, selectedConfig)
    
    If missingParams.Count > 0 Then
        Dim msg As String = "The following parameters were not found in the assembly:" & vbCrLf & vbCrLf
        For Each p As String In missingParams
            msg &= "  - " & p & vbCrLf
        Next
        msg &= vbCrLf & "Continue anyway? (Parameters will be skipped)"
        
        If MessageBox.Show(msg, "Release Variant - Missing Parameters", _
                           MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then
            Exit Sub
        End If
    End If
    
    ' Determine release folder - project_root/r
    Dim releaseFolder As String = System.IO.Path.Combine(projectRoot, "r")
    
    ' Save if there are unsaved changes (silently)
    If asmDoc.Dirty Then
        asmDoc.Save()
    End If
    
    ' Perform the release (master assembly stays open)
    Dim logMessages As New List(Of String)
    
    Logger.Info("Starting release of: " & selectedConfig.ConfigName)
    
    ' Get iLogic automation for running rules
    Dim iLogicAuto As Object = Nothing
    Try
        iLogicAuto = iLogicVb.Automation
    Catch
    End Try
    
    Dim releasedPath As String = VariantReleaseLib.ReleaseVariant( _
        app, _
        asmPath, _
        selectedConfig.ConfigName, _
        selectedConfig.PartNumber, _
        selectedConfig.Parameters, _
        releaseFolder, _
        True, ' Include drawings
        logMessages, _
        iLogicAuto)
    
    ' Output log messages to iLogic Logger
    Logger.Info("=== Release Variant Log ===")
    For Each msg As String In logMessages
        If msg.StartsWith("ERROR") OrElse msg.Contains("ERROR:") Then
            Logger.Error(msg)
        ElseIf msg.StartsWith("Warning") OrElse msg.Contains("Warning:") Then
            Logger.Warn(msg)
        Else
            Logger.Info(msg)
        End If
    Next
    
    ' Write log file
    If Not String.IsNullOrEmpty(releasedPath) Then
        Dim variantFolder As String = System.IO.Path.GetDirectoryName(releasedPath)
        VariantReleaseLib.WriteLogFile(variantFolder, selectedConfig.ConfigName, logMessages)
    End If
    
    ' Reopen the master assembly if it was closed during release
    Dim masterStillOpen As Boolean = False
    For Each openDoc As Document In app.Documents
        If openDoc.FullFileName.Equals(asmPath, StringComparison.OrdinalIgnoreCase) Then
            masterStillOpen = True
            Exit For
        End If
    Next
    
    If Not masterStillOpen Then
        app.Documents.Open(asmPath, True)
    End If
    
    ' Show result
    If Not String.IsNullOrEmpty(releasedPath) Then
        Logger.Info("=== Release Complete: " & releasedPath & " ===")
        
        Dim successMsg As String = "Variant released successfully!" & vbCrLf & vbCrLf & _
                                   "Location: " & releasedPath & vbCrLf & vbCrLf & _
                                   "Open the released assembly now?"
        
        If MessageBox.Show(successMsg, "Release Variant - Complete", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            app.Documents.Open(releasedPath, True)
        End If
    Else
        Logger.Error("=== Release Failed ===")
        
        Dim errorMsg As String = "Release failed. See iLogic Log for details."
        MessageBox.Show(errorMsg, "Release Variant - Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End If
End Sub
