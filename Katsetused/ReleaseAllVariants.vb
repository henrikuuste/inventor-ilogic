' ============================================================================
' ReleaseAllVariants - Release All Variants from Excel Configuration
' 
' Creates independent copies for ALL variants defined in the Excel file.
' Each variant gets its own dated folder with updated references and parameters.
'
' Usage:
' 1. Open the master assembly in Inventor
' 2. Create an Excel file with variant configurations
' 3. Run this rule
' 4. All variants will be created in the /release folder
'
' Excel File Format:
' Place a file named "AssemblyName_Variants.xlsx" next to the assembly.
' | VariantName | PartNumber | Param1 | Param2 | ... |
' |-------------|------------|--------|--------|-----|
' | Variant1    | PN-001     | 100    | 200    | ... |
' | Variant2    | PN-002     | 150    | 250    | ... |
'
' Special columns (optional):
' | _SkipDrawings | Set to "Yes" to skip drawing copy for this variant
' | _Description  | Custom description for iProperty
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
        MessageBox.Show("No active document. Please open the master assembly.", "Release All Variants")
        Exit Sub
    End If
    
    If doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MessageBox.Show("This rule only works with assembly documents (.iam)." & vbCrLf & _
                        "Please open the master assembly and run again.", "Release All Variants")
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
        MessageBox.Show("Error reading Excel file:" & vbCrLf & ex.Message, "Release All Variants")
        Exit Sub
    End Try
    
    If configs Is Nothing OrElse configs.Count = 0 Then
        MessageBox.Show("No configurations found in the Excel file.", "Release All Variants")
        Exit Sub
    End If
    
    ' Show confirmation dialog with list of configurations
    Dim configList As String = ""
    For Each cfg As ExcelReaderLib.ReleaseConfig In configs
        configList &= "  - " & cfg.ConfigName & " (" & cfg.PartNumber & ")" & vbCrLf
    Next
    
    Dim releaseFolder As String = System.IO.Path.Combine(projectRoot, "r")
    
    Dim confirmMsg As String = "Release ALL configurations (" & configs.Count & " total):" & vbCrLf & vbCrLf & _
                               configList & vbCrLf & _
                               "Release folder: " & releaseFolder & vbCrLf & vbCrLf & _
                               "This may take several minutes. Continue?"
    
    If MessageBox.Show(confirmMsg, "Release All Variants - Confirm", MessageBoxButtons.YesNo) <> DialogResult.Yes Then
        Exit Sub
    End If
    
    ' Validate parameters for all configurations
    Dim allMissingParams As New HashSet(Of String)
    For Each cfg As ExcelReaderLib.ReleaseConfig In configs
        Dim missing As List(Of String) = ExcelReaderLib.ValidateParameters(asmDoc, cfg)
        For Each p As String In missing
            allMissingParams.Add(p)
        Next
    Next
    
    If allMissingParams.Count > 0 Then
        Dim msg As String = "The following parameters were not found in the assembly:" & vbCrLf & vbCrLf
        For Each p As String In allMissingParams
            msg &= "  - " & p & vbCrLf
        Next
        msg &= vbCrLf & "Continue anyway? (Missing parameters will be skipped)"
        
        If MessageBox.Show(msg, "Release All Variants - Missing Parameters", _
                           MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then
            Exit Sub
        End If
    End If
    
    ' Save and close the master document
    If asmDoc.Dirty Then
        Dim saveResult As DialogResult = MessageBox.Show( _
            "The assembly has unsaved changes. Save before releasing?", _
            "Release All Variants", MessageBoxButtons.YesNoCancel)
        
        If saveResult = DialogResult.Cancel Then
            Exit Sub
        ElseIf saveResult = DialogResult.Yes Then
            asmDoc.Save()
        End If
    End If
    
    asmDoc.Close(True)
    
    ' Process each configuration
    Dim successCount As Integer = 0
    Dim failCount As Integer = 0
    Dim allLogs As New List(Of String)
    
    allLogs.Add("========================================")
    allLogs.Add("Batch Release Started: " & DateTime.Now.ToString())
    allLogs.Add("Master Assembly: " & asmPath)
    allLogs.Add("Total Configurations: " & configs.Count)
    allLogs.Add("========================================")
    allLogs.Add("")
    
    For i As Integer = 0 To configs.Count - 1
        Dim cfg As ExcelReaderLib.ReleaseConfig = configs(i)
        
        allLogs.Add("----------------------------------------")
        allLogs.Add("Configuration " & (i + 1) & " of " & configs.Count & ": " & cfg.ConfigName)
        allLogs.Add("----------------------------------------")
        
        ' Check if drawings should be skipped
        Dim includeDrawings As Boolean = True
        If cfg.Parameters.ContainsKey("_SkipDrawings") Then
            Dim skipValue As String = cfg.Parameters("_SkipDrawings")
            If skipValue.Equals("Yes", StringComparison.OrdinalIgnoreCase) OrElse _
               skipValue.Equals("True", StringComparison.OrdinalIgnoreCase) OrElse _
               skipValue.Equals("1", StringComparison.OrdinalIgnoreCase) Then
                includeDrawings = False
            End If
        End If
        
        ' Release this configuration
        Dim logMessages As New List(Of String)
        
        ' Get iLogic automation for running rules
        Dim iLogicAuto As Object = Nothing
        Try
            iLogicAuto = iLogicVb.Automation
        Catch
        End Try
        
        Dim releasedPath As String = VariantReleaseLib.ReleaseVariant( _
            app, _
            asmPath, _
            cfg.ConfigName, _
            cfg.PartNumber, _
            cfg.Parameters, _
            releaseFolder, _
            includeDrawings, _
            logMessages, _
            iLogicAuto)
        
        ' Add configuration logs to the batch log and iLogic Logger
        For Each msg As String In logMessages
            allLogs.Add("  " & msg)
            ' Output to iLogic Logger
            If msg.StartsWith("ERROR") OrElse msg.Contains("ERROR:") Then
                Logger.Error(msg)
            ElseIf msg.StartsWith("Warning") OrElse msg.Contains("Warning:") Then
                Logger.Warn(msg)
            Else
                Logger.Info(msg)
            End If
        Next
        
        If Not String.IsNullOrEmpty(releasedPath) Then
            successCount += 1
            allLogs.Add("  => SUCCESS: " & releasedPath)
            Logger.Info("=> SUCCESS: " & releasedPath)
            
            ' Write individual log file
            Dim variantFolder As String = System.IO.Path.GetDirectoryName(releasedPath)
            VariantReleaseLib.WriteLogFile(variantFolder, cfg.ConfigName, logMessages)
        Else
            failCount += 1
            allLogs.Add("  => FAILED")
            Logger.Error("=> FAILED: " & cfg.ConfigName)
        End If
        
        allLogs.Add("")
    Next
    
    ' Write batch log
    allLogs.Add("========================================")
    allLogs.Add("Batch Release Complete: " & DateTime.Now.ToString())
    allLogs.Add("Successful: " & successCount)
    allLogs.Add("Failed: " & failCount)
    allLogs.Add("========================================")
    
    ' Save batch log
    Try
        Dim batchLogPath As String = System.IO.Path.Combine(releaseFolder, "BatchRelease_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".log")
        If Not System.IO.Directory.Exists(releaseFolder) Then
            System.IO.Directory.CreateDirectory(releaseFolder)
        End If
        System.IO.File.WriteAllText(batchLogPath, String.Join(vbCrLf, allLogs.ToArray()))
    Catch
    End Try
    
    ' Show summary
    Dim summaryMsg As String = "Batch release complete!" & vbCrLf & vbCrLf & _
                               "Successful: " & successCount & " of " & configs.Count & vbCrLf & _
                               "Failed: " & failCount & vbCrLf & vbCrLf & _
                               "Release folder: " & releaseFolder & vbCrLf & vbCrLf & _
                               "Open release folder?"
    
    If MessageBox.Show(summaryMsg, "Release All Variants - Complete", MessageBoxButtons.YesNo) = DialogResult.Yes Then
        Process.Start("explorer.exe", releaseFolder)
    End If
End Sub
