' ============================================================================
' ExcelReaderLib - Excel File Reading for Variant Configuration
' 
' Provides functions to read variant tables from Excel files.
' Uses Excel COM Interop (requires Excel to be installed).
'
' Usage: AddVbFile "Lib/ExcelReaderLib.vb"
'
' Excel Table Format:
' | VariantName | PartNumber | Param1 | Param2 | ... |
' |-------------|------------|--------|--------|-----|
' | Variant1    | PN-001     | 100    | 200    | ... |
' | Variant2    | PN-002     | 150    | 250    | ... |
'
' First row is headers (parameter names)
' First column is variant name (used for folder/file naming)
' Second column is part number for the assembly
' Remaining columns are parameter names and values
'
' Ref: https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=110f3019-404c-4fc4-8b5d-7a3143f129da
' ============================================================================

Imports System.Collections.Generic

Public Module ExcelReaderLib

    ' ============================================================================
    ' SECTION 1: Data Structures
    ' ============================================================================

    ''' <summary>
    ''' Represents a single release configuration read from Excel.
    ''' </summary>
    Public Class ReleaseConfig
        Public Property ConfigName As String = ""
        Public Property PartNumber As String = ""
        Public Property Parameters As Dictionary(Of String, String) = New Dictionary(Of String, String)
        
        ''' <summary>
        ''' Get a parameter value, or default if not found.
        ''' </summary>
        Public Function GetParameter(paramName As String, Optional defaultValue As String = "") As String
            If Parameters.ContainsKey(paramName) Then
                Return Parameters(paramName)
            End If
            Return defaultValue
        End Function
        
        ''' <summary>
        ''' Get a parameter value as a Double.
        ''' </summary>
        Public Function GetParameterAsDouble(paramName As String, Optional defaultValue As Double = 0) As Double
            If Parameters.ContainsKey(paramName) Then
                Dim result As Double
                If Double.TryParse(Parameters(paramName), result) Then
                    Return result
                End If
            End If
            Return defaultValue
        End Function
    End Class

    ' ============================================================================
    ' SECTION 2: Excel Reading Functions
    ' ============================================================================

    ''' <summary>
    ''' Read all configurations from an Excel file.
    ''' Returns a list of ReleaseConfig objects.
    ''' </summary>
    ''' <param name="excelPath">Full path to the Excel file</param>
    ''' <param name="sheetName">Optional sheet name (defaults to first sheet)</param>
    ''' <returns>List of release configurations</returns>
    Public Function ReadVariantTable(excelPath As String, Optional sheetName As String = "") As List(Of ReleaseConfig)
        Dim configs As New List(Of ReleaseConfig)
        
        If Not System.IO.File.Exists(excelPath) Then
            Throw New System.IO.FileNotFoundException("Excel file not found: " & excelPath)
        End If
        
        Dim excelApp As Object = Nothing
        Dim workbook As Object = Nothing
        Dim worksheet As Object = Nothing
        
        Try
            ' Create Excel application instance
            excelApp = CreateObject("Excel.Application")
            excelApp.Visible = False
            excelApp.DisplayAlerts = False
            
            ' Open workbook
            workbook = excelApp.Workbooks.Open(excelPath, ReadOnly:=True)
            
            ' Get worksheet
            If sheetName = "" Then
                worksheet = workbook.Sheets(1)
            Else
                worksheet = workbook.Sheets(sheetName)
            End If
            
            ' Find the used range
            Dim usedRange As Object = worksheet.UsedRange
            Dim rowCount As Integer = usedRange.Rows.Count
            Dim colCount As Integer = usedRange.Columns.Count
            
            If rowCount < 2 OrElse colCount < 2 Then
                Throw New Exception("Excel table must have at least 2 rows (header + data) and 2 columns (VariantName + PartNumber)")
            End If
            
            ' Read header row (row 1)
            Dim headers As New List(Of String)
            For col As Integer = 1 To colCount
                Dim cellValue As Object = worksheet.Cells(1, col).Value
                If cellValue IsNot Nothing Then
                    headers.Add(cellValue.ToString().Trim())
                Else
                    headers.Add("")
                End If
            Next
            
            ' Read data rows (starting from row 2)
            For row As Integer = 2 To rowCount
                Dim cfg As New ReleaseConfig()
                
                ' First column is VariantName (stored as ConfigName)
                Dim configName As Object = worksheet.Cells(row, 1).Value
                If configName Is Nothing OrElse String.IsNullOrWhiteSpace(configName.ToString()) Then
                    Continue For ' Skip empty rows
                End If
                cfg.ConfigName = configName.ToString().Trim()
                
                ' Second column is PartNumber
                Dim partNum As Object = worksheet.Cells(row, 2).Value
                If partNum IsNot Nothing Then
                    cfg.PartNumber = partNum.ToString().Trim()
                End If
                
                ' Remaining columns are parameters
                For col As Integer = 3 To colCount
                    Dim headerName As String = If(col <= headers.Count, headers(col - 1), "")
                    If String.IsNullOrWhiteSpace(headerName) Then Continue For
                    
                    ' Store all columns (including special ones starting with underscore)
                    Dim cellValue As Object = worksheet.Cells(row, col).Value
                    If cellValue IsNot Nothing Then
                        cfg.Parameters(headerName) = cellValue.ToString().Trim()
                    End If
                Next
                
                configs.Add(cfg)
            Next
            
        Finally
            ' Clean up Excel objects
            If worksheet IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet)
            End If
            If workbook IsNot Nothing Then
                workbook.Close(False)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            End If
            If excelApp IsNot Nothing Then
                excelApp.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
            End If
        End Try
        
        Return configs
    End Function

    ''' <summary>
    ''' Get a specific configuration by name.
    ''' </summary>
    Public Function GetConfigByName(configs As List(Of ReleaseConfig), configName As String) As ReleaseConfig
        For Each cfg As ReleaseConfig In configs
            If cfg.ConfigName.Equals(configName, StringComparison.OrdinalIgnoreCase) Then
                Return cfg
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Get list of configuration names from the list.
    ''' </summary>
    Public Function GetConfigNames(configs As List(Of ReleaseConfig)) As List(Of String)
        Dim names As New List(Of String)
        For Each cfg As ReleaseConfig In configs
            names.Add(cfg.ConfigName)
        Next
        Return names
    End Function

    ' ============================================================================
    ' SECTION 3: Validation Functions
    ' ============================================================================

    ''' <summary>
    ''' Validate that all parameters in the configuration exist in the document.
    ''' Returns a list of missing parameter names.
    ''' </summary>
    Public Function ValidateParameters(doc As Inventor.Document, cfg As ReleaseConfig) As List(Of String)
        Dim missing As New List(Of String)
        
        Dim params As Inventor.Parameters = Nothing
        
        If doc.DocumentType = Inventor.DocumentTypeEnum.kAssemblyDocumentObject Then
            params = CType(doc, Inventor.AssemblyDocument).ComponentDefinition.Parameters
        ElseIf doc.DocumentType = Inventor.DocumentTypeEnum.kPartDocumentObject Then
            params = CType(doc, Inventor.PartDocument).ComponentDefinition.Parameters
        Else
            Return missing ' Can't validate non-model documents
        End If
        
        For Each kvp As KeyValuePair(Of String, String) In cfg.Parameters
            ' Skip special columns (start with underscore)
            If kvp.Key.StartsWith("_") Then Continue For
            
            ' Try to find the parameter
            Dim found As Boolean = False
            Try
                Dim param As Inventor.Parameter = params.Item(kvp.Key)
                found = True
            Catch
                found = False
            End Try
            
            If Not found Then
                missing.Add(kvp.Key)
            End If
        Next
        
        Return missing
    End Function

    ''' <summary>
    ''' Find the Excel file for variants, looking in standard locations.
    ''' </summary>
    Public Function FindVariantsExcelFile(asmDoc As Inventor.AssemblyDocument) As String
        Dim asmPath As String = asmDoc.FullFileName
        Dim asmFolder As String = System.IO.Path.GetDirectoryName(asmPath)
        Dim asmName As String = System.IO.Path.GetFileNameWithoutExtension(asmPath)
        
        ' Search patterns in priority order
        Dim searchPatterns As String() = {
            asmName & "_Variants.xlsx",
            asmName & "_Variants.xls",
            "Variants.xlsx",
            "Variants.xls"
        }
        
        ' Search in assembly folder and parent folder
        Dim searchFolders As String() = {
            asmFolder,
            System.IO.Path.GetDirectoryName(asmFolder)
        }
        
        For Each folder As String In searchFolders
            If String.IsNullOrEmpty(folder) Then Continue For
            For Each pattern As String In searchPatterns
                Dim filePath As String = System.IO.Path.Combine(folder, pattern)
                If System.IO.File.Exists(filePath) Then
                    Return filePath
                End If
            Next
        Next
        
        Return ""
    End Function

    ' ============================================================================
    ' SECTION 4: User Interface Helpers
    ' ============================================================================

    ''' <summary>
    ''' Show a selection dialog for choosing a configuration.
    ''' Returns the selected ReleaseConfig, or Nothing if cancelled.
    ''' </summary>
    Public Function ShowConfigSelectionDialog(configs As List(Of ReleaseConfig)) As ReleaseConfig
        If configs Is Nothing OrElse configs.Count = 0 Then
            Return Nothing
        End If
        
        Dim frm As New System.Windows.Forms.Form()
        frm.Text = "Select Configuration"
        frm.Width = 400
        frm.Height = 300
        frm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        frm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        frm.MaximizeBox = False
        frm.MinimizeBox = False
        
        Dim lbl As New System.Windows.Forms.Label()
        lbl.Text = "Select configuration to release:"
        lbl.Left = 10
        lbl.Top = 10
        lbl.Width = 360
        frm.Controls.Add(lbl)
        
        Dim lst As New System.Windows.Forms.ListBox()
        lst.Name = "lstConfigs"
        lst.Left = 10
        lst.Top = 35
        lst.Width = 360
        lst.Height = 180
        For Each cfg As ReleaseConfig In configs
            lst.Items.Add(cfg.ConfigName & " (" & cfg.PartNumber & ")")
        Next
        If lst.Items.Count > 0 Then
            lst.SelectedIndex = 0
        End If
        frm.Controls.Add(lst)
        
        Dim btnOK As New System.Windows.Forms.Button()
        btnOK.Text = "OK"
        btnOK.Left = 200
        btnOK.Top = 225
        btnOK.Width = 80
        btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        frm.Controls.Add(btnOK)
        
        Dim btnCancel As New System.Windows.Forms.Button()
        btnCancel.Text = "Cancel"
        btnCancel.Left = 290
        btnCancel.Top = 225
        btnCancel.Width = 80
        btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        frm.Controls.Add(btnCancel)
        
        frm.AcceptButton = btnOK
        frm.CancelButton = btnCancel
        
        Dim result As System.Windows.Forms.DialogResult = frm.ShowDialog()
        
        If result = System.Windows.Forms.DialogResult.OK AndAlso lst.SelectedIndex >= 0 Then
            Return configs(lst.SelectedIndex)
        End If
        
        Return Nothing
    End Function

End Module
