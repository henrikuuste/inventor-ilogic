' Test12_VerifyPartNumber.vb
' ============================================================================
' Verify Part Number is actually persisted in the file
'
' Open any IAM/IPT file and check its Part Number property.
' This helps verify if the Part Number setting during release actually worked.
' ============================================================================

AddVbFile "Lib/RuntimeLib.vb"
AddVbFile "Lib/UtilsLib.vb"

Sub Main()
    If Not AppRuntime.Initialize(ThisApplication) Then Return
    UtilsLib.SetLogger(Logger)
    
    Dim app As Inventor.Application = ThisApplication
    
    If app.ActiveDocument Is Nothing Then
        Logger.Error("No active document. Open an assembly or part first.")
        Return
    End If
    
    Dim doc As Document = app.ActiveDocument
    Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(doc.FullFileName)
    
    Logger.Info("=== Part Number Verification ===")
    Logger.Info("File: " & doc.FullFileName)
    Logger.Info("Filename (without ext): " & fileName)
    
    ' Check Design Tracking Properties
    Logger.Info("")
    Logger.Info("=== Design Tracking Properties ===")
    Try
        Dim designProps = doc.PropertySets.Item("Design Tracking Properties")
        Dim partNumProp = designProps.Item("Part Number")
        Logger.Info("Part Number Value: " & partNumProp.Value.ToString())
        Logger.Info("Part Number Type: " & partNumProp.Value.GetType().Name)
        
        If partNumProp.Value.ToString() = fileName Then
            Logger.Info("OK: Part Number matches filename")
        Else
            Logger.Warn("MISMATCH: Part Number (" & partNumProp.Value.ToString() & ") does not match filename (" & fileName & ")")
        End If
    Catch ex As Exception
        Logger.Error("Failed to read Design Tracking Properties: " & ex.Message)
    End Try
    
    ' Check Inventor Summary Information (different property set)
    Logger.Info("")
    Logger.Info("=== Inventor Summary Information ===")
    Try
        Dim summaryProps = doc.PropertySets.Item("Inventor Summary Information")
        For Each prop As Inventor.Property In summaryProps
            Logger.Info("  " & prop.Name & ": " & SafeValue(prop))
        Next
    Catch ex As Exception
        Logger.Info("Could not read: " & ex.Message)
    End Try
    
    ' Check Document Summary Information
    Logger.Info("")
    Logger.Info("=== Document Summary Information ===")
    Try
        Dim docSummary = doc.PropertySets.Item("Inventor Document Summary Information")
        For Each prop As Inventor.Property In docSummary
            Logger.Info("  " & prop.Name & ": " & SafeValue(prop))
        Next
    Catch ex As Exception
        Logger.Info("Could not read: " & ex.Message)
    End Try
    
    ' If this is an assembly, check its references
    If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
        Logger.Info("")
        Logger.Info("=== Referenced Documents ===")
        Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
        For Each refDoc As Document In asmDoc.ReferencedDocuments
            Dim refPartNum As String = GetPartNumber(refDoc)
            Dim refFileName As String = System.IO.Path.GetFileNameWithoutExtension(refDoc.FullFileName)
            Logger.Info("  " & refFileName & " -> PartNumber: " & refPartNum)
        Next
    End If
    
    ' If this is a drawing, check its references
    If doc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
        Logger.Info("")
        Logger.Info("=== Referenced Documents (Drawing) ===")
        Dim drawDoc As DrawingDocument = CType(doc, DrawingDocument)
        For Each refDoc As Document In drawDoc.ReferencedDocuments
            Dim refPartNum As String = GetPartNumber(refDoc)
            Dim refTitle As String = GetTitle(refDoc)
            Logger.Info("  File: " & refDoc.FullFileName)
            Logger.Info("    PartNumber: " & refPartNum & ", Title: " & refTitle)
        Next
        
        ' Also check file descriptors
        Logger.Info("")
        Logger.Info("=== File Descriptors (Drawing) ===")
        For i As Integer = 1 To drawDoc.File.ReferencedFileDescriptors.Count
            Dim fd As FileDescriptor = drawDoc.File.ReferencedFileDescriptors.Item(i)
            Logger.Info("  FD: " & fd.FullFileName)
        Next
    End If
    
    Logger.Info("")
    Logger.Info("=== Test complete ===")
End Sub

Function GetPartNumber(doc As Document) As String
    Try
        Return doc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
    Catch
        Return "(unknown)"
    End Try
End Function

Function GetTitle(doc As Document) As String
    Try
        Return doc.PropertySets.Item("Inventor Summary Information").Item("Title").Value.ToString()
    Catch
        Return "(unknown)"
    End Try
End Function

Function SafeValue(prop As Inventor.Property) As String
    Try
        If prop.Value Is Nothing Then Return "(null)"
        Return prop.Value.ToString()
    Catch
        Return "(error reading)"
    End Try
End Function
