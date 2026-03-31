' ============================================================================
' TestSaveCopyAs - Test SaveAs with copy flag AND ApprenticeServer
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    
    Dim info As New System.Text.StringBuilder()
    info.AppendLine("=== TEST: SaveAs and ApprenticeServer ===")
    info.AppendLine()
    
    ' Test 1: Check what save methods exist on Document
    info.AppendLine("=== Test 1: Available Save methods on Document ===")
    Dim doc As Document = app.ActiveDocument
    If doc Is Nothing Then
        info.AppendLine("  No document open")
        ShowResult(info.ToString())
        Exit Sub
    End If
    
    info.AppendLine("  Document: " & doc.FullFileName)
    info.AppendLine()
    
    ' Check for various save methods
    info.AppendLine("  Checking Save methods:")
    
    ' SaveAs
    info.AppendLine("    .Save() - exists")
    info.AppendLine("    .SaveAs(fileName, saveCopyAs) - testing...")
    
    ' Check if SaveAs accepts two parameters
    Try
        ' Don't actually call it, just document that we'll test it
        info.AppendLine("    Will test SaveAs with saveCopyAs=True parameter")
    Catch : End Try
    info.AppendLine()
    
    ' Test 2: Check ApprenticeServer
    info.AppendLine("=== Test 2: ApprenticeServer availability ===")
    Try
        ' Try to create ApprenticeServer
        Dim apprentice As Object = CreateObject("Inventor.ApprenticeServerComponent")
        info.AppendLine("  ApprenticeServer created successfully!")
        info.AppendLine("  Type: " & TypeName(apprentice))
        
        ' Check available methods
        info.AppendLine()
        info.AppendLine("  Checking ApprenticeServer methods:")
        
        ' Open
        Try
            info.AppendLine("    .Open(fileName) - checking...")
        Catch : End Try
        
        ' FileSaveAs
        Try
            Dim fsa As Object = CallByName(apprentice, "FileSaveAs", CallType.Get)
            info.AppendLine("    .FileSaveAs: " & TypeName(fsa))
        Catch ex As Exception
            info.AppendLine("    .FileSaveAs: " & ex.Message)
        End Try
        
        ' FileManager
        Try
            Dim fm As Object = CallByName(apprentice, "FileManager", CallType.Get)
            info.AppendLine("    .FileManager: " & TypeName(fm))
        Catch ex As Exception
            info.AppendLine("    .FileManager: " & ex.Message)
        End Try
        
        ' DesignProjectManager  
        Try
            Dim dpm As Object = CallByName(apprentice, "DesignProjectManager", CallType.Get)
            info.AppendLine("    .DesignProjectManager: " & TypeName(dpm))
        Catch ex As Exception
            info.AppendLine("    .DesignProjectManager: " & ex.Message)
        End Try
        
        ' Close apprentice
        Try
            CallByName(apprentice, "Close", CallType.Method)
        Catch : End Try
        
    Catch ex As Exception
        info.AppendLine("  ApprenticeServer error: " & ex.Message)
        info.AppendLine()
        info.AppendLine("  Note: ApprenticeServer may need to be installed separately")
    End Try
    info.AppendLine()
    
    ' Test 3: Try Document.SaveAs with copy flag
    info.AppendLine("=== Test 3: Document.SaveAs with copy flag ===")
    If doc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
        Dim partDoc As PartDocument = CType(doc, PartDocument)
        
        Dim testPath As String = System.IO.Path.Combine( _
            System.IO.Path.GetDirectoryName(doc.FullFileName), _
            "_TestCopy", _
            System.IO.Path.GetFileName(doc.FullFileName))
        
        ' Create folder
        Dim testFolder As String = System.IO.Path.GetDirectoryName(testPath)
        If Not System.IO.Directory.Exists(testFolder) Then
            System.IO.Directory.CreateDirectory(testFolder)
        End If
        
        info.AppendLine("  Test path: " & testPath)
        
        ' Try SaveAs
        Try
            ' SaveAs(fileName, saveCopyAs) where saveCopyAs=True means save a copy
            doc.SaveAs(testPath, True)
            info.AppendLine("  SaveAs(path, True) succeeded!")
            
            ' Check if original document is still the original
            info.AppendLine("  Current document: " & doc.FullFileName)
            
            ' Open the copy and check references
            info.AppendLine()
            info.AppendLine("  Opening copy to check references...")
            Dim copyDoc As Document = app.Documents.Open(testPath, False)
            
            For Each refDoc As Document In copyDoc.ReferencedDocuments
                info.AppendLine("    Referenced: " & refDoc.FullFileName)
            Next
            
            copyDoc.Close(True)
            
        Catch ex As Exception
            info.AppendLine("  SaveAs error: " & ex.Message)
        End Try
    End If
    info.AppendLine()
    
    ' Test 4: Check if there's a way to use Inventor's Pack and Go programmatically
    info.AppendLine("=== Test 4: Pack and Go API ===")
    Try
        Dim packAndGo As Object = CallByName(app, "PackAndGo", CallType.Get)
        info.AppendLine("  PackAndGo: " & TypeName(packAndGo))
    Catch ex As Exception
        info.AppendLine("  PackAndGo: " & ex.Message)
    End Try
    
    ' Try through FileManager
    Try
        Dim fm As FileManager = app.FileManager
        info.AppendLine("  FileManager available")
        
        ' Check for PackAndGo method
        Try
            Dim pag As Object = CallByName(fm, "CreatePackAndGoOptions", CallType.Get)
            info.AppendLine("    .CreatePackAndGoOptions: " & TypeName(pag))
        Catch ex As Exception
            info.AppendLine("    .CreatePackAndGoOptions: " & ex.Message)
        End Try
    Catch ex As Exception
        info.AppendLine("  FileManager error: " & ex.Message)
    End Try
    
    ShowResult(info.ToString())
End Sub

Sub ShowResult(text As String)
    Dim resultForm As New System.Windows.Forms.Form()
    resultForm.Text = "Save Methods Test"
    resultForm.Width = 900
    resultForm.Height = 700
    resultForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    
    Dim txt As New System.Windows.Forms.TextBox()
    txt.Multiline = True
    txt.ScrollBars = System.Windows.Forms.ScrollBars.Both
    txt.Dock = System.Windows.Forms.DockStyle.Fill
    txt.Text = text
    txt.ReadOnly = True
    
    resultForm.Controls.Add(txt)
    resultForm.ShowDialog()
End Sub
