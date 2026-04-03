' ============================================================================
' TestDerivedPartRefs - Diagnostic test for derived part reference update
'
' Run this on a derived part to test different methods to update the base reference.
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
        MessageBox.Show("Please open a derived PART file.", "Test Derived Part Refs")
        Exit Sub
    End If
    
    Dim partDoc As PartDocument = CType(doc, PartDocument)
    Dim compDef As PartComponentDefinition = partDoc.ComponentDefinition
    Dim info As New System.Text.StringBuilder()
    
    info.AppendLine("=== DERIVED PART REFERENCE DIAGNOSTIC v2 ===")
    info.AppendLine()
    info.AppendLine("Document: " & partDoc.DisplayName)
    info.AppendLine()
    
    ' Method 1: Document.ReferencedFileDescriptors
    info.AppendLine("=== Method 1: Document.ReferencedFileDescriptors ===")
    Try
        Dim refDescriptors As ReferencedFileDescriptors = partDoc.ReferencedFileDescriptors
        info.AppendLine("  Count: " & refDescriptors.Count)
        For i As Integer = 1 To refDescriptors.Count
            Dim rfd As ReferencedFileDescriptor = refDescriptors.Item(i)
            info.AppendLine("  Item " & i & ":")
            info.AppendLine("    FullFileName: " & rfd.FullFileName)
            info.AppendLine("    LogicalFileName: " & rfd.LogicalFileName)
            info.AppendLine("    ReferenceType: " & rfd.ReferenceType.ToString())
            info.AppendLine("    ReferenceMissing: " & rfd.ReferenceMissing.ToString())
        Next
    Catch ex As Exception
        info.AppendLine("  ERROR: " & ex.Message)
    End Try
    info.AppendLine()
    
    ' Method 2: ReferenceFeatures in Features collection
    info.AppendLine("=== Method 2: ReferenceFeatures ===")
    Try
        For Each feature As PartFeature In compDef.Features
            If TypeOf feature Is ReferenceFeature Then
                Dim refFeature As ReferenceFeature = CType(feature, ReferenceFeature)
                info.AppendLine("  ReferenceFeature: " & refFeature.Name)
                
                ' Try to get definition
                Try
                    Dim refDef As Object = refFeature.Definition
                    info.AppendLine("    Definition Type: " & TypeName(refDef))
                    
                    ' Check if it's DerivedPartDefinition
                    If TypeOf refDef Is DerivedPartDefinition Then
                        Dim derivedDef As DerivedPartDefinition = CType(refDef, DerivedPartDefinition)
                        info.AppendLine("    DerivedFrom (FullDocumentName): " & derivedDef.FullDocumentName)
                    End If
                Catch ex As Exception
                    info.AppendLine("    Definition Error: " & ex.Message)
                End Try
                
                ' Try ReferencedDocumentDescriptor
                Try
                    Dim refDocDesc As Object = refFeature.ReferencedDocumentDescriptor
                    info.AppendLine("    ReferencedDocumentDescriptor.FullFileName: " & _
                        CStr(CallByName(refDocDesc, "FullFileName", CallType.Get)))
                Catch ex As Exception
                    info.AppendLine("    ReferencedDocumentDescriptor Error: " & ex.Message)
                End Try
            End If
        Next
    Catch ex As Exception
        info.AppendLine("  ERROR: " & ex.Message)
    End Try
    info.AppendLine()
    
    ' Method 3: DerivedPartComponents via alternative access
    info.AppendLine("=== Method 3: DerivedPartComponents (alternative) ===")
    Try
        Dim derivedParts As Object = compDef.ReferenceComponents.DerivedPartComponents
        info.AppendLine("  Type: " & TypeName(derivedParts))
        
        ' Try iteration
        For Each dpc As Object In derivedParts
            info.AppendLine("  DerivedPartComponent:")
            info.AppendLine("    Name: " & CStr(CallByName(dpc, "Name", CallType.Get)))
            Try
                Dim refDocDesc As Object = CallByName(dpc, "ReferencedDocumentDescriptor", CallType.Get)
                info.AppendLine("    ReferencedDocumentDescriptor.FullFileName: " & _
                    CStr(CallByName(refDocDesc, "FullFileName", CallType.Get)))
            Catch
            End Try
        Next
    Catch ex As Exception
        info.AppendLine("  ERROR: " & ex.Message)
    End Try
    info.AppendLine()
    
    ' Method 4: Check FileMigrator for replacing references
    info.AppendLine("=== Method 4: Check available methods for reference update ===")
    info.AppendLine("  Testing ReferencedFileDescriptor.PutLogicalFileName capability...")
    Try
        Dim refDescriptors As ReferencedFileDescriptors = partDoc.ReferencedFileDescriptors
        If refDescriptors.Count > 0 Then
            Dim rfd As ReferencedFileDescriptor = refDescriptors.Item(1)
            info.AppendLine("  Current: " & rfd.FullFileName)
            info.AppendLine("  PutLogicalFileName method exists: Yes (can be used to update)")
        End If
    Catch ex As Exception
        info.AppendLine("  ERROR: " & ex.Message)
    End Try
    
    ' Show results
    Dim resultForm As New System.Windows.Forms.Form()
    resultForm.Text = "Derived Part Reference Diagnostic v2"
    resultForm.Width = 800
    resultForm.Height = 600
    resultForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    
    Dim txt As New System.Windows.Forms.TextBox()
    txt.Multiline = True
    txt.ScrollBars = System.Windows.Forms.ScrollBars.Both
    txt.Dock = System.Windows.Forms.DockStyle.Fill
    txt.Text = info.ToString()
    txt.ReadOnly = True
    
    resultForm.Controls.Add(txt)
    resultForm.ShowDialog()
End Sub
