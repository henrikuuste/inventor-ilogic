' Copyright (c) 2026 Henri Kuuste
' ============================================================================
' TestMasterDependencies - Diagnostic script for master dependency tree
'
' Run this on an assembly to discover ALL masters (internal and external),
' build a dependency graph, and output to the iLogic log.
'
' Purpose:
' - Verify we can discover all masters regardless of folder location
' - Understand master-to-master relationships (derivation, projected geometry)
' - Identify intermediate assemblies used for projected geometry
' - Provide baseline for multi-master release system implementation
'
' Usage:
' 1. Open the main assembly of a base element
' 2. Run this script
' 3. Review the dependency tree in the iLogic log
' ============================================================================

Sub Main()
    Dim app As Inventor.Application = ThisApplication
    Dim doc As Document = app.ActiveDocument
    
    If doc Is Nothing OrElse doc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
        MessageBox.Show("Ava aluselemendi põhikoost.", "TestMasterDependencies")
        Return
    End If
    
    Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
    Logger.Info("=== MASTER DEPENDENCY DIAGNOSTIC ===")
    Logger.Info("Assembly: " & asmDoc.DisplayName)
    Logger.Info("Path: " & asmDoc.FullFileName)
    Logger.Info("")
    
    ' Determine source root (element folder)
    Dim sourceRoot As String = FindElementSourceRoot(asmDoc.FullFileName)
    Logger.Info("Source Root: " & sourceRoot)
    Logger.Info("")
    
    ' Step 1: Collect all parts from the assembly
    Logger.Info("=== STEP 1: Discovering all parts ===")
    Dim allParts As New Dictionary(Of String, PartDocument)(StringComparer.OrdinalIgnoreCase)
    CollectAllParts(app, asmDoc, allParts)
    Logger.Info("Found " & allParts.Count & " unique parts")
    Logger.Info("")
    
    ' Step 2: Identify derived parts and their masters
    Logger.Info("=== STEP 2: Identifying derived parts and masters ===")
    Dim derivedParts As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) ' part -> master
    Dim allMasters As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Dim internalMasters As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Dim externalMasters As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    
    For Each kvp In allParts
        Dim partPath As String = kvp.Key
        Dim partDoc As PartDocument = kvp.Value
        Dim masterPath As String = GetMasterPath(partDoc)
        
        If Not String.IsNullOrEmpty(masterPath) Then
            derivedParts(partPath) = masterPath
            allMasters.Add(masterPath)
            
            If IsInsideFolder(masterPath, sourceRoot) Then
                internalMasters.Add(masterPath)
                Logger.Info("  DERIVED (internal): " & System.IO.Path.GetFileName(partPath))
                Logger.Info("    -> Master: " & System.IO.Path.GetFileName(masterPath))
            Else
                externalMasters.Add(masterPath)
                Logger.Info("  DERIVED (EXTERNAL): " & System.IO.Path.GetFileName(partPath))
                Logger.Info("    -> Master: " & masterPath)
            End If
        End If
    Next
    
    Logger.Info("")
    Logger.Info("Summary: " & derivedParts.Count & " derived parts")
    Logger.Info("  Internal masters: " & internalMasters.Count)
    Logger.Info("  External masters: " & externalMasters.Count)
    Logger.Info("")
    
    ' Step 3: Recursively discover master dependencies
    Logger.Info("=== STEP 3: Discovering master dependencies (recursive) ===")
    Dim masterDependencies As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
    Dim intermediateAssemblies As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Dim discoveredMasters As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Dim candidateAssemblies As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase) ' asm -> list of masters it contains
    
    ' Track actual projected geometry chains: (SourcePart, ViaAssembly, TargetPart)
    ' This proves that assembly X is used to project geometry from part A to part B
    Dim projectedGeomChains As New List(Of Tuple(Of String, String, String)) ' (SourcePart, ViaAssembly, TargetPart)
    Dim assembliesWithProjectedGeom As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    
    ' Track unresolved occurrence names from projected geometry references
    ' Format: (OccurrenceName, TargetMasterPartNumber, SourceMasterPath)
    ' Example: ("Selg - Eskiis Multibody (000130):1", "000130", "...\000131.ipt")
    Dim unresolvedOccurrences As New List(Of Tuple(Of String, String, String))
    
    ' Start with known masters and discover recursively
    Dim toProcess As New Queue(Of String)
    For Each m In allMasters
        toProcess.Enqueue(m)
    Next
    
    Do While toProcess.Count > 0
        Dim masterPath As String = toProcess.Dequeue()
        If discoveredMasters.Contains(masterPath) Then Continue Do
        discoveredMasters.Add(masterPath)
        
        Logger.Info("  Analyzing: " & masterPath)
        
        ' Get dependencies of this master
        Dim deps As New List(Of String)
        Try
            Dim masterDoc As Document = app.Documents.Open(masterPath, False)
            
            If masterDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                ' Check for DerivedPartComponents
                Dim partMaster As PartDocument = CType(masterDoc, PartDocument)
                Dim refComps = partMaster.ComponentDefinition.ReferenceComponents
                
                ' DerivedPartComponents - derives from another part (or assembly)
                Logger.Info("    DerivedPartComponents.Count: " & refComps.DerivedPartComponents.Count)
                For Each dpc As DerivedPartComponent In refComps.DerivedPartComponents
                    Try
                        Dim refFile As String = dpc.ReferencedFile.FullFileName
                        Dim defType As String = "unknown"
                        Try
                            defType = TypeName(dpc.Definition)
                        Catch : End Try
                        
                        deps.Add(refFile)
                        Logger.Info("    -> DerivedPartComponent: " & System.IO.Path.GetFileName(refFile))
                        Logger.Info("       Definition type: " & defType)
                        Logger.Info("       ReferencedFile ext: " & System.IO.Path.GetExtension(refFile))
                        
                        ' Try to get more details about the derivation
                        Try
                            Dim dpDef As Object = dpc.Definition
                            ' Check FullDocumentName - might show assembly context
                            Try
                                Dim fullDocName As String = CStr(CallByName(dpDef, "FullDocumentName", CallType.Get))
                                If fullDocName.ToLower().Contains(".iam") Then
                                    Logger.Info("       FullDocumentName: " & fullDocName)
                                    Logger.Info("       ** Contains assembly in derivation path! **")
                                    ' Extract the assembly path
                                    Dim asmMatch As String = ExtractAssemblyPath(fullDocName)
                                    If Not String.IsNullOrEmpty(asmMatch) Then
                                        intermediateAssemblies.Add(asmMatch)
                                        If Not deps.Contains(asmMatch) Then deps.Add(asmMatch)
                                        If Not discoveredMasters.Contains(asmMatch) AndAlso Not toProcess.Contains(asmMatch) Then
                                            toProcess.Enqueue(asmMatch)
                                        End If
                                    End If
                                End If
                            Catch : End Try
                        Catch : End Try
                        
                        ' Check if this is actually an assembly reference disguised as DerivedPartComponent
                        If refFile.ToLower().EndsWith(".iam") Then
                            intermediateAssemblies.Add(refFile)
                            Logger.Info("       ** This is an ASSEMBLY (intermediate) **")
                        End If
                        
                        If Not discoveredMasters.Contains(refFile) AndAlso Not toProcess.Contains(refFile) Then
                            toProcess.Enqueue(refFile)
                        End If
                    Catch ex As Exception
                        Logger.Warn("       Error reading DPC: " & ex.Message)
                    End Try
                Next
                
                ' DerivedAssemblyComponents - derives from an assembly (projected geometry)
                Logger.Info("    DerivedAssemblyComponents.Count: " & refComps.DerivedAssemblyComponents.Count)
                For Each dac As DerivedAssemblyComponent In refComps.DerivedAssemblyComponents
                    Try
                        Dim refFile As String = dac.ReferencedFile.FullFileName
                        deps.Add(refFile)
                        intermediateAssemblies.Add(refFile)
                        Logger.Info("    -> DerivedAssemblyComponent: " & System.IO.Path.GetFileName(refFile))
                        If Not discoveredMasters.Contains(refFile) AndAlso Not toProcess.Contains(refFile) Then
                            toProcess.Enqueue(refFile)
                        End If
                    Catch ex As Exception
                        Logger.Warn("       Error reading DAC: " & ex.Message)
                    End Try
                Next
                
                ' Check ReferencedDocuments - this shows ALL referenced files
                Logger.Info("    ReferencedDocuments.Count: " & partMaster.ReferencedDocuments.Count)
                For Each refDoc As Document In partMaster.ReferencedDocuments
                    Dim refPath As String = refDoc.FullFileName
                    Dim docType As String = If(refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject, "ASM", "PRT")
                    If Not deps.Contains(refPath) Then
                        deps.Add(refPath)
                        Logger.Info("    -> ReferencedDocument [" & docType & "]: " & System.IO.Path.GetFileName(refPath))
                        If refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                            intermediateAssemblies.Add(refPath)
                            Logger.Info("       ** INTERMEDIATE ASSEMBLY detected via ReferencedDocuments **")
                        End If
                        If Not discoveredMasters.Contains(refPath) AndAlso Not toProcess.Contains(refPath) Then
                            toProcess.Enqueue(refPath)
                        End If
                    End If
                Next
                
                ' Initialize variables for tracking projected geometry
                Dim compDef As PartComponentDefinition = partMaster.ComponentDefinition
                Dim projectedGeomFound As Boolean = False
                Dim allProjectedSources As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Dim projectedChains As New List(Of String)
                
                ' Check File.ReferencedFileDescriptors - this contains FullDocumentName with assembly context
                Logger.Info("    ReferencedFileDescriptors.Count: " & partMaster.File.ReferencedFileDescriptors.Count)
                For i As Integer = 1 To partMaster.File.ReferencedFileDescriptors.Count
                    Try
                        Dim fd As FileDescriptor = partMaster.File.ReferencedFileDescriptors.Item(i)
                        Dim refPath As String = fd.FullFileName
                        Dim ext As String = System.IO.Path.GetExtension(refPath).ToLower()
                        Logger.Info("    -> FileDescriptor [" & i & "]: " & System.IO.Path.GetFileName(refPath))
                        Logger.Info("       ReferenceType: " & fd.ReferenceType.ToString())
                        
                        ' Check FullDocumentName - crucial for assembly context
                        Dim fullDocName As String = ""
                        Try : fullDocName = fd.FullDocumentName : Catch : End Try
                        If Not String.IsNullOrEmpty(fullDocName) Then
                            Logger.Info("       FullDocumentName: " & fullDocName)
                            
                            ' Extract assembly context if present
                            ' FullDocumentName format: "C:\path\assembly.iam>C:\path\part.ipt" or "|" separated
                            If fullDocName <> refPath AndAlso (fullDocName.Contains("|") OrElse fullDocName.Contains(">") OrElse fullDocName.ToLower().Contains(".iam")) Then
                                Dim asmPath As String = ExtractAssemblyPath(fullDocName)
                                If Not String.IsNullOrEmpty(asmPath) Then
                                    Logger.Info("       ** ASSEMBLY CONTEXT FOUND: " & asmPath)
                                    intermediateAssemblies.Add(asmPath)
                                    allProjectedSources.Add(asmPath)
                                    allProjectedSources.Add(refPath)
                                    projectedGeomChains.Add(Tuple.Create(refPath, asmPath, masterPath))
                                    assembliesWithProjectedGeom.Add(asmPath)
                                    projectedGeomFound = True
                                    Logger.Info("       *** CHAIN: " & System.IO.Path.GetFileName(refPath) & " -> via " & System.IO.Path.GetFileName(asmPath) & " -> " & System.IO.Path.GetFileName(masterPath))
                                End If
                            End If
                        End If
                        
                        If Not deps.Contains(refPath) Then
                            deps.Add(refPath)
                            If ext = ".iam" Then
                                intermediateAssemblies.Add(refPath)
                                Logger.Info("       ** INTERMEDIATE ASSEMBLY detected via FileDescriptor **")
                            End If
                            If Not discoveredMasters.Contains(refPath) AndAlso Not toProcess.Contains(refPath) Then
                                toProcess.Enqueue(refPath)
                            End If
                        End If
                    Catch : End Try
                Next
                
                ' Check for in-context/adaptive references (projected geometry through assembly)
                ' Part opened standalone won't have ContainingOccurrence - we need assembly context
                ' Summarize sketch references without per-entity logging
                Logger.Info("    Checking sketches (summary)...")
                Dim sketchesWithRefs As Integer = 0
                Dim totalRefEntities As Integer = 0
                
                For Each sketch As PlanarSketch In compDef.Sketches
                    Try
                        Dim sketchRefCount As Integer = 0
                        For Each entity As SketchEntity In sketch.SketchEntities
                            Try
                                Dim isRef As Boolean = False
                                Try : isRef = CBool(CallByName(entity, "Reference", CallType.Get)) : Catch : End Try
                                If isRef Then sketchRefCount += 1
                            Catch : End Try
                        Next
                        If sketchRefCount > 0 Then
                            sketchesWithRefs += 1
                            totalRefEntities += sketchRefCount
                        End If
                    Catch : End Try
                Next
                Logger.Info("      Sketches with references: " & sketchesWithRefs & " (total ref entities: " & totalRefEntities & ")")
                
                ' Method 2: Check ReferenceFeatures at part level
                Logger.Info("    Checking part-level ReferenceFeatures...")
                Try
                    Dim featureCount As Integer = 0
                    Dim totalFeatures As Integer = compDef.Features.Count
                    Dim featureTypes As New Dictionary(Of String, Integer)
                    
                    For Each feature As PartFeature In compDef.Features
                        Dim fType As String = TypeName(feature)
                        If Not featureTypes.ContainsKey(fType) Then featureTypes(fType) = 0
                        featureTypes(fType) += 1
                        
                        ' Check for Reference in name even if not ReferenceFeature type
                        If feature.Name.StartsWith("Reference") AndAlso feature.Name.Contains("(") Then
                            Logger.Info("      Found: " & feature.Name & " [" & fType & "]")
                        End If
                        
                        If TypeOf feature Is ReferenceFeature Then
                            featureCount += 1
                            Dim refFeature As ReferenceFeature = CType(feature, ReferenceFeature)
                            projectedGeomFound = True
                            Logger.Info("      ReferenceFeature: " & refFeature.Name)
                            
                            Try
                                Dim refDocDesc As Object = refFeature.ReferencedDocumentDescriptor
                                If refDocDesc IsNot Nothing Then
                                    Dim refPath As String = CStr(CallByName(refDocDesc, "FullFileName", CallType.Get))
                                    Logger.Info("        FullFileName: " & refPath)
                                    
                                    ' Check FullDocumentName - may contain assembly context
                                    Dim fullDocName As String = ""
                                    Try : fullDocName = CStr(CallByName(refDocDesc, "FullDocumentName", CallType.Get)) : Catch : End Try
                                    If Not String.IsNullOrEmpty(fullDocName) AndAlso fullDocName <> refPath Then
                                        Logger.Info("        FullDocumentName: " & fullDocName)
                                        ' Extract assembly from FullDocumentName
                                        Dim asmPath As String = ExtractAssemblyPath(fullDocName)
                                        If Not String.IsNullOrEmpty(asmPath) Then
                                            Logger.Info("        ** Assembly context: " & asmPath)
                                            allProjectedSources.Add(asmPath)
                                            allProjectedSources.Add(refPath)
                                            projectedGeomChains.Add(Tuple.Create(refPath, asmPath, masterPath))
                                            assembliesWithProjectedGeom.Add(asmPath)
                                            Logger.Info("        *** CHAIN: " & System.IO.Path.GetFileName(refPath) & " -> via " & System.IO.Path.GetFileName(asmPath) & " -> " & System.IO.Path.GetFileName(masterPath))
                                        End If
                                    Else
                                        allProjectedSources.Add(refPath)
                                    End If
                                End If
                            Catch : End Try
                        End If
                    Next
                    If featureCount = 0 Then
                        Logger.Info("      (no ReferenceFeatures found)")
                        Logger.Info("      Feature types in part (" & totalFeatures & " total): " & String.Join(", ", featureTypes.Select(Function(kv) kv.Key & "=" & kv.Value)))
                    End If
                Catch : End Try
                
                ' Method 2b: Check ReferencedDocumentDescriptors - try to find assembly context
                ' DocumentDescriptor may have different property names
                Logger.Info("    Checking Document.ReferencedDocumentDescriptors...")
                Try
                    Dim rddCount As Integer = 0
                    For Each rdd As Object In partMaster.ReferencedDocumentDescriptors
                        rddCount += 1
                        ' Try different property names
                        Dim rddFile As String = ""
                        Try : rddFile = CStr(CallByName(rdd, "FullDocumentName", CallType.Get)) : Catch : End Try
                        If String.IsNullOrEmpty(rddFile) Then
                            Try : rddFile = CStr(CallByName(rdd, "DisplayName", CallType.Get)) : Catch : End Try
                        End If
                        If String.IsNullOrEmpty(rddFile) Then
                            rddFile = TypeName(rdd)
                        End If
                        Logger.Info("      [" & rddCount & "] " & rddFile)
                    Next
                    If rddCount = 0 Then
                        Logger.Info("      (none)")
                    End If
                Catch ex As Exception
                    Logger.Info("      (skipped - " & ex.Message & ")")
                End Try
                
                ' Skip ReferencedOccurrences - requires assembly context which we handle in Step 3a
                
                ' Method 4: Check part's File.ReferencingFiles (inverse reference)
                ' NOTE: ReferencingFiles returns ALL assemblies that have ever referenced this part,
                ' including old versions stored by Vault. We MUST filter out old versions here.
                Logger.Info("    Checking File.ReferencingFiles...")
                Try
                    Dim refingFiles As FilesEnumerator = partMaster.File.ReferencingFiles
                    For Each rf As File In refingFiles
                        Dim rfPath As String = rf.FullFileName
                        If rfPath.ToLower().EndsWith(".iam") Then
                            ' CRITICAL: Skip Vault old versions - they are not actively in use
                            If IsVaultOldVersion(rfPath) Then
                                Logger.Info("      (skipped old version: " & System.IO.Path.GetFileName(rfPath) & ")")
                                Continue For
                            End If
                            
                            Logger.Info("      Part is referenced by assembly: " & System.IO.Path.GetFileName(rfPath))
                            ' Add to a temporary collection - we'll filter after discovering all masters
                            If Not candidateAssemblies.ContainsKey(rfPath) Then
                                candidateAssemblies(rfPath) = New List(Of String)
                            End If
                            candidateAssemblies(rfPath).Add(masterPath)
                        End If
                    Next
                Catch : End Try
                
                ' Log all discovered projected sources
                If allProjectedSources.Count > 0 Then
                    Logger.Info("    PROJECTED GEOMETRY SOURCES:")
                    For Each src In allProjectedSources
                        Dim ext As String = System.IO.Path.GetExtension(src).ToLower()
                        Dim location As String = If(IsInsideFolder(src, sourceRoot), "[INT]", "[EXT]")
                        Logger.Info("      " & location & " " & src)
                        
                        If ext = ".iam" Then
                            intermediateAssemblies.Add(src)
                            If Not deps.Contains(src) Then deps.Add(src)
                            If Not discoveredMasters.Contains(src) AndAlso Not toProcess.Contains(src) Then
                                toProcess.Enqueue(src)
                            End If
                        ElseIf ext = ".ipt" Then
                            If Not deps.Contains(src) Then deps.Add(src)
                            If Not discoveredMasters.Contains(src) AndAlso Not toProcess.Contains(src) Then
                                toProcess.Enqueue(src)
                            End If
                        End If
                    Next
                ElseIf Not projectedGeomFound Then
                    Logger.Info("      (no projected geometry found)")
                End If
                
            ElseIf masterDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                ' Intermediate assembly - check what it references
                Logger.Info("    (This is an ASSEMBLY)")
                Dim asmMaster As AssemblyDocument = CType(masterDoc, AssemblyDocument)
                For Each refDoc As Document In asmMaster.AllReferencedDocuments
                    Dim refPath As String = refDoc.FullFileName
                    deps.Add(refPath)
                    Logger.Info("    -> Assembly contains: " & System.IO.Path.GetFileName(refPath))
                    ' Queue .ipt files as potential masters
                    If refPath.ToLower().EndsWith(".ipt") Then
                        If Not discoveredMasters.Contains(refPath) AndAlso Not toProcess.Contains(refPath) Then
                            toProcess.Enqueue(refPath)
                        End If
                    End If
                Next
            End If
            
            masterDependencies(masterPath) = deps
            
        Catch ex As Exception
            Logger.Warn("    ERROR: " & ex.Message)
            masterDependencies(masterPath) = deps
        End Try
    Loop
    
    Logger.Info("")
    Logger.Info("Total discovered masters: " & discoveredMasters.Count)
    Logger.Info("")
    
    ' Step 3a: Analyze assemblies in context to find actual projected geometry chains
    ' Open each candidate assembly and check for in-place references between masters
    Logger.Info("=== STEP 3a: Analyzing assemblies for projected geometry ===")
    
    ' CRITICAL: Always ensure the root assembly is analyzed, regardless of ReferencingFiles results
    ' ReferencingFiles may not return the assembly for all masters (e.g., if they're in sub-assemblies)
    Dim rootAsmPath As String = asmDoc.FullFileName
    If Not candidateAssemblies.ContainsKey(rootAsmPath) Then
        candidateAssemblies(rootAsmPath) = New List(Of String)
    End If
    ' Add all discovered masters to the root assembly's list
    For Each masterPath In discoveredMasters
        If Not candidateAssemblies(rootAsmPath).Contains(masterPath) Then
            candidateAssemblies(rootAsmPath).Add(masterPath)
        End If
    Next
    Logger.Info("  Ensured root assembly " & System.IO.Path.GetFileName(rootAsmPath) & " is analyzed with " & candidateAssemblies(rootAsmPath).Count & " masters")
    
    For Each kvp In candidateAssemblies
        Dim asmPath As String = kvp.Key
        Dim containedMasters As List(Of String) = kvp.Value
        
        ' Safety check: Skip Vault old versions (should already be filtered, but double-check)
        If IsVaultOldVersion(asmPath) Then Continue For
        
        ' Need at least 2 masters to have projected geom between them
        ' (But we're now more inclusive by ensuring root assembly has all masters)
        If containedMasters.Count < 2 Then Continue For
        
        Logger.Info("  Opening assembly: " & System.IO.Path.GetFileName(asmPath))
        Logger.Info("    Masters to check: " & String.Join(", ", containedMasters.Select(Function(m) System.IO.Path.GetFileName(m))))
        Try
            Dim candAsm As AssemblyDocument = CType(app.Documents.Open(asmPath, True), AssemblyDocument)
            
            ' CRITICAL: Activate the assembly so browser nodes can be accessed
            candAsm.Activate()
            app.ActiveView.Fit()
            Logger.Info("    Assembly activated for browser node access")
            
            ' === COMPREHENSIVE ASSEMBLY-LEVEL ANALYSIS ===
            Logger.Info("    === ASSEMBLY-LEVEL ANALYSIS ===")
            
            ' Check assembly's ReferencedDocuments
            Logger.Info("    ReferencedDocuments:")
            Try
                For Each refDoc As Document In candAsm.ReferencedDocuments
                    Logger.Info("      -> " & System.IO.Path.GetFileName(refDoc.FullFileName))
                Next
            Catch ex As Exception
                Logger.Info("      (error: " & ex.Message & ")")
            End Try
            
            ' Check assembly's ReferencedDocumentDescriptors
            Logger.Info("    ReferencedDocumentDescriptors:")
            Try
                For Each rdd As Object In candAsm.ReferencedDocumentDescriptors
                    Try
                        Dim rddPath As String = CStr(CallByName(rdd, "FullDocumentName", CallType.Get))
                        Dim refType As String = ""
                        Try : refType = CStr(CallByName(rdd, "ReferenceType", CallType.Get)) : Catch : End Try
                        Logger.Info("      -> " & System.IO.Path.GetFileName(rddPath) & If(refType <> "", " [" & refType & "]", ""))
                    Catch : End Try
                Next
            Catch ex As Exception
                Logger.Info("      (error: " & ex.Message & ")")
            End Try
            
            ' Check assembly's File.ReferencedFileDescriptors for FullDocumentName (assembly context)
            Logger.Info("    File.ReferencedFileDescriptors (checking FullDocumentName):")
            Try
                For i As Integer = 1 To candAsm.File.ReferencedFileDescriptors.Count
                    Try
                        Dim fd As FileDescriptor = candAsm.File.ReferencedFileDescriptors.Item(i)
                        Dim fdPath As String = fd.FullFileName
                        Dim fullDocName As String = ""
                        Try : fullDocName = fd.FullDocumentName : Catch : End Try
                        Logger.Info("      -> " & System.IO.Path.GetFileName(fdPath))
                        If Not String.IsNullOrEmpty(fullDocName) AndAlso fullDocName <> fdPath Then
                            Logger.Info("         FullDocumentName: " & fullDocName)
                        End If
                    Catch : End Try
                Next
            Catch ex As Exception
                Logger.Info("      (error: " & ex.Message & ")")
            End Try
            
            ' Check assembly constraints for cross-part references
            Logger.Info("    Assembly Constraints:")
            Try
                Dim asmDef As AssemblyComponentDefinition = candAsm.ComponentDefinition
                Dim constraintCount As Integer = asmDef.Constraints.Count
                Logger.Info("      Total constraints: " & constraintCount)
                ' Log first few constraints
                Dim cIdx As Integer = 0
                For Each constr As Object In asmDef.Constraints
                    cIdx = cIdx + 1
                    If cIdx > 5 Then Exit For
                    Try
                        Logger.Info("        [" & cIdx & "] " & TypeName(constr))
                    Catch : End Try
                Next
            Catch ex As Exception
                Logger.Info("      (error: " & ex.Message & ")")
            End Try
            
            ' Check for DocumentInterests (links between documents)
            Logger.Info("    DocumentInterests:")
            Try
                Dim docInterests As Object = CallByName(candAsm, "DocumentInterests", CallType.Get)
                If docInterests IsNot Nothing Then
                    Dim diCount As Integer = CInt(CallByName(docInterests, "Count", CallType.Get))
                    Logger.Info("      Count: " & diCount)
                    For idx As Integer = 1 To Math.Min(diCount, 10)
                        Try
                            Dim di As Object = CallByName(docInterests, "Item", CallType.Method, idx)
                            Dim diName As String = CStr(CallByName(di, "Name", CallType.Get))
                            Logger.Info("      [" & idx & "] " & diName)
                        Catch : End Try
                    Next
                End If
            Catch ex As Exception
                Logger.Info("      (not available: " & ex.Message & ")")
            End Try
            
            ' Check for AttributeSets on the assembly
            Logger.Info("    Assembly AttributeSets:")
            Try
                Dim attrSets As AttributeSets = candAsm.AttributeSets
                Logger.Info("      Count: " & attrSets.Count)
                For Each attrSet As AttributeSet In attrSets
                    Logger.Info("      -> " & attrSet.Name & " (" & attrSet.Count & " attributes)")
                Next
            Catch ex As Exception
                Logger.Info("      (error: " & ex.Message & ")")
            End Try
            
            ' Check ReferenceKeyManager
            Logger.Info("    ReferenceKeyManager:")
            Try
                Dim rkm As Object = CallByName(candAsm, "ReferenceKeyManager", CallType.Get)
                If rkm IsNot Nothing Then
                    Logger.Info("      Available: Yes")
                    ' Try to get context count or other info
                    Try
                        Dim contextCount As Integer = CInt(CallByName(rkm, "ContextCount", CallType.Get))
                        Logger.Info("      ContextCount: " & contextCount)
                    Catch : End Try
                End If
            Catch ex As Exception
                Logger.Info("      (error: " & ex.Message & ")")
            End Try
            
            ' Build a map of all occurrences (including nested) by document path
            Dim occByPath As New Dictionary(Of String, ComponentOccurrence)(StringComparer.OrdinalIgnoreCase)
            Dim allOccurrences As New List(Of ComponentOccurrence)
            
            Logger.Info("    === OCCURRENCES ===")
            For Each occ As ComponentOccurrence In candAsm.ComponentDefinition.Occurrences.AllLeafOccurrences
                Try
                    allOccurrences.Add(occ)
                    Dim occPath As String = occ.Definition.Document.FullFileName
                    Logger.Info("      " & occ.Name & " -> " & System.IO.Path.GetFileName(occPath))
                    
                    ' Check if this occurrence has any ReferencedOccurrences (links to other occurrences)
                    Try
                        Dim refOccs As Object = CallByName(occ, "ReferencedOccurrences", CallType.Get)
                        If refOccs IsNot Nothing Then
                            Dim refOccsCount As Integer = CInt(CallByName(refOccs, "Count", CallType.Get))
                            If refOccsCount > 0 Then
                                Logger.Info("        ** ReferencedOccurrences: " & refOccsCount)
                            End If
                        End If
                    Catch : End Try
                    
                    If Not occByPath.ContainsKey(occPath) Then
                        occByPath(occPath) = occ
                    End If
                Catch ex As Exception
                    Logger.Info("      " & occ.Name & " -> (error: " & ex.Message & ")")
                End Try
            Next
            Logger.Info("    Total occurrences: " & allOccurrences.Count)
            Logger.Info("    Unique parts: " & occByPath.Count)
            
            ' For each master in this assembly, find its occurrence
            For Each masterPath As String In containedMasters
                Dim masterOcc As ComponentOccurrence = Nothing
                If occByPath.ContainsKey(masterPath) Then
                    masterOcc = occByPath(masterPath)
                End If
                
                If masterOcc Is Nothing Then
                    Logger.Info("    Master NOT FOUND in assembly: " & System.IO.Path.GetFileName(masterPath))
                    Continue For
                End If
                Logger.Info("    Checking: " & System.IO.Path.GetFileName(masterPath) & " (occ: " & masterOcc.Name & ")")
                
                ' Get the part definition
                Dim partDef As PartComponentDefinition = CType(masterOcc.Definition, PartComponentDefinition)
                Dim foundRefsInPart As Integer = 0
                Dim totalRefEntities As Integer = 0
                Dim refEntWithSource As Integer = 0
                
                ' First check part-level ReferenceComponents collection
                Logger.Info("      === PART-LEVEL ReferenceComponents ===")
                Try
                    Dim refComps As ReferenceComponents = partDef.ReferenceComponents
                    Logger.Info("        DerivedPartComponents: " & refComps.DerivedPartComponents.Count)
                    Logger.Info("        DerivedAssemblyComponents: " & refComps.DerivedAssemblyComponents.Count)
                    
                    ' Check for other reference collections
                    Dim rcColNames As String() = {"DerivedPartComponents", "DerivedAssemblyComponents", 
                        "WorkSurfaceReferences", "ProjectedEdges", "ProjectedSketches", "ImportedComponents"}
                    For Each rcName In rcColNames
                        Try
                            Dim rcCol As Object = CallByName(refComps, rcName, CallType.Get)
                            If rcCol IsNot Nothing Then
                                Dim cnt As Integer = 0
                                Try : cnt = CInt(CallByName(rcCol, "Count", CallType.Get)) : Catch : End Try
                                If cnt > 0 Then
                                    Logger.Info("        " & rcName & ": " & cnt)
                                    ' List first few items
                                    For i As Integer = 1 To Math.Min(cnt, 5)
                                        Try
                                            Dim item As Object = CallByName(rcCol, "Item", CallType.Method, i)
                                            Dim itemName As String = ""
                                            Try : itemName = CStr(CallByName(item, "Name", CallType.Get)) : Catch : itemName = TypeName(item) : End Try
                                            Logger.Info("          [" & i & "] " & itemName)
                                            
                                            ' Try to get referenced document
                                            Try
                                                Dim refFile As Object = CallByName(item, "ReferencedFile", CallType.Get)
                                                If refFile IsNot Nothing Then
                                                    Dim rfPath As String = CStr(CallByName(refFile, "FullFileName", CallType.Get))
                                                    Logger.Info("              RefFile: " & System.IO.Path.GetFileName(rfPath))
                                                End If
                                            Catch : End Try
                                            
                                            ' For DerivedPartComponent, check Definition for includes
                                            If TypeName(item).Contains("DerivedPart") Then
                                                Try
                                                    Dim dpDef As Object = CallByName(item, "Definition", CallType.Get)
                                                    If dpDef IsNot Nothing Then
                                                        ' Check IncludeSketches
                                                        Dim inclSk As Object = CallByName(dpDef, "IncludeSketches", CallType.Get)
                                                        If inclSk IsNot Nothing Then
                                                            Dim skCnt As Integer = CInt(CallByName(inclSk, "Count", CallType.Get))
                                                            If skCnt > 0 Then Logger.Info("              IncludeSketches: " & skCnt)
                                                        End If
                                                    End If
                                                Catch : End Try
                                            End If
                                        Catch : End Try
                                    Next
                                End If
                            End If
                        Catch : End Try
                    Next
                Catch ex As Exception
                    Logger.Warn("        Error: " & ex.Message)
                End Try
                
                ' Analyze ALL sketches for projected geometry (browser node analysis)
                Logger.Info("      === ANALYZING SKETCHES FOR PROJECTED GEOMETRY ===")
                For Each sk As PlanarSketch In partDef.Sketches
                    ' Check if this sketch has any reference entities before doing detailed analysis
                    Dim hasRefs As Boolean = False
                    Try
                        For Each entity As SketchEntity In sk.SketchEntities
                            Try
                                Dim isRef As Boolean = False
                                Try : isRef = CBool(CallByName(entity, "Reference", CallType.Get)) : Catch : End Try
                                If isRef Then hasRefs = True : Exit For
                            Catch : End Try
                        Next
                    Catch : End Try
                    If Not hasRefs Then Continue For ' Skip sketches without reference entities
                    
                    Try
                        ' Log which sketch we're analyzing
                        Logger.Info("        --- Analyzing sketch: " & sk.Name & " ---")
                        Logger.Info("        Sketch type: " & TypeName(sk))
                        Logger.Info("        Entity counts:")
                        Logger.Info("          SketchEntities: " & sk.SketchEntities.Count)
                        Logger.Info("          SketchPoints: " & sk.SketchPoints.Count)
                        Logger.Info("          SketchLines: " & sk.SketchLines.Count)
                        Logger.Info("          SketchArcs: " & sk.SketchArcs.Count)
                        Logger.Info("          SketchSplines: " & sk.SketchSplines.Count)
                        
                        ' Try all possible collections that might contain Reference items
                        Dim colsToTry As String() = {"ReferenceComponents", "ExternalSketchPoints", 
                            "ProjectedCuts", "GeometricConstraints", "DimensionConstraints",
                            "IncludeEntities", "IntersectEntities", "ProjectedEntities", 
                            "SketchImages", "SketchBlocks", "SketchFixedSplines", "SketchEquationCurves",
                            "SketchOGSCurves", "OffsetEntities"}
                        
                        For Each colName In colsToTry
                            Try
                                Dim col As Object = CallByName(sk, colName, CallType.Get)
                                If col IsNot Nothing Then
                                    Dim cnt As Integer = 0
                                    Try : cnt = CInt(CallByName(col, "Count", CallType.Get)) : Catch : End Try
                                    If cnt > 0 Then Logger.Info("          " & colName & ": " & cnt)
                                End If
                            Catch : End Try
                        Next
                        
                        ' Explore sketch-level properties that might reveal projected geometry sources
                        Logger.Info("        === SKETCH PROPERTIES EXPLORATION ===")
                        
                        ' Check for Adaptive property (mentioned in forums)
                        Try
                            Dim isAdaptive As Boolean = CBool(CallByName(sk, "Adaptive", CallType.Get))
                            Logger.Info("          Adaptive: " & isAdaptive)
                        Catch : End Try
                        
                        ' Check sketch's ReferencedFiles
                        Try
                            Dim skRefFiles As Object = CallByName(sk, "ReferencedFiles", CallType.Get)
                            If skRefFiles IsNot Nothing Then
                                Dim rfCount As Integer = CInt(CallByName(skRefFiles, "Count", CallType.Get))
                                Logger.Info("          ReferencedFiles: " & rfCount)
                            End If
                        Catch : End Try
                        
                        ' Check sketch's Parent for clues
                        Try
                            Dim skParent As Object = sk.Parent
                            Logger.Info("          Parent: " & TypeName(skParent))
                        Catch : End Try
                        
                        ' Check for ReferenceComponents on the sketch
                        Try
                            Dim skRefComps As Object = CallByName(sk, "ReferenceComponents", CallType.Get)
                            If skRefComps IsNot Nothing Then
                                Dim rcCount As Integer = CInt(CallByName(skRefComps, "Count", CallType.Get))
                                Logger.Info("          Sketch ReferenceComponents: " & rcCount)
                            End If
                        Catch : End Try
                        
                        ' Check for ExternalReferences on the sketch
                        Try
                            Dim extRefs As Object = CallByName(sk, "ExternalReferences", CallType.Get)
                            If extRefs IsNot Nothing Then
                                Dim erCount As Integer = CInt(CallByName(extRefs, "Count", CallType.Get))
                                Logger.Info("          ExternalReferences: " & erCount)
                                For erIdx As Integer = 1 To Math.Min(erCount, 5)
                                    Try
                                        Dim er As Object = CallByName(extRefs, "Item", CallType.Method, erIdx)
                                        Logger.Info("            [" & erIdx & "] " & TypeName(er))
                                        ' Try to get source from external reference
                                        Try
                                            Dim erDoc As Object = CallByName(er, "ReferencedDocument", CallType.Get)
                                            If erDoc IsNot Nothing Then
                                                Dim erPath As String = CStr(CallByName(erDoc, "FullFileName", CallType.Get))
                                                Logger.Info("                ReferencedDocument: " & System.IO.Path.GetFileName(erPath))
                                            End If
                                        Catch : End Try
                                    Catch : End Try
                                Next
                            End If
                        Catch : End Try
                        
                        ' Check various other collections
                        Dim sketchCollections As String() = {"IncludeGeometry", "ProjectedCuts", "IntersectGeometry", 
                            "AssociativeGeometry", "ExternalSketchGeometry", "AdaptiveGeometry"}
                        For Each colName In sketchCollections
                            Try
                                Dim col As Object = CallByName(sk, colName, CallType.Get)
                                If col IsNot Nothing Then
                                    Dim colCount As Integer = CInt(CallByName(col, "Count", CallType.Get))
                                    If colCount > 0 Then
                                        Logger.Info("          " & colName & ": " & colCount)
                                        ' Sample first item
                                        Try
                                            Dim item As Object = CallByName(col, "Item", CallType.Method, 1)
                                            Logger.Info("            [1] " & TypeName(item))
                                            ' Try to get occurrence info
                                            Try
                                                Dim itemOcc As Object = CallByName(item, "ContainingOccurrence", CallType.Get)
                                                If itemOcc IsNot Nothing Then
                                                    Logger.Info("            ** ContainingOccurrence: " & CStr(CallByName(itemOcc, "Name", CallType.Get)))
                                                End If
                                            Catch : End Try
                                        Catch : End Try
                                    End If
                                End If
                            Catch : End Try
                        Next
                        
                        ' Check AttributeSets on the sketch
                        Try
                            Dim skAttrs As AttributeSets = sk.AttributeSets
                            If skAttrs.Count > 0 Then
                                Logger.Info("          Sketch AttributeSets: " & skAttrs.Count)
                                For Each attrSet As AttributeSet In skAttrs
                                    Logger.Info("            -> " & attrSet.Name)
                                Next
                            End If
                        Catch : End Try
                        
                        ' Check if sketch has any child features (browser items under sketch)
                        Logger.Info("        Checking for browser children:")
                        Try
                            Dim browserPane As Object = app.ActiveDocument.BrowserPanes.ActivePane
                            Dim skNode As Object = browserPane.GetBrowserNodeFromObject(sk)
                            If skNode IsNot Nothing Then
                                Logger.Info("          Sketch browser node found")
                                
                                ' Try to get child nodes - enumerate all properties to understand them
                                Try
                                    Dim childNodes As BrowserNodesEnumerator = skNode.BrowserNodes
                                    If childNodes IsNot Nothing Then
                                        Dim childCount As Integer = childNodes.Count
                                        Logger.Info("          Child nodes: " & childCount)
                                        Dim i As Integer = 0
                                        For Each childNode As BrowserNode In childNodes
                                            i += 1
                                            If i > 15 Then Exit For
                                            Try
                                                Dim nodeType As String = TypeName(childNode)
                                                Logger.Info("            [" & i & "] NodeType: " & nodeType)
                                                
                                                ' Get BrowserNodeDefinition
                                                Dim childDef As BrowserNodeDefinition = childNode.BrowserNodeDefinition
                                                
                                                If childDef IsNot Nothing Then
                                                    Dim defType As String = TypeName(childDef)
                                                    Logger.Info("                DefType: " & defType)
                                                    
                                                    ' Try Label property
                                                    Dim label As String = ""
                                                    Try : label = childDef.Label : Catch : End Try
                                                    If Not String.IsNullOrEmpty(label) Then
                                                        Logger.Info("                Label: " & label)
                                                        
                                                        ' Parse the label to extract source occurrence
                                                        ' Format: "ReferenceXX (OccurrenceName)" e.g. "Reference83 (Selg - Eskiis Multibody (000130):1)"
                                                        If label.StartsWith("Reference") AndAlso label.Contains("(") Then
                                                            ' Extract the occurrence name from the label
                                                            Dim occNameFromLabel As String = label.Substring(label.IndexOf("(") + 1)
                                                            If occNameFromLabel.EndsWith(")") Then occNameFromLabel = occNameFromLabel.Substring(0, occNameFromLabel.Length - 1)
                                                            
                                                            Logger.Info("                   OccName from label: " & occNameFromLabel)
                                                            
                                                            ' Now find this occurrence in the assembly and verify the link
                                                            Dim foundOcc As ComponentOccurrence = Nothing
                                                            Dim verifiedSourcePath As String = ""
                                                            
                                                            ' Search all occurrences in the assembly
                                                            For Each occ As ComponentOccurrence In candAsm.ComponentDefinition.Occurrences.AllLeafOccurrences
                                                                Try
                                                                    If occ.Name = occNameFromLabel Then
                                                                        foundOcc = occ
                                                                        verifiedSourcePath = occ.Definition.Document.FullFileName
                                                                        Exit For
                                                                    End If
                                                                Catch : End Try
                                                            Next
                                                            
                                                            If foundOcc IsNot Nothing Then
                                                                Logger.Info("                   ** VERIFIED in assembly:")
                                                                Logger.Info("                      Occurrence: " & foundOcc.Name)
                                                                Logger.Info("                      Document: " & System.IO.Path.GetFileName(verifiedSourcePath))
                                                                Logger.Info("                      Full path: " & verifiedSourcePath)
                                                                
                                                                ' Log additional occurrence info
                                                                Try
                                                                    Logger.Info("                      Visible: " & foundOcc.Visible)
                                                                    Logger.Info("                      Grounded: " & foundOcc.Grounded)
                                                                Catch : End Try
                                                                
                                                                ' This is the actual source - verify it's a cross-master reference
                                                                If Not verifiedSourcePath.Equals(masterPath, StringComparison.OrdinalIgnoreCase) Then
                                                                    ' Check if this chain already exists (avoid duplicates)
                                                                    Dim chainExists As Boolean = projectedGeomChains.Any(Function(c) _
                                                                        c.Item1.Equals(verifiedSourcePath, StringComparison.OrdinalIgnoreCase) AndAlso _
                                                                        c.Item2.Equals(asmPath, StringComparison.OrdinalIgnoreCase) AndAlso _
                                                                        c.Item3.Equals(masterPath, StringComparison.OrdinalIgnoreCase))
                                                                    If Not chainExists Then
                                                                        Logger.Info("                   *** CROSS-MASTER REFERENCE VERIFIED ***")
                                                                        projectedGeomChains.Add(Tuple.Create(verifiedSourcePath, asmPath, masterPath))
                                                                        assembliesWithProjectedGeom.Add(asmPath)
                                                                    End If
                                                                Else
                                                                    Logger.Info("                   (same part, self-reference)")
                                                                End If
                                                            Else
                                                                Logger.Info("                   (occurrence not found by exact name - trying fallback)")
                                                                
                                                                ' Fallback: extract part number from occurrence name and match by filename
                                                                ' Two formats possible:
                                                                ' 1. With description: "Selg - Eskiis Multibody (000130):1" -> extract "000130" from inner ()
                                                                ' 2. Default (no description): "000130:1" -> extract "000130" before colon
                                                                Dim sourcePartNum As String = ""
                                                                
                                                                If occNameFromLabel.Contains("(") Then
                                                                    ' Format: "Description (000130):1" - extract from inner parentheses
                                                                    Dim innerStart As Integer = occNameFromLabel.LastIndexOf("(")
                                                                    Dim innerEnd As Integer = occNameFromLabel.IndexOf(")", innerStart)
                                                                    If innerEnd > innerStart Then
                                                                        sourcePartNum = occNameFromLabel.Substring(innerStart + 1, innerEnd - innerStart - 1)
                                                                        Logger.Info("                   Format: with description")
                                                                    End If
                                                                ElseIf occNameFromLabel.Contains(":") Then
                                                                    ' Format: "000130:1" - extract before colon
                                                                    sourcePartNum = occNameFromLabel.Split(":"c)(0)
                                                                    Logger.Info("                   Format: default (no description)")
                                                                Else
                                                                    ' Just use the whole thing as the part number
                                                                    sourcePartNum = occNameFromLabel
                                                                    Logger.Info("                   Format: plain filename")
                                                                End If
                                                                
                                                                If Not String.IsNullOrEmpty(sourcePartNum) Then
                                                                    Logger.Info("                   Extracted part#: " & sourcePartNum)
                                                                    
                                                                    ' Find in occByPath
                                                                    For Each kvp2 In occByPath
                                                                        If System.IO.Path.GetFileNameWithoutExtension(kvp2.Key) = sourcePartNum Then
                                                                            verifiedSourcePath = kvp2.Key
                                                                            Logger.Info("                   Matched: " & verifiedSourcePath)
                                                                            Exit For
                                                                        End If
                                                                    Next
                                                                End If
                                                                
                                                                If Not String.IsNullOrEmpty(verifiedSourcePath) AndAlso Not verifiedSourcePath.Equals(masterPath, StringComparison.OrdinalIgnoreCase) Then
                                                                    Dim chainExists As Boolean = projectedGeomChains.Any(Function(c) _
                                                                        c.Item1.Equals(verifiedSourcePath, StringComparison.OrdinalIgnoreCase) AndAlso _
                                                                        c.Item2.Equals(asmPath, StringComparison.OrdinalIgnoreCase) AndAlso _
                                                                        c.Item3.Equals(masterPath, StringComparison.OrdinalIgnoreCase))
                                                                    If Not chainExists Then
                                                                        Logger.Info("                   *** CROSS-MASTER REFERENCE (fallback) ***")
                                                                        projectedGeomChains.Add(Tuple.Create(verifiedSourcePath, asmPath, masterPath))
                                                                        assembliesWithProjectedGeom.Add(asmPath)
                                                                    End If
                                                                Else
                                                                    ' Occurrence not found in THIS assembly - record it for later search
                                                                    ' The projected geometry was created in a DIFFERENT assembly
                                                                    If Not String.IsNullOrEmpty(sourcePartNum) Then
                                                                        Dim alreadyRecorded As Boolean = unresolvedOccurrences.Any(Function(u) _
                                                                            u.Item1.Equals(occNameFromLabel, StringComparison.OrdinalIgnoreCase) AndAlso _
                                                                            u.Item3.Equals(masterPath, StringComparison.OrdinalIgnoreCase))
                                                                        If Not alreadyRecorded Then
                                                                            unresolvedOccurrences.Add(Tuple.Create(occNameFromLabel, sourcePartNum, masterPath))
                                                                            Logger.Info("                   ** UNRESOLVED - will search for assembly with this occurrence **")
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                    
                                                    ' If NativeBrowserNodeDefinition, try to get NativeObject from it
                                                    If TypeOf childDef Is NativeBrowserNodeDefinition Then
                                                        Dim nativeDef As NativeBrowserNodeDefinition = CType(childDef, NativeBrowserNodeDefinition)
                                                        Try
                                                            Dim nativeObj As Object = nativeDef.NativeObject
                                                            If nativeObj IsNot Nothing Then
                                                                Logger.Info("                NativeObject: " & TypeName(nativeObj))
                                                                
                                                                ' Try to get name/info
                                                                Try
                                                                    Dim objName As String = CStr(CallByName(nativeObj, "Name", CallType.Get))
                                                                    Logger.Info("                  Name: " & objName)
                                                                Catch : End Try
                                                                
                                                                ' For edges/faces, check parent
                                                                Try
                                                                    Dim parent As Object = CallByName(nativeObj, "Parent", CallType.Get)
                                                                    If parent IsNot Nothing Then
                                                                        Logger.Info("                  Parent: " & TypeName(parent))
                                                                    End If
                                                                Catch : End Try
                                                                
                                                                ' Check for ContainingOccurrence
                                                                Try
                                                                    Dim contOcc As Object = CallByName(nativeObj, "ContainingOccurrence", CallType.Get)
                                                                    If contOcc IsNot Nothing Then
                                                                        Dim occName As String = CStr(CallByName(contOcc, "Name", CallType.Get))
                                                                        Logger.Info("                  ** OCCURRENCE: " & occName)
                                                                    End If
                                                                Catch : End Try
                                                            End If
                                                        Catch ex As Exception
                                                            Logger.Info("                NativeObject error: " & ex.Message)
                                                        End Try
                                                    End If
                                                End If
                                                
                                                ' Also try NativeObject directly on the node
                                                Try
                                                    Dim directNative As Object = childNode.NativeObject
                                                    If directNative IsNot Nothing Then
                                                        Logger.Info("                DirectNative: " & TypeName(directNative))
                                                    End If
                                                Catch : End Try
                                            Catch ex As Exception
                                                Logger.Info("            [" & i & "] Error: " & ex.Message)
                                            End Try
                                        Next childNode
                                    End If
                                Catch ex As Exception
                                    Logger.Info("          BrowserNodes error: " & ex.Message)
                                End Try
                            End If
                        Catch ex As Exception
                            Logger.Info("          Browser access error: " & ex.Message)
                        End Try
                        
                        ' Check ALL SketchEntities for any reference-related properties
                        Logger.Info("        Scanning ALL SketchEntities for references:")
                        Dim entIdx As Integer = 0
                        Dim refEntsFound As Integer = 0
                        For Each se As SketchEntity In sk.SketchEntities
                            Try
                                entIdx = entIdx + 1
                                Dim seType As String = TypeName(se)
                                
                                ' Check for ReferencedEntity
                                Dim refEnt As Object = Nothing
                                Try : refEnt = CallByName(se, "ReferencedEntity", CallType.Get) : Catch : End Try
                                
                                ' Check for Reference property
                                Dim isRef As Boolean = False
                                Try : isRef = CBool(CallByName(se, "Reference", CallType.Get)) : Catch : End Try
                                
                                ' Check for AssociativeID (indicates associative geometry)
                                Dim assocId As Integer = 0
                                Try : assocId = CInt(CallByName(se, "AssociativeID", CallType.Get)) : Catch : End Try
                                
                                If refEnt IsNot Nothing OrElse isRef OrElse assocId > 0 Then
                                    refEntsFound = refEntsFound + 1
                                    Dim seName As String = ""
                                    Try : seName = se.Name : Catch : seName = seType & entIdx : End Try
                                    
                                    Dim info As String = seName & " [" & seType & "]"
                                    If isRef Then info = info & " Ref=True"
                                    If assocId > 0 Then info = info & " AssocID=" & assocId
                                    If refEntsFound <= 5 Then
                                        Logger.Info("          " & info)
                                        
                                        If refEnt IsNot Nothing Then
                                            Logger.Info("              ReferencedEntity: " & TypeName(refEnt))
                                            ' Try ContainingOccurrence
                                            Try
                                                Dim contOcc As Object = CallByName(refEnt, "ContainingOccurrence", CallType.Get)
                                                If contOcc IsNot Nothing Then
                                                    Logger.Info("              ** ContainingOccurrence: " & CStr(CallByName(contOcc, "Name", CallType.Get)))
                                                End If
                                            Catch : End Try
                                        End If
                                        
                                        ' Try GetReferenceKey and decode with ReferenceKeyManager
                                        Try
                                            Dim keyBytes(0 To 255) As Byte
                                            CallByName(se, "GetReferenceKey", CallType.Method, keyBytes, 0)
                                            
                                            ' Try to decode with ReferenceKeyManager.KeyToString
                                            Try
                                                Dim rkm As Object = candAsm.ReferenceKeyManager
                                                Dim keyStr As String = CStr(CallByName(rkm, "KeyToString", CallType.Method, keyBytes))
                                                Logger.Info("              RefKey (string): " & Left(keyStr, 60) & "...")
                                            Catch : End Try
                                            
                                            ' Try to bind the key back to an object in the assembly context
                                            Try
                                                Dim rkm As Object = candAsm.ReferenceKeyManager
                                                Dim boundObj As Object = Nothing
                                                boundObj = CallByName(rkm, "BindKeyToObject", CallType.Method, keyBytes, 0)
                                                If boundObj IsNot Nothing Then
                                                    Logger.Info("              BoundObject: " & TypeName(boundObj))
                                                    ' Try to get occurrence from bound object
                                                    Try
                                                        Dim boundOcc As Object = CallByName(boundObj, "ContainingOccurrence", CallType.Get)
                                                        If boundOcc IsNot Nothing Then
                                                            Logger.Info("              ** BOUND TO OCC: " & CStr(CallByName(boundOcc, "Name", CallType.Get)))
                                                        End If
                                                    Catch : End Try
                                                End If
                                            Catch ex As Exception
                                                Logger.Info("              BindKey error: " & ex.Message)
                                            End Try
                                        Catch : End Try
                                        
                                        ' Try to access various properties that might reveal source
                                        Dim propsToTry As String() = {"SourceEntity", "SourceGeometry", "SourceObject", 
                                            "AssociativeObject", "LinkedObject", "ParentObject", "OriginObject",
                                            "ReferenceComponent", "AssociativeGeometry", "Definition"}
                                        For Each propName In propsToTry
                                            Try
                                                Dim propVal As Object = CallByName(se, propName, CallType.Get)
                                                If propVal IsNot Nothing Then
                                                    Logger.Info("              " & propName & ": " & TypeName(propVal))
                                                End If
                                            Catch : End Try
                                        Next
                                        
                                        ' Check if this is a Proxy object
                                        If seType.EndsWith("Proxy") Then
                                            Logger.Info("              ** IS PROXY **")
                                            ' Try NativeObject
                                            Try
                                                Dim nativeObj As Object = CallByName(se, "NativeObject", CallType.Get)
                                                If nativeObj IsNot Nothing Then
                                                    Logger.Info("              NativeObject: " & TypeName(nativeObj))
                                                End If
                                            Catch : End Try
                                        End If
                                    End If
                                End If
                            Catch : End Try
                        Next
                        Logger.Info("          Total entities with ref properties: " & refEntsFound & " of " & entIdx)
                        
                        ' Check SketchLines with Reference=True - log ALL of them with full details
                        Logger.Info("        SketchLines with Reference=True:")
                        Dim refLineIdx As Integer = 0
                        For Each sl As SketchLine In sk.SketchLines
                            Try
                                Dim isRef As Boolean = False
                                Try : isRef = CBool(CallByName(sl, "Reference", CallType.Get)) : Catch : End Try
                                If isRef Then
                                    refLineIdx += 1
                                    Dim lineName As String = ""
                                    Try : lineName = sl.Name : Catch : lineName = "Line" & refLineIdx : End Try
                                    
                                    Dim refEnt As Object = Nothing
                                    Try : refEnt = CallByName(sl, "ReferencedEntity", CallType.Get) : Catch : End Try
                                    
                                    If refEnt IsNot Nothing Then
                                        Dim refType As String = TypeName(refEnt)
                                        Logger.Info("          [" & refLineIdx & "] " & lineName & " -> " & refType)
                                        
                                        ' Check for ContainingOccurrence on the referenced entity
                                        Dim contOcc As Object = Nothing
                                        Try : contOcc = CallByName(refEnt, "ContainingOccurrence", CallType.Get) : Catch : End Try
                                        If contOcc IsNot Nothing Then
                                            Dim occName As String = CStr(CallByName(contOcc, "Name", CallType.Get))
                                            Logger.Info("              ContainingOccurrence: " & occName)
                                        End If
                                        
                                        ' Trace parent chain
                                        Dim parent As Object = refEnt
                                        Dim depth As Integer = 0
                                        Do While parent IsNot Nothing AndAlso depth < 5
                                            Try
                                                parent = CallByName(parent, "Parent", CallType.Get)
                                                If parent IsNot Nothing Then
                                                    depth += 1
                                                    Dim pType As String = TypeName(parent)
                                                    Dim pInfo As String = pType
                                                    Try
                                                        Dim pName As String = CStr(CallByName(parent, "Name", CallType.Get))
                                                        pInfo &= " (" & pName & ")"
                                                    Catch : End Try
                                                    Try
                                                        Dim pPath As String = CStr(CallByName(parent, "FullFileName", CallType.Get))
                                                        pInfo &= " -> " & System.IO.Path.GetFileName(pPath)
                                                    Catch : End Try
                                                    Logger.Info("              Parent[" & depth & "]: " & pInfo)
                                                End If
                                            Catch
                                                Exit Do
                                            End Try
                                        Loop
                                    Else
                                        Logger.Info("          [" & refLineIdx & "] " & lineName & " -> (no ReferencedEntity)")
                                    End If
                                End If
                            Catch : End Try
                        Next
                        If refLineIdx = 0 Then Logger.Info("          (none)")
                        
                        ' Check SketchSplines with Reference=True
                        Logger.Info("        SketchSplines with Reference=True:")
                        Dim refSplineIdx As Integer = 0
                        For Each ss As SketchSpline In sk.SketchSplines
                            Try
                                Dim isRef As Boolean = False
                                Try : isRef = CBool(CallByName(ss, "Reference", CallType.Get)) : Catch : End Try
                                If isRef Then
                                    refSplineIdx += 1
                                    Dim refEnt As Object = Nothing
                                    Try : refEnt = CallByName(ss, "ReferencedEntity", CallType.Get) : Catch : End Try
                                    If refEnt IsNot Nothing AndAlso refSplineIdx <= 5 Then
                                        Logger.Info("          [" & refSplineIdx & "] -> " & TypeName(refEnt))
                                    End If
                                End If
                            Catch : End Try
                        Next
                        Logger.Info("          Total: " & refSplineIdx)
                        
                    Catch ex As Exception
                        Logger.Warn("        Error: " & ex.Message)
                    End Try
                    Exit For ' Only analyze first "Põhi J" sketch
                Next
                
                For Each sketch As PlanarSketch In partDef.Sketches
                    For Each entity As SketchEntity In sketch.SketchEntities
                        Try
                            Dim isRef As Boolean = False
                            Try : isRef = CBool(CallByName(entity, "Reference", CallType.Get)) : Catch : End Try
                            If Not isRef Then Continue For
                            totalRefEntities += 1
                            
                            ' Try multiple ways to get source info
                            Dim refEnt As Object = Nothing
                            Try : refEnt = CallByName(entity, "ReferencedEntity", CallType.Get) : Catch : End Try
                            If refEnt Is Nothing Then Continue For
                            
                            Dim refTypeName As String = TypeName(refEnt)
                            Dim sourcePath As String = Nothing
                            Dim sourceOccName As String = Nothing
                            
                            ' Method 1: ContainingOccurrence (for proxy objects)
                            Try
                                Dim containingOcc As Object = CallByName(refEnt, "ContainingOccurrence", CallType.Get)
                                If containingOcc IsNot Nothing Then
                                    sourceOccName = CStr(CallByName(containingOcc, "Name", CallType.Get))
                                    Dim occDef As Object = CallByName(containingOcc, "Definition", CallType.Get)
                                    Dim sourceDoc As Object = CallByName(occDef, "Document", CallType.Get)
                                    sourcePath = CStr(CallByName(sourceDoc, "FullFileName", CallType.Get))
                                End If
                            Catch : End Try
                            
                            ' Method 2: NativeObject -> Parent chain
                            If sourcePath Is Nothing Then
                                Try
                                    Dim nativeObj As Object = CallByName(refEnt, "NativeObject", CallType.Get)
                                    If nativeObj IsNot Nothing Then
                                        Dim parent As Object = CallByName(nativeObj, "Parent", CallType.Get)
                                        If parent IsNot Nothing Then
                                            Dim parentDoc As Object = CallByName(parent, "Document", CallType.Get)
                                            If parentDoc IsNot Nothing Then
                                                sourcePath = CStr(CallByName(parentDoc, "FullFileName", CallType.Get))
                                            End If
                                        End If
                                    End If
                                Catch : End Try
                            End If
                            
                            ' Method 3: Direct Parent chain
                            If sourcePath Is Nothing Then
                                Try
                                    Dim parent As Object = CallByName(refEnt, "Parent", CallType.Get)
                                    If parent IsNot Nothing Then
                                        Dim parentDoc As Object = Nothing
                                        Try : parentDoc = CallByName(parent, "Document", CallType.Get) : Catch : End Try
                                        If parentDoc IsNot Nothing Then
                                            sourcePath = CStr(CallByName(parentDoc, "FullFileName", CallType.Get))
                                        End If
                                    End If
                                Catch : End Try
                            End If
                            
                            ' Track if we found any source
                            If sourcePath IsNot Nothing Then
                                refEntWithSource += 1
                                ' Log first few detailed refs per part
                                If refEntWithSource <= 5 Then
                                    Logger.Info("        Ref in '" & sketch.Name & "': " & refTypeName)
                                    Logger.Info("          Source: " & System.IO.Path.GetFileName(sourcePath))
                                    If sourceOccName IsNot Nothing Then Logger.Info("          Occurrence: " & sourceOccName)
                                End If
                            End If
                            
                            ' Log if we found a reference to another master
                            If sourcePath IsNot Nothing AndAlso Not sourcePath.Equals(masterPath, StringComparison.OrdinalIgnoreCase) Then
                                If discoveredMasters.Contains(sourcePath) Then
                                    If Not projectedGeomChains.Any(Function(c) c.Item1.Equals(sourcePath, StringComparison.OrdinalIgnoreCase) AndAlso c.Item2.Equals(asmPath, StringComparison.OrdinalIgnoreCase) AndAlso c.Item3.Equals(masterPath, StringComparison.OrdinalIgnoreCase)) Then
                                        projectedGeomChains.Add(Tuple.Create(sourcePath, asmPath, masterPath))
                                        assembliesWithProjectedGeom.Add(asmPath)
                                        Logger.Info("      *** CHAIN: " & System.IO.Path.GetFileName(sourcePath) & " -> " & System.IO.Path.GetFileName(masterPath))
                                        Logger.Info("          Sketch: " & sketch.Name & ", Type: " & refTypeName)
                                        If sourceOccName IsNot Nothing Then Logger.Info("          Occurrence: " & sourceOccName)
                                        foundRefsInPart += 1
                                    End If
                                End If
                            End If
                        Catch : End Try
                    Next
                Next
                
                Logger.Info("      Ref entities: " & totalRefEntities & ", with source: " & refEntWithSource & ", cross-master: " & foundRefsInPart)
            Next
        Catch ex As Exception
            Logger.Warn("    Error: " & ex.Message)
        End Try
    Next
    Logger.Info("")
    
    ' Print discovered projected geometry chains
    Logger.Info("=== PROJECTED GEOMETRY CHAINS FOUND ===")
    If projectedGeomChains.Count > 0 Then
        For Each chain In projectedGeomChains
            Logger.Info("  " & System.IO.Path.GetFileName(chain.Item1) & " -> (via " & System.IO.Path.GetFileName(chain.Item2) & ") -> " & System.IO.Path.GetFileName(chain.Item3))
        Next
    Else
        Logger.Info("  (no projected geometry chains detected)")
        Logger.Info("  Falling back to dependency-based assembly detection")
    End If
    Logger.Info("")
    
    ' Step 3b: Filter candidate assemblies
    ' An assembly is needed if:
    ' 1. It was explicitly found in a projected geometry chain, OR
    ' 2. It contains a master AND one of that master's dependents (for dependency graph purposes)
    Logger.Info("=== STEP 3b: Filtering candidate assemblies ===")
    Logger.Info("Candidate assemblies found: " & candidateAssemblies.Count)
    Logger.Info("Assemblies with actual projected geometry: " & assembliesWithProjectedGeom.Count)
    
    For Each kvp In candidateAssemblies
        Dim asmPath As String = kvp.Key
        Dim containedMasters As List(Of String) = kvp.Value
        
        Logger.Info("  Checking: " & System.IO.Path.GetFileName(asmPath))
        Logger.Info("    Contains masters: " & String.Join(", ", containedMasters.Select(Function(m) System.IO.Path.GetFileName(m))))
        
        ' First check: is this assembly explicitly in a projected geometry chain?
        Dim hasProjectedGeom As Boolean = assembliesWithProjectedGeom.Contains(asmPath)
        If hasProjectedGeom Then
            Logger.Info("    ** HAS ACTUAL PROJECTED GEOMETRY - definitely needed")
            intermediateAssemblies.Add(asmPath)
            Continue For
        End If
        
        ' Second check: does this assembly bridge a dependency (containment-based)
        Dim isNeeded As Boolean = False
        For Each master In containedMasters
            ' Check if any other contained master depends on this one
            For Each otherMaster In containedMasters
                If Not otherMaster.Equals(master, StringComparison.OrdinalIgnoreCase) Then
                    ' Check if otherMaster depends on master
                    If masterDependencies.ContainsKey(otherMaster) Then
                        If masterDependencies(otherMaster).Any(Function(dep) dep.Equals(master, StringComparison.OrdinalIgnoreCase)) Then
                            isNeeded = True
                            Logger.Info("    -> NEEDED (dependency bridge): " & System.IO.Path.GetFileName(otherMaster) & " depends on " & System.IO.Path.GetFileName(master))
                            Exit For
                        End If
                    End If
                End If
            Next
            If isNeeded Then Exit For
        Next
        
        If isNeeded Then
            intermediateAssemblies.Add(asmPath)
        Else
            Logger.Info("    -> NOT NEEDED (just contains masters, no actual references)")
        End If
    Next
    
    Logger.Info("")
    Logger.Info("Intermediate assemblies (for geometry propagation): " & intermediateAssemblies.Count)
    Logger.Info("")
    
    ' Step 3c: Find intermediate assemblies by testing where projected geometry RESOLVES
    ' The correct intermediate assembly is the one where the projected geometry references
    ' actually work (the occurrence is found by exact name match).
    Logger.Info("=== STEP 3c: Finding intermediate assemblies (resolution test) ===")
    Dim containingAssemblies As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
    
    ' Build list of master dependency pairs that need intermediate assemblies
    ' Source 1: Unresolved occurrences from browser node analysis
    ' Source 2: Derivation chain from masterDependencies (fallback when masters aren't in root assembly)
    Dim dependencyPairs As New List(Of Tuple(Of String, String)) ' (DependentMaster, SourceMaster)
    
    ' First, add pairs from unresolved occurrences
    For Each unres In unresolvedOccurrences
        ' unres = (OccurrenceName, PartNumber, DependentMasterPath)
        For Each m In discoveredMasters
            If System.IO.Path.GetFileNameWithoutExtension(m) = unres.Item2 Then
                Dim pair = Tuple.Create(unres.Item3, m)
                If Not dependencyPairs.Any(Function(p) p.Item1.Equals(pair.Item1, StringComparison.OrdinalIgnoreCase) AndAlso p.Item2.Equals(pair.Item2, StringComparison.OrdinalIgnoreCase)) Then
                    dependencyPairs.Add(pair)
                End If
                Exit For
            End If
        Next
    Next
    
    ' Second, add pairs from derivation chain (fallback for when masters aren't in root assembly)
    ' This ensures we find intermediate assemblies even without browser node analysis
    For Each kvp In masterDependencies
        Dim dependentMaster As String = kvp.Key
        For Each dep In kvp.Value
            ' Only consider .ipt dependencies (not .iam)
            If dep.ToLower().EndsWith(".ipt") AndAlso discoveredMasters.Contains(dep) Then
                Dim pair = Tuple.Create(dependentMaster, dep)
                If Not dependencyPairs.Any(Function(p) p.Item1.Equals(pair.Item1, StringComparison.OrdinalIgnoreCase) AndAlso p.Item2.Equals(pair.Item2, StringComparison.OrdinalIgnoreCase)) Then
                    dependencyPairs.Add(pair)
                    Logger.Info("  Added from derivation chain: " & System.IO.Path.GetFileName(dependentMaster) & " -> " & System.IO.Path.GetFileName(dep))
                End If
            End If
        Next
    Next
    
    If dependencyPairs.Count = 0 Then
        Logger.Info("  No dependency pairs to search for")
    Else
        Logger.Info("  Dependency pairs to resolve:")
        For Each pair In dependencyPairs
            Logger.Info("    " & System.IO.Path.GetFileName(pair.Item1) & " depends on " & System.IO.Path.GetFileName(pair.Item2))
        Next
        
        ' Derive product family root
        Dim productFamilyRoot As String = ""
        Try
            Dim asmFolder As String = System.IO.Path.GetDirectoryName(asmDoc.FullFileName)
            Dim parts() As String = asmFolder.Split(System.IO.Path.DirectorySeparatorChar)
            For i As Integer = parts.Length - 1 To 0 Step -1
                If parts(i).Equals("Aluselemendid", StringComparison.OrdinalIgnoreCase) Then
                    productFamilyRoot = String.Join(System.IO.Path.DirectorySeparatorChar, parts.Take(i))
                    Exit For
                End If
            Next
        Catch
        End Try
        
        If String.IsNullOrEmpty(productFamilyRoot) Then
            Logger.Warn("    Could not determine product family root folder")
        Else
            Logger.Info("    Product family root: " & productFamilyRoot)
            
            Try
                Dim allAssemblies = System.IO.Directory.GetFiles(productFamilyRoot, "*.iam", System.IO.SearchOption.AllDirectories)
                Logger.Info("    Scanning " & allAssemblies.Length & " assemblies...")
                
                Dim currentAsmPath As String = asmDoc.FullFileName
                Dim testedCount As Integer = 0
                
                For Each iamFile In allAssemblies
                    ' Skip already processed or invalid
                    If iamFile.Equals(currentAsmPath, StringComparison.OrdinalIgnoreCase) Then Continue For
                    If intermediateAssemblies.Contains(iamFile) Then Continue For
                    If IsVaultOldVersion(iamFile) Then Continue For
                    
                    Try
                        ' First, quick check: does this assembly contain BOTH masters of any dependency pair?
                        ' Open invisibly first for the quick containment check
                        Dim extAsm As AssemblyDocument = CType(app.Documents.Open(iamFile, False), AssemblyDocument)
                        Dim containedFiles As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                        For Each refDoc As Document In extAsm.AllReferencedDocuments
                            containedFiles.Add(refDoc.FullFileName)
                        Next
                        
                        ' Check each dependency pair
                        For Each pair In dependencyPairs
                            ' Skip if this pair is already resolved
                            If projectedGeomChains.Any(Function(c) c.Item1.Equals(pair.Item2, StringComparison.OrdinalIgnoreCase) AndAlso c.Item3.Equals(pair.Item1, StringComparison.OrdinalIgnoreCase)) Then
                                Continue For
                            End If
                            
                            ' Does assembly contain BOTH the dependent and source master?
                            Dim hasDep As Boolean = containedFiles.Contains(pair.Item1)
                            Dim hasSrc As Boolean = containedFiles.Contains(pair.Item2)
                            If Not hasDep OrElse Not hasSrc Then
                                ' Log why we're skipping (only for first few to avoid spam)
                                If testedCount < 3 Then
                                    Logger.Info("  Skipping " & System.IO.Path.GetFileName(iamFile) & ": hasDep=" & hasDep & " hasSrc=" & hasSrc)
                                End If
                                Continue For
                            End If
                            
                            testedCount += 1
                            Logger.Info("  Testing: " & System.IO.Path.GetFileName(iamFile))
                            Logger.Info("    Dependent: " & pair.Item1)
                            Logger.Info("    Source: " & pair.Item2)
                            
                            Try
                                ' Activate assembly for browser node access
                                ' Must activate to access browser panes
                                extAsm.Activate()
                                app.ActiveView.Fit()
                            Catch actEx As Exception
                                Logger.Warn("    Failed to activate: " & actEx.Message)
                                Continue For
                            End Try
                            
                            ' Build occurrence map
                            Dim occByName As New Dictionary(Of String, ComponentOccurrence)(StringComparer.OrdinalIgnoreCase)
                            Try
                                For Each occ As ComponentOccurrence In extAsm.ComponentDefinition.Occurrences.AllLeafOccurrences
                                    Try
                                        occByName(occ.Name) = occ
                                    Catch : End Try
                                Next
                            Catch occEx As Exception
                                Logger.Warn("    Failed to get occurrences: " & occEx.Message)
                                Continue For
                            End Try
                            Logger.Info("    Occurrences in assembly: " & occByName.Count)
                            
                            ' Find the dependent master occurrence
                            Dim dependentOcc As ComponentOccurrence = Nothing
                            For Each kvp In occByName
                                Try
                                    If kvp.Value.Definition.Document.FullFileName.Equals(pair.Item1, StringComparison.OrdinalIgnoreCase) Then
                                        dependentOcc = kvp.Value
                                        Logger.Info("    Found dependent occurrence: " & kvp.Key)
                                        Exit For
                                    End If
                                Catch : End Try
                            Next
                            
                            If dependentOcc Is Nothing Then
                                Logger.Info("    Dependent master not found as occurrence")
                                Continue For
                            End If
                            
                            ' Check if projected geometry references RESOLVE in this assembly
                            Dim partDef As PartComponentDefinition = Nothing
                            Try
                                partDef = CType(dependentOcc.Definition, PartComponentDefinition)
                            Catch castEx As Exception
                                Logger.Warn("    Failed to get PartComponentDefinition: " & castEx.Message)
                                Continue For
                            End Try
                            
                            Dim referencesResolved As Boolean = False
                            Dim sourceOccName As String = ""
                            
                            ' Find the browser node for this occurrence
                            ' In assembly context, the part's sketches are under the occurrence node
                            ' Match by occurrence name since COM object comparison with Is may not work
                            Dim occBrowserNode As BrowserNode = Nothing
                            Dim dependentOccName As String = dependentOcc.Name
                            Logger.Info("    Looking for browser node of: " & dependentOccName)
                            Try
                                For Each pane As BrowserPane In app.ActiveDocument.BrowserPanes
                                    Try
                                        For Each topNode As BrowserNode In pane.TopNode.BrowserNodes
                                            Try
                                                Dim nodeLabel As String = topNode.BrowserNodeDefinition.Label
                                                ' Match occurrence name (e.g., "Nurk - Eskiis (000131):1")
                                                If nodeLabel.Equals(dependentOccName, StringComparison.OrdinalIgnoreCase) Then
                                                    occBrowserNode = topNode
                                                    Exit For
                                                End If
                                            Catch : End Try
                                        Next
                                    Catch : End Try
                                    If occBrowserNode IsNot Nothing Then Exit For
                                Next
                            Catch : End Try
                            
                            If occBrowserNode Is Nothing Then
                                Logger.Info("    Occurrence browser node not found - trying sketch approach")
                            Else
                                Logger.Info("    Found occurrence browser node: " & occBrowserNode.BrowserNodeDefinition.Label)
                            End If
                            
                            ' Check sketches for browser nodes with cross-part references
                            ' Method 1: Look through the occurrence's browser nodes for sketches
                            Dim sketchesChecked As Integer = 0
                            
                            If occBrowserNode IsNot Nothing Then
                                ' Recursively find sketch nodes under the occurrence
                                For Each childNode As BrowserNode In occBrowserNode.BrowserNodes
                                    Try
                                        Dim childLabel As String = childNode.BrowserNodeDefinition.Label
                                        
                                        ' Check if this node has reference children
                                        For Each refNode As BrowserNode In childNode.BrowserNodes
                                            Try
                                                Dim refLabel As String = refNode.BrowserNodeDefinition.Label
                                                If refLabel.StartsWith("Reference") AndAlso refLabel.Contains("(") Then
                                                    sketchesChecked += 1
                                                    
                                                    ' Extract occurrence name from label
                                                    Dim parenStart As Integer = refLabel.IndexOf("(")
                                                    Dim parenEnd As Integer = refLabel.LastIndexOf(")")
                                                    If parenEnd > parenStart Then
                                                        Dim occNameFromLabel As String = refLabel.Substring(parenStart + 1, parenEnd - parenStart - 1)
                                                        
                                                        ' Does this occurrence exist in the assembly?
                                                        If occByName.ContainsKey(occNameFromLabel) Then
                                                            Dim refOcc As ComponentOccurrence = occByName(occNameFromLabel)
                                                            Dim refPath As String = refOcc.Definition.Document.FullFileName
                                                            
                                                            ' Does it point to our source master?
                                                            If refPath.Equals(pair.Item2, StringComparison.OrdinalIgnoreCase) Then
                                                                referencesResolved = True
                                                                sourceOccName = occNameFromLabel
                                                                Logger.Info("    ** REFERENCE RESOLVED: " & refLabel & " -> " & occNameFromLabel)
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Catch : End Try
                                        Next
                                        
                                        If referencesResolved Then Exit For
                                    Catch : End Try
                                Next
                            End If
                            
                            ' Method 2: If no occurrence browser node, try the old approach with sketches
                            If Not referencesResolved AndAlso occBrowserNode Is Nothing Then
                                For Each sk As PlanarSketch In partDef.Sketches
                                    Try
                                        ' Try to find this sketch's browser node by matching name
                                        Dim sketchName As String = sk.Name
                                        Dim sketchNode As BrowserNode = Nothing
                                        
                                        Try
                                            For Each pane As BrowserPane In app.ActiveDocument.BrowserPanes
                                                Try
                                                    For Each topNode As BrowserNode In pane.TopNode.BrowserNodes
                                                        Try
                                                            Dim nodeLabel As String = topNode.BrowserNodeDefinition.Label
                                                            If nodeLabel.Contains(sketchName) Then
                                                                sketchNode = topNode
                                                                Exit For
                                                            End If
                                                        Catch : End Try
                                                    Next
                                                Catch : End Try
                                                If sketchNode IsNot Nothing Then Exit For
                                            Next
                                        Catch : End Try
                                        
                                        If sketchNode Is Nothing Then Continue For
                                        sketchesChecked += 1
                                        
                                        ' Check child nodes for reference labels
                                        For Each childNode As BrowserNode In sketchNode.BrowserNodes
                                            Try
                                                Dim label As String = childNode.BrowserNodeDefinition.Label
                                                If label.StartsWith("Reference") AndAlso label.Contains("(") Then
                                                    Dim parenStart As Integer = label.IndexOf("(")
                                                    Dim parenEnd As Integer = label.LastIndexOf(")")
                                                    If parenEnd > parenStart Then
                                                        Dim occNameFromLabel As String = label.Substring(parenStart + 1, parenEnd - parenStart - 1)
                                                        
                                                        If occByName.ContainsKey(occNameFromLabel) Then
                                                            Dim refOcc As ComponentOccurrence = occByName(occNameFromLabel)
                                                            Dim refPath As String = refOcc.Definition.Document.FullFileName
                                                            
                                                            If refPath.Equals(pair.Item2, StringComparison.OrdinalIgnoreCase) Then
                                                                referencesResolved = True
                                                                sourceOccName = occNameFromLabel
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Catch : End Try
                                        Next
                                        
                                        If referencesResolved Then Exit For
                                    Catch : End Try
                                Next
                            End If
                            
                            Logger.Info("    Sketches/nodes checked: " & sketchesChecked)
                            
                            If referencesResolved Then
                                Logger.Info("    ** VERIFIED: References resolve in this assembly!")
                                Logger.Info("       Source occurrence: " & sourceOccName)
                                intermediateAssemblies.Add(iamFile)
                                assembliesWithProjectedGeom.Add(iamFile)
                                projectedGeomChains.Add(Tuple.Create(pair.Item2, iamFile, pair.Item1))
                                Logger.Info("       *** CHAIN: " & System.IO.Path.GetFileName(pair.Item2) & " --(via " & System.IO.Path.GetFileName(iamFile) & ")--> " & System.IO.Path.GetFileName(pair.Item1))
                            Else
                                Logger.Info("    (references do not resolve)")
                            End If
                        Next
                    Catch openEx As Exception
                        Logger.Warn("  Error testing " & System.IO.Path.GetFileName(iamFile) & ": " & openEx.Message)
                    End Try
                Next
                
                Logger.Info("    Tested " & testedCount & " candidate assemblies")
            Catch ex As Exception
                Logger.Warn("    Error searching product family: " & ex.Message)
            End Try
        End If
    End If
    
    Logger.Info("")
    Logger.Info("=== INTERMEDIATE ASSEMBLIES FOUND ===")
    Logger.Info("Total: " & intermediateAssemblies.Count)
    If intermediateAssemblies.Count > 0 Then
        For Each intAsm In intermediateAssemblies
            Dim loc As String = If(IsInsideFolder(intAsm, sourceRoot), "[INT]", "[EXT]")
            Logger.Info("  " & loc & " " & System.IO.Path.GetFileName(intAsm))
        Next
    End If
    Logger.Info("")
    
    Logger.Info("=== PROJECTED GEOMETRY CHAINS ===")
    If projectedGeomChains.Count > 0 Then
        For Each chain In projectedGeomChains
            Logger.Info("  " & System.IO.Path.GetFileName(chain.Item1) & " --(via " & System.IO.Path.GetFileName(chain.Item2) & ")--> " & System.IO.Path.GetFileName(chain.Item3))
        Next
    Else
        Logger.Info("  (none found)")
    End If
    Logger.Info("")
    
    ' Step 4: Classify all discovered files
    Logger.Info("=== STEP 4: Classification ===")
    Dim internalCount As Integer = 0
    Dim externalCount As Integer = 0
    
    For Each m In discoveredMasters
        If IsInsideFolder(m, sourceRoot) Then
            internalCount += 1
            Logger.Info("  [INTERNAL] " & System.IO.Path.GetFileName(m))
        Else
            externalCount += 1
            Logger.Info("  [EXTERNAL] " & m)
        End If
    Next
    
    Logger.Info("")
    Logger.Info("Internal: " & internalCount)
    Logger.Info("External: " & externalCount)
    Logger.Info("")
    
    ' Step 5: Build and display dependency tree
    Logger.Info("=== STEP 5: Dependency Tree ===")
    
    ' Find root masters (those with no dependencies or only external dependencies)
    Dim rootMasters As New List(Of String)
    For Each m In discoveredMasters
        If Not masterDependencies.ContainsKey(m) OrElse masterDependencies(m).Count = 0 Then
            rootMasters.Add(m)
        Else
            ' Check if all dependencies are non-master files
            Dim hasMAsterDep As Boolean = False
            For Each dep In masterDependencies(m)
                If discoveredMasters.Contains(dep) Then
                    hasMAsterDep = True
                    Exit For
                End If
            Next
            If Not hasMAsterDep Then
                rootMasters.Add(m)
            End If
        End If
    Next
    
    Logger.Info("Root masters (no dependencies): " & rootMasters.Count)
    For Each root In rootMasters
        PrintDependencyTree(root, masterDependencies, discoveredMasters, 0, New HashSet(Of String))
    Next
    
    ' Step 6: Suggest copy order (topological sort)
    Logger.Info("")
    Logger.Info("=== STEP 6: Suggested Copy Order (topological) ===")
    
    ' Add intermediate assemblies to the set for sorting
    Dim allFilesToCopy As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    For Each m In discoveredMasters
        allFilesToCopy.Add(m)
    Next
    For Each a In intermediateAssemblies
        allFilesToCopy.Add(a)
        ' Add assembly dependencies: assembly depends on the masters it contains
        If containingAssemblies.ContainsKey(a) Then
            If Not masterDependencies.ContainsKey(a) Then
                masterDependencies(a) = New List(Of String)
            End If
            For Each contained In containingAssemblies(a)
                If Not masterDependencies(a).Contains(contained) Then
                    masterDependencies(a).Add(contained)
                End If
            Next
        End If
    Next
    
    Dim copyOrder As List(Of String) = TopologicalSort(allFilesToCopy, masterDependencies)
    Dim orderNum As Integer = 1
    For Each m In copyOrder
        Dim location As String = If(IsInsideFolder(m, sourceRoot), "[INT]", "[EXT]")
        Dim mType As String = If(m.ToLower().EndsWith(".iam"), "(asm)", "(ipt)")
        Logger.Info("  " & orderNum & ". " & location & " " & mType & " " & System.IO.Path.GetFileName(m))
        orderNum += 1
    Next
    
    ' Summary
    Logger.Info("")
    Logger.Info("=== SUMMARY ===")
    Logger.Info("Master parts to copy: " & discoveredMasters.Count)
    Logger.Info("Intermediate assemblies to copy: " & intermediateAssemblies.Count)
    Logger.Info("Total files: " & allFilesToCopy.Count)
    Logger.Info("")
    
    Dim internalParts As Integer = 0
    Dim externalParts As Integer = 0
    Dim internalAsm As Integer = 0
    Dim externalAsm As Integer = 0
    
    For Each f In allFilesToCopy
        Dim isInternal As Boolean = IsInsideFolder(f, sourceRoot)
        Dim isAsm As Boolean = f.ToLower().EndsWith(".iam")
        If isAsm Then
            If isInternal Then internalAsm += 1 Else externalAsm += 1
        Else
            If isInternal Then internalParts += 1 Else externalParts += 1
        End If
    Next
    
    Logger.Info("Internal parts: " & internalParts)
    Logger.Info("External parts: " & externalParts)
    Logger.Info("Internal assemblies: " & internalAsm)
    Logger.Info("External assemblies: " & externalAsm)
    
    If externalParts > 0 OrElse externalAsm > 0 Then
        Logger.Info("")
        Logger.Info("** EXTERNAL FILES (will be copied to each element): **")
        For Each f In allFilesToCopy
            If Not IsInsideFolder(f, sourceRoot) Then
                Logger.Info("  " & f)
            End If
        Next
    End If
    
    ' Show projected geometry chains summary
    Logger.Info("")
    Logger.Info("** PROJECTED GEOMETRY REFERENCE CHAINS: **")
    If projectedGeomChains.Count > 0 Then
        Logger.Info("These chains show actual sketch/geometry references through assemblies:")
        For Each chain In projectedGeomChains
            Logger.Info("  " & System.IO.Path.GetFileName(chain.Item1) & " --(via " & System.IO.Path.GetFileName(chain.Item2) & ")--> " & System.IO.Path.GetFileName(chain.Item3))
        Next
    Else
        Logger.Info("  (no projected geometry chains found - intermediate assemblies determined by dependency relationships)")
    End If
    
    Logger.Info("")
    Logger.Info("=== DIAGNOSTIC COMPLETE ===")
    
    MessageBox.Show("Diagnostic complete. Check iLogic log for results.", "TestMasterDependencies")
End Sub

''' <summary>
''' Find element source root (Aluselemendid/{ElementName}/)
''' </summary>
Function FindElementSourceRoot(asmPath As String) As String
    Dim folder As String = System.IO.Path.GetDirectoryName(asmPath)
    
    Do While Not String.IsNullOrEmpty(folder)
        Dim parentName As String = System.IO.Path.GetFileName(System.IO.Path.GetDirectoryName(folder))
        If parentName.Equals("Aluselemendid", StringComparison.OrdinalIgnoreCase) OrElse _
           parentName.Equals("Alusmoodulid", StringComparison.OrdinalIgnoreCase) Then
            Return folder
        End If
        folder = System.IO.Path.GetDirectoryName(folder)
    Loop
    
    Return System.IO.Path.GetDirectoryName(asmPath)
End Function

''' <summary>
''' Check if a path is inside a folder
''' </summary>
Function IsInsideFolder(filePath As String, folder As String) As Boolean
    If String.IsNullOrEmpty(folder) OrElse String.IsNullOrEmpty(filePath) Then Return False
    folder = folder.TrimEnd("\"c) & "\"
    Return filePath.StartsWith(folder, StringComparison.OrdinalIgnoreCase)
End Function

''' <summary>
''' Collect all parts from an assembly (including nested)
''' </summary>
Sub CollectAllParts(app As Inventor.Application, asmDoc As AssemblyDocument, parts As Dictionary(Of String, PartDocument))
    For Each refDoc As Document In asmDoc.AllReferencedDocuments
        If refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
            If Not parts.ContainsKey(refDoc.FullFileName) Then
                parts.Add(refDoc.FullFileName, CType(refDoc, PartDocument))
            End If
        End If
    Next
End Sub

''' <summary>
''' Get the master path from a derived part
''' </summary>
Function GetMasterPath(partDoc As PartDocument) As String
    Try
        Dim dpcs = partDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents
        If dpcs.Count > 0 AndAlso partDoc.ReferencedDocuments.Count > 0 Then
            Return partDoc.ReferencedDocuments.Item(1).FullFileName
        End If
    Catch
    End Try
    Return Nothing
End Function

''' <summary>
''' Print dependency tree recursively
''' </summary>
Sub PrintDependencyTree(node As String, deps As Dictionary(Of String, List(Of String)), allNodes As HashSet(Of String), indent As Integer, visited As HashSet(Of String))
    If visited.Contains(node) Then
        Logger.Info(New String(" "c, indent * 2) & "-> " & System.IO.Path.GetFileName(node) & " (circular ref)")
        Return
    End If
    visited.Add(node)
    
    Logger.Info(New String(" "c, indent * 2) & "-> " & System.IO.Path.GetFileName(node))
    
    If deps.ContainsKey(node) Then
        For Each dep In deps(node)
            If allNodes.Contains(dep) Then
                PrintDependencyTree(dep, deps, allNodes, indent + 1, visited)
            End If
        Next
    End If
    
    visited.Remove(node)
End Sub

''' <summary>
''' Topological sort of masters by dependency
''' </summary>
Function TopologicalSort(nodes As HashSet(Of String), deps As Dictionary(Of String, List(Of String))) As List(Of String)
    Dim result As New List(Of String)
    Dim visited As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Dim tempMark As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    
    For Each node In nodes
        If Not visited.Contains(node) Then
            TopologicalVisit(node, deps, nodes, visited, tempMark, result)
        End If
    Next
    
    Return result
End Function

Sub TopologicalVisit(node As String, deps As Dictionary(Of String, List(Of String)), allNodes As HashSet(Of String), _
                     visited As HashSet(Of String), tempMark As HashSet(Of String), result As List(Of String))
    If tempMark.Contains(node) Then Return ' Cycle detected, skip
    If visited.Contains(node) Then Return
    
    tempMark.Add(node)
    
    If deps.ContainsKey(node) Then
        For Each dep In deps(node)
            If allNodes.Contains(dep) Then
                TopologicalVisit(dep, deps, allNodes, visited, tempMark, result)
            End If
        Next
    End If
    
    tempMark.Remove(node)
    visited.Add(node)
    result.Add(node)
End Sub

''' <summary>
''' Check if a file path is a Vault old version (either in OldVersions folder or has version suffix)
''' Vault stores old versions in OldVersions folders and names them like 000114.0025.iam
''' </summary>
Function IsVaultOldVersion(filePath As String) As Boolean
    If String.IsNullOrEmpty(filePath) Then Return False
    
    ' Check 1: File is in an OldVersions folder
    If filePath.ToLower().Contains("\oldversions\") Then
        Return True
    End If
    
    ' Check 2: Filename has version suffix (e.g., 000114.0025.iam -> version is .0025)
    Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(filePath)
    If fileName.Contains(".") Then
        Dim parts() As String = fileName.Split("."c)
        If parts.Length > 1 Then
            ' Check if the last part is numeric (version number)
            Dim lastPart As String = parts(parts.Length - 1)
            Dim versionNum As Integer
            If Integer.TryParse(lastPart, versionNum) Then
                Return True
            End If
        End If
    End If
    
    Return False
End Function

''' <summary>
''' Check if an assembly is in the main assembly's reference tree
''' </summary>
Function IsInMainAssemblyTree(asmPath As String, mainAsm As AssemblyDocument) As Boolean
    ' Check if asmPath is the main assembly itself
    If asmPath.Equals(mainAsm.FullFileName, StringComparison.OrdinalIgnoreCase) Then
        Return True
    End If
    
    ' Check if asmPath is in the main assembly's AllReferencedDocuments
    Try
        For Each refDoc As Document In mainAsm.AllReferencedDocuments
            If refDoc.FullFileName.Equals(asmPath, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next
    Catch
    End Try
    
    ' Also check ReferencedFileDescriptors for suppressed assemblies
    Try
        For i As Integer = 1 To mainAsm.File.ReferencedFileDescriptors.Count
            Dim fd As FileDescriptor = mainAsm.File.ReferencedFileDescriptors.Item(i)
            If fd.FullFileName.Equals(asmPath, StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next
    Catch
    End Try
    
    Return False
End Function

''' <summary>
''' Extract assembly path from FullDocumentName which may contain assembly context
''' Format could be: "C:\path\assembly.iam|C:\path\part.ipt" or similar
''' </summary>
Function ExtractAssemblyPath(fullDocName As String) As String
    If String.IsNullOrEmpty(fullDocName) Then Return Nothing
    
    ' FullDocumentName might contain multiple paths separated by | for in-context references
    Dim parts() As String = fullDocName.Split("|"c)
    For Each p In parts
        Dim trimmed As String = p.Trim()
        If trimmed.ToLower().EndsWith(".iam") Then
            Return trimmed
        End If
    Next
    
    ' Also check for assembly paths in the middle of the string
    Dim iamPos As Integer = fullDocName.ToLower().IndexOf(".iam")
    If iamPos > 0 Then
        ' Find the start of the path (look backwards for drive letter or UNC)
        Dim startPos As Integer = 0
        For i As Integer = iamPos - 1 To 0 Step -1
            Dim c As Char = fullDocName(i)
            If c = "|"c OrElse c = ">"c OrElse c = "<"c Then
                startPos = i + 1
                Exit For
            End If
        Next
        Return fullDocName.Substring(startPos, iamPos + 4 - startPos)
    End If
    
    Return Nothing
End Function

''' <summary>
''' Extract assembly path from an OccurrencePath string by finding the assembly
''' that contains the referenced occurrence.
''' OccurrencePath format examples: "000114:1", "Selg - Eskiis Multibody (000130):1"
''' </summary>
Function ExtractAssemblyFromOccurrencePath(occPath As String, app As Inventor.Application) As String
    If String.IsNullOrEmpty(occPath) Then Return Nothing
    
    ' The occurrence path contains occurrence names - we need to find the assembly
    ' by looking through open assemblies
    Try
        For Each doc As Document In app.Documents
            If doc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                Dim asmDoc As AssemblyDocument = CType(doc, AssemblyDocument)
                Try
                    ' Check if this assembly has an occurrence matching part of the path
                    For Each occ As ComponentOccurrence In asmDoc.ComponentDefinition.Occurrences
                        If occPath.Contains(occ.Name) Then
                            Return asmDoc.FullFileName
                        End If
                    Next
                Catch : End Try
            End If
        Next
    Catch : End Try
    
    Return Nothing
End Function
