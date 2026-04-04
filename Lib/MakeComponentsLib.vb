' ============================================================================
' MakeComponentsLib - Core logic for creating derived parts from solid bodies
' 
' Provides functions to:
' - Detect axes for solid bodies (using BoundingBoxStockLib patterns)
' - Create derived parts from single bodies
' - Set iProperties on new parts
' - Assign materials to parts
' - Place components in assembly
'
' Usage: AddVbFile "Lib/MakeComponentsLib.vb"
'
' Note: Logger is not available in library modules.
'       Pass a List(Of String) to collect log messages.
' ============================================================================

Imports Inventor

Public Module MakeComponentsLib

    ' ============================================================================
    ' Body Info Class - holds detected properties for each body
    ' ============================================================================
    
    Public Class BodyInfo
        Public Body As SurfaceBody
        Public Name As String
        Public ThicknessVector As String
        Public ThicknessValue As Double
        Public WidthVector As String
        Public WidthValue As Double
        Public LengthVector As String
        Public LengthValue As Double
        Public ConvertToSheetMetal As Boolean
        Public MaterialName As String
        Public Selected As Boolean
        
        ' Part reference - path to created part (if exists)
        Public CreatedPartPath As String
        Public PartExists As Boolean
        
        ' Geometry signature for matching across body renames
        ' Format: "V:{volume};F:{faceCount};A:{surfaceArea}"
        Public Signature As String
        
        Public Sub New(b As SurfaceBody)
            Body = b
            Name = b.Name
            Selected = True
            ConvertToSheetMetal = False
            MaterialName = ""
            CreatedPartPath = ""
            PartExists = False
            Signature = ComputeBodySignature(b)
        End Sub
    End Class
    
    ' Compute a geometry signature for a body (for matching across renames)
    Public Function ComputeBodySignature(body As SurfaceBody) As String
        Try
            Dim volume As Double = 0
            Dim area As Double = 0
            Dim faceCount As Integer = body.Faces.Count
            
            Try : volume = body.Volume(0.001) : Catch : End Try
            Try : area = body.SurfaceArea(0.001) : Catch : End Try
            
            ' Round to reduce floating point differences
            Return String.Format("V:{0:F4};F:{1};A:{2:F4}", volume * 1000000, faceCount, area * 10000)
        Catch
            Return ""
        End Try
    End Function
    
    ' ============================================================================
    ' Body Data Storage - persist settings in master document
    ' ============================================================================
    
    Private Const PROP_PREFIX As String = "LK_Body_"
    Private Const GENERAL_PREFIX As String = "LK_General_"
    
    ' Stored body data class (what we save to master document)
    Public Class StoredBodyData
        Public Name As String
        Public Signature As String
        Public ThicknessVector As String
        Public WidthVector As String
        Public LengthVector As String
        Public ConvertToSheetMetal As Boolean
        Public MaterialName As String
        Public CreatedPartPath As String
    End Class
    
    ' General settings class (non-body-specific settings)
    Public Class GeneralSettings
        Public ProjectName As String
        Public Template As String
        Public Subfolder As String
        Public AssemblyAction As String  ' NONE, CREATE, UPDATE
        Public AssemblyPath As String
        
        Public Sub New()
            ProjectName = ""
            Template = "Part.ipt"
            Subfolder = "Detailid"
            AssemblyAction = "NONE"
            AssemblyPath = ""
        End Sub
    End Class
    
    ' Save general settings to master document
    Public Sub SaveGeneralSettings(masterDoc As PartDocument, _
                                   settings As GeneralSettings, _
                                   logs As System.Collections.Generic.List(Of String))
        Try
            Dim userProps As PropertySet = masterDoc.PropertySets.Item("Inventor User Defined Properties")
            
            SetOrAddProperty(userProps, GENERAL_PREFIX & "Project", settings.ProjectName)
            SetOrAddProperty(userProps, GENERAL_PREFIX & "Template", settings.Template)
            SetOrAddProperty(userProps, GENERAL_PREFIX & "Subfolder", settings.Subfolder)
            SetOrAddProperty(userProps, GENERAL_PREFIX & "AsmAction", settings.AssemblyAction)
            SetOrAddProperty(userProps, GENERAL_PREFIX & "AsmPath", settings.AssemblyPath)
            
            logs.Add("MakeComponentsLib: Saved general settings")
        Catch ex As Exception
            logs.Add("MakeComponentsLib: Failed to save general settings: " & ex.Message)
        End Try
    End Sub
    
    ' Load general settings from master document
    Public Function LoadGeneralSettings(masterDoc As PartDocument, _
                                        logs As System.Collections.Generic.List(Of String)) As GeneralSettings
        Dim settings As New GeneralSettings()
        
        Try
            Dim userProps As PropertySet = masterDoc.PropertySets.Item("Inventor User Defined Properties")
            
            settings.ProjectName = GetPropertyValue(userProps, GENERAL_PREFIX & "Project", "")
            settings.Template = GetPropertyValue(userProps, GENERAL_PREFIX & "Template", "Part.ipt")
            settings.Subfolder = GetPropertyValue(userProps, GENERAL_PREFIX & "Subfolder", "Detailid")
            settings.AssemblyAction = GetPropertyValue(userProps, GENERAL_PREFIX & "AsmAction", "NONE")
            settings.AssemblyPath = GetPropertyValue(userProps, GENERAL_PREFIX & "AsmPath", "")
            
            ' Check if stored assembly still exists
            If Not String.IsNullOrEmpty(settings.AssemblyPath) Then
                If System.IO.File.Exists(settings.AssemblyPath) Then
                    logs.Add("MakeComponentsLib: Found existing assembly: " & _
                             System.IO.Path.GetFileName(settings.AssemblyPath))
                Else
                    logs.Add("MakeComponentsLib: Stored assembly not found, resetting")
                    settings.AssemblyPath = ""
                    settings.AssemblyAction = "NONE"
                End If
            End If
            
            If Not String.IsNullOrEmpty(settings.ProjectName) Then
                logs.Add("MakeComponentsLib: Loaded general settings - Project: " & settings.ProjectName)
            End If
        Catch ex As Exception
            logs.Add("MakeComponentsLib: No general settings found or error: " & ex.Message)
        End Try
        
        Return settings
    End Function
    
    ' Save body data to master document properties
    Public Sub SaveBodyDataToMaster(masterDoc As PartDocument, _
                                    bodies As System.Collections.Generic.List(Of BodyInfo), _
                                    logs As System.Collections.Generic.List(Of String))
        Try
            Dim userProps As PropertySet = masterDoc.PropertySets.Item("Inventor User Defined Properties")
            
            ' Clear old properties
            ClearBodyProperties(userProps)
            
            ' Save count
            SetOrAddProperty(userProps, PROP_PREFIX & "Count", bodies.Count.ToString())
            
            ' Save each body's data
            For i As Integer = 0 To bodies.Count - 1
                Dim bi As BodyInfo = bodies(i)
                Dim prefix As String = PROP_PREFIX & i.ToString() & "_"
                
                SetOrAddProperty(userProps, prefix & "Name", bi.Name)
                SetOrAddProperty(userProps, prefix & "Sig", bi.Signature)
                SetOrAddProperty(userProps, prefix & "TAxis", bi.ThicknessVector)
                SetOrAddProperty(userProps, prefix & "WAxis", bi.WidthVector)
                SetOrAddProperty(userProps, prefix & "LAxis", bi.LengthVector)
                SetOrAddProperty(userProps, prefix & "SM", If(bi.ConvertToSheetMetal, "1", "0"))
                SetOrAddProperty(userProps, prefix & "Mat", bi.MaterialName)
                SetOrAddProperty(userProps, prefix & "Part", bi.CreatedPartPath)
            Next
            
            logs.Add("MakeComponentsLib: Saved data for " & bodies.Count & " bodies to master")
        Catch ex As Exception
            logs.Add("MakeComponentsLib: Failed to save body data: " & ex.Message)
        End Try
    End Sub
    
    ' Load body data from master document properties
    Public Function LoadBodyDataFromMaster(masterDoc As PartDocument, _
                                           logs As System.Collections.Generic.List(Of String)) As System.Collections.Generic.List(Of StoredBodyData)
        Dim result As New System.Collections.Generic.List(Of StoredBodyData)
        
        Try
            Dim userProps As PropertySet = masterDoc.PropertySets.Item("Inventor User Defined Properties")
            
            Dim countStr As String = GetPropertyValue(userProps, PROP_PREFIX & "Count", "0")
            Dim count As Integer = 0
            Integer.TryParse(countStr, count)
            
            For i As Integer = 0 To count - 1
                Dim prefix As String = PROP_PREFIX & i.ToString() & "_"
                
                Dim data As New StoredBodyData()
                data.Name = GetPropertyValue(userProps, prefix & "Name", "")
                data.Signature = GetPropertyValue(userProps, prefix & "Sig", "")
                data.ThicknessVector = GetPropertyValue(userProps, prefix & "TAxis", "")
                data.WidthVector = GetPropertyValue(userProps, prefix & "WAxis", "")
                data.LengthVector = GetPropertyValue(userProps, prefix & "LAxis", "")
                data.ConvertToSheetMetal = GetPropertyValue(userProps, prefix & "SM", "0") = "1"
                data.MaterialName = GetPropertyValue(userProps, prefix & "Mat", "")
                data.CreatedPartPath = GetPropertyValue(userProps, prefix & "Part", "")
                
                If Not String.IsNullOrEmpty(data.Name) Then
                    result.Add(data)
                End If
            Next
            
            logs.Add("MakeComponentsLib: Loaded data for " & result.Count & " bodies from master")
        Catch ex As Exception
            logs.Add("MakeComponentsLib: No stored body data found or error: " & ex.Message)
        End Try
        
        Return result
    End Function
    
    ' Match current bodies with stored data (by name first, then by signature)
    Public Sub ApplyStoredDataToBodies(bodies As System.Collections.Generic.List(Of BodyInfo), _
                                       storedData As System.Collections.Generic.List(Of StoredBodyData), _
                                       logs As System.Collections.Generic.List(Of String))
        Dim matchedIndices As New System.Collections.Generic.HashSet(Of Integer)
        
        ' First pass: match by name
        For Each bi As BodyInfo In bodies
            For j As Integer = 0 To storedData.Count - 1
                If matchedIndices.Contains(j) Then Continue For
                
                Dim sd As StoredBodyData = storedData(j)
                If bi.Name.Equals(sd.Name, StringComparison.OrdinalIgnoreCase) Then
                    ApplyStoredData(bi, sd, logs)
                    matchedIndices.Add(j)
                    Exit For
                End If
            Next
        Next
        
        ' Second pass: match unmatched bodies by signature
        For Each bi As BodyInfo In bodies
            If bi.PartExists Then Continue For ' Already matched
            
            For j As Integer = 0 To storedData.Count - 1
                If matchedIndices.Contains(j) Then Continue For
                
                Dim sd As StoredBodyData = storedData(j)
                If Not String.IsNullOrEmpty(bi.Signature) AndAlso _
                   bi.Signature.Equals(sd.Signature, StringComparison.OrdinalIgnoreCase) Then
                    logs.Add("MakeComponentsLib: Matched '" & bi.Name & "' to stored '" & sd.Name & "' by signature")
                    ApplyStoredData(bi, sd, logs)
                    matchedIndices.Add(j)
                    Exit For
                End If
            Next
        Next
    End Sub
    
    Private Sub ApplyStoredData(bi As BodyInfo, sd As StoredBodyData, _
                                logs As System.Collections.Generic.List(Of String))
        ' Apply stored axis settings if available
        If Not String.IsNullOrEmpty(sd.ThicknessVector) Then
            bi.ThicknessVector = sd.ThicknessVector
            bi.WidthVector = sd.WidthVector
            bi.LengthVector = sd.LengthVector
        End If
        
        bi.ConvertToSheetMetal = sd.ConvertToSheetMetal
        bi.MaterialName = sd.MaterialName
        bi.CreatedPartPath = sd.CreatedPartPath
        
        ' Check if part exists on disk
        If Not String.IsNullOrEmpty(sd.CreatedPartPath) Then
            bi.PartExists = System.IO.File.Exists(sd.CreatedPartPath)
            If bi.PartExists Then
                ' Default to NOT selected for existing parts (user must opt-in to recreate)
                bi.Selected = False
                logs.Add("MakeComponentsLib: Body '" & bi.Name & "' has existing part: " & _
                         System.IO.Path.GetFileName(sd.CreatedPartPath))
            End If
        End If
    End Sub
    
    Private Sub ClearBodyProperties(userProps As PropertySet)
        Dim toRemove As New System.Collections.Generic.List(Of String)
        
        For Each prop As Inventor.Property In userProps
            If prop.Name.StartsWith(PROP_PREFIX) Then
                toRemove.Add(prop.Name)
            End If
        Next
        
        For Each propName As String In toRemove
            Try
                userProps.Item(propName).Delete()
            Catch
            End Try
        Next
    End Sub
    
    Private Sub SetOrAddProperty(userProps As PropertySet, name As String, value As String)
        Try
            userProps.Item(name).Value = value
        Catch
            Try
                userProps.Add(value, name)
            Catch
            End Try
        End Try
    End Sub
    
    Private Function GetPropertyValue(userProps As PropertySet, name As String, defaultValue As String) As String
        Try
            Return CStr(userProps.Item(name).Value)
        Catch
            Return defaultValue
        End Try
    End Function
    
    ' ============================================================================
    ' Axis Detection (adapted from BoundingBoxStockLib)
    ' ============================================================================
    
    ' Detect axes for a single solid body
    ' Uses face normals to find thickness (smallest extent), then computes perpendicular axes
    Public Sub DetectAxesForBody(body As SurfaceBody, _
                                 ByRef thicknessVec As String, ByRef thicknessVal As Double, _
                                 ByRef widthVec As String, ByRef widthVal As Double, _
                                 ByRef lengthVec As String, ByRef lengthVal As Double, _
                                 logs As System.Collections.Generic.List(Of String))
        Dim checkedNormals As New System.Collections.Generic.List(Of String)
        Dim bestNormalX As Double = 0, bestNormalY As Double = 0, bestNormalZ As Double = 0
        Dim minExtent As Double = Double.MaxValue
        Dim foundNormal As Boolean = False
        
        ' Find the face normal direction with the smallest extent (thickness)
        For Each face As Face In body.Faces
            Dim nx As Double = 0, ny As Double = 0, nz As Double = 0
            If GetFaceNormal(face, nx, ny, nz) Then
                Dim len As Double = Math.Sqrt(nx * nx + ny * ny + nz * nz)
                If len > 0.0001 Then
                    nx /= len : ny /= len : nz /= len
                End If
                
                ' Make normal canonical (always point in "positive" direction)
                If nx < -0.0001 OrElse (Math.Abs(nx) < 0.0001 AndAlso ny < -0.0001) OrElse _
                   (Math.Abs(nx) < 0.0001 AndAlso Math.Abs(ny) < 0.0001 AndAlso nz < -0.0001) Then
                    nx = -nx : ny = -ny : nz = -nz
                End If
                
                Dim normalKey As String = Math.Round(nx, 3).ToString() & "," & _
                                          Math.Round(ny, 3).ToString() & "," & _
                                          Math.Round(nz, 3).ToString()
                
                If checkedNormals.Contains(normalKey) Then Continue For
                checkedNormals.Add(normalKey)
                
                Dim extent As Double = GetOrientedExtentForBody(body, nx, ny, nz)
                If extent > 0 AndAlso extent < minExtent Then
                    minExtent = extent
                    bestNormalX = nx
                    bestNormalY = ny
                    bestNormalZ = nz
                    foundNormal = True
                End If
            End If
        Next
        
        If Not foundNormal Then
            logs.Add("MakeComponentsLib: Could not detect axes for '" & body.Name & "'")
            Exit Sub
        End If
        
        ' Set thickness
        thicknessVal = minExtent
        thicknessVec = SimplifyAxisVector(bestNormalX, bestNormalY, bestNormalZ)
        
        ' Compute perpendicular vectors for width and length
        Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
        Dim lx As Double = 0, ly As Double = 0, lz As Double = 0
        ComputePerpendicularVectors(bestNormalX, bestNormalY, bestNormalZ, wx, wy, wz, lx, ly, lz)
        
        ' Measure extents along perpendicular axes
        Dim widthExtent As Double = GetOrientedExtentForBody(body, wx, wy, wz)
        Dim lengthExtent As Double = GetOrientedExtentForBody(body, lx, ly, lz)
        
        ' Assign width (smaller) and length (larger)
        If lengthExtent >= widthExtent Then
            widthVal = widthExtent
            lengthVal = lengthExtent
            widthVec = SimplifyAxisVector(wx, wy, wz)
            lengthVec = SimplifyAxisVector(lx, ly, lz)
        Else
            widthVal = lengthExtent
            lengthVal = widthExtent
            widthVec = SimplifyAxisVector(lx, ly, lz)
            lengthVec = SimplifyAxisVector(wx, wy, wz)
        End If
        
        logs.Add("MakeComponentsLib: Detected axes for '" & body.Name & "' - T:" & _
                 FormatNumber(thicknessVal * 10, 2) & " W:" & FormatNumber(widthVal * 10, 2) & _
                 " L:" & FormatNumber(lengthVal * 10, 2))
    End Sub
    
    ' Compute two perpendicular vectors to a given normal
    Private Sub ComputePerpendicularVectors(nx As Double, ny As Double, nz As Double, _
                                            ByRef wx As Double, ByRef wy As Double, ByRef wz As Double, _
                                            ByRef lx As Double, ByRef ly As Double, ByRef lz As Double)
        ' Find a reference vector not parallel to normal
        Dim refX As Double = 1, refY As Double = 0, refZ As Double = 0
        Dim dot As Double = nx * refX + ny * refY + nz * refZ
        If Math.Abs(dot) > 0.9 Then
            refX = 0 : refY = 1 : refZ = 0
        End If
        
        ' Cross product: w = n x ref
        wx = ny * refZ - nz * refY
        wy = nz * refX - nx * refZ
        wz = nx * refY - ny * refX
        
        ' Normalize w
        Dim wLen As Double = Math.Sqrt(wx * wx + wy * wy + wz * wz)
        If wLen > 0.0001 Then
            wx /= wLen : wy /= wLen : wz /= wLen
        End If
        
        ' Cross product: l = n x w
        lx = ny * wz - nz * wy
        ly = nz * wx - nx * wz
        lz = nx * wy - ny * wx
        
        ' Normalize l
        Dim lLen As Double = Math.Sqrt(lx * lx + ly * ly + lz * lz)
        If lLen > 0.0001 Then
            lx /= lLen : ly /= lLen : lz /= lLen
        End If
    End Sub
    
    ' Convert vector to string, simplifying to X/Y/Z if aligned with principal axis
    Private Function SimplifyAxisVector(vx As Double, vy As Double, vz As Double) As String
        If Math.Abs(vx) > 0.9998 Then Return "X"
        If Math.Abs(vy) > 0.9998 Then Return "Y"
        If Math.Abs(vz) > 0.9998 Then Return "Z"
        Return VectorToString(vx, vy, vz)
    End Function
    
    ' Get all bodies with detected axes
    Public Function GetBodiesWithAxes(partDoc As PartDocument, _
                                      logs As System.Collections.Generic.List(Of String)) As System.Collections.Generic.List(Of BodyInfo)
        Dim result As New System.Collections.Generic.List(Of BodyInfo)
        
        For Each body As SurfaceBody In partDoc.ComponentDefinition.SurfaceBodies
            Dim info As New BodyInfo(body)
            DetectAxesForBody(body, _
                              info.ThicknessVector, info.ThicknessValue, _
                              info.WidthVector, info.WidthValue, _
                              info.LengthVector, info.LengthValue, _
                              logs)
            result.Add(info)
        Next
        
        Return result
    End Function
    
    ' ============================================================================
    ' Part Derivation
    ' ============================================================================
    
    ' Create a new part document from template
    Public Function CreatePartFromTemplate(app As Inventor.Application, templatePath As String, _
                                           logs As System.Collections.Generic.List(Of String)) As PartDocument
        Try
            Dim newDoc As Document = app.Documents.Add(DocumentTypeEnum.kPartDocumentObject, templatePath, True)
            logs.Add("MakeComponentsLib: Created part from template: " & templatePath)
            Return CType(newDoc, PartDocument)
        Catch ex As Exception
            logs.Add("MakeComponentsLib: Failed to create part: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ' Find template in templates folder
    Public Function FindTemplate(app As Inventor.Application, templateName As String, _
                                 logs As System.Collections.Generic.List(Of String)) As String
        Try
            Dim templatesPath As String = app.DesignProjectManager.ActiveDesignProject.TemplatesPath
            logs.Add("MakeComponentsLib: Templates folder: " & templatesPath)
            
            Dim candidates() As String = { _
                templateName, _
                templateName & ".ipt", _
                "Part.ipt", _
                "Sheet Metal.ipt", _
                "SheetMetal Part.ipt" _
            }
            
            For Each candidate As String In candidates
                Dim fullPath As String = System.IO.Path.Combine(templatesPath, candidate)
                If System.IO.File.Exists(fullPath) Then
                    logs.Add("MakeComponentsLib: Found template: " & fullPath)
                    Return fullPath
                End If
            Next
            
            logs.Add("MakeComponentsLib: No template found matching '" & templateName & "'")
            Return ""
        Catch ex As Exception
            logs.Add("MakeComponentsLib: Error finding template: " & ex.Message)
            Return ""
        End Try
    End Function
    
    ' Helper to exclude all entities from a derived part entity collection
    ' Wrapped in Try/Catch since some collections may not exist in iLogic context
    Private Sub ExcludeAllDerivedEntities(entities As Object, _
                                          logs As System.Collections.Generic.List(Of String), _
                                          entityType As String)
        Try
            Dim count As Integer = 0
            For Each dpe As DerivedPartEntity In entities
                dpe.IncludeEntity = False
                count += 1
            Next
            If count > 0 Then
                logs.Add("MakeComponentsLib: Excluded " & count & " " & entityType)
            End If
        Catch
            ' Collection may not exist or not be accessible in iLogic context
        End Try
    End Sub
    
    ' Derive a single body from a multi-body part
    ' Note: DerivedPartUniformScaleDef does NOT have IncludeAllWorkSurfaces, IncludeAllParameters etc.
    ' Only DeriveStyle and individual entity IncludeEntity work in iLogic
    Public Function DeriveBodyAsNewPart(masterDoc As PartDocument, _
                                        targetBodyName As String, _
                                        newPartDoc As PartDocument, _
                                        logs As System.Collections.Generic.List(Of String)) As Boolean
        Try
            Dim dpcs As DerivedPartComponents = newPartDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents
            Dim dpDef As DerivedPartUniformScaleDef = dpcs.CreateUniformScaleDef(masterDoc.FullDocumentName)
            
            logs.Add("MakeComponentsLib: Created DerivedPartUniformScaleDef, solids count: " & dpDef.Solids.Count)
            
            Dim included As Integer = 0
            Dim excluded As Integer = 0
            
            For Each dpe As DerivedPartEntity In dpDef.Solids
                Dim bodyName As String = ""
                Try
                    Dim refEntity As Object = dpe.ReferencedEntity
                    If TypeOf refEntity Is SurfaceBody Then
                        bodyName = CType(refEntity, SurfaceBody).Name
                    End If
                Catch
                End Try
                
                If bodyName = targetBodyName Then
                    dpe.IncludeEntity = True
                    included += 1
                    logs.Add("MakeComponentsLib: Including body: '" & bodyName & "'")
                Else
                    dpe.IncludeEntity = False
                    excluded += 1
                End If
            Next
            
            logs.Add("MakeComponentsLib: Included: " & included & ", Excluded: " & excluded)
            
            If included = 0 Then
                logs.Add("MakeComponentsLib: No bodies matched target '" & targetBodyName & "'")
                Return False
            End If
            
            ' Exclude sketches, work features, surfaces, and parameters to derive only the solid body
            ' Each call wrapped in Try/Catch because property access itself may fail if property doesn't exist
            Try : ExcludeAllDerivedEntities(dpDef.Sketches3D, logs, "3D sketches") : Catch : End Try
            Try : ExcludeAllDerivedEntities(dpDef.Sketches, logs, "sketches") : Catch : End Try
            Try : ExcludeAllDerivedEntities(dpDef.WorkFeatures, logs, "work features") : Catch : End Try
            Try : ExcludeAllDerivedEntities(dpDef.Surfaces, logs, "surfaces") : Catch : End Try
            Try : ExcludeAllDerivedEntities(dpDef.Parameters, logs, "parameters") : Catch : End Try
            
            dpDef.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyWithSeams
            
            dpcs.Add(dpDef)
            newPartDoc.Update()
            logs.Add("MakeComponentsLib: Derivation complete for '" & targetBodyName & "'")
            Return True
        Catch ex As Exception
            logs.Add("MakeComponentsLib: Derivation failed: " & ex.Message)
            Return False
        End Try
    End Function
    
    ' Update derivation in an existing part (delete old derivation and recreate)
    Public Sub UpdateDerivedPart(masterDoc As PartDocument, _
                                 targetBodyName As String, _
                                 existingPartDoc As PartDocument, _
                                 logs As System.Collections.Generic.List(Of String))
        Try
            Dim compDef As PartComponentDefinition = existingPartDoc.ComponentDefinition
            Dim dpcs As DerivedPartComponents = compDef.ReferenceComponents.DerivedPartComponents
            
            ' Delete existing derived components
            Dim toDelete As New System.Collections.Generic.List(Of DerivedPartComponent)
            For Each dpc As DerivedPartComponent In dpcs
                toDelete.Add(dpc)
            Next
            
            For Each dpc As DerivedPartComponent In toDelete
                Try
                    dpc.Delete()
                Catch
                End Try
            Next
            
            If toDelete.Count > 0 Then
                logs.Add("MakeComponentsLib: Deleted " & toDelete.Count & " existing derivation(s)")
            End If
            
            ' Also delete existing solid bodies (they came from derivation)
            Dim bodiesToDelete As New System.Collections.Generic.List(Of SurfaceBody)
            For Each body As SurfaceBody In compDef.SurfaceBodies
                bodiesToDelete.Add(body)
            Next
            
            ' Delete via features if possible
            For Each feature As Object In compDef.Features
                Try
                    If TypeOf feature Is PartFeature Then
                        Dim pf As PartFeature = CType(feature, PartFeature)
                        ' Skip work features and sketches
                        If TypeOf feature Is DerivedPartComponent Then
                            pf.Delete()
                        End If
                    End If
                Catch
                End Try
            Next
            
            ' Recreate derivation
            Dim dpDef As DerivedPartUniformScaleDef = dpcs.CreateUniformScaleDef(masterDoc.FullDocumentName)
            
            Dim included As Integer = 0
            For Each dpe As DerivedPartEntity In dpDef.Solids
                Dim bodyName As String = ""
                Try
                    Dim refEntity As Object = dpe.ReferencedEntity
                    If TypeOf refEntity Is SurfaceBody Then
                        bodyName = CType(refEntity, SurfaceBody).Name
                    End If
                Catch
                End Try
                
                If bodyName = targetBodyName Then
                    dpe.IncludeEntity = True
                    included += 1
                Else
                    dpe.IncludeEntity = False
                End If
            Next
            
            If included = 0 Then
                logs.Add("MakeComponentsLib: Warning - no bodies matched '" & targetBodyName & "' during update")
            End If
            
            ' Exclude sketches, work features, surfaces, and parameters to derive only the solid body
            ' Each call wrapped in Try/Catch because property access itself may fail if property doesn't exist
            Try : ExcludeAllDerivedEntities(dpDef.Sketches3D, logs, "3D sketches") : Catch : End Try
            Try : ExcludeAllDerivedEntities(dpDef.Sketches, logs, "sketches") : Catch : End Try
            Try : ExcludeAllDerivedEntities(dpDef.WorkFeatures, logs, "work features") : Catch : End Try
            Try : ExcludeAllDerivedEntities(dpDef.Surfaces, logs, "surfaces") : Catch : End Try
            Try : ExcludeAllDerivedEntities(dpDef.Parameters, logs, "parameters") : Catch : End Try
            
            dpDef.DeriveStyle = DerivedComponentStyleEnum.kDeriveAsSingleBodyWithSeams
            dpcs.Add(dpDef)
            existingPartDoc.Update()
            
            logs.Add("MakeComponentsLib: Updated derivation for '" & targetBodyName & "'")
        Catch ex As Exception
            logs.Add("MakeComponentsLib: Failed to update derivation: " & ex.Message)
        End Try
    End Sub
    
    ' ============================================================================
    ' iProperties
    ' ============================================================================
    
    ' Set iProperties on a part document
    ' If partNumber is empty, Part Number property is left unchanged (Vault will assign)
    Public Sub SetPartProperties(partDoc As PartDocument, _
                                 projectName As String, _
                                 description As String, _
                                 partNumber As String, _
                                 logs As System.Collections.Generic.List(Of String))
        Try
            Dim designProps As PropertySet = partDoc.PropertySets.Item("Design Tracking Properties")
            
            SetPropertyValue(designProps, "Project", projectName)
            SetPropertyValue(designProps, "Description", description)
            
            ' Only set Part Number if provided - let Vault assign if empty
            If Not String.IsNullOrEmpty(partNumber) Then
                SetPropertyValue(designProps, "Part Number", partNumber)
            End If
            
            logs.Add("MakeComponentsLib: Set properties - Project: " & projectName & ", Desc: " & description)
        Catch ex As Exception
            logs.Add("MakeComponentsLib: Failed to set properties: " & ex.Message)
        End Try
    End Sub
    
    ' Set custom dimension properties (Thickness, Width, Length)
    Public Sub SetDimensionProperties(partDoc As PartDocument, _
                                      thickness As Double, _
                                      width As Double, _
                                      length As Double, _
                                      logs As System.Collections.Generic.List(Of String))
        Try
            Dim userProps As PropertySet = partDoc.PropertySets.Item("Inventor User Defined Properties")
            
            Dim uom As UnitsOfMeasure = partDoc.UnitsOfMeasure
            Dim thicknessMm As String = uom.GetStringFromValue(thickness, "mm")
            Dim widthMm As String = uom.GetStringFromValue(width, "mm")
            Dim lengthMm As String = uom.GetStringFromValue(length, "mm")
            
            SetOrAddCustomProperty(userProps, "Thickness", thicknessMm)
            SetOrAddCustomProperty(userProps, "Width", widthMm)
            SetOrAddCustomProperty(userProps, "Length", lengthMm)
            
            logs.Add("MakeComponentsLib: Set dimensions - T:" & thicknessMm & " W:" & widthMm & " L:" & lengthMm)
        Catch ex As Exception
            logs.Add("MakeComponentsLib: Failed to set dimensions: " & ex.Message)
        End Try
    End Sub
    
    Private Sub SetPropertyValue(propSet As PropertySet, propName As String, value As String)
        Try
            propSet.Item(propName).Value = value
        Catch
        End Try
    End Sub
    
    Private Sub SetOrAddCustomProperty(propSet As PropertySet, propName As String, value As String)
        Try
            propSet.Item(propName).Value = value
        Catch
            Try
                propSet.Add(value, propName)
            Catch
            End Try
        End Try
    End Sub
    
    ' ============================================================================
    ' Material Assignment
    ' ============================================================================
    
    ' Get available materials as a list from the document's Materials collection
    ' Note: app.Assets does not work in iLogic, use partDoc.Materials instead
    Public Function GetAvailableMaterials(partDoc As PartDocument, _
                                          logs As System.Collections.Generic.List(Of String)) As System.Collections.Generic.List(Of String)
        Dim materials As New System.Collections.Generic.List(Of String)
        
        logs.Add("MakeComponentsLib: Enumerating materials from document...")
        
        Try
            For Each mat As Material In partDoc.Materials
                If Not materials.Contains(mat.Name) Then
                    materials.Add(mat.Name)
                End If
            Next
            
            logs.Add("MakeComponentsLib: Found " & materials.Count & " materials in document")
            
            ' Sort materials alphabetically
            materials.Sort()
            
        Catch ex As Exception
            logs.Add("MakeComponentsLib: Error enumerating materials: " & ex.Message)
        End Try
        
        Return materials
    End Function
    
    ' Assign material to a part
    Public Sub AssignMaterial(partDoc As PartDocument, materialName As String, _
                              logs As System.Collections.Generic.List(Of String))
        If String.IsNullOrEmpty(materialName) Then Return
        
        Try
            Dim material As Material = partDoc.Materials.Item(materialName)
            partDoc.ComponentDefinition.Material = material
            logs.Add("MakeComponentsLib: Assigned material '" & materialName & "'")
        Catch
            Try
                For Each mat As Material In partDoc.Materials
                    If mat.Name.IndexOf(materialName, StringComparison.OrdinalIgnoreCase) >= 0 Then
                        partDoc.ComponentDefinition.Material = mat
                        logs.Add("MakeComponentsLib: Assigned material '" & mat.Name & "'")
                        Return
                    End If
                Next
                logs.Add("MakeComponentsLib: Material '" & materialName & "' not found")
            Catch ex As Exception
                logs.Add("MakeComponentsLib: Error assigning material: " & ex.Message)
            End Try
        End Try
    End Sub
    
    ' ============================================================================
    ' Assembly Operations
    ' ============================================================================
    
    ' Create a new assembly document
    Public Function CreateAssembly(app As Inventor.Application, templatePath As String, _
                                   logs As System.Collections.Generic.List(Of String)) As AssemblyDocument
        Try
            Dim newDoc As Document
            If Not String.IsNullOrEmpty(templatePath) AndAlso System.IO.File.Exists(templatePath) Then
                newDoc = app.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, templatePath, True)
            Else
                newDoc = app.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, , True)
            End If
            logs.Add("MakeComponentsLib: Created new assembly")
            Return CType(newDoc, AssemblyDocument)
        Catch ex As Exception
            logs.Add("MakeComponentsLib: Failed to create assembly: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ' Place a component in assembly (grounded at origin)
    Public Function PlaceComponentGrounded(asmDoc As AssemblyDocument, partPath As String, _
                                           logs As System.Collections.Generic.List(Of String)) As ComponentOccurrence
        Try
            Dim tg As TransientGeometry = asmDoc.ComponentDefinition.Application.TransientGeometry
            Dim origin As Matrix = tg.CreateMatrix()
            Dim occ As ComponentOccurrence = asmDoc.ComponentDefinition.Occurrences.Add(partPath, origin)
            occ.Grounded = True
            logs.Add("MakeComponentsLib: Placed component: " & System.IO.Path.GetFileName(partPath))
            Return occ
        Catch ex As Exception
            logs.Add("MakeComponentsLib: Failed to place component: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ' Set iProperties on an assembly document
    Public Sub SetAssemblyProperties(asmDoc As AssemblyDocument, _
                                     projectName As String, _
                                     logs As System.Collections.Generic.List(Of String))
        Try
            Dim designProps As PropertySet = asmDoc.PropertySets.Item("Design Tracking Properties")
            
            ' Set Project property
            Try
                designProps.Item("Project").Value = projectName
            Catch
                Try
                    designProps.Add(projectName, "Project")
                Catch
                End Try
            End Try
            
            logs.Add("MakeComponentsLib: Set assembly project to '" & projectName & "'")
        Catch ex As Exception
            logs.Add("MakeComponentsLib: Failed to set assembly properties: " & ex.Message)
        End Try
    End Sub
    
    ' ============================================================================
    ' Path/Project Utilities
    ' ============================================================================
    
    ' Extract project name from path */Tooted/[ProjectName]/...
    Public Function ExtractProjectName(filePath As String) As String
        Try
            Dim parts() As String = filePath.Split(System.IO.Path.DirectorySeparatorChar)
            For i As Integer = 0 To parts.Length - 2
                If parts(i).Equals("Tooted", StringComparison.OrdinalIgnoreCase) Then
                    Return parts(i + 1)
                End If
            Next
        Catch
        End Try
        Return ""
    End Function
    
    ' Create subfolder if it doesn't exist (local file system only)
    Public Function EnsureSubfolder(basePath As String, subfolderName As String) As String
        Try
            Dim subPath As String = System.IO.Path.Combine(basePath, subfolderName)
            If Not System.IO.Directory.Exists(subPath) Then
                System.IO.Directory.CreateDirectory(subPath)
            End If
            Return subPath
        Catch
            Return basePath
        End Try
    End Function
    
    ' Create subfolder in both local file system AND Vault
    ' Parameters:
    '   basePath - The parent folder path (local file system)
    '   subfolderName - Name of the subfolder to create
    '   vaultConn - Vault connection object (from VaultNumberingLib.GetVaultConnection)
    '   workspaceRoot - Local workspace root path (maps to $/ in Vault)
    '   logs - List to collect log messages
    ' Returns: The full path of the created subfolder
    Public Function EnsureSubfolderWithVault(basePath As String, _
                                             subfolderName As String, _
                                             vaultConn As Object, _
                                             workspaceRoot As String, _
                                             logs As System.Collections.Generic.List(Of String)) As String
        ' Create local folder first
        Dim localPath As String = EnsureSubfolder(basePath, subfolderName)
        
        ' If Vault is connected and workspace is known, also create in Vault
        If vaultConn IsNot Nothing AndAlso Not String.IsNullOrEmpty(workspaceRoot) Then
            Dim vaultPath As String = VaultNumberingLib.ConvertLocalPathToVaultPath(localPath, workspaceRoot)
            
            If Not String.IsNullOrEmpty(vaultPath) Then
                Dim vaultFolder As Object = VaultNumberingLib.EnsureVaultFolder(vaultConn, vaultPath, logs)
                If vaultFolder IsNot Nothing Then
                    logs.Add("MakeComponentsLib: Vault folder ready: " & vaultPath)
                Else
                    logs.Add("MakeComponentsLib: Could not create Vault folder (local only): " & vaultPath)
                End If
            Else
                logs.Add("MakeComponentsLib: Path not in workspace, skipping Vault folder creation")
            End If
        Else
            logs.Add("MakeComponentsLib: No Vault connection or workspace, local folder only")
        End If
        
        Return localPath
    End Function
    
    ' ============================================================================
    ' Helper Functions
    ' ============================================================================
    
    Public Function GetFaceNormal(face As Face, ByRef nx As Double, ByRef ny As Double, ByRef nz As Double) As Boolean
        Try
            Dim geom As Object = face.Geometry
            If TypeOf geom Is Plane Then
                Dim plane As Plane = CType(geom, Plane)
                Dim normal As UnitVector = plane.Normal
                nx = normal.X
                ny = normal.Y
                nz = normal.Z
                Return True
            End If
        Catch
        End Try
        Return False
    End Function
    
    Public Function GetOrientedExtentForBody(body As SurfaceBody, dirX As Double, dirY As Double, dirZ As Double) As Double
        Dim minProj As Double = Double.MaxValue
        Dim maxProj As Double = Double.MinValue
        
        Try
            For Each vertex As Vertex In body.Vertices
                Dim pt As Point = vertex.Point
                Dim proj As Double = pt.X * dirX + pt.Y * dirY + pt.Z * dirZ
                If proj < minProj Then minProj = proj
                If proj > maxProj Then maxProj = proj
            Next
        Catch
        End Try
        
        If minProj = Double.MaxValue Then Return 0
        Return maxProj - minProj
    End Function
    
    Public Function VectorToString(vx As Double, vy As Double, vz As Double) As String
        Return "V:" & vx.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture) & "," & _
                      vy.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture) & "," & _
                      vz.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture)
    End Function

End Module
