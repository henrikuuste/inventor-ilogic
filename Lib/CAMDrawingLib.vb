' ============================================================================
' CAMDrawingLib - Library for CAM drawing automation
' 
' Provides functions to:
' - Create drawings from template with 1:1 scale views
' - Place all 6 orthographic views in T-layout
' - Resize sheets to fit part extents
' - Add extent dimensions
' - Export to DWG/DXF (2010 format)
'
' Usage: 
'   In calling script:
'     AddVbFile "Lib/CAMDrawingLib.vb"
'
' Note: Logger is not available in library modules.
'       Pass a List(Of String) to collect log messages.
' ============================================================================

Imports Inventor
Imports System.Collections.Generic

Public Module CAMDrawingLib

    ' DWG Translator AddIn GUID
    Private Const DWG_ADDIN_GUID As String = "{C24E3AC2-122E-11D5-8E91-0010B541CD80}"
    
    ' Default template name
    Private Const DEFAULT_TEMPLATE As String = "Drawing.1.1.idw"
    
    ' Sheet Metal GUID
    Private Const SHEET_METAL_GUID As String = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"

    ' ============================================================================
    ' Base View Orientation Detection
    ' ============================================================================
    
    ' Determine the appropriate base view orientation for a part
    ' Priority: 1) Sheet metal flat pattern, 2) BB_ThicknessAxis property, 3) Default front
    Public Function DetermineBaseViewOrientation(partDoc As PartDocument, _
                                                  logs As List(Of String)) As ViewOrientationTypeEnum
        ' Check if sheet metal with flat pattern
        If IsSheetMetal(partDoc) AndAlso HasFlatPattern(partDoc) Then
            logs.Add("CAMDrawingLib: Part is sheet metal with flat pattern - using Top view")
            Return ViewOrientationTypeEnum.kTopViewOrientation  ' Flat pattern uses top view + SetFlatPatternView
        End If
        
        ' Check for BB_ThicknessAxis custom property
        Dim thicknessAxis As String = GetCustomPropertyValue(partDoc, "BB_ThicknessAxis", "")
        
        If Not String.IsNullOrEmpty(thicknessAxis) Then
            logs.Add("CAMDrawingLib: Found BB_ThicknessAxis = " & thicknessAxis)
            Return GetViewOrientationFromThicknessAxis(thicknessAxis, logs)
        End If
        
        ' Default to front view
        logs.Add("CAMDrawingLib: Using default front view orientation")
        Return ViewOrientationTypeEnum.kFrontViewOrientation
    End Function
    
    ' Check if part is sheet metal
    Public Function IsSheetMetal(partDoc As PartDocument) As Boolean
        Return partDoc.SubType = SHEET_METAL_GUID
    End Function
    
    ' Check if sheet metal part has a flat pattern
    Public Function HasFlatPattern(partDoc As PartDocument) As Boolean
        If Not IsSheetMetal(partDoc) Then Return False
        Try
            Dim smCompDef As SheetMetalComponentDefinition = CType(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
            Return smCompDef.HasFlatPattern
        Catch
            Return False
        End Try
    End Function
    
    ' Create a flat pattern view for sheet metal
    ' Uses AddBaseView with SheetMetalFoldedModel = False to create flat pattern view
    ' See: https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=DrawingViews_AddBaseView
    Public Function CreateFlatPatternView(sheet As Sheet, _
                                          partDoc As PartDocument, _
                                          app As Inventor.Application, _
                                          position As Point2d, _
                                          logs As List(Of String)) As DrawingView
        If Not IsSheetMetal(partDoc) OrElse Not HasFlatPattern(partDoc) Then
            logs.Add("CAMDrawingLib: Part is not sheet metal with flat pattern")
            Return Nothing
        End If
        
        Dim view As DrawingView = Nothing
        Dim flatPatternCreated As Boolean = False
        
        ' Log info about the part for debugging
        logs.Add("CAMDrawingLib: Part full path: " & partDoc.FullDocumentName)
        logs.Add("CAMDrawingLib: Part SubType: " & partDoc.SubType)
        
        
        ' SOLUTION: Use NameValueMap with named parameter AdditionalOptions
        ' See: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-to-create-drawing-with-flat-pattern-view/td-p/13367792
        ' Use kDefaultViewOrientation - this uses the flat pattern's natural orientation (looking down at the flat sheet)
        Try
            Dim viewOptions As NameValueMap = app.TransientObjects.CreateNameValueMap()
            viewOptions.Add("SheetMetalFoldedModel", False)
            
            view = sheet.DrawingViews.AddBaseView( _
                partDoc, _
                position, _
                1.0, _
                ViewOrientationTypeEnum.kDefaultViewOrientation, _
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                AdditionalOptions := viewOptions)
            
            logs.Add("CAMDrawingLib: Created flat pattern view with kDefaultViewOrientation")
            flatPatternCreated = CheckIfFlatPattern(view, logs)
        Catch ex As Exception
            logs.Add("CAMDrawingLib: kDefaultViewOrientation method failed: " & ex.Message)
        End Try
        
        ' Fallback: Try with kCurrentViewOrientation
        If Not flatPatternCreated Then
            Try
                Dim viewOptions As NameValueMap = app.TransientObjects.CreateNameValueMap()
                viewOptions.Add("SheetMetalFoldedModel", False)
                
                view = sheet.DrawingViews.AddBaseView( _
                    partDoc, _
                    position, _
                    1.0, _
                    ViewOrientationTypeEnum.kCurrentViewOrientation, _
                    DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                    AdditionalOptions := viewOptions)
                
                logs.Add("CAMDrawingLib: Created flat pattern view with kCurrentViewOrientation")
                flatPatternCreated = CheckIfFlatPattern(view, logs)
            Catch ex As Exception
                logs.Add("CAMDrawingLib: kCurrentViewOrientation method failed: " & ex.Message)
            End Try
        End If
        
        ' Last fallback: Create regular view if flat pattern fails
        If Not flatPatternCreated Then
            Try
                view = sheet.DrawingViews.AddBaseView( _
                    partDoc, _
                    position, _
                    1.0, _
                    ViewOrientationTypeEnum.kTopViewOrientation, _
                    DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
                
                logs.Add("CAMDrawingLib: Created fallback folded view (flat pattern not available)")
            Catch ex As Exception
                logs.Add("CAMDrawingLib: Fallback view creation also failed: " & ex.Message)
            End Try
        End If
        
        ' Fallback: just create a regular view
        If view Is Nothing Then
            Try
                view = sheet.DrawingViews.AddBaseView( _
                    partDoc, _
                    position, _
                    1.0, _
                    ViewOrientationTypeEnum.kTopViewOrientation, _
                    DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                    Nothing, Nothing)
                logs.Add("CAMDrawingLib: Created fallback folded view")
            Catch ex As Exception
                logs.Add("CAMDrawingLib: Fallback also failed: " & ex.Message)
            End Try
        End If
        
        ' Name the view
        If view IsNot Nothing Then
            Try : view.Name = "Flat Pattern" : Catch : End Try
        End If
        
        If flatPatternCreated Then
            logs.Add("CAMDrawingLib: Flat pattern view created successfully")
        Else
            logs.Add("CAMDrawingLib: WARNING - Could not create flat pattern view")
        End If
        
        Return view
    End Function
    
    ' Helper to check if view is flat pattern
    Private Function CheckIfFlatPattern(view As DrawingView, logs As List(Of String)) As Boolean
        If view Is Nothing Then Return False
        Try
            Dim isFP As Boolean = view.IsFlatPatternView
            logs.Add("CAMDrawingLib: IsFlatPatternView = " & isFP.ToString())
            Return isFP
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Could not check IsFlatPatternView: " & ex.Message)
            Return False
        End Try
    End Function
    
    ' Add projected views for flat pattern - 4 edge views (no bottom needed, same as top)
    ' Layout: Front above, Back below, Left on left, Right on right
    ' Named consistently with regular parts workflow
    Private Function AddFlatPatternProjectedViews(sheet As Sheet, _
                                                   baseView As DrawingView, _
                                                   partDoc As PartDocument, _
                                                   app As Inventor.Application, _
                                                   views As List(Of DrawingView), _
                                                   dimSpace As Double, _
                                                   logs As List(Of String)) As List(Of DrawingView)
        ' Get actual drawing view dimensions (in cm, at 1:1 scale)
        ' Use the view's Width and Height properties which reflect the actual on-sheet size
        Dim viewWidth As Double = baseView.Width    ' Horizontal extent on sheet (cm)
        Dim viewHeight As Double = baseView.Height  ' Vertical extent on sheet (cm)
        
        ' Get thickness for edge view height calculation
        Dim smCompDef As SheetMetalComponentDefinition = CType(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
        Dim thickness As Double = smCompDef.Thickness.Value  ' Material thickness in cm
        
        logs.Add("CAMDrawingLib: View dimensions on sheet: " & (viewWidth * 10).ToString("F1") & " x " & (viewHeight * 10).ToString("F1") & " mm, thickness: " & (thickness * 10).ToString("F2") & " mm")
        
        Dim baseX As Double = baseView.Position.X
        Dim baseY As Double = baseView.Position.Y
        
        ' Add Front view (above base - looking at edge)
        Try
            Dim frontY As Double = baseY + viewHeight / 2 + dimSpace + thickness / 2
            Dim frontView As DrawingView = sheet.DrawingViews.AddProjectedView( _
                baseView, _
                app.TransientGeometry.CreatePoint2d(baseX, frontY), _
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                Nothing)
            Try : frontView.Name = "Front" : Catch : End Try
            views.Add(frontView)
            logs.Add("CAMDrawingLib: Front edge view created")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to create Front view: " & ex.Message)
        End Try
        
        ' Add Back view (below base - opposite edge)
        Try
            Dim backY As Double = baseY - viewHeight / 2 - dimSpace - thickness / 2
            Dim backView As DrawingView = sheet.DrawingViews.AddProjectedView( _
                baseView, _
                app.TransientGeometry.CreatePoint2d(baseX, backY), _
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                Nothing)
            Try : backView.Name = "Back" : Catch : End Try
            views.Add(backView)
            logs.Add("CAMDrawingLib: Back edge view created")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to create Back view: " & ex.Message)
        End Try
        
        ' Add Left view (left of base - looking at edge)
        Try
            Dim leftX As Double = baseX - viewWidth / 2 - dimSpace - thickness / 2
            Dim leftView As DrawingView = sheet.DrawingViews.AddProjectedView( _
                baseView, _
                app.TransientGeometry.CreatePoint2d(leftX, baseY), _
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                Nothing)
            Try : leftView.Name = "Left" : Catch : End Try
            views.Add(leftView)
            logs.Add("CAMDrawingLib: Left edge view created")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to create Left view: " & ex.Message)
        End Try
        
        ' Add Right view (right of base - opposite edge)
        Try
            Dim rightX As Double = baseX + viewWidth / 2 + dimSpace + thickness / 2
            Dim rightView As DrawingView = sheet.DrawingViews.AddProjectedView( _
                baseView, _
                app.TransientGeometry.CreatePoint2d(rightX, baseY), _
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                Nothing)
            Try : rightView.Name = "Right" : Catch : End Try
            views.Add(rightView)
            logs.Add("CAMDrawingLib: Right edge view created")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to create Right view: " & ex.Message)
        End Try
        
        ' No Bottom view needed - for flat pattern, bottom is identical to top (Flat Pattern)
        logs.Add("CAMDrawingLib: Flat pattern views complete (5 total: Flat Pattern + 4 edges)")
        
        Return views
    End Function
    
    ' Get the opposite view orientation (for the 6th view in T-layout)
    ' When base is Front, opposite is Back; when base is Right, opposite is Left; etc.
    Public Function GetOppositeViewOrientation(baseOrientation As ViewOrientationTypeEnum) As ViewOrientationTypeEnum
        Select Case baseOrientation
            Case ViewOrientationTypeEnum.kFrontViewOrientation
                Return ViewOrientationTypeEnum.kBackViewOrientation
            Case ViewOrientationTypeEnum.kBackViewOrientation
                Return ViewOrientationTypeEnum.kFrontViewOrientation
            Case ViewOrientationTypeEnum.kRightViewOrientation
                Return ViewOrientationTypeEnum.kLeftViewOrientation
            Case ViewOrientationTypeEnum.kLeftViewOrientation
                Return ViewOrientationTypeEnum.kRightViewOrientation
            Case ViewOrientationTypeEnum.kTopViewOrientation
                Return ViewOrientationTypeEnum.kBottomViewOrientation
            Case ViewOrientationTypeEnum.kBottomViewOrientation
                Return ViewOrientationTypeEnum.kTopViewOrientation
            Case Else
                Return ViewOrientationTypeEnum.kBackViewOrientation
        End Select
    End Function
    
    ' Get human-readable name for view orientation
    Public Function GetViewOrientationName(orientation As ViewOrientationTypeEnum) As String
        Select Case orientation
            Case ViewOrientationTypeEnum.kFrontViewOrientation
                Return "Front"
            Case ViewOrientationTypeEnum.kBackViewOrientation
                Return "Back"
            Case ViewOrientationTypeEnum.kTopViewOrientation
                Return "Top"
            Case ViewOrientationTypeEnum.kBottomViewOrientation
                Return "Bottom"
            Case ViewOrientationTypeEnum.kLeftViewOrientation
                Return "Left"
            Case ViewOrientationTypeEnum.kRightViewOrientation
                Return "Right"
            Case ViewOrientationTypeEnum.kIsoTopRightViewOrientation
                Return "Iso"
            Case Else
                Return orientation.ToString().Replace("kViewOrientation", "").Replace("k", "")
        End Select
    End Function
    
    ' Convert BB_ThicknessAxis to view orientation
    ' The thickness axis is the direction we look ALONG to see the flat face
    ' Returns Nothing for custom vectors (requires arbitrary camera)
    Private Function GetViewOrientationFromThicknessAxis(thicknessAxis As String, _
                                                          logs As List(Of String)) As ViewOrientationTypeEnum
        ' Check for simple axis format (X, Y, Z)
        Select Case thicknessAxis.ToUpper()
            Case "X"
                logs.Add("CAMDrawingLib: Thickness along X - using Right view")
                Return ViewOrientationTypeEnum.kRightViewOrientation
            Case "Y"
                logs.Add("CAMDrawingLib: Thickness along Y - using Front view")
                Return ViewOrientationTypeEnum.kFrontViewOrientation
            Case "Z"
                logs.Add("CAMDrawingLib: Thickness along Z - using Top view")
                Return ViewOrientationTypeEnum.kTopViewOrientation
        End Select
        
        ' Check for vector format "V:x,y,z"
        If thicknessAxis.StartsWith("V:") Then
            Dim vx As Double = 0, vy As Double = 0, vz As Double = 0
            If ParseVectorComponents(thicknessAxis, vx, vy, vz) Then
                ' Check if vector is axis-aligned (within tolerance)
                Dim tolerance As Double = 0.001
                Dim absX As Double = Math.Abs(vx)
                Dim absY As Double = Math.Abs(vy)
                Dim absZ As Double = Math.Abs(vz)
                
                ' Check for axis-aligned vectors
                If absY < tolerance AndAlso absZ < tolerance Then
                    logs.Add("CAMDrawingLib: Thickness vector along X - using Right view")
                    Return ViewOrientationTypeEnum.kRightViewOrientation
                ElseIf absX < tolerance AndAlso absZ < tolerance Then
                    logs.Add("CAMDrawingLib: Thickness vector along Y - using Front view")
                    Return ViewOrientationTypeEnum.kFrontViewOrientation
                ElseIf absX < tolerance AndAlso absY < tolerance Then
                    logs.Add("CAMDrawingLib: Thickness vector along Z - using Top view")
                    Return ViewOrientationTypeEnum.kTopViewOrientation
                End If
                
                ' Vector is NOT axis-aligned - need arbitrary camera
                logs.Add("CAMDrawingLib: Thickness vector is custom (" & thicknessAxis & ") - needs arbitrary camera")
                Return ViewOrientationTypeEnum.kArbitraryViewOrientation
            End If
        End If
        
        ' Default to front view
        logs.Add("CAMDrawingLib: Could not parse thickness axis, using Front view")
        Return ViewOrientationTypeEnum.kFrontViewOrientation
    End Function
    
    ' Check if thickness axis requires arbitrary camera (custom vector not axis-aligned)
    Public Function NeedsArbitraryCamera(partDoc As PartDocument) As Boolean
        Dim thicknessAxis As String = GetCustomPropertyValue(partDoc, "BB_ThicknessAxis", "")
        If String.IsNullOrEmpty(thicknessAxis) Then Return False
        If Not thicknessAxis.StartsWith("V:") Then Return False
        
        Dim vx As Double = 0, vy As Double = 0, vz As Double = 0
        If Not ParseVectorComponents(thicknessAxis, vx, vy, vz) Then Return False
        
        ' Check if not axis-aligned
        Dim tolerance As Double = 0.001
        Dim absX As Double = Math.Abs(vx)
        Dim absY As Double = Math.Abs(vy)
        Dim absZ As Double = Math.Abs(vz)
        
        ' Axis-aligned if two components are near zero
        If absY < tolerance AndAlso absZ < tolerance Then Return False
        If absX < tolerance AndAlso absZ < tolerance Then Return False
        If absX < tolerance AndAlso absY < tolerance Then Return False
        
        Return True
    End Function
    
    ' Create a camera for arbitrary view direction based on thickness axis
    ' The camera looks ALONG the thickness axis direction
    Public Function CreateArbitraryCameraFromThicknessAxis(partDoc As PartDocument, _
                                                            app As Inventor.Application, _
                                                            logs As List(Of String)) As Camera
        Dim thicknessAxis As String = GetCustomPropertyValue(partDoc, "BB_ThicknessAxis", "")
        If String.IsNullOrEmpty(thicknessAxis) OrElse Not thicknessAxis.StartsWith("V:") Then
            logs.Add("CAMDrawingLib: No valid vector thickness axis found")
            Return Nothing
        End If
        
        Dim vx As Double = 0, vy As Double = 0, vz As Double = 0
        If Not ParseVectorComponents(thicknessAxis, vx, vy, vz) Then
            logs.Add("CAMDrawingLib: Failed to parse thickness vector")
            Return Nothing
        End If
        
        logs.Add("CAMDrawingLib: Creating arbitrary camera for vector (" & _
                 vx.ToString("F4") & ", " & vy.ToString("F4") & ", " & vz.ToString("F4") & ")")
        
        ' Normalize the view direction vector
        Dim length As Double = Math.Sqrt(vx * vx + vy * vy + vz * vz)
        If length < 0.0001 Then
            logs.Add("CAMDrawingLib: Vector length too small")
            Return Nothing
        End If
        vx /= length
        vy /= length
        vz /= length
        
        ' Get part center from bounding box
        Dim partBox As Box = partDoc.ComponentDefinition.RangeBox
        Dim centerX As Double = (partBox.MinPoint.X + partBox.MaxPoint.X) / 2
        Dim centerY As Double = (partBox.MinPoint.Y + partBox.MaxPoint.Y) / 2
        Dim centerZ As Double = (partBox.MinPoint.Z + partBox.MaxPoint.Z) / 2
        
        ' Calculate a reasonable eye distance
        Dim extentX As Double = partBox.MaxPoint.X - partBox.MinPoint.X
        Dim extentY As Double = partBox.MaxPoint.Y - partBox.MinPoint.Y
        Dim extentZ As Double = partBox.MaxPoint.Z - partBox.MinPoint.Z
        Dim maxExtent As Double = Math.Max(extentX, Math.Max(extentY, extentZ))
        Dim eyeDistance As Double = maxExtent * 2
        
        ' Eye position = center + direction * distance (looking FROM this point)
        Dim eyeX As Double = centerX + vx * eyeDistance
        Dim eyeY As Double = centerY + vy * eyeDistance
        Dim eyeZ As Double = centerZ + vz * eyeDistance
        
        ' Calculate up vector (perpendicular to view direction)
        ' Try to use world Z as reference, but if view direction is along Z, use Y
        Dim upX As Double, upY As Double, upZ As Double
        
        If Math.Abs(vz) > 0.9 Then
            ' View direction is nearly along Z, use Y as up reference
            upX = 0 : upY = 1 : upZ = 0
        Else
            ' Use Z as up reference
            upX = 0 : upY = 0 : upZ = 1
        End If
        
        ' Make up vector perpendicular to view direction using cross product
        ' right = view x up
        Dim rightX As Double = vy * upZ - vz * upY
        Dim rightY As Double = vz * upX - vx * upZ
        Dim rightZ As Double = vx * upY - vy * upX
        
        ' Normalize right vector
        Dim rightLen As Double = Math.Sqrt(rightX * rightX + rightY * rightY + rightZ * rightZ)
        If rightLen > 0.0001 Then
            rightX /= rightLen
            rightY /= rightLen
            rightZ /= rightLen
        End If
        
        ' up = right x view (perpendicular to both)
        upX = rightY * vz - rightZ * vy
        upY = rightZ * vx - rightX * vz
        upZ = rightX * vy - rightY * vx
        
        ' Normalize up vector
        Dim upLen As Double = Math.Sqrt(upX * upX + upY * upY + upZ * upZ)
        If upLen > 0.0001 Then
            upX /= upLen
            upY /= upLen
            upZ /= upLen
        End If
        
        Dim tg As TransientGeometry = app.TransientGeometry
        Dim camera As Camera = Nothing
        
        Try
            camera = app.TransientObjects.CreateCamera()
            logs.Add("CAMDrawingLib: Camera object created")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to create Camera object: " & ex.Message)
            Return Nothing
        End Try
        
        Try
            camera.Eye = tg.CreatePoint(eyeX, eyeY, eyeZ)
            logs.Add("CAMDrawingLib: Eye set to (" & eyeX.ToString("F2") & ", " & eyeY.ToString("F2") & ", " & eyeZ.ToString("F2") & ")")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to set Eye: " & ex.Message)
            Return Nothing
        End Try
        
        Try
            camera.Target = tg.CreatePoint(centerX, centerY, centerZ)
            logs.Add("CAMDrawingLib: Target set to (" & centerX.ToString("F2") & ", " & centerY.ToString("F2") & ", " & centerZ.ToString("F2") & ")")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to set Target: " & ex.Message)
            Return Nothing
        End Try
        
        Try
            camera.UpVector = tg.CreateUnitVector(upX, upY, upZ)
            logs.Add("CAMDrawingLib: UpVector set to (" & upX.ToString("F4") & ", " & upY.ToString("F4") & ", " & upZ.ToString("F4") & ")")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to set UpVector: " & ex.Message)
            ' Try setting individual components instead
            Try
                logs.Add("CAMDrawingLib: Trying alternative up vector approach...")
                Dim upVec As UnitVector = tg.CreateUnitVector(upX, upY, upZ)
                camera.UpVector = upVec
                logs.Add("CAMDrawingLib: UpVector set successfully via UnitVector")
            Catch ex2 As Exception
                logs.Add("CAMDrawingLib: Alternative also failed: " & ex2.Message)
                Return Nothing
            End Try
        End Try
        
        Try
            camera.ViewOrientationType = ViewOrientationTypeEnum.kArbitraryViewOrientation
            logs.Add("CAMDrawingLib: ViewOrientationType set to kArbitraryViewOrientation")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to set ViewOrientationType: " & ex.Message)
            ' This might be read-only, continue anyway
        End Try
        
        Try
            camera.Perspective = False
            logs.Add("CAMDrawingLib: Perspective set to False (orthographic)")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to set Perspective: " & ex.Message)
            ' Continue anyway
        End Try
        
        logs.Add("CAMDrawingLib: Arbitrary camera created successfully")
        Return camera
    End Function
    
    ' Create a camera for the OPPOSITE view direction (looking from the other side)
    Public Function CreateOppositeCameraFromThicknessAxis(partDoc As PartDocument, _
                                                           app As Inventor.Application, _
                                                           logs As List(Of String)) As Camera
        Dim thicknessAxis As String = GetCustomPropertyValue(partDoc, "BB_ThicknessAxis", "")
        If String.IsNullOrEmpty(thicknessAxis) OrElse Not thicknessAxis.StartsWith("V:") Then
            logs.Add("CAMDrawingLib: No valid vector thickness axis found for opposite camera")
            Return Nothing
        End If
        
        Dim vx As Double = 0, vy As Double = 0, vz As Double = 0
        If Not ParseVectorComponents(thicknessAxis, vx, vy, vz) Then
            logs.Add("CAMDrawingLib: Failed to parse thickness vector for opposite camera")
            Return Nothing
        End If
        
        ' NEGATE the direction for opposite view
        vx = -vx
        vy = -vy
        vz = -vz
        
        logs.Add("CAMDrawingLib: Creating opposite camera for vector (" & _
                 vx.ToString("F4") & ", " & vy.ToString("F4") & ", " & vz.ToString("F4") & ")")
        
        ' Normalize the view direction vector
        Dim length As Double = Math.Sqrt(vx * vx + vy * vy + vz * vz)
        If length < 0.0001 Then
            logs.Add("CAMDrawingLib: Vector length too small")
            Return Nothing
        End If
        vx /= length
        vy /= length
        vz /= length
        
        ' Get part center from bounding box
        Dim partBox As Box = partDoc.ComponentDefinition.RangeBox
        Dim centerX As Double = (partBox.MinPoint.X + partBox.MaxPoint.X) / 2
        Dim centerY As Double = (partBox.MinPoint.Y + partBox.MaxPoint.Y) / 2
        Dim centerZ As Double = (partBox.MinPoint.Z + partBox.MaxPoint.Z) / 2
        
        ' Calculate a reasonable eye distance
        Dim extentX As Double = partBox.MaxPoint.X - partBox.MinPoint.X
        Dim extentY As Double = partBox.MaxPoint.Y - partBox.MinPoint.Y
        Dim extentZ As Double = partBox.MaxPoint.Z - partBox.MinPoint.Z
        Dim maxExtent As Double = Math.Max(extentX, Math.Max(extentY, extentZ))
        Dim eyeDistance As Double = maxExtent * 2
        
        ' Eye position = center + direction * distance
        Dim eyeX As Double = centerX + vx * eyeDistance
        Dim eyeY As Double = centerY + vy * eyeDistance
        Dim eyeZ As Double = centerZ + vz * eyeDistance
        
        ' Calculate up vector (same logic as base camera)
        Dim upX As Double, upY As Double, upZ As Double
        
        If Math.Abs(vz) > 0.9 Then
            upX = 0 : upY = 1 : upZ = 0
        Else
            upX = 0 : upY = 0 : upZ = 1
        End If
        
        ' Make up vector perpendicular
        Dim rightX As Double = vy * upZ - vz * upY
        Dim rightY As Double = vz * upX - vx * upZ
        Dim rightZ As Double = vx * upY - vy * upX
        
        Dim rightLen As Double = Math.Sqrt(rightX * rightX + rightY * rightY + rightZ * rightZ)
        If rightLen > 0.0001 Then
            rightX /= rightLen
            rightY /= rightLen
            rightZ /= rightLen
        End If
        
        upX = rightY * vz - rightZ * vy
        upY = rightZ * vx - rightX * vz
        upZ = rightX * vy - rightY * vx
        
        Dim upLen As Double = Math.Sqrt(upX * upX + upY * upY + upZ * upZ)
        If upLen > 0.0001 Then
            upX /= upLen
            upY /= upLen
            upZ /= upLen
        End If
        
        Dim tg As TransientGeometry = app.TransientGeometry
        Dim camera As Camera = Nothing
        
        Try
            camera = app.TransientObjects.CreateCamera()
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to create opposite Camera object: " & ex.Message)
            Return Nothing
        End Try
        
        Try
            camera.Eye = tg.CreatePoint(eyeX, eyeY, eyeZ)
            camera.Target = tg.CreatePoint(centerX, centerY, centerZ)
            camera.UpVector = tg.CreateUnitVector(upX, upY, upZ)
            logs.Add("CAMDrawingLib: Opposite camera configured")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to configure opposite camera: " & ex.Message)
            Return Nothing
        End Try
        
        Return camera
    End Function
    
    ' Parse vector format "V:x,y,z"
    Private Function ParseVectorComponents(axis As String, ByRef vx As Double, ByRef vy As Double, ByRef vz As Double) As Boolean
        If Not axis.StartsWith("V:") Then Return False
        Try
            Dim parts() As String = axis.Substring(2).Split(","c)
            If parts.Length <> 3 Then Return False
            vx = Double.Parse(parts(0), System.Globalization.CultureInfo.InvariantCulture)
            vy = Double.Parse(parts(1), System.Globalization.CultureInfo.InvariantCulture)
            vz = Double.Parse(parts(2), System.Globalization.CultureInfo.InvariantCulture)
            Return True
        Catch
            Return False
        End Try
    End Function
    
    ' Get custom property value from part
    Private Function GetCustomPropertyValue(partDoc As PartDocument, propName As String, defaultValue As String) As String
        Try
            Dim propSets As PropertySets = partDoc.PropertySets
            Dim userProps As PropertySet = propSets.Item("Inventor User Defined Properties")
            Return CStr(userProps.Item(propName).Value)
        Catch
            Return defaultValue
        End Try
    End Function

    ' ============================================================================
    ' Drawing Creation
    ' ============================================================================
    
    ' Create a new drawing from template (auto-find template)
    Public Function CreateDrawingFromTemplate(app As Inventor.Application, _
                                              logs As List(Of String)) As DrawingDocument
        Dim templatePath As String = FindDrawingTemplate(app, logs)
        Return CreateDrawingFromTemplate(app, templatePath, logs)
    End Function
    
    ' Create a new drawing from specific template
    Public Function CreateDrawingFromTemplate(app As Inventor.Application, _
                                              templatePath As String, _
                                              logs As List(Of String)) As DrawingDocument
        If String.IsNullOrEmpty(templatePath) Then
            logs.Add("CAMDrawingLib: Template path is empty")
            Return Nothing
        End If
        
        If Not System.IO.File.Exists(templatePath) Then
            logs.Add("CAMDrawingLib: Template not found: " & templatePath)
            Return Nothing
        End If
        
        Try
            Dim drawDoc As DrawingDocument = CType( _
                app.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, templatePath, True), _
                DrawingDocument)
            logs.Add("CAMDrawingLib: Drawing created from template")
            Return drawDoc
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to create drawing: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ' Find the drawing template
    Public Function FindDrawingTemplate(app As Inventor.Application, _
                                        logs As List(Of String)) As String
        ' Try specific template first
        Try
            Dim templatePath As String = app.FileManager.GetTemplateFile(DocumentTypeEnum.kDrawingDocumentObject)
            If Not String.IsNullOrEmpty(templatePath) Then
                Dim templateFolder As String = System.IO.Path.GetDirectoryName(templatePath)
                
                ' Look for our specific template
                Dim specificPath As String = System.IO.Path.Combine(templateFolder, DEFAULT_TEMPLATE)
                If System.IO.File.Exists(specificPath) Then
                    logs.Add("CAMDrawingLib: Found template: " & specificPath)
                    Return specificPath
                End If
                
                ' Fall back to default template
                If System.IO.File.Exists(templatePath) Then
                    logs.Add("CAMDrawingLib: Using default template: " & templatePath)
                    Return templatePath
                End If
            End If
        Catch
        End Try
        
        ' Try project-specific templates
        Try
            Dim project As DesignProject = app.DesignProjectManager.ActiveDesignProject
            Dim templatesPath As String = project.TemplatesPath
            If Not String.IsNullOrEmpty(templatesPath) Then
                Dim specificPath As String = System.IO.Path.Combine(templatesPath, DEFAULT_TEMPLATE)
                If System.IO.File.Exists(specificPath) Then
                    logs.Add("CAMDrawingLib: Found project template: " & specificPath)
                    Return specificPath
                End If
                
                ' Look for any .idw template
                Dim idwFiles() As String = System.IO.Directory.GetFiles(templatesPath, "*.idw")
                If idwFiles.Length > 0 Then
                    logs.Add("CAMDrawingLib: Using first available template: " & idwFiles(0))
                    Return idwFiles(0)
                End If
            End If
        Catch
        End Try
        
        logs.Add("CAMDrawingLib: No drawing template found")
        Return ""
    End Function
    
    ' Open an existing drawing document
    Public Function OpenExistingDrawing(app As Inventor.Application, _
                                        drawingPath As String, _
                                        logs As List(Of String)) As DrawingDocument
        If String.IsNullOrEmpty(drawingPath) Then
            logs.Add("CAMDrawingLib: Drawing path is empty")
            Return Nothing
        End If
        
        If Not System.IO.File.Exists(drawingPath) Then
            logs.Add("CAMDrawingLib: Drawing not found: " & drawingPath)
            Return Nothing
        End If
        
        Try
            Dim drawDoc As DrawingDocument = CType( _
                app.Documents.Open(drawingPath, True), _
                DrawingDocument)
            logs.Add("CAMDrawingLib: Opened drawing: " & drawingPath)
            Return drawDoc
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to open drawing: " & ex.Message)
            Return Nothing
        End Try
    End Function
    
    ' Find drawing template by name
    Public Function FindDrawingTemplate(app As Inventor.Application, _
                                        templateName As String, _
                                        logs As List(Of String)) As String
        If String.IsNullOrEmpty(templateName) Then
            templateName = DEFAULT_TEMPLATE
        End If
        
        ' Try FileManager template path first
        Try
            Dim defaultPath As String = app.FileManager.GetTemplateFile(DocumentTypeEnum.kDrawingDocumentObject)
            If Not String.IsNullOrEmpty(defaultPath) Then
                Dim templateFolder As String = System.IO.Path.GetDirectoryName(defaultPath)
                
                ' Look for specific template
                Dim specificPath As String = System.IO.Path.Combine(templateFolder, templateName)
                If System.IO.File.Exists(specificPath) Then
                    logs.Add("CAMDrawingLib: Found template: " & specificPath)
                    Return specificPath
                End If
                
                ' Fall back to default
                If System.IO.File.Exists(defaultPath) Then
                    logs.Add("CAMDrawingLib: Using default template: " & defaultPath)
                    Return defaultPath
                End If
            End If
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Error accessing FileManager: " & ex.Message)
        End Try
        
        ' Try project templates path
        Try
            Dim project As DesignProject = app.DesignProjectManager.ActiveDesignProject
            Dim templatesPath As String = project.TemplatesPath
            If Not String.IsNullOrEmpty(templatesPath) Then
                Dim specificPath As String = System.IO.Path.Combine(templatesPath, templateName)
                If System.IO.File.Exists(specificPath) Then
                    logs.Add("CAMDrawingLib: Found template in project: " & specificPath)
                    Return specificPath
                End If
                
                ' Look for any .idw
                Dim idwFiles() As String = System.IO.Directory.GetFiles(templatesPath, "*.idw")
                If idwFiles.Length > 0 Then
                    logs.Add("CAMDrawingLib: Using first available template: " & idwFiles(0))
                    Return idwFiles(0)
                End If
            End If
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Error accessing project templates: " & ex.Message)
        End Try
        
        logs.Add("CAMDrawingLib: No drawing template found")
        Return ""
    End Function

    ' ============================================================================
    ' Sheet Management
    ' ============================================================================
    
    ' Resize sheet to custom dimensions (in mm)
    Public Sub ResizeSheet(sheet As Sheet, _
                           widthMm As Double, _
                           heightMm As Double, _
                           logs As List(Of String))
        Dim widthCm As Double = widthMm / 10
        Dim heightCm As Double = heightMm / 10
        
        logs.Add("CAMDrawingLib: Attempting to resize sheet to " & widthMm.ToString("F1") & " x " & heightMm.ToString("F1") & " mm")
        logs.Add("CAMDrawingLib: Current sheet size: " & (sheet.Width * 10).ToString("F1") & " x " & (sheet.Height * 10).ToString("F1") & " mm")
        
        ' Ensure sheet is set to custom size first
        Try
            If sheet.Size <> DrawingSheetSizeEnum.kCustomDrawingSheetSize Then
                logs.Add("CAMDrawingLib: Setting sheet to custom size mode")
                sheet.Size = DrawingSheetSizeEnum.kCustomDrawingSheetSize
            End If
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Could not set custom size mode: " & ex.Message)
        End Try
        
        Try
            ' Method 1: Try Resize method
            sheet.Resize(DrawingSheetSizeEnum.kCustomDrawingSheetSize, widthCm, heightCm)
            
            Dim actualWidth As Double = sheet.Width * 10
            Dim actualHeight As Double = sheet.Height * 10
            logs.Add("CAMDrawingLib: Sheet resized (Resize method) to " & actualWidth.ToString("F1") & " x " & actualHeight.ToString("F1") & " mm")
        Catch ex1 As Exception
            logs.Add("CAMDrawingLib: Resize method failed: " & ex1.Message)
            
            ' Method 2: Try direct property assignment
            ' NOTE: Setting Width may actually affect Height in some cases - try both orders
            Try
                logs.Add("CAMDrawingLib: Before any changes: " & (sheet.Width * 10).ToString("F1") & " x " & (sheet.Height * 10).ToString("F1") & " mm")
                
                ' Try setting Height first, then Width
                sheet.Height = heightCm
                logs.Add("CAMDrawingLib: After setting Height=" & (heightCm * 10).ToString("F1") & ": " & (sheet.Width * 10).ToString("F1") & " x " & (sheet.Height * 10).ToString("F1") & " mm")
                
                sheet.Width = widthCm
                logs.Add("CAMDrawingLib: After setting Width=" & (widthCm * 10).ToString("F1") & ": " & (sheet.Width * 10).ToString("F1") & " x " & (sheet.Height * 10).ToString("F1") & " mm")
                
                ' Check if we need to swap (in case Width/Height are swapped in landscape mode)
                If Math.Abs(sheet.Width * 10 - widthMm) > 1 AndAlso Math.Abs(sheet.Height * 10 - widthMm) < 1 Then
                    logs.Add("CAMDrawingLib: Detected Width/Height swap - sheet may be in portrait mode, trying swap...")
                    sheet.Width = heightCm
                    sheet.Height = widthCm
                    logs.Add("CAMDrawingLib: After swap: " & (sheet.Width * 10).ToString("F1") & " x " & (sheet.Height * 10).ToString("F1") & " mm")
                End If
                
                logs.Add("CAMDrawingLib: Sheet resized (direct) to " & (sheet.Width * 10).ToString("F1") & " x " & (sheet.Height * 10).ToString("F1") & " mm")
            Catch ex2 As Exception
                logs.Add("CAMDrawingLib: Failed to resize sheet: " & ex2.Message)
            End Try
        End Try
    End Sub
    
    ' Calculate sheet size from actual view bounds + dimension space + border padding
    ' This is the preferred method - call AFTER creating views and adding dimensions
    ' Returns (width, height) in mm
    ' 
    ' dimensionOffsetMm: space reserved for dimensions on each view (default DIMENSION_OFFSET * 10)
    '                    Applied to right and bottom of EACH view for extent dimensions
    ' borderPaddingMm:   minimum border around all content (default 15mm)
    Public Function CalculateSheetSizeFromViews(views As List(Of DrawingView), _
                                                 logs As List(Of String), _
                                                 Optional dimensionOffsetMm As Double = -1, _
                                                 Optional borderPaddingMm As Double = 15) As Double()
        If views Is Nothing OrElse views.Count = 0 Then
            logs.Add("CAMDrawingLib: No views to calculate size from")
            Return New Double() {100, 80}  ' Minimum default
        End If
        
        ' Use default dimension offset if not specified
        If dimensionOffsetMm < 0 Then
            dimensionOffsetMm = DIMENSION_OFFSET * 10  ' Convert cm constant to mm
        End If
        
        Dim dimOffsetCm As Double = dimensionOffsetMm / 10  ' Convert mm to cm for internal calculations
        
        ' Find bounding box of all views INCLUDING dimension space
        ' Dimension space is added to right (for vertical dim) and bottom (for horizontal dim) of each view
        Dim minX As Double = Double.MaxValue
        Dim minY As Double = Double.MaxValue
        Dim maxX As Double = Double.MinValue
        Dim maxY As Double = Double.MinValue
        
        For Each view As DrawingView In views
            ' Use GetViewBoundsWithDimensions to get bounds including dim space
            Dim bounds() As Double = GetViewBoundsWithDimensions(view, dimOffsetCm)
            Dim vLeft As Double = bounds(0)
            Dim vRight As Double = bounds(1)
            Dim vBottom As Double = bounds(2)
            Dim vTop As Double = bounds(3)
            
            If vLeft < minX Then minX = vLeft
            If vRight > maxX Then maxX = vRight
            If vBottom < minY Then minY = vBottom
            If vTop > maxY Then maxY = vTop
        Next
        
        ' Content bounds in mm (views + dimension space)
        Dim contentWidth As Double = (maxX - minX) * 10
        Dim contentHeight As Double = (maxY - minY) * 10
        
        logs.Add("CAMDrawingLib: Views bounding box (incl. " & FormatNumber(dimensionOffsetMm, 0) & "mm dim space): " & _
                 FormatNumber(contentWidth, 1) & " x " & FormatNumber(contentHeight, 1) & " mm")
        
        ' Add border padding on all sides
        Dim sheetWidth As Double = contentWidth + borderPaddingMm * 2
        Dim sheetHeight As Double = contentHeight + borderPaddingMm * 2
        
        ' Minimum practical size
        sheetWidth = Math.Max(sheetWidth, 100)
        sheetHeight = Math.Max(sheetHeight, 80)
        
        logs.Add("CAMDrawingLib: Sheet size with " & FormatNumber(borderPaddingMm, 0) & "mm border: " & _
                 FormatNumber(sheetWidth, 1) & " x " & FormatNumber(sheetHeight, 1) & " mm")
        
        Return New Double() {sheetWidth, sheetHeight}
    End Function
    
    ' DEPRECATED: Old signature for backward compatibility
    ' Use the new signature with named parameters instead
    Public Function CalculateSheetSizeFromViewsLegacy(views As List(Of DrawingView), _
                                                       paddingPercent As Double, _
                                                       logs As List(Of String)) As Double()
        If views Is Nothing OrElse views.Count = 0 Then
            logs.Add("CAMDrawingLib: No views to calculate size from")
            Return New Double() {100, 80}
        End If
        
        Dim minX As Double = Double.MaxValue
        Dim minY As Double = Double.MaxValue
        Dim maxX As Double = Double.MinValue
        Dim maxY As Double = Double.MinValue
        
        For Each view As DrawingView In views
            Dim vLeft As Double = view.Position.X - view.Width / 2
            Dim vRight As Double = view.Position.X + view.Width / 2
            Dim vBottom As Double = view.Position.Y - view.Height / 2
            Dim vTop As Double = view.Position.Y + view.Height / 2
            
            If vLeft < minX Then minX = vLeft
            If vRight > maxX Then maxX = vRight
            If vBottom < minY Then minY = vBottom
            If vTop > maxY Then maxY = vTop
        Next
        
        Dim viewsWidth As Double = (maxX - minX) * 10
        Dim viewsHeight As Double = (maxY - minY) * 10
        
        Dim sheetWidth As Double = viewsWidth * (1 + 2 * paddingPercent)
        Dim sheetHeight As Double = viewsHeight * (1 + 2 * paddingPercent)
        
        sheetWidth = Math.Max(sheetWidth, 100)
        sheetHeight = Math.Max(sheetHeight, 80)
        
        logs.Add("CAMDrawingLib: (Legacy) Sheet size with " & FormatNumber(paddingPercent * 100, 0) & "% padding: " & _
                 FormatNumber(sheetWidth, 1) & " x " & FormatNumber(sheetHeight, 1) & " mm")
        
        Return New Double() {sheetWidth, sheetHeight}
    End Function
    
    ' Fit sheet to views and center them
    ' This handles the correct order: move views first, then resize sheet
    ' dimensionOffsetMm: space reserved for dimensions (default DIMENSION_OFFSET * 10 = 25mm)
    ' borderPaddingMm: minimum border around all views (default 10mm)
    Public Sub FitSheetToViews(sheet As Sheet, _
                                views As List(Of DrawingView), _
                                app As Inventor.Application, _
                                logs As List(Of String), _
                                Optional dimensionOffsetMm As Double = -1, _
                                Optional borderPaddingMm As Double = 10)
        If views Is Nothing OrElse views.Count = 0 Then
            logs.Add("CAMDrawingLib: No views to fit sheet to")
            Return
        End If
        
        ' Calculate required sheet size
        Dim requiredSize() As Double = CalculateSheetSizeFromViews(views, logs, dimensionOffsetMm, borderPaddingMm)
        Dim requiredWidthCm As Double = requiredSize(0) / 10
        Dim requiredHeightCm As Double = requiredSize(1) / 10
        
        ' Get dimension offset in cm for bounds calculation
        Dim dimOffsetCm As Double
        If dimensionOffsetMm < 0 Then
            dimOffsetCm = DIMENSION_OFFSET
        Else
            dimOffsetCm = dimensionOffsetMm / 10
        End If
        
        ' Find current view bounds INCLUDING dimension space using helper function
        Dim minX As Double = Double.MaxValue
        Dim minY As Double = Double.MaxValue
        Dim maxX As Double = Double.MinValue
        Dim maxY As Double = Double.MinValue
        
        For Each view As DrawingView In views
            Dim bounds() As Double = GetViewBoundsWithDimensions(view, dimOffsetCm)
            
            If bounds(0) < minX Then minX = bounds(0)  ' left
            If bounds(1) > maxX Then maxX = bounds(1)  ' right
            If bounds(2) < minY Then minY = bounds(2)  ' bottom
            If bounds(3) > maxY Then maxY = bounds(3)  ' top
        Next
        
        ' Calculate offset to move views to be centered on the NEW sheet size
        Dim viewsCenterX As Double = (minX + maxX) / 2
        Dim viewsCenterY As Double = (minY + maxY) / 2
        Dim newSheetCenterX As Double = requiredWidthCm / 2
        Dim newSheetCenterY As Double = requiredHeightCm / 2
        
        Dim offsetX As Double = newSheetCenterX - viewsCenterX
        Dim offsetY As Double = newSheetCenterY - viewsCenterY
        
        ' Move all views to fit within new sheet bounds
        For Each view As DrawingView In views
            Dim newPos As Point2d = app.TransientGeometry.CreatePoint2d( _
                view.Position.X + offsetX, _
                view.Position.Y + offsetY)
            view.Position = newPos
        Next
        
        logs.Add("CAMDrawingLib: Views moved (offset: " & FormatNumber(offsetX * 10, 1) & ", " & _
                 FormatNumber(offsetY * 10, 1) & " mm)")
        
        ' NOW resize the sheet (views are within bounds)
        ResizeSheet(sheet, requiredSize(0), requiredSize(1), logs)
    End Sub
    
    ' Center all views on the sheet (current sheet size)
    Public Sub CenterViewsOnSheet(sheet As Sheet, _
                                   views As List(Of DrawingView), _
                                   app As Inventor.Application, _
                                   logs As List(Of String))
        If views Is Nothing OrElse views.Count = 0 Then
            logs.Add("CAMDrawingLib: No views to center")
            Return
        End If
        
        ' Find current view bounds
        Dim minX As Double = Double.MaxValue
        Dim minY As Double = Double.MaxValue
        Dim maxX As Double = Double.MinValue
        Dim maxY As Double = Double.MinValue
        
        For Each view As DrawingView In views
            Dim vLeft As Double = view.Position.X - view.Width / 2
            Dim vRight As Double = view.Position.X + view.Width / 2
            Dim vBottom As Double = view.Position.Y - view.Height / 2
            Dim vTop As Double = view.Position.Y + view.Height / 2
            
            If vLeft < minX Then minX = vLeft
            If vRight > maxX Then maxX = vRight
            If vBottom < minY Then minY = vBottom
            If vTop > maxY Then maxY = vTop
        Next
        
        ' Calculate offset to center views
        Dim viewsCenterX As Double = (minX + maxX) / 2
        Dim viewsCenterY As Double = (minY + maxY) / 2
        Dim sheetCenterX As Double = sheet.Width / 2
        Dim sheetCenterY As Double = sheet.Height / 2
        
        Dim offsetX As Double = sheetCenterX - viewsCenterX
        Dim offsetY As Double = sheetCenterY - viewsCenterY
        
        ' Move all views
        For Each view As DrawingView In views
            Dim newPos As Point2d = app.TransientGeometry.CreatePoint2d( _
                view.Position.X + offsetX, _
                view.Position.Y + offsetY)
            view.Position = newPos
        Next
        
        logs.Add("CAMDrawingLib: Views centered (offset: " & FormatNumber(offsetX * 10, 1) & ", " & _
                 FormatNumber(offsetY * 10, 1) & " mm)")
    End Sub
    
    ' Calculate required sheet size based on part extents (DEPRECATED - use CalculateSheetSizeFromViews)
    ' Returns (width, height) in mm
    ' Handles sheet metal flat patterns (single view) and regular parts (T-layout)
    Public Function CalculateSheetSize(partDoc As PartDocument, _
                                       dimSpaceMm As Double, _
                                       logs As List(Of String)) As Double()
        ' Check for sheet metal flat pattern
        If IsSheetMetal(partDoc) AndAlso HasFlatPattern(partDoc) Then
            Return CalculateFlatPatternSheetSize(partDoc, dimSpaceMm, logs)
        End If
        
        Dim partBox As Box = partDoc.ComponentDefinition.RangeBox
        
        ' Part dimensions in mm
        Dim xSize As Double = (partBox.MaxPoint.X - partBox.MinPoint.X) * 10
        Dim ySize As Double = (partBox.MaxPoint.Y - partBox.MinPoint.Y) * 10
        Dim zSize As Double = (partBox.MaxPoint.Z - partBox.MinPoint.Z) * 10
        
        logs.Add("CAMDrawingLib: Part size: " & FormatNumber(xSize, 1) & " x " & _
                 FormatNumber(ySize, 1) & " x " & FormatNumber(zSize, 1) & " mm")
        
        ' T-layout: 
        ' Horizontal: Left + Front + Right + Back = Y + X + Y + X = 2X + 2Y + 5*dimSpace
        ' Vertical: Top + Front + Bottom = Y + Z + Y = 2Y + Z + 4*dimSpace
        ' But views share edges, so:
        ' Width: 4 views horizontally at most = 2*X + 2*Y + 5*dimSpace
        ' Height: 3 views vertically = Z + 2*Y + 4*dimSpace (but Y is depth shown in top/bottom)
        
        ' Simplified calculation for T-layout:
        ' Width = Left(Y) + Front(X) + Right(Y) + Back(X) + margins = 2X + 2Y + 5*dimSpace
        ' Height = Top(Y) + Front(Z) + Bottom(Y) + margins = 2Y + Z + 4*dimSpace
        
        ' Actually, let's think about what each view shows:
        ' Front/Back: width=X, height=Z
        ' Left/Right: width=Y, height=Z
        ' Top/Bottom: width=X, height=Y
        
        ' T-layout horizontal: Left(Y) + space + Front(X) + space + Right(Y) + space + Back(X) + space (margins)
        ' Total width: Y + X + Y + X + 5*dimSpace = 2X + 2Y + 5*dimSpace
        
        ' T-layout vertical: Top(Y) + space + Front(Z) + space + Bottom(Y) + space (margins)
        ' Total height: Y + Z + Y + 4*dimSpace = 2Y + Z + 4*dimSpace
        
        Dim sheetWidth As Double = 2 * xSize + 2 * ySize + 5 * dimSpaceMm
        Dim sheetHeight As Double = 2 * ySize + zSize + 4 * dimSpaceMm
        
        ' Add margins
        sheetWidth += 2 * dimSpaceMm
        sheetHeight += 2 * dimSpaceMm
        
        ' Minimum practical size
        sheetWidth = Math.Max(sheetWidth, 100)
        sheetHeight = Math.Max(sheetHeight, 80)
        
        logs.Add("CAMDrawingLib: Calculated sheet size: " & FormatNumber(sheetWidth, 1) & " x " & _
                 FormatNumber(sheetHeight, 1) & " mm")
        
        Return New Double() {sheetWidth, sheetHeight}
    End Function
    
    ' Calculate sheet size for sheet metal flat pattern (single view)
    Private Function CalculateFlatPatternSheetSize(partDoc As PartDocument, _
                                                    dimSpaceMm As Double, _
                                                    logs As List(Of String)) As Double()
        Dim fpWidth As Double = 100  ' Default fallback in mm
        Dim fpHeight As Double = 80
        
        Try
            Dim smCompDef As SheetMetalComponentDefinition = CType(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
            If smCompDef.HasFlatPattern Then
                ' Get flat pattern extent from the FlatPattern property
                Dim fp As FlatPattern = smCompDef.FlatPattern
                
                ' Width and Length are Double properties (in cm)
                fpWidth = fp.Width * 10   ' Convert cm to mm
                fpHeight = fp.Length * 10 ' Convert cm to mm
                
                logs.Add("CAMDrawingLib: Flat pattern size: " & FormatNumber(fpWidth, 1) & " x " & _
                         FormatNumber(fpHeight, 1) & " mm")
            End If
        Catch ex As Exception
            ' Fall back to bounding box if flat pattern dimensions not accessible
            logs.Add("CAMDrawingLib: Could not get flat pattern dimensions: " & ex.Message)
            Dim partBox As Box = partDoc.ComponentDefinition.RangeBox
            fpWidth = (partBox.MaxPoint.X - partBox.MinPoint.X) * 10
            fpHeight = (partBox.MaxPoint.Y - partBox.MinPoint.Y) * 10
            logs.Add("CAMDrawingLib: Using bounding box: " & FormatNumber(fpWidth, 1) & " x " & _
                     FormatNumber(fpHeight, 1) & " mm")
        End Try
        
        ' For single view: view size + margins for dimensions
        Dim sheetWidth As Double = fpWidth + 4 * dimSpaceMm
        Dim sheetHeight As Double = fpHeight + 4 * dimSpaceMm
        
        ' Minimum practical size
        sheetWidth = Math.Max(sheetWidth, 100)
        sheetHeight = Math.Max(sheetHeight, 80)
        
        logs.Add("CAMDrawingLib: Sheet size for flat pattern: " & FormatNumber(sheetWidth, 1) & " x " & _
                 FormatNumber(sheetHeight, 1) & " mm")
        
        Return New Double() {sheetWidth, sheetHeight}
    End Function

    ' ============================================================================
    ' View Placement
    ' ============================================================================
    
    ' Add all 6 orthographic views at 1:1 scale in T-layout
    ' Returns list of created views
    ' Handles: sheet metal (flat pattern), BB_ThicknessAxis orientation, default front
    Public Function AddAllViews(sheet As Sheet, _
                                partDoc As PartDocument, _
                                app As Inventor.Application, _
                                logs As List(Of String)) As List(Of DrawingView)
        Dim views As New List(Of DrawingView)
        
        ' Get part dimensions for spacing calculation
        Dim partBox As Box = partDoc.ComponentDefinition.RangeBox
        Dim xSize As Double = (partBox.MaxPoint.X - partBox.MinPoint.X)  ' in cm
        Dim ySize As Double = (partBox.MaxPoint.Y - partBox.MinPoint.Y)
        Dim zSize As Double = (partBox.MaxPoint.Z - partBox.MinPoint.Z)
        
        ' Dimension spacing (in cm)
        ' Must account for: dimension line (DIMENSION_OFFSET) + gap between dimension and next view
        ' Minimum DIMENSION_OFFSET + 1.5cm = 4.0cm (40mm), or 10% of max dimension for larger parts
        Dim maxExtent As Double = Math.Max(xSize, Math.Max(ySize, zSize))
        Dim dimSpace As Double = Math.Max(DIMENSION_OFFSET + 1.5, maxExtent * 0.10)  ' At least 40mm, or 10% of max
        
        ' Determine base view orientation
        Dim baseOrientation As ViewOrientationTypeEnum = DetermineBaseViewOrientation(partDoc, logs)
        
        ' Check for sheet metal flat pattern
        Dim useSheetMetalFlatPattern As Boolean = IsSheetMetal(partDoc) AndAlso HasFlatPattern(partDoc)
        If useSheetMetalFlatPattern Then
            logs.Add("CAMDrawingLib: Will use sheet metal flat pattern view")
        End If
        
        ' Calculate front view position (center of sheet, adjusted for T-layout)
        ' Front view should be positioned so all views fit
        Dim frontX As Double = sheet.Width / 2 - (ySize / 2)  ' Shift left to make room for right + back
        Dim frontY As Double = sheet.Height / 2
        
        ' Add base view (either flat pattern or oriented view)
        Dim baseView As DrawingView = Nothing
        Try
            If useSheetMetalFlatPattern Then
                ' For sheet metal, create flat pattern view
                baseView = CreateFlatPatternView(sheet, partDoc, app, _
                    app.TransientGeometry.CreatePoint2d(frontX, frontY), logs)
                
                If baseView IsNot Nothing Then
                    views.Add(baseView)
                    logs.Add("CAMDrawingLib: Flat pattern base view added")
                Else
                    ' Fall back to regular view creation if flat pattern failed
                    logs.Add("CAMDrawingLib: Flat pattern failed, falling back to regular view")
                    useSheetMetalFlatPattern = False
                End If
            End If
            
            If Not useSheetMetalFlatPattern Then
                ' Check if we need arbitrary camera for custom thickness axis
                Dim baseName As String = "Front"
                
                If baseOrientation = ViewOrientationTypeEnum.kArbitraryViewOrientation Then
                    ' Use arbitrary camera for custom vector thickness axis
                    Dim camera As Camera = CreateArbitraryCameraFromThicknessAxis(partDoc, app, logs)
                    
                    If camera IsNot Nothing Then
                        Try
                            ' Try method 1: Pass camera as 7th parameter
                            baseView = sheet.DrawingViews.AddBaseView( _
                                partDoc, _
                                app.TransientGeometry.CreatePoint2d(frontX, frontY), _
                                1.0, _
                                ViewOrientationTypeEnum.kArbitraryViewOrientation, _
                                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                                "", _
                                camera)
                            baseName = "Custom"
                            logs.Add("CAMDrawingLib: Created base view with arbitrary camera (method 1)")
                        Catch ex1 As Exception
                            logs.Add("CAMDrawingLib: Method 1 failed: " & ex1.Message)
                            
                            Try
                                ' Try method 2: Use NameValueMap
                                Dim viewOptions As NameValueMap = app.TransientObjects.CreateNameValueMap()
                                viewOptions.Add("Camera", camera)
                                
                                baseView = sheet.DrawingViews.AddBaseView( _
                                    partDoc, _
                                    app.TransientGeometry.CreatePoint2d(frontX, frontY), _
                                    1.0, _
                                    ViewOrientationTypeEnum.kArbitraryViewOrientation, _
                                    DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                                    AdditionalOptions := viewOptions)
                                baseName = "Custom"
                                logs.Add("CAMDrawingLib: Created base view with arbitrary camera (method 2)")
                            Catch ex2 As Exception
                                logs.Add("CAMDrawingLib: Method 2 failed: " & ex2.Message)
                                
                                ' Try method 3: Set view orientation type from camera
                                Try
                                    baseView = sheet.DrawingViews.AddBaseView( _
                                        partDoc, _
                                        app.TransientGeometry.CreatePoint2d(frontX, frontY), _
                                        1.0, _
                                        camera.ViewOrientationType, _
                                        DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                                        Nothing, _
                                        camera)
                                    baseName = "Custom"
                                    logs.Add("CAMDrawingLib: Created base view with arbitrary camera (method 3)")
                                Catch ex3 As Exception
                                    logs.Add("CAMDrawingLib: Method 3 failed: " & ex3.Message)
                                End Try
                            End Try
                        End Try
                    End If
                    
                    If baseView Is Nothing Then
                        ' Fall back to front view if all methods failed
                        logs.Add("CAMDrawingLib: All arbitrary camera methods failed, falling back to Front view")
                        baseOrientation = ViewOrientationTypeEnum.kFrontViewOrientation
                    End If
                End If
                
                ' Create standard orientation view if not already created
                If baseView Is Nothing Then
                    baseView = sheet.DrawingViews.AddBaseView( _
                        partDoc, _
                        app.TransientGeometry.CreatePoint2d(frontX, frontY), _
                        1.0, _
                        baseOrientation, _
                        DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                        Nothing, Nothing)
                    
                    baseName = GetViewOrientationName(baseOrientation)
                End If
                Try : baseView.Name = baseName : Catch : End Try
                
                views.Add(baseView)
                logs.Add("CAMDrawingLib: Base view created: " & baseName)
            End If
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to create base view: " & ex.Message)
            Return views
        End Try
        
        ' Handle flat pattern views differently - only need edge views (4 sides), no bottom view
        If useSheetMetalFlatPattern Then
            Return AddFlatPatternProjectedViews(sheet, baseView, partDoc, app, views, dimSpace, logs)
        End If
        
        ' For regular parts: Use actual view bounds for positioning
        Dim baseX As Double = baseView.Position.X
        Dim baseY As Double = baseView.Position.Y
        Dim baseWidth As Double = baseView.Width    ' Actual view width on sheet (cm)
        Dim baseHeight As Double = baseView.Height  ' Actual view height on sheet (cm)
        
        logs.Add("CAMDrawingLib: Base view dimensions: " & (baseWidth * 10).ToString("F1") & " x " & (baseHeight * 10).ToString("F1") & " mm")
        
        ' Add Top view (above base)
        Dim topView As DrawingView = Nothing
        Try
            ' Create at estimated position, then reposition based on actual view size
            Dim topY As Double = baseY + baseHeight / 2 + dimSpace + baseWidth / 2  ' Estimate depth as similar to width
            topView = sheet.DrawingViews.AddProjectedView( _
                baseView, _
                app.TransientGeometry.CreatePoint2d(baseX, topY), _
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                Nothing)
            Try : topView.Name = "Top" : Catch : End Try
            
            ' Reposition based on actual Top view height
            Dim actualTopY As Double = baseY + baseHeight / 2 + dimSpace + topView.Height / 2
            topView.Position = app.TransientGeometry.CreatePoint2d(baseX, actualTopY)
            
            views.Add(topView)
            logs.Add("CAMDrawingLib: Top view created")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to create Top view: " & ex.Message)
        End Try
        
        ' Add Bottom view (below base)
        Dim bottomView As DrawingView = Nothing
        Try
            Dim bottomY As Double = baseY - baseHeight / 2 - dimSpace - baseWidth / 2
            bottomView = sheet.DrawingViews.AddProjectedView( _
                baseView, _
                app.TransientGeometry.CreatePoint2d(baseX, bottomY), _
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                Nothing)
            Try : bottomView.Name = "Bottom" : Catch : End Try
            
            ' Reposition based on actual Bottom view height
            Dim actualBottomY As Double = baseY - baseHeight / 2 - dimSpace - bottomView.Height / 2
            bottomView.Position = app.TransientGeometry.CreatePoint2d(baseX, actualBottomY)
            
            views.Add(bottomView)
            logs.Add("CAMDrawingLib: Bottom view created")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to create Bottom view: " & ex.Message)
        End Try
        
        ' Add Left view (left of base)
        Dim leftView As DrawingView = Nothing
        Try
            Dim leftX As Double = baseX - baseWidth / 2 - dimSpace - baseHeight / 2
            leftView = sheet.DrawingViews.AddProjectedView( _
                baseView, _
                app.TransientGeometry.CreatePoint2d(leftX, baseY), _
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                Nothing)
            Try : leftView.Name = "Left" : Catch : End Try
            
            ' Reposition based on actual Left view width
            Dim actualLeftX As Double = baseX - baseWidth / 2 - dimSpace - leftView.Width / 2
            leftView.Position = app.TransientGeometry.CreatePoint2d(actualLeftX, baseY)
            
            views.Add(leftView)
            logs.Add("CAMDrawingLib: Left view created")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to create Left view: " & ex.Message)
        End Try
        
        ' Add Right view (right of base)
        Dim rightView As DrawingView = Nothing
        Try
            Dim rightX As Double = baseX + baseWidth / 2 + dimSpace + baseHeight / 2
            rightView = sheet.DrawingViews.AddProjectedView( _
                baseView, _
                app.TransientGeometry.CreatePoint2d(rightX, baseY), _
                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                Nothing)
            Try : rightView.Name = "Right" : Catch : End Try
            
            ' Reposition based on actual Right view width
            Dim actualRightX As Double = baseX + baseWidth / 2 + dimSpace + rightView.Width / 2
            rightView.Position = app.TransientGeometry.CreatePoint2d(actualRightX, baseY)
            
            views.Add(rightView)
            logs.Add("CAMDrawingLib: Right view created")
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Failed to create Right view: " & ex.Message)
        End Try
        
        ' Add opposite view (right of Right view)
        ' This must be a separate BASE view - projected views only work for orthogonal projections
        If rightView IsNot Nothing Then
            Try
                ' Position to the right of Right view
                Dim oppositeX As Double = rightView.Position.X + rightView.Width / 2 + dimSpace + baseWidth / 2
                Dim oppositeView As DrawingView = Nothing
                Dim oppName As String = "Back"
                
                ' For arbitrary orientations, create opposite camera
                If baseOrientation = ViewOrientationTypeEnum.kArbitraryViewOrientation Then
                    Dim oppositeCamera As Camera = CreateOppositeCameraFromThicknessAxis(partDoc, app, logs)
                    
                    If oppositeCamera IsNot Nothing Then
                        Try
                            oppositeView = sheet.DrawingViews.AddBaseView( _
                                partDoc, _
                                app.TransientGeometry.CreatePoint2d(oppositeX, baseY), _
                                1.0, _
                                ViewOrientationTypeEnum.kArbitraryViewOrientation, _
                                DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                                "", _
                                oppositeCamera)
                            oppName = "Opposite"
                            logs.Add("CAMDrawingLib: Created opposite view with inverted camera")
                        Catch ex As Exception
                            logs.Add("CAMDrawingLib: Opposite camera view failed: " & ex.Message)
                        End Try
                    End If
                End If
                
                ' Fall back to standard opposite orientation
                If oppositeView Is Nothing Then
                    Dim oppositeOrientation As ViewOrientationTypeEnum = GetOppositeViewOrientation(baseOrientation)
                    
                    oppositeView = sheet.DrawingViews.AddBaseView( _
                        partDoc, _
                        app.TransientGeometry.CreatePoint2d(oppositeX, baseY), _
                        1.0, _
                        oppositeOrientation, _
                        DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, _
                        Nothing, Nothing)
                    
                    oppName = GetViewOrientationName(oppositeOrientation)
                End If
                
                Try : oppositeView.Name = oppName : Catch : End Try
                
                ' Reposition based on actual opposite view width
                Dim actualOppX As Double = rightView.Position.X + rightView.Width / 2 + dimSpace + oppositeView.Width / 2
                oppositeView.Position = app.TransientGeometry.CreatePoint2d(actualOppX, baseY)
                
                views.Add(oppositeView)
                logs.Add("CAMDrawingLib: Opposite view created: " & oppName)
            Catch ex As Exception
                logs.Add("CAMDrawingLib: Failed to create opposite view: " & ex.Message)
            End Try
        End If
        
        logs.Add("CAMDrawingLib: Total views created: " & views.Count)
        Return views
    End Function
    
    ' Reposition existing views with new spacing
    Public Sub RepositionViews(sheet As Sheet, _
                               dimSpaceMm As Double, _
                               logs As List(Of String))
        If sheet.DrawingViews.Count = 0 Then
            logs.Add("CAMDrawingLib: No views to reposition")
            Return
        End If
        
        ' Find the base view (typically first view or Front)
        Dim baseView As DrawingView = sheet.DrawingViews.Item(1)
        logs.Add("CAMDrawingLib: Using base view: " & baseView.Name)
        
        ' This is a simplified repositioning - in practice, would need to identify 
        ' which view is which orientation and reposition accordingly
        logs.Add("CAMDrawingLib: View repositioning - use AddAllViews for new drawings")
    End Sub

    ' ============================================================================
    ' Dimensions
    ' ============================================================================
    
    ' Dimension offset from view edge (in cm) - space between view boundary and dimension line
    ' Default 2.5cm (25mm) provides reasonable spacing for dimension text
    Public Const DIMENSION_OFFSET As Double = 2.5
    
    ' Get the bounds of a view including dimensions that would be added to it
    ' Returns array: (left, right, bottom, top) in internal units (cm)
    Public Function GetViewBoundsWithDimensions(view As DrawingView, _
                                                 offsetCm As Double) As Double()
        Dim viewLeft As Double = view.Position.X - view.Width / 2
        Dim viewRight As Double = view.Position.X + view.Width / 2 + offsetCm  ' Vertical dim on right
        Dim viewBottom As Double = view.Position.Y - view.Height / 2 - offsetCm  ' Horizontal dim below
        Dim viewTop As Double = view.Position.Y + view.Height / 2
        
        Return New Double() {viewLeft, viewRight, viewBottom, viewTop}
    End Function
    
    ' Add extent dimensions to a view
    ' Finds outermost geometry and creates horizontal/vertical dimensions
    ' offsetCm: offset from VIEW BOUNDS in cm (default DIMENSION_OFFSET)
    ' 
    ' The dimension LINE is placed outside the view bounds (not curve extremes)
    ' to ensure dimensions never overlap the rendered view content.
    Public Sub AddExtentDimensions(sheet As Sheet, _
                                   view As DrawingView, _
                                   app As Inventor.Application, _
                                   logs As List(Of String), _
                                   Optional offsetCm As Double = DIMENSION_OFFSET)
        ' Get view bounds (the actual rendering area)
        Dim viewLeft As Double = view.Position.X - view.Width / 2
        Dim viewRight As Double = view.Position.X + view.Width / 2
        Dim viewBottom As Double = view.Position.Y - view.Height / 2
        Dim viewTop As Double = view.Position.Y + view.Height / 2
        
        ' Find extreme points by iterating drawing curves
        ' We need these to attach dimensions to the actual geometry
        Dim minX As Double = Double.MaxValue
        Dim maxX As Double = Double.MinValue
        Dim minY As Double = Double.MaxValue
        Dim maxY As Double = Double.MinValue
        
        Dim leftCurve As DrawingCurve = Nothing
        Dim rightCurve As DrawingCurve = Nothing
        Dim topCurve As DrawingCurve = Nothing
        Dim bottomCurve As DrawingCurve = Nothing
        
        ' Track which point type is at each extreme
        Dim leftIntent As PointIntentEnum = PointIntentEnum.kStartPointIntent
        Dim rightIntent As PointIntentEnum = PointIntentEnum.kStartPointIntent
        Dim topIntent As PointIntentEnum = PointIntentEnum.kStartPointIntent
        Dim bottomIntent As PointIntentEnum = PointIntentEnum.kStartPointIntent
        
        Dim curveCount As Integer = 0
        
        For Each curve As DrawingCurve In view.DrawingCurves
            curveCount += 1
            
            ' Process start point
            Try
                Dim startPt As Point2d = curve.StartPoint
                If startPt IsNot Nothing Then
                    If startPt.X < minX Then
                        minX = startPt.X
                        leftCurve = curve
                        leftIntent = PointIntentEnum.kStartPointIntent
                    End If
                    If startPt.X > maxX Then
                        maxX = startPt.X
                        rightCurve = curve
                        rightIntent = PointIntentEnum.kStartPointIntent
                    End If
                    If startPt.Y < minY Then
                        minY = startPt.Y
                        bottomCurve = curve
                        bottomIntent = PointIntentEnum.kStartPointIntent
                    End If
                    If startPt.Y > maxY Then
                        maxY = startPt.Y
                        topCurve = curve
                        topIntent = PointIntentEnum.kStartPointIntent
                    End If
                End If
            Catch
            End Try
            
            ' Process end point
            Try
                Dim endPt As Point2d = curve.EndPoint
                If endPt IsNot Nothing Then
                    If endPt.X < minX Then
                        minX = endPt.X
                        leftCurve = curve
                        leftIntent = PointIntentEnum.kEndPointIntent
                    End If
                    If endPt.X > maxX Then
                        maxX = endPt.X
                        rightCurve = curve
                        rightIntent = PointIntentEnum.kEndPointIntent
                    End If
                    If endPt.Y < minY Then
                        minY = endPt.Y
                        bottomCurve = curve
                        bottomIntent = PointIntentEnum.kEndPointIntent
                    End If
                    If endPt.Y > maxY Then
                        maxY = endPt.Y
                        topCurve = curve
                        topIntent = PointIntentEnum.kEndPointIntent
                    End If
                End If
            Catch
            End Try
            
            ' Process mid point (important for arcs where extremes may be at midpoint)
            Try
                Dim midPt As Point2d = curve.MidPoint
                If midPt IsNot Nothing Then
                    If midPt.X < minX Then
                        minX = midPt.X
                        leftCurve = curve
                        leftIntent = PointIntentEnum.kMidPointIntent
                    End If
                    If midPt.X > maxX Then
                        maxX = midPt.X
                        rightCurve = curve
                        rightIntent = PointIntentEnum.kMidPointIntent
                    End If
                    If midPt.Y < minY Then
                        minY = midPt.Y
                        bottomCurve = curve
                        bottomIntent = PointIntentEnum.kMidPointIntent
                    End If
                    If midPt.Y > maxY Then
                        maxY = midPt.Y
                        topCurve = curve
                        topIntent = PointIntentEnum.kMidPointIntent
                    End If
                End If
            Catch
            End Try
        Next
        
        logs.Add("CAMDrawingLib: Found " & curveCount & " curves, extent: " & _
                 FormatNumber((maxX - minX) * 10, 1) & " x " & FormatNumber((maxY - minY) * 10, 1) & " mm")
        logs.Add("CAMDrawingLib: View bounds: " & FormatNumber(view.Width * 10, 1) & " x " & _
                 FormatNumber(view.Height * 10, 1) & " mm at (" & FormatNumber(view.Position.X * 10, 1) & ", " & _
                 FormatNumber(view.Position.Y * 10, 1) & ")")
        
        ' Create horizontal dimension (below view)
        ' Place dimension line OUTSIDE the view bounds, not curve bounds
        If leftCurve IsNot Nothing AndAlso rightCurve IsNot Nothing Then
            Try
                Dim leftGeomIntent As GeometryIntent = sheet.CreateGeometryIntent(leftCurve, leftIntent)
                Dim rightGeomIntent As GeometryIntent = sheet.CreateGeometryIntent(rightCurve, rightIntent)
                
                ' Position dimension line below VIEW bounds (not curve bounds)
                Dim dimY As Double = viewBottom - offsetCm
                Dim dimPoint As Point2d = app.TransientGeometry.CreatePoint2d((minX + maxX) / 2, dimY)
                
                Dim hDim As GeneralDimension = sheet.DrawingDimensions.GeneralDimensions.AddLinear( _
                    dimPoint, leftGeomIntent, rightGeomIntent, _
                    DimensionTypeEnum.kHorizontalDimensionType)
                
                If hDim IsNot Nothing Then
                    logs.Add("CAMDrawingLib: Horizontal dimension created at Y=" & FormatNumber(dimY * 10, 1) & "mm")
                End If
            Catch ex As Exception
                logs.Add("CAMDrawingLib: Horizontal dimension failed: " & ex.Message)
            End Try
        Else
            logs.Add("CAMDrawingLib: Cannot create horizontal dimension - missing curve(s)")
        End If
        
        ' Create vertical dimension (right of view)
        ' Place dimension line OUTSIDE the view bounds, not curve bounds
        If topCurve IsNot Nothing AndAlso bottomCurve IsNot Nothing Then
            Try
                Dim bottomGeomIntent As GeometryIntent = sheet.CreateGeometryIntent(bottomCurve, bottomIntent)
                Dim topGeomIntent As GeometryIntent = sheet.CreateGeometryIntent(topCurve, topIntent)
                
                ' Position dimension line to the right of VIEW bounds (not curve bounds)
                Dim dimX As Double = viewRight + offsetCm
                Dim dimPoint As Point2d = app.TransientGeometry.CreatePoint2d(dimX, (minY + maxY) / 2)
                
                Dim vDim As GeneralDimension = sheet.DrawingDimensions.GeneralDimensions.AddLinear( _
                    dimPoint, bottomGeomIntent, topGeomIntent, _
                    DimensionTypeEnum.kVerticalDimensionType)
                
                If vDim IsNot Nothing Then
                    logs.Add("CAMDrawingLib: Vertical dimension created at X=" & FormatNumber(dimX * 10, 1) & "mm")
                End If
            Catch ex As Exception
                logs.Add("CAMDrawingLib: Vertical dimension failed: " & ex.Message)
            End Try
        Else
            logs.Add("CAMDrawingLib: Cannot create vertical dimension - missing curve(s)")
        End If
    End Sub
    
    ' Add extent dimensions to all views on a sheet
    ' View spacing (dimSpace) should account for dimension offset to prevent overlapping
    Public Sub AddExtentDimensionsToAllViews(sheet As Sheet, _
                                             app As Inventor.Application, _
                                             logs As List(Of String))
        logs.Add("CAMDrawingLib: Adding dimensions to " & sheet.DrawingViews.Count & " views...")
        For i As Integer = 1 To sheet.DrawingViews.Count
            Dim view As DrawingView = sheet.DrawingViews.Item(i)
            logs.Add("CAMDrawingLib: Adding dimensions to view " & i & " (" & view.Name & ")")
            AddExtentDimensions(sheet, view, app, logs)
        Next
    End Sub
    
    ' Add extent dimensions only to the base view (first view on sheet)
    ' This is the recommended approach to avoid overlapping dimensions
    Public Sub AddExtentDimensionsToBaseView(sheet As Sheet, _
                                              app As Inventor.Application, _
                                              logs As List(Of String))
        If sheet.DrawingViews.Count = 0 Then
            logs.Add("CAMDrawingLib: No views on sheet")
            Return
        End If
        
        Dim baseView As DrawingView = sheet.DrawingViews.Item(1)
        logs.Add("CAMDrawingLib: Adding dimensions to base view: " & baseView.Name)
        AddExtentDimensions(sheet, baseView, app, logs)
    End Sub
    
    ' Add extent dimensions to specific views by index (1-based)
    Public Sub AddExtentDimensionsToViews(sheet As Sheet, _
                                          viewIndices As List(Of Integer), _
                                          app As Inventor.Application, _
                                          logs As List(Of String))
        For Each idx As Integer In viewIndices
            If idx >= 1 AndAlso idx <= sheet.DrawingViews.Count Then
                Dim view As DrawingView = sheet.DrawingViews.Item(idx)
                logs.Add("CAMDrawingLib: Adding dimensions to view " & idx & " (" & view.Name & ")")
                AddExtentDimensions(sheet, view, app, logs)
            Else
                logs.Add("CAMDrawingLib: View index " & idx & " out of range")
            End If
        Next
    End Sub
    
    ' Remove all dimensions from a sheet
    Public Sub RemoveAllDimensions(sheet As Sheet, _
                                   logs As List(Of String))
        Dim count As Integer = 0
        
        ' Remove general dimensions
        While sheet.DrawingDimensions.GeneralDimensions.Count > 0
            Try
                sheet.DrawingDimensions.GeneralDimensions.Item(1).Delete()
                count += 1
            Catch
                Exit While
            End Try
        End While
        
        logs.Add("CAMDrawingLib: Removed " & count & " dimensions")
    End Sub
    
    ' Refresh dimensions on a sheet (remove and recreate)
    Public Sub RefreshDimensions(sheet As Sheet, _
                                 app As Inventor.Application, _
                                 logs As List(Of String))
        RemoveAllDimensions(sheet, logs)
        AddExtentDimensionsToAllViews(sheet, app, logs)
    End Sub

    ' ============================================================================
    ' Export
    ' ============================================================================
    
    ' Export drawing to DWG or DXF
    ' format: "DWG" or "DXF"
    ' Note: For DXF, uses simple SaveAs method. For DWG, uses TranslatorAddIn.
    Public Sub ExportToDwgOrDxf(app As Inventor.Application, _
                                drawDoc As DrawingDocument, _
                                outputPath As String, _
                                format As String, _
                                logs As List(Of String))
        If drawDoc Is Nothing Then
            logs.Add("CAMDrawingLib: No drawing to export")
            Return
        End If
        
        ' Determine file type and extension
        Dim isDxf As Boolean = format.ToUpper() = "DXF"
        Dim ext As String = If(isDxf, ".dxf", ".dwg")
        If Not outputPath.ToLower().EndsWith(ext) Then
            outputPath = System.IO.Path.ChangeExtension(outputPath, ext)
        End If
        
        ' Ensure output directory exists
        Dim outputDir As String = System.IO.Path.GetDirectoryName(outputPath)
        If Not System.IO.Directory.Exists(outputDir) Then
            logs.Add("CAMDrawingLib: Output directory does not exist: " & outputDir)
            Return
        End If
        
        ' Try simple SaveAs method first (works for both DWG and DXF on unsaved documents)
        Try
            drawDoc.SaveAs(outputPath, True)
            If System.IO.File.Exists(outputPath) Then
                Dim fileInfo As New System.IO.FileInfo(outputPath)
                logs.Add("CAMDrawingLib: Exported " & If(isDxf, "DXF", "DWG") & " to " & outputPath & " (" & fileInfo.Length & " bytes)")
                Return
            Else
                logs.Add("CAMDrawingLib: SaveAs completed but file not found")
            End If
        Catch ex As Exception
            logs.Add("CAMDrawingLib: SaveAs failed: " & ex.Message)
            If Not isDxf Then
                logs.Add("CAMDrawingLib: Trying translator method for DWG...")
            End If
        End Try
        
        ' For DWG, if SaveAs failed, save the drawing first then use translator
        If Not isDxf Then
            ' Check if drawing needs to be saved first
            If String.IsNullOrEmpty(drawDoc.FullDocumentName) Then
                ' Drawing not saved yet - trigger save (Vault dialog will appear)
                Try
                    logs.Add("CAMDrawingLib: Drawing not saved, triggering save...")
                    drawDoc.Save()
                    
                    ' Check if save succeeded
                    If String.IsNullOrEmpty(drawDoc.FullDocumentName) Then
                        logs.Add("CAMDrawingLib: Save was cancelled or failed")
                        Return
                    End If
                    logs.Add("CAMDrawingLib: Drawing saved to: " & drawDoc.FullDocumentName)
                Catch ex As Exception
                    logs.Add("CAMDrawingLib: Failed to save drawing: " & ex.Message)
                    Return
                End Try
            End If
            
            ' Now export using translator (document is saved)
            ExportWithTranslator(app, drawDoc, outputPath, False, logs)
        End If
    End Sub
    
    ' Internal helper: Export using DWG TranslatorAddIn
    Private Sub ExportWithTranslator(app As Inventor.Application, _
                                     drawDoc As DrawingDocument, _
                                     outputPath As String, _
                                     isDxf As Boolean, _
                                     logs As List(Of String))
        ' Enable silent operation to suppress export dialogs
        Dim previousSilentOperation As Boolean = app.SilentOperation
        app.SilentOperation = True
        
        ' Get DWG translator
        Dim dwgAddin As TranslatorAddIn = Nothing
        Try
            dwgAddin = CType(app.ApplicationAddIns.ItemById(DWG_ADDIN_GUID), TranslatorAddIn)
        Catch ex As Exception
            logs.Add("CAMDrawingLib: DWG translator not found: " & ex.Message)
            app.SilentOperation = previousSilentOperation
            Return
        End Try
        
        ' Set up export context
        Dim context As TranslationContext = app.TransientObjects.CreateTranslationContext
        context.Type = IOMechanismEnum.kFileBrowseIOMechanism
        
        ' Create options
        Dim options As NameValueMap = app.TransientObjects.CreateNameValueMap
        
        ' Check if translator has options and populate defaults
        Dim hasOptions As Boolean = False
        Try
            hasOptions = dwgAddin.HasSaveCopyAsOptions(drawDoc, context, options)
            If hasOptions Then
                logs.Add("CAMDrawingLib: Translator has options available")
            End If
        Catch ex As Exception
            logs.Add("CAMDrawingLib: HasSaveCopyAsOptions failed: " & ex.Message)
        End Try
        
        ' Configure export options
        If hasOptions Then
            Try
                ' Set DWG version to 2010 (AutoCAD 2010-2012 format)
                ' Known values: 23=2000, 24=2004, 25=2007, 26=2010, 27=2013, 28=2018
                options.Value("DwgVersion") = 26
                logs.Add("CAMDrawingLib: Set DWG version to 2010")
            Catch
                Try
                    options.Add("DwgVersion", 26)
                    logs.Add("CAMDrawingLib: Added DWG version = 2010")
                Catch
                    logs.Add("CAMDrawingLib: Could not set DWG version")
                End Try
            End Try
            
            ' Set sheet range to export all sheets
            Try
                options.Value("Sheet_Range") = PrintRangeEnum.kPrintAllSheets
            Catch
                Try
                    options.Add("Sheet_Range", PrintRangeEnum.kPrintAllSheets)
                Catch
                End Try
            End Try
        End If
        
        ' Set up data medium with output path
        Dim dataMedium As DataMedium = app.TransientObjects.CreateDataMedium
        dataMedium.FileName = outputPath
        
        ' Export
        Try
            dwgAddin.SaveCopyAs(drawDoc, context, options, dataMedium)
            
            If System.IO.File.Exists(outputPath) Then
                Dim fileInfo As New System.IO.FileInfo(outputPath)
                logs.Add("CAMDrawingLib: Exported " & If(isDxf, "DXF", "DWG") & " to " & outputPath & " (" & fileInfo.Length & " bytes)")
            Else
                logs.Add("CAMDrawingLib: Export completed but file not found at " & outputPath)
            End If
        Catch ex As Exception
            logs.Add("CAMDrawingLib: Export failed: " & ex.Message)
            logs.Add("CAMDrawingLib: Debug - Output path: " & outputPath)
            logs.Add("CAMDrawingLib: Debug - DrawDoc FullName: " & drawDoc.FullDocumentName)
            logs.Add("CAMDrawingLib: Debug - HasOptions: " & hasOptions.ToString())
        Finally
            ' Restore silent operation state
            app.SilentOperation = previousSilentOperation
        End Try
    End Sub
    
    ' Export with additional options for future INI/template support
    Public Sub ExportWithOptions(app As Inventor.Application, _
                                 drawDoc As DrawingDocument, _
                                 outputPath As String, _
                                 format As String, _
                                 iniFilePath As String, _
                                 acaTemplatePath As String, _
                                 logs As List(Of String))
        ' For now, just call the basic export
        ' Future: parse INI file and apply settings
        ' Future: use AutoCAD template (.dwt)
        
        If Not String.IsNullOrEmpty(iniFilePath) Then
            logs.Add("CAMDrawingLib: INI file support not yet implemented: " & iniFilePath)
        End If
        
        If Not String.IsNullOrEmpty(acaTemplatePath) Then
            logs.Add("CAMDrawingLib: AutoCAD template support not yet implemented: " & acaTemplatePath)
        End If
        
        ExportToDwgOrDxf(app, drawDoc, outputPath, format, logs)
    End Sub

    ' ============================================================================
    ' Utility Functions
    ' ============================================================================
    
    ' Get the referenced part document from a drawing
    Public Function GetReferencedPartDocument(drawDoc As DrawingDocument, _
                                              logs As List(Of String)) As PartDocument
        If drawDoc.ReferencedDocuments.Count = 0 Then
            logs.Add("CAMDrawingLib: Drawing has no referenced documents")
            Return Nothing
        End If
        
        For Each refDoc As Document In drawDoc.ReferencedDocuments
            If refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject Then
                logs.Add("CAMDrawingLib: Found referenced part: " & refDoc.DisplayName)
                Return CType(refDoc, PartDocument)
            End If
        Next
        
        logs.Add("CAMDrawingLib: No part document found in references")
        Return Nothing
    End Function
    
    ' Calculate dimension spacing based on part size
    Public Function CalculateDimSpacing(partDoc As PartDocument) As Double
        Dim partBox As Box = partDoc.ComponentDefinition.RangeBox
        Dim maxExtent As Double = Math.Max( _
            partBox.MaxPoint.X - partBox.MinPoint.X, _
            Math.Max(partBox.MaxPoint.Y - partBox.MinPoint.Y, _
                     partBox.MaxPoint.Z - partBox.MinPoint.Z))
        
        ' 5% of max extent, minimum 30mm, maximum 100mm
        Dim spacing As Double = Math.Max(30, Math.Min(100, maxExtent * 10 * 0.05))
        Return spacing
    End Function

End Module
