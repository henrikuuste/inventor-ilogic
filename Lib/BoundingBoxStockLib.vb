' ============================================================================
' BoundingBoxStockLib - Shared Library for Bounding Box Stock Calculations
' 
' This module contains shared functions used by both:
' - BoundingBoxStock.vb (standalone part processing)
' - BoundingBoxStockBatch.vb (batch processing from assembly)
'
' Usage: AddVbFile "BoundingBoxStockLib.vb"
' ============================================================================

Imports Inventor
Imports System.Windows.Forms

Public Module BoundingBoxStockLib

    ' ============================================================================
    ' Main processing function - called by both standalone and batch scripts
    ' Returns: "OK" = success, "SKIP" = user skipped, "CANCEL" = user cancelled
    ' iLogicAuto: Pass iLogicVb.Automation from the calling script
    ' useEstonian: If True, show Estonian UI text
    ' ============================================================================
    Public Function ProcessPartDocument(ByVal app As Inventor.Application, ByVal partDoc As PartDocument, _
                                        ByVal formTitle As String, ByVal showSkipButton As Boolean, _
                                        ByVal iLogicAuto As Object, Optional ByVal useEstonian As Boolean = False) As String
        ' Get bounding box dimensions
        Dim xSize As Double = 0
        Dim ySize As Double = 0
        Dim zSize As Double = 0
        GetBoundingBoxSizes(partDoc, xSize, ySize, zSize)

        ' Run the UI loop
        Dim axisConfig As String = RunConfigLoop(app, partDoc, xSize, ySize, zSize, formTitle, showSkipButton, useEstonian)

        If axisConfig = "" Then
            Return "CANCEL" ' User cancelled
        ElseIf axisConfig = "SKIP" Then
            Return "SKIP" ' User skipped this part
        End If

        ' Parse the axis config (format: "T:Z|W:X|L:Y" or "T:V:x,y,z|W:V:x,y,z|L:V:x,y,z")
        Dim thicknessAxis As String = ""
        Dim widthAxis As String = ""
        Dim lengthAxis As String = ""
        ParseAxisConfig(axisConfig, thicknessAxis, widthAxis, lengthAxis)

        ' Create/update the iProperties and document rule
        CreateOrUpdateRule(partDoc, thicknessAxis, widthAxis, lengthAxis, iLogicAuto)

        Return "OK"
    End Function

    Public Sub GetBoundingBoxSizes(ByVal partDoc As PartDocument, ByRef xSize As Double, ByRef ySize As Double, ByRef zSize As Double)
        Dim rangebox As Box = partDoc.ComponentDefinition.RangeBox
        xSize = rangebox.MaxPoint.X - rangebox.MinPoint.X
        ySize = rangebox.MaxPoint.Y - rangebox.MinPoint.Y
        zSize = rangebox.MaxPoint.Z - rangebox.MinPoint.Z
    End Sub

    Public Function RunConfigLoop(ByVal app As Inventor.Application, ByVal partDoc As PartDocument, _
                                  ByVal xSize As Double, ByVal ySize As Double, ByVal zSize As Double, _
                                  ByVal formTitle As String, ByVal showSkipButton As Boolean, _
                                  Optional ByVal useEstonian As Boolean = False) As String
        ' Try to read existing axis configuration from iProperties
        Dim thicknessAxis As String = GetCustomPropertyValue(partDoc, "BB_ThicknessAxis", "")
        Dim widthAxis As String = GetCustomPropertyValue(partDoc, "BB_WidthAxis", "")
        Dim lengthAxis As String = ""
        Dim customAxisDesc As String = "" ' Description when custom axis is picked

        ' If no existing config, use defaults: Z = Thickness, then Width = longer of X/Y
        If thicknessAxis = "" Then
            thicknessAxis = "Z"
            AssignWidthLength(thicknessAxis, xSize, ySize, zSize, widthAxis, lengthAxis)
        ElseIf IsVectorFormat(thicknessAxis) Then
            ' Vector format - compute perpendicular vectors for width/length
            Dim tx As Double = 0, ty As Double = 0, tz As Double = 0
            Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
            Dim lx As Double = 0, ly As Double = 0, lz As Double = 0
            If ParseVectorComponents(thicknessAxis, tx, ty, tz) Then
                ComputePerpendicularVectors(tx, ty, tz, wx, wy, wz, lx, ly, lz)
                widthAxis = VectorToString(wx, wy, wz)
                lengthAxis = VectorToString(lx, ly, lz)
                customAxisDesc = "Custom (" & FormatVectorDesc(tx, ty, tz) & ")"
            Else
                thicknessAxis = "Z"
                AssignWidthLength(thicknessAxis, xSize, ySize, zSize, widthAxis, lengthAxis)
            End If
        Else
            ' Simple axis format - length is the remaining axis
            lengthAxis = GetRemainingAxis(thicknessAxis, widthAxis)
        End If

        Dim keepGoing As Boolean = True
        Do While keepGoing
            Dim action As String = ""
            Dim result As DialogResult = ShowConfigForm(app, partDoc, xSize, ySize, zSize, _
                                                         thicknessAxis, widthAxis, lengthAxis, _
                                                         customAxisDesc, formTitle, showSkipButton, action, useEstonian)

            If result = DialogResult.Cancel Then
                Return ""
            ElseIf result = DialogResult.OK Then
                Return "T:" & thicknessAxis & "|W:" & widthAxis & "|L:" & lengthAxis
            ElseIf result = DialogResult.No Then
                ' Skip button pressed (using No for skip)
                Return "SKIP"
            ElseIf action = "PICK_PLANE" Then
                ' Pick face/work plane for thickness direction
                Dim planeDesc As String = ""
                Dim pickedVector As String = PickPlaneForThickness(app, planeDesc, useEstonian)
                If pickedVector <> "" Then
                    thicknessAxis = pickedVector
                    customAxisDesc = planeDesc
                    ' Compute perpendicular vectors for width/length
                    Dim tx As Double = 0, ty As Double = 0, tz As Double = 0
                    Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
                    Dim lx As Double = 0, ly As Double = 0, lz As Double = 0
                    ParseVectorComponents(pickedVector, tx, ty, tz)
                    ComputePerpendicularVectors(tx, ty, tz, wx, wy, wz, lx, ly, lz)
                    widthAxis = VectorToString(wx, wy, wz)
                    lengthAxis = VectorToString(lx, ly, lz)
                End If
            ElseIf action = "PICK_THICKNESS" Then
                ' Legacy: pick edge/axis (maps to principal axis)
                Dim pickedAxis As String = PickAxisWithDescription(app, customAxisDesc)
                If pickedAxis <> "" Then
                    thicknessAxis = pickedAxis
                    AssignWidthLength(thicknessAxis, xSize, ySize, zSize, widthAxis, lengthAxis)
                End If
            ElseIf action = "SELECT_AXIS" Then
                customAxisDesc = "" ' Clear custom when selecting standard axis
                AssignWidthLength(thicknessAxis, xSize, ySize, zSize, widthAxis, lengthAxis)
            ElseIf action = "FLIP" Then
                Dim temp As String = widthAxis
                widthAxis = lengthAxis
                lengthAxis = temp
            End If
        Loop

        Return ""
    End Function

    Public Sub AssignWidthLength(ByVal thicknessAxis As String, ByVal xSize As Double, ByVal ySize As Double, ByVal zSize As Double, _
                                 ByRef widthAxis As String, ByRef lengthAxis As String)
        Dim axis1 As String = ""
        Dim axis2 As String = ""
        Dim size1 As Double = 0
        Dim size2 As Double = 0

        If thicknessAxis = "X" Then
            axis1 = "Y" : size1 = ySize
            axis2 = "Z" : size2 = zSize
        ElseIf thicknessAxis = "Y" Then
            axis1 = "X" : size1 = xSize
            axis2 = "Z" : size2 = zSize
        Else ' Z
            axis1 = "X" : size1 = xSize
            axis2 = "Y" : size2 = ySize
        End If

        If size1 >= size2 Then
            widthAxis = axis1
            lengthAxis = axis2
        Else
            widthAxis = axis2
            lengthAxis = axis1
        End If
    End Sub

    Public Function ShowConfigForm(ByVal app As Inventor.Application, ByVal partDoc As PartDocument, _
                                   ByVal xSize As Double, ByVal ySize As Double, ByVal zSize As Double, _
                                   ByRef thicknessAxis As String, ByVal widthAxis As String, ByVal lengthAxis As String, _
                                   ByVal customAxisDesc As String, ByVal formTitle As String, ByVal showSkipButton As Boolean, _
                                   ByRef action As String, Optional ByVal useEstonian As Boolean = False) As DialogResult
        action = ""
        Dim uom As UnitsOfMeasure = partDoc.UnitsOfMeasure

        ' Estonian/English UI text
        Dim txtThicknessAxisLabel As String = If(useEstonian, "Paksuse telg:", "Thickness Axis:")
        Dim txtThicknessLabel As String = If(useEstonian, "Paksus:", "Thickness:")
        Dim txtWidthLabel As String = If(useEstonian, "Laius", "Width")
        Dim txtLengthLabel As String = If(useEstonian, "Pikkus", "Length")
        Dim txtFlipButton As String = If(useEstonian, "Vaheta laius/pikkus", "Flip Width/Length")
        Dim txtSkipButton As String = If(useEstonian, "Jäta vahele", "Skip")
        Dim txtCancelButton As String = If(useEstonian, "Tühista", "Cancel")
        Dim txtCustom As String = If(useEstonian, "kohandatud", "custom")
        Dim txtAxis As String = If(useEstonian, "telg", "axis")
        Dim txtDefaultTitle As String = If(useEstonian, "Mõõdud", "Bounding Box Stock Sizes")

        ' Calculate display values - use OBB for vector format
        Dim thicknessValue As Double = 0
        Dim widthValue As Double = 0
        Dim lengthValue As Double = 0

        If IsVectorFormat(thicknessAxis) Then
            GetOrientedSizes(partDoc, thicknessAxis, widthAxis, lengthAxis, thicknessValue, widthValue, lengthValue)
        Else
            thicknessValue = GetAxisSize(thicknessAxis, xSize, ySize, zSize)
            widthValue = GetAxisSize(widthAxis, xSize, ySize, zSize)
            lengthValue = GetAxisSize(lengthAxis, xSize, ySize, zSize)
        End If

        Dim thicknessStr As String = uom.GetStringFromValue(thicknessValue, uom.LengthUnits)
        Dim widthStr As String = uom.GetStringFromValue(widthValue, uom.LengthUnits)
        Dim lengthStr As String = uom.GetStringFromValue(lengthValue, uom.LengthUnits)

        ' Determine if we have a custom pick (either from description or from vector format)
        Dim isCustomPick As Boolean = (customAxisDesc <> "") OrElse IsVectorFormat(thicknessAxis)

        ' Create form
        Dim frm As New System.Windows.Forms.Form()
        frm.Text = If(formTitle <> "", formTitle, txtDefaultTitle)
        frm.StartPosition = FormStartPosition.CenterScreen
        frm.FormBorderStyle = FormBorderStyle.FixedDialog
        frm.MaximizeBox = False
        frm.MinimizeBox = False
        frm.AutoSize = True
        frm.AutoSizeMode = AutoSizeMode.GrowAndShrink
        frm.Padding = New System.Windows.Forms.Padding(10)

        ' Main layout panel
        Dim layout As New System.Windows.Forms.TableLayoutPanel()
        layout.AutoSize = True
        layout.AutoSizeMode = AutoSizeMode.GrowAndShrink
        layout.ColumnCount = 3
        layout.RowCount = 5
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 100))
        layout.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
        layout.CellBorderStyle = TableLayoutPanelCellBorderStyle.None
        frm.Controls.Add(layout)

        ' Row 0: Thickness Axis
        Dim lblThicknessAxis As New System.Windows.Forms.Label()
        lblThicknessAxis.Text = txtThicknessAxisLabel
        lblThicknessAxis.AutoSize = True
        lblThicknessAxis.Anchor = AnchorStyles.Left
        layout.Controls.Add(lblThicknessAxis, 0, 0)

        Dim cboAxis As New System.Windows.Forms.ComboBox()
        cboAxis.Name = "cboAxis"
        cboAxis.Width = 80
        cboAxis.DropDownStyle = ComboBoxStyle.DropDownList
        cboAxis.Items.Add("X")
        cboAxis.Items.Add("Y")
        cboAxis.Items.Add("Z")
        cboAxis.Items.Add("Custom")
        cboAxis.SelectedItem = If(isCustomPick, "Custom", thicknessAxis)
        cboAxis.Anchor = AnchorStyles.Left
        layout.Controls.Add(cboAxis, 1, 0)

        Dim txtCustomAxis As New System.Windows.Forms.TextBox()
        txtCustomAxis.Name = "txtCustomAxis"
        txtCustomAxis.Width = 180
        txtCustomAxis.ReadOnly = True
        txtCustomAxis.Text = If(isCustomPick, customAxisDesc, "")
        txtCustomAxis.Anchor = AnchorStyles.Left
        layout.Controls.Add(txtCustomAxis, 2, 0)

        ' Row 1: Thickness value
        Dim lblThickness As New System.Windows.Forms.Label()
        lblThickness.Text = txtThicknessLabel
        lblThickness.AutoSize = True
        lblThickness.Anchor = AnchorStyles.Left
        layout.Controls.Add(lblThickness, 0, 1)

        Dim txtThickness As New System.Windows.Forms.TextBox()
        txtThickness.Text = thicknessStr
        txtThickness.ReadOnly = True
        txtThickness.Dock = DockStyle.Fill
        layout.Controls.Add(txtThickness, 1, 1)

        ' Row 2: Width
        Dim lblWidth As New System.Windows.Forms.Label()
        If IsVectorFormat(widthAxis) Then
            lblWidth.Text = txtWidthLabel & " (" & txtCustom & "):"
        Else
            lblWidth.Text = txtWidthLabel & " (" & widthAxis & " " & txtAxis & "):"
        End If
        lblWidth.AutoSize = True
        lblWidth.Anchor = AnchorStyles.Left
        layout.Controls.Add(lblWidth, 0, 2)

        Dim txtWidth As New System.Windows.Forms.TextBox()
        txtWidth.Text = widthStr
        txtWidth.ReadOnly = True
        txtWidth.Dock = DockStyle.Fill
        layout.Controls.Add(txtWidth, 1, 2)

        Dim btnFlip As New System.Windows.Forms.Button()
        btnFlip.Text = txtFlipButton
        btnFlip.AutoSize = True
        btnFlip.DialogResult = DialogResult.Ignore
        btnFlip.Anchor = AnchorStyles.Left
        layout.Controls.Add(btnFlip, 2, 2)

        ' Row 3: Length
        Dim lblLength As New System.Windows.Forms.Label()
        If IsVectorFormat(lengthAxis) Then
            lblLength.Text = txtLengthLabel & " (" & txtCustom & "):"
        Else
            lblLength.Text = txtLengthLabel & " (" & lengthAxis & " " & txtAxis & "):"
        End If
        lblLength.AutoSize = True
        lblLength.Anchor = AnchorStyles.Left
        layout.Controls.Add(lblLength, 0, 3)

        Dim txtLength As New System.Windows.Forms.TextBox()
        txtLength.Text = lengthStr
        txtLength.ReadOnly = True
        txtLength.Dock = DockStyle.Fill
        layout.Controls.Add(txtLength, 1, 3)

        ' Row 4: Buttons
        Dim buttonPanel As New System.Windows.Forms.FlowLayoutPanel()
        buttonPanel.AutoSize = True
        buttonPanel.FlowDirection = FlowDirection.RightToLeft
        buttonPanel.Anchor = AnchorStyles.Right
        layout.Controls.Add(buttonPanel, 0, 4)
        layout.SetColumnSpan(buttonPanel, 3)

        Dim btnCancel As New System.Windows.Forms.Button()
        btnCancel.Text = txtCancelButton
        btnCancel.Width = 80
        btnCancel.Height = 28
        btnCancel.DialogResult = DialogResult.Cancel
        frm.CancelButton = btnCancel
        buttonPanel.Controls.Add(btnCancel)

        Dim btnOK As New System.Windows.Forms.Button()
        btnOK.Text = "OK"
        btnOK.Width = 80
        btnOK.Height = 28
        btnOK.DialogResult = DialogResult.OK
        frm.AcceptButton = btnOK
        buttonPanel.Controls.Add(btnOK)

        If showSkipButton Then
            Dim btnSkip As New System.Windows.Forms.Button()
            btnSkip.Text = txtSkipButton
            btnSkip.Width = 80
            btnSkip.Height = 28
            btnSkip.DialogResult = DialogResult.No
            buttonPanel.Controls.Add(btnSkip)
        End If

        cboAxis.Tag = frm
        AddHandler cboAxis.SelectedIndexChanged, AddressOf OnAxisComboChanged

        Dim result As DialogResult = frm.ShowDialog()
        Dim selectedAxis As String = CStr(cboAxis.SelectedItem)

        If result = DialogResult.Ignore Then
            action = "FLIP"
        ElseIf result = DialogResult.Yes Then
            thicknessAxis = selectedAxis
            action = "SELECT_AXIS"
        ElseIf result = DialogResult.Abort Then
            action = "PICK_PLANE"
        ElseIf result = DialogResult.Retry Then
            action = "PICK_THICKNESS"
        End If

        Return result
    End Function

    Public Sub OnAxisComboChanged(sender As Object, e As EventArgs)
        Dim cbo As System.Windows.Forms.ComboBox = CType(sender, System.Windows.Forms.ComboBox)
        Dim frm As System.Windows.Forms.Form = CType(cbo.Tag, System.Windows.Forms.Form)
        Dim selected As String = CStr(cbo.SelectedItem)

        If selected = "Custom" Then
            ' Use Abort for PICK_PLANE action (face/work plane selection)
            frm.DialogResult = DialogResult.Abort
            frm.Close()
        ElseIf selected = "X" OrElse selected = "Y" OrElse selected = "Z" Then
            frm.DialogResult = DialogResult.Yes
            frm.Close()
        End If
    End Sub

    Public Function GetAxisSize(ByVal axis As String, ByVal xSize As Double, ByVal ySize As Double, ByVal zSize As Double) As Double
        If axis = "X" Then Return xSize
        If axis = "Y" Then Return ySize
        Return zSize
    End Function

    ' ============================================================================
    ' Calculate oriented bounding box sizes for vector-based axes
    ' ============================================================================
    Public Sub GetOrientedSizes(ByVal partDoc As PartDocument, ByVal thicknessAxis As String, ByVal widthAxis As String, ByVal lengthAxis As String, _
                                ByRef thicknessSize As Double, ByRef widthSize As Double, ByRef lengthSize As Double)
        Dim tx As Double = 0, ty As Double = 0, tz As Double = 0
        Dim wx As Double = 0, wy As Double = 0, wz As Double = 0
        Dim lx As Double = 0, ly As Double = 0, lz As Double = 0

        ' Parse thickness vector
        If Not ParseVectorComponents(thicknessAxis, tx, ty, tz) Then
            thicknessSize = 0 : widthSize = 0 : lengthSize = 0
            Return
        End If

        ' Parse or compute width/length vectors
        If IsVectorFormat(widthAxis) Then
            ParseVectorComponents(widthAxis, wx, wy, wz)
            ' Length = cross(thickness, width)
            lx = ty * wz - tz * wy
            ly = tz * wx - tx * wz
            lz = tx * wy - ty * wx
        Else
            ' Compute perpendicular vectors
            ComputePerpendicularVectors(tx, ty, tz, wx, wy, wz, lx, ly, lz)
        End If

        ' Calculate extents by projecting all vertices
        thicknessSize = GetOrientedExtent(partDoc, tx, ty, tz)
        widthSize = GetOrientedExtent(partDoc, wx, wy, wz)
        lengthSize = GetOrientedExtent(partDoc, lx, ly, lz)
    End Sub

    Public Function GetOrientedExtent(ByVal partDoc As PartDocument, ByVal dirX As Double, ByVal dirY As Double, ByVal dirZ As Double) As Double
        Dim minProj As Double = Double.MaxValue
        Dim maxProj As Double = Double.MinValue

        Try
            For Each body As SurfaceBody In partDoc.ComponentDefinition.SurfaceBodies
                For Each vertex As Vertex In body.Vertices
                    Dim pt As Point = vertex.Point
                    Dim proj As Double = pt.X * dirX + pt.Y * dirY + pt.Z * dirZ
                    If proj < minProj Then minProj = proj
                    If proj > maxProj Then maxProj = proj
                Next
            Next
        Catch
        End Try

        If minProj = Double.MaxValue Then Return 0
        Return maxProj - minProj
    End Function

    Public Function PickAxisWithDescription(ByVal app As Inventor.Application, ByRef axisDesc As String) As String
        axisDesc = ""
        
        Dim selFilter As SelectionFilterEnum = SelectionFilterEnum.kAllLinearEntities
        Dim selectedObj As Object = Nothing

        Try
            selectedObj = app.CommandManager.Pick( _
                selFilter, _
                "Select an edge or work axis to define the thickness direction:")
        Catch
            Return ""
        End Try

        If selectedObj Is Nothing Then
            Return ""
        End If

        Dim dirVector As Vector = Nothing
        Dim objName As String = ""

        If TypeOf selectedObj Is Edge Then
            Dim edge As Edge = CType(selectedObj, Edge)
            dirVector = GetEdgeDirection(app, edge)
            objName = "Edge"
        ElseIf TypeOf selectedObj Is WorkAxis Then
            Dim workAxis As WorkAxis = CType(selectedObj, WorkAxis)
            dirVector = GetWorkAxisDirection(app, workAxis)
            Try
                objName = workAxis.Name
            Catch
                objName = "Work Axis"
            End Try
        Else
            MessageBox.Show("Please select a linear edge or work axis.", "Bounding Box Stock")
            Return ""
        End If

        If dirVector Is Nothing Then
            Return ""
        End If

        Dim principalAxis As String = GetPrincipalAxis(dirVector)
        axisDesc = principalAxis & " axis (from " & objName & ")"
        
        Return principalAxis
    End Function

    ' ============================================================================
    ' Pick a face or work plane for thickness direction (returns vector format)
    ' Returns: "V:x,y,z" string or "" if cancelled
    ' ============================================================================
    Public Function PickPlaneForThickness(ByVal app As Inventor.Application, ByRef planeDesc As String, _
                                          Optional ByVal useEstonian As Boolean = False) As String
        planeDesc = ""

        ' Estonian/English text
        Dim txtPrompt As String = If(useEstonian, "Vali pind või töötasapind paksuse suuna määramiseks:", _
                                     "Select a face or work plane to define the thickness direction:")
        Dim txtTitle As String = If(useEstonian, "Mõõdud", "Bounding Box Stock")
        Dim txtFaceError As String = If(useEstonian, "Valitud pinnalt ei saanud normaali. Valige tasapinnaline pind.", _
                                        "Could not get normal from selected face. Please select a planar face.")
        Dim txtPlaneError As String = If(useEstonian, "Valitud töötasapinnalt ei saanud normaali.", _
                                         "Could not get normal from selected work plane.")
        Dim txtSelectError As String = If(useEstonian, "Valige pind või töötasapind.", _
                                          "Please select a face or work plane.")
        Dim txtFace As String = If(useEstonian, "Pind", "Face")
        Dim txtWorkPlane As String = If(useEstonian, "Töötasapind", "Work Plane")
        
        Dim selFilter As SelectionFilterEnum = SelectionFilterEnum.kAllPlanarEntities
        Dim selectedObj As Object = Nothing

        Try
            selectedObj = app.CommandManager.Pick(selFilter, txtPrompt)
        Catch
            Return ""
        End Try

        If selectedObj Is Nothing Then
            Return ""
        End If

        Dim normalX As Double = 0, normalY As Double = 0, normalZ As Double = 0
        Dim objName As String = ""

        If TypeOf selectedObj Is Face Then
            Dim face As Face = CType(selectedObj, Face)
            If Not GetFaceNormal(face, normalX, normalY, normalZ) Then
                MessageBox.Show(txtFaceError, txtTitle)
                Return ""
            End If
            objName = txtFace
        ElseIf TypeOf selectedObj Is WorkPlane Then
            Dim workPlane As WorkPlane = CType(selectedObj, WorkPlane)
            If Not GetWorkPlaneNormal(workPlane, normalX, normalY, normalZ) Then
                MessageBox.Show(txtPlaneError, txtTitle)
                Return ""
            End If
            Try
                objName = txtWorkPlane & ": " & workPlane.Name
            Catch
                objName = txtWorkPlane
            End Try
        Else
            MessageBox.Show(txtSelectError, txtTitle)
            Return ""
        End If

        planeDesc = objName & " (" & FormatVectorDesc(normalX, normalY, normalZ) & ")"
        Return VectorToString(normalX, normalY, normalZ)
    End Function

    Public Function GetFaceNormal(ByVal face As Face, ByRef nx As Double, ByRef ny As Double, ByRef nz As Double) As Boolean
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

    Public Function GetWorkPlaneNormal(ByVal workPlane As WorkPlane, ByRef nx As Double, ByRef ny As Double, ByRef nz As Double) As Boolean
        Try
            Dim plane As Plane = workPlane.Plane
            Dim normal As UnitVector = plane.Normal
            nx = normal.X
            ny = normal.Y
            nz = normal.Z
            Return True
        Catch
        End Try
        Return False
    End Function

    Public Function FormatVectorDesc(ByVal vx As Double, ByVal vy As Double, ByVal vz As Double) As String
        ' Format vector as readable description like "0.71, 0.71, 0.00"
        Return vx.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) & ", " & _
               vy.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) & ", " & _
               vz.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)
    End Function

    Public Function GetEdgeDirection(ByVal app As Inventor.Application, ByVal edge As Edge) As Vector
        Try
            Dim geom As Object = edge.Geometry
            If TypeOf geom Is Line Then
                Dim line As Line = CType(geom, Line)
                Dim dir As UnitVector = line.Direction
                Dim vec As Vector = app.TransientGeometry.CreateVector(dir.X, dir.Y, dir.Z)
                Return vec
            ElseIf TypeOf geom Is LineSegment Then
                Dim lineSeg As LineSegment = CType(geom, LineSegment)
                Dim dir As UnitVector = lineSeg.Direction
                Dim vec As Vector = app.TransientGeometry.CreateVector(dir.X, dir.Y, dir.Z)
                Return vec
            End If
        Catch
        End Try
        Return Nothing
    End Function

    Public Function GetWorkAxisDirection(ByVal app As Inventor.Application, ByVal workAxis As WorkAxis) As Vector
        Try
            Dim line As Line = workAxis.Line
            Dim dir As UnitVector = line.Direction
            Dim vec As Vector = app.TransientGeometry.CreateVector(dir.X, dir.Y, dir.Z)
            Return vec
        Catch
        End Try
        Return Nothing
    End Function

    Public Function GetPrincipalAxis(ByVal vec As Vector) As String
        Dim absX As Double = Math.Abs(vec.X)
        Dim absY As Double = Math.Abs(vec.Y)
        Dim absZ As Double = Math.Abs(vec.Z)

        If absX >= absY AndAlso absX >= absZ Then
            Return "X"
        ElseIf absY >= absX AndAlso absY >= absZ Then
            Return "Y"
        Else
            Return "Z"
        End If
    End Function

    Public Function GetRemainingAxis(ByVal axis1 As String, ByVal axis2 As String) As String
        ' Returns the third axis given two axes
        If (axis1 = "X" AndAlso axis2 = "Y") OrElse (axis1 = "Y" AndAlso axis2 = "X") Then
            Return "Z"
        ElseIf (axis1 = "X" AndAlso axis2 = "Z") OrElse (axis1 = "Z" AndAlso axis2 = "X") Then
            Return "Y"
        Else
            Return "X"
        End If
    End Function

    ' ============================================================================
    ' Vector Format Utility Functions (for OBB support)
    ' Format: "V:x,y,z" where x,y,z are unit vector components
    ' ============================================================================

    Public Function IsVectorFormat(ByVal axis As String) As Boolean
        Return axis IsNot Nothing AndAlso axis.StartsWith("V:")
    End Function

    Public Function ParseVectorComponents(ByVal axis As String, ByRef vx As Double, ByRef vy As Double, ByRef vz As Double) As Boolean
        If Not IsVectorFormat(axis) Then Return False
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

    Public Function VectorToString(ByVal vx As Double, ByVal vy As Double, ByVal vz As Double) As String
        Return "V:" & vx.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture) & "," & _
                      vy.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture) & "," & _
                      vz.ToString("0.######", System.Globalization.CultureInfo.InvariantCulture)
    End Function

    Public Sub ComputePerpendicularVectors(ByVal tx As Double, ByVal ty As Double, ByVal tz As Double, _
                                           ByRef wx As Double, ByRef wy As Double, ByRef wz As Double, _
                                           ByRef lx As Double, ByRef ly As Double, ByRef lz As Double)
        ' Find a vector perpendicular to thickness vector
        ' Use the axis most perpendicular to thickness as a reference
        Dim refX As Double = 0, refY As Double = 0, refZ As Double = 0
        If Math.Abs(tx) <= Math.Abs(ty) AndAlso Math.Abs(tx) <= Math.Abs(tz) Then
            refX = 1 : refY = 0 : refZ = 0
        ElseIf Math.Abs(ty) <= Math.Abs(tz) Then
            refX = 0 : refY = 1 : refZ = 0
        Else
            refX = 0 : refY = 0 : refZ = 1
        End If

        ' Width = cross(thickness, reference) then normalize
        wx = ty * refZ - tz * refY
        wy = tz * refX - tx * refZ
        wz = tx * refY - ty * refX
        Dim wLen As Double = Math.Sqrt(wx * wx + wy * wy + wz * wz)
        If wLen > 0.0001 Then
            wx = wx / wLen : wy = wy / wLen : wz = wz / wLen
        End If

        ' Length = cross(thickness, width) then normalize
        lx = ty * wz - tz * wy
        ly = tz * wx - tx * wz
        lz = tx * wy - ty * wx
        Dim lLen As Double = Math.Sqrt(lx * lx + ly * ly + lz * lz)
        If lLen > 0.0001 Then
            lx = lx / lLen : ly = ly / lLen : lz = lz / lLen
        End If
    End Sub

    Public Sub ParseAxisConfig(ByVal config As String, ByRef thicknessAxis As String, ByRef widthAxis As String, ByRef lengthAxis As String)
        ' Split on "|" to handle vector format (which contains commas)
        Dim parts() As String = config.Split("|"c)
        For Each part As String In parts
            ' Split only on first ":" to preserve vector format "V:x,y,z"
            Dim colonPos As Integer = part.IndexOf(":"c)
            If colonPos > 0 Then
                Dim key As String = part.Substring(0, colonPos)
                Dim value As String = part.Substring(colonPos + 1)
                If key = "T" Then thicknessAxis = value
                If key = "W" Then widthAxis = value
                If key = "L" Then lengthAxis = value
            End If
        Next
    End Sub

    Public Sub CreateOrUpdateRule(ByVal partDoc As PartDocument, ByVal thicknessAxis As String, ByVal widthAxis As String, ByVal lengthAxis As String, ByVal iLogicAuto As Object)
        ' Only store thickness and width - length is determined by the other two
        SetCustomProperty(partDoc, "BB_ThicknessAxis", thicknessAxis)
        SetCustomProperty(partDoc, "BB_WidthAxis", widthAxis)

        Dim ruleName As String = "Uuenda mõõdud"
        Dim ruleText As String = BuildRuleText()

        ' Delete old BoundingBoxStockUpdate rule if it exists (migration to new name)
        Try
            Dim oldRule As Object = iLogicAuto.GetRule(partDoc, "BoundingBoxStockUpdate")
            If oldRule IsNot Nothing Then
                iLogicAuto.DeleteRule(partDoc, "BoundingBoxStockUpdate")
            End If
        Catch
        End Try

        Dim existingRule As Object = Nothing

        Try
            existingRule = iLogicAuto.GetRule(partDoc, ruleName)
        Catch
            existingRule = Nothing
        End Try

        If existingRule IsNot Nothing Then
            existingRule.Text = ruleText
        Else
            iLogicAuto.AddRule(partDoc, ruleName, ruleText)
        End If

        iLogicAuto.RunRule(partDoc, ruleName)
    End Sub

    Public Sub SetCustomProperty(ByVal doc As Document, ByVal propName As String, ByVal propValue As String)
        Dim propSet As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
        Try
            propSet.Item(propName).Value = propValue
        Catch
            propSet.Add(propValue, propName)
        End Try
    End Sub

    Public Function GetCustomPropertyValue(ByVal doc As Document, ByVal propName As String, ByVal defaultValue As String) As String
        Try
            Dim propSet As PropertySet = doc.PropertySets.Item("Inventor User Defined Properties")
            Return CStr(propSet.Item(propName).Value)
        Catch
            Return defaultValue
        End Try
    End Function

    Public Function BuildRuleText() As String
        Dim sb As New System.Text.StringBuilder()
        
        sb.AppendLine("' Auto-generated rule: Updates Width, Length, Thickness iProperties from bounding box")
        sb.AppendLine("' Supports both simple axis (X/Y/Z) and vector format (V:x,y,z) for oriented bounding box")
        sb.AppendLine("' Override by creating parameters: WidthOverride, LengthOverride, ThicknessOverride")
        sb.AppendLine("")
        sb.AppendLine("Sub Main()")
        sb.AppendLine("    Dim partDoc As PartDocument = CType(ThisDoc.Document, PartDocument)")
        sb.AppendLine("")
        sb.AppendLine("    ' Read axis configuration from iProperties")
        sb.AppendLine("    Dim thicknessAxis As String = GetCustomProp(partDoc, ""BB_ThicknessAxis"", ""Z"")")
        sb.AppendLine("    Dim widthAxis As String = GetCustomProp(partDoc, ""BB_WidthAxis"", ""X"")")
        sb.AppendLine("")
        sb.AppendLine("    Dim thicknessVal As Double = 0")
        sb.AppendLine("    Dim widthVal As Double = 0")
        sb.AppendLine("    Dim lengthVal As Double = 0")
        sb.AppendLine("")
        sb.AppendLine("    If IsVectorFormat(thicknessAxis) Then")
        sb.AppendLine("        ' Oriented Bounding Box calculation")
        sb.AppendLine("        Dim tx As Double = 0, ty As Double = 0, tz As Double = 0")
        sb.AppendLine("        Dim wx As Double = 0, wy As Double = 0, wz As Double = 0")
        sb.AppendLine("        Dim lx As Double = 0, ly As Double = 0, lz As Double = 0")
        sb.AppendLine("")
        sb.AppendLine("        ParseVectorComponents(thicknessAxis, tx, ty, tz)")
        sb.AppendLine("")
        sb.AppendLine("        If IsVectorFormat(widthAxis) Then")
        sb.AppendLine("            ParseVectorComponents(widthAxis, wx, wy, wz)")
        sb.AppendLine("            ' Length = cross(thickness, width)")
        sb.AppendLine("            lx = ty * wz - tz * wy")
        sb.AppendLine("            ly = tz * wx - tx * wz")
        sb.AppendLine("            lz = tx * wy - ty * wx")
        sb.AppendLine("        Else")
        sb.AppendLine("            ' Compute perpendicular vectors")
        sb.AppendLine("            ComputePerpendicularVectors(tx, ty, tz, wx, wy, wz, lx, ly, lz)")
        sb.AppendLine("        End If")
        sb.AppendLine("")
        sb.AppendLine("        thicknessVal = GetOrientedExtent(partDoc, tx, ty, tz)")
        sb.AppendLine("        widthVal = GetOrientedExtent(partDoc, wx, wy, wz)")
        sb.AppendLine("        lengthVal = GetOrientedExtent(partDoc, lx, ly, lz)")
        sb.AppendLine("    Else")
        sb.AppendLine("        ' Standard axis-aligned bounding box (fast path)")
        sb.AppendLine("        Dim rangebox As Box = partDoc.ComponentDefinition.RangeBox")
        sb.AppendLine("        Dim xSize As Double = rangebox.MaxPoint.X - rangebox.MinPoint.X")
        sb.AppendLine("        Dim ySize As Double = rangebox.MaxPoint.Y - rangebox.MinPoint.Y")
        sb.AppendLine("        Dim zSize As Double = rangebox.MaxPoint.Z - rangebox.MinPoint.Z")
        sb.AppendLine("")
        sb.AppendLine("        Dim lengthAxis As String = GetRemainingAxis(thicknessAxis, widthAxis)")
        sb.AppendLine("        thicknessVal = GetAxisValue(thicknessAxis, xSize, ySize, zSize)")
        sb.AppendLine("        widthVal = GetAxisValue(widthAxis, xSize, ySize, zSize)")
        sb.AppendLine("        lengthVal = GetAxisValue(lengthAxis, xSize, ySize, zSize)")
        sb.AppendLine("    End If")
        sb.AppendLine("")
        sb.AppendLine("    ' Check for overrides")
        sb.AppendLine("    thicknessVal = GetOverrideOrValue(partDoc, ""ThicknessOverride"", thicknessVal)")
        sb.AppendLine("    widthVal = GetOverrideOrValue(partDoc, ""WidthOverride"", widthVal)")
        sb.AppendLine("    lengthVal = GetOverrideOrValue(partDoc, ""LengthOverride"", lengthVal)")
        sb.AppendLine("")
        sb.AppendLine("    ' Update iProperties (always in mm with 1 decimal place for consistency)")
        sb.AppendLine("    ' Internal values are in cm, convert to mm (* 10)")
        sb.AppendLine("    Dim thicknessStr As String = FormatDimensionMm(thicknessVal * 10)")
        sb.AppendLine("    Dim widthStr As String = FormatDimensionMm(widthVal * 10)")
        sb.AppendLine("    Dim lengthStr As String = FormatDimensionMm(lengthVal * 10)")
        sb.AppendLine("")
        sb.AppendLine("    SetCustomProp(partDoc, ""Thickness"", thicknessStr)")
        sb.AppendLine("    SetCustomProp(partDoc, ""Width"", widthStr)")
        sb.AppendLine("    SetCustomProp(partDoc, ""Length"", lengthStr)")
        sb.AppendLine("End Sub")
        sb.AppendLine("")
        sb.AppendLine("' ============================================================================")
        sb.AppendLine("' Vector format utilities")
        sb.AppendLine("' ============================================================================")
        sb.AppendLine("")
        sb.AppendLine("Function IsVectorFormat(axis As String) As Boolean")
        sb.AppendLine("    Return axis IsNot Nothing AndAlso axis.StartsWith(""V:"")")
        sb.AppendLine("End Function")
        sb.AppendLine("")
        sb.AppendLine("Sub ParseVectorComponents(axis As String, ByRef vx As Double, ByRef vy As Double, ByRef vz As Double)")
        sb.AppendLine("    If Not IsVectorFormat(axis) Then Exit Sub")
        sb.AppendLine("    Try")
        sb.AppendLine("        Dim parts() As String = axis.Substring(2).Split("",""c)")
        sb.AppendLine("        If parts.Length = 3 Then")
        sb.AppendLine("            vx = Double.Parse(parts(0), System.Globalization.CultureInfo.InvariantCulture)")
        sb.AppendLine("            vy = Double.Parse(parts(1), System.Globalization.CultureInfo.InvariantCulture)")
        sb.AppendLine("            vz = Double.Parse(parts(2), System.Globalization.CultureInfo.InvariantCulture)")
        sb.AppendLine("        End If")
        sb.AppendLine("    Catch")
        sb.AppendLine("    End Try")
        sb.AppendLine("End Sub")
        sb.AppendLine("")
        sb.AppendLine("Sub ComputePerpendicularVectors(tx As Double, ty As Double, tz As Double, _")
        sb.AppendLine("                                ByRef wx As Double, ByRef wy As Double, ByRef wz As Double, _")
        sb.AppendLine("                                ByRef lx As Double, ByRef ly As Double, ByRef lz As Double)")
        sb.AppendLine("    ' Find reference axis most perpendicular to thickness")
        sb.AppendLine("    Dim refX As Double = 0, refY As Double = 0, refZ As Double = 0")
        sb.AppendLine("    If Math.Abs(tx) <= Math.Abs(ty) AndAlso Math.Abs(tx) <= Math.Abs(tz) Then")
        sb.AppendLine("        refX = 1 : refY = 0 : refZ = 0")
        sb.AppendLine("    ElseIf Math.Abs(ty) <= Math.Abs(tz) Then")
        sb.AppendLine("        refX = 0 : refY = 1 : refZ = 0")
        sb.AppendLine("    Else")
        sb.AppendLine("        refX = 0 : refY = 0 : refZ = 1")
        sb.AppendLine("    End If")
        sb.AppendLine("")
        sb.AppendLine("    ' Width = cross(thickness, reference) normalized")
        sb.AppendLine("    wx = ty * refZ - tz * refY")
        sb.AppendLine("    wy = tz * refX - tx * refZ")
        sb.AppendLine("    wz = tx * refY - ty * refX")
        sb.AppendLine("    Dim wLen As Double = Math.Sqrt(wx * wx + wy * wy + wz * wz)")
        sb.AppendLine("    If wLen > 0.0001 Then")
        sb.AppendLine("        wx = wx / wLen : wy = wy / wLen : wz = wz / wLen")
        sb.AppendLine("    End If")
        sb.AppendLine("")
        sb.AppendLine("    ' Length = cross(thickness, width) normalized")
        sb.AppendLine("    lx = ty * wz - tz * wy")
        sb.AppendLine("    ly = tz * wx - tx * wz")
        sb.AppendLine("    lz = tx * wy - ty * wx")
        sb.AppendLine("    Dim lLen As Double = Math.Sqrt(lx * lx + ly * ly + lz * lz)")
        sb.AppendLine("    If lLen > 0.0001 Then")
        sb.AppendLine("        lx = lx / lLen : ly = ly / lLen : lz = lz / lLen")
        sb.AppendLine("    End If")
        sb.AppendLine("End Sub")
        sb.AppendLine("")
        sb.AppendLine("Function GetOrientedExtent(partDoc As PartDocument, dirX As Double, dirY As Double, dirZ As Double) As Double")
        sb.AppendLine("    Dim minProj As Double = Double.MaxValue")
        sb.AppendLine("    Dim maxProj As Double = Double.MinValue")
        sb.AppendLine("")
        sb.AppendLine("    Try")
        sb.AppendLine("        For Each body As SurfaceBody In partDoc.ComponentDefinition.SurfaceBodies")
        sb.AppendLine("            For Each vertex As Vertex In body.Vertices")
        sb.AppendLine("                Dim pt As Point = vertex.Point")
        sb.AppendLine("                Dim proj As Double = pt.X * dirX + pt.Y * dirY + pt.Z * dirZ")
        sb.AppendLine("                If proj < minProj Then minProj = proj")
        sb.AppendLine("                If proj > maxProj Then maxProj = proj")
        sb.AppendLine("            Next")
        sb.AppendLine("        Next")
        sb.AppendLine("    Catch")
        sb.AppendLine("    End Try")
        sb.AppendLine("")
        sb.AppendLine("    If minProj = Double.MaxValue Then Return 0")
        sb.AppendLine("    Return maxProj - minProj")
        sb.AppendLine("End Function")
        sb.AppendLine("")
        sb.AppendLine("' ============================================================================")
        sb.AppendLine("' Standard axis utilities (for backward compatibility)")
        sb.AppendLine("' ============================================================================")
        sb.AppendLine("")
        sb.AppendLine("Function GetAxisValue(axis As String, xSize As Double, ySize As Double, zSize As Double) As Double")
        sb.AppendLine("    If axis = ""X"" Then Return xSize")
        sb.AppendLine("    If axis = ""Y"" Then Return ySize")
        sb.AppendLine("    Return zSize")
        sb.AppendLine("End Function")
        sb.AppendLine("")
        sb.AppendLine("Function GetRemainingAxis(axis1 As String, axis2 As String) As String")
        sb.AppendLine("    If (axis1 = ""X"" AndAlso axis2 = ""Y"") OrElse (axis1 = ""Y"" AndAlso axis2 = ""X"") Then Return ""Z""")
        sb.AppendLine("    If (axis1 = ""X"" AndAlso axis2 = ""Z"") OrElse (axis1 = ""Z"" AndAlso axis2 = ""X"") Then Return ""Y""")
        sb.AppendLine("    Return ""X""")
        sb.AppendLine("End Function")
        sb.AppendLine("")
        sb.AppendLine("' ============================================================================")
        sb.AppendLine("' Property and override utilities")
        sb.AppendLine("' ============================================================================")
        sb.AppendLine("")
        sb.AppendLine("Function GetCustomProp(doc As Document, propName As String, defaultVal As String) As String")
        sb.AppendLine("    Try")
        sb.AppendLine("        Return CStr(doc.PropertySets.Item(""Inventor User Defined Properties"").Item(propName).Value)")
        sb.AppendLine("    Catch")
        sb.AppendLine("        Return defaultVal")
        sb.AppendLine("    End Try")
        sb.AppendLine("End Function")
        sb.AppendLine("")
        sb.AppendLine("Function GetOverrideOrValue(doc As PartDocument, paramName As String, calcValue As Double) As Double")
        sb.AppendLine("    Try")
        sb.AppendLine("        Dim param As Parameter = doc.ComponentDefinition.Parameters.Item(paramName)")
        sb.AppendLine("        If param IsNot Nothing Then")
        sb.AppendLine("            Return param.Value")
        sb.AppendLine("        End If")
        sb.AppendLine("    Catch")
        sb.AppendLine("    End Try")
        sb.AppendLine("    Return calcValue")
        sb.AppendLine("End Function")
        sb.AppendLine("")
        sb.AppendLine("Sub SetCustomProp(doc As Document, propName As String, propValue As String)")
        sb.AppendLine("    Dim propSet As PropertySet = doc.PropertySets.Item(""Inventor User Defined Properties"")")
        sb.AppendLine("    Try")
        sb.AppendLine("        propSet.Item(propName).Value = propValue")
        sb.AppendLine("    Catch")
        sb.AppendLine("        propSet.Add(propValue, propName)")
        sb.AppendLine("    End Try")
        sb.AppendLine("End Sub")
        sb.AppendLine("")
        sb.AppendLine("Function FormatDimensionMm(valueMm As Double) As String")
        sb.AppendLine("    Return valueMm.ToString(""0.0"", System.Globalization.CultureInfo.InvariantCulture) & "" mm""")
        sb.AppendLine("End Function")

        Return sb.ToString()
    End Function

End Module

