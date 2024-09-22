Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : StrokeWin
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A Panel for setting properties on the strokes of selected objects as well as 
''' the application wide stroke contained in the Session 
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class StrokeWin
    Inherits System.Windows.Forms.UserControl

#Region "Constructors"

    Private _CBStrokeChanged As New GDIObjects.Session.StrokeChanged(AddressOf OnstrokeChanged)



    Private Sub onStrokeChanged(ByVal s As Object, ByVal e As EventArgs)
        updatestroke(Session.Stroke)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        populate()
        updatestroke(Session.Stroke)
        'Add any initialization after the InitializeComponent() call

         GDIObjects.Session.setStrokeChangedCallBack(_CBStrokeChanged)
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populates the stroke window's controls.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub populate()
        With cboDash
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(Drawing2D.DashStyle)))
            .EndUpdate()
        End With

        With cboStartCap
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(Drawing2D.LineCap)))
            .EndUpdate()

        End With

        With cboLineJoin
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(Drawing2D.LineJoin)))
            .EndUpdate()
        End With

        With cboEndCap
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(Drawing2D.LineCap)))
            .EndUpdate()
        End With


        With cboAlignment
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(Drawing2D.PenAlignment)))
            .EndUpdate()
        End With
    End Sub

#End Region

#Region "Refresh"

 



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Refreshes the panel to display the latest stroke properties.
    ''' </summary>
    ''' <param name="stroke"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub updatestroke(ByVal stroke As GDIStroke)
        Trace.WriteLineIf(App.TraceOn, "StrokeWin.updatestroke")
        cboDash.SelectedItem = System.Enum.GetName(GetType(Drawing2D.DashStyle), stroke.DashStyle)
        cboStartCap.SelectedItem = System.Enum.GetName(GetType(Drawing2D.LineCap), stroke.Startcap)
        cboEndCap.SelectedItem = System.Enum.GetName(GetType(Drawing2D.LineCap), stroke.Endcap)
        cboLineJoin.SelectedItem = System.Enum.GetName(GetType(Drawing2D.LineJoin), stroke.LineJoin)
        cboAlignment.SelectedItem = System.Enum.GetName(GetType(Drawing2D.PenAlignment), stroke.Alignment)

        nudDashOffset.Text = CStr(stroke.DashOffset)
        nudWidth.Text = CStr(stroke.Width)

        picColor.BackColor = stroke.Color

    End Sub
#End Region

#Region " Windows Form Designer generated code "




    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Private WithEvents Label6 As System.Windows.Forms.Label
    Private WithEvents cboEndCap As System.Windows.Forms.ComboBox
    Private WithEvents Label4 As System.Windows.Forms.Label
    Private WithEvents cboStartCap As System.Windows.Forms.ComboBox
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents cboDash As System.Windows.Forms.ComboBox
    Private WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents picColor As System.Windows.Forms.PictureBox
    Private WithEvents nudWidth As System.Windows.Forms.NumericUpDown
    Private WithEvents Label5 As System.Windows.Forms.Label
    Private WithEvents cboLineJoin As System.Windows.Forms.ComboBox
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents nudDashOffset As System.Windows.Forms.NumericUpDown
    Private WithEvents cboAlignment As System.Windows.Forms.ComboBox
    Private WithEvents Label7 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label6 = New System.Windows.Forms.Label
        Me.cboEndCap = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cboStartCap = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboDash = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.picColor = New System.Windows.Forms.PictureBox
        Me.nudWidth = New System.Windows.Forms.NumericUpDown
        Me.Label5 = New System.Windows.Forms.Label
        Me.cboLineJoin = New System.Windows.Forms.ComboBox
        Me.nudDashOffset = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboAlignment = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        CType(Me.nudWidth, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudDashOffset, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(5, 111)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 16)
        Me.Label6.TabIndex = 27
        Me.Label6.Text = "End Cap"
        '
        'cboEndCap
        '
        Me.cboEndCap.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEndCap.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboEndCap.Location = New System.Drawing.Point(7, 123)
        Me.cboEndCap.Name = "cboEndCap"
        Me.cboEndCap.Size = New System.Drawing.Size(137, 22)
        Me.cboEndCap.TabIndex = 26
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(5, 76)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 16)
        Me.Label4.TabIndex = 25
        Me.Label4.Text = "Start Cap"
        '
        'cboStartCap
        '
        Me.cboStartCap.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStartCap.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboStartCap.Location = New System.Drawing.Point(7, 88)
        Me.cboStartCap.Name = "cboStartCap"
        Me.cboStartCap.Size = New System.Drawing.Size(137, 22)
        Me.cboStartCap.TabIndex = 24
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 16)
        Me.Label2.TabIndex = 23
        Me.Label2.Text = "Dash Style"
        '
        'cboDash
        '
        Me.cboDash.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDash.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboDash.Location = New System.Drawing.Point(8, 56)
        Me.cboDash.Name = "cboDash"
        Me.cboDash.Size = New System.Drawing.Size(88, 22)
        Me.cboDash.TabIndex = 22
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(88, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 16)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Width"
        '
        'picColor
        '
        Me.picColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picColor.Location = New System.Drawing.Point(10, 5)
        Me.picColor.Name = "picColor"
        Me.picColor.Size = New System.Drawing.Size(30, 30)
        Me.picColor.TabIndex = 29
        Me.picColor.TabStop = False
        '
        'nudWidth
        '
        Me.nudWidth.DecimalPlaces = 3
        Me.nudWidth.Location = New System.Drawing.Point(88, 16)
        Me.nudWidth.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.nudWidth.Minimum = New Decimal(New Integer() {1, 0, 0, 196608})
        Me.nudWidth.Name = "nudWidth"
        Me.nudWidth.Size = New System.Drawing.Size(56, 20)
        Me.nudWidth.TabIndex = 30
        Me.nudWidth.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(8, 146)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(64, 16)
        Me.Label5.TabIndex = 33
        Me.Label5.Text = "Line Join"
        '
        'cboLineJoin
        '
        Me.cboLineJoin.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLineJoin.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLineJoin.Location = New System.Drawing.Point(8, 159)
        Me.cboLineJoin.Name = "cboLineJoin"
        Me.cboLineJoin.Size = New System.Drawing.Size(137, 22)
        Me.cboLineJoin.TabIndex = 34
        '
        'nudDashOffset
        '
        Me.nudDashOffset.DecimalPlaces = 1
        Me.nudDashOffset.Location = New System.Drawing.Point(96, 56)
        Me.nudDashOffset.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.nudDashOffset.Name = "nudDashOffset"
        Me.nudDashOffset.Size = New System.Drawing.Size(48, 20)
        Me.nudDashOffset.TabIndex = 36
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(96, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 16)
        Me.Label3.TabIndex = 35
        Me.Label3.Text = "Offset"
        '
        'cboAlignment
        '
        Me.cboAlignment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAlignment.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboAlignment.Location = New System.Drawing.Point(8, 201)
        Me.cboAlignment.Name = "cboAlignment"
        Me.cboAlignment.Size = New System.Drawing.Size(137, 22)
        Me.cboAlignment.TabIndex = 38
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(8, 185)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(64, 16)
        Me.Label7.TabIndex = 37
        Me.Label7.Text = "Alignment"
        '
        'StrokeWin
        '
        Me.Controls.Add(Me.cboEndCap)
        Me.Controls.Add(Me.cboStartCap)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboDash)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.picColor)
        Me.Controls.Add(Me.nudWidth)
        Me.Controls.Add(Me.cboLineJoin)
        Me.Controls.Add(Me.nudDashOffset)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboAlignment)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Name = "StrokeWin"
        Me.Size = New System.Drawing.Size(160, 232)
        CType(Me.nudWidth, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudDashOffset, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region




    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            GDIObjects.Session.removeStrokeCallback(_CBStrokeChanged)
        End If
        MyBase.Dispose(disposing)
    End Sub

#Region "Event Handlers"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to clicks on the color picture box, launching the appropriate color picker.
    ''' Assigns the selected color to the stroke.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub picColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picColor.Click

        Dim tempColor As Color = Utility.PickColor(Session.Stroke.Color)
        If Not Color.op_Equality(tempColor, Color.Empty) Then
            picColor.BackColor = tempColor
            Session.Stroke.Color = tempColor
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the dash style of strokes
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboDash_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDash.SelectedValueChanged
        Dim dStyle As Drawing2D.DashStyle = DirectCast(System.Enum.GetValues(GetType(Drawing2D.DashStyle)).GetValue(cboDash.SelectedIndex), Drawing2D.DashStyle)
        If Not Session.Stroke.DashStyle = dStyle Then
            Session.Stroke.DashStyle = dStyle
        End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the alignment style of strokes
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboAlignment_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboAlignment.SelectedValueChanged
        Dim eAlignment As Drawing2D.PenAlignment = DirectCast(System.Enum.GetValues(GetType(Drawing2D.PenAlignment)).GetValue(cboAlignment.SelectedIndex), Drawing2D.PenAlignment)
        If Not Session.Stroke.Alignment = eAlignment Then
            Session.Stroke.Alignment = eAlignment
        End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the linjoin type
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboLineJoin_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLineJoin.SelectedValueChanged
        Dim lJoin As Drawing2D.LineJoin = DirectCast(System.Enum.GetValues(GetType(Drawing2D.LineJoin)).GetValue(cboLineJoin.SelectedIndex), Drawing2D.LineJoin)

        If Not Session.Stroke.LineJoin = lJoin Then
            Session.Stroke.LineJoin = lJoin
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the end cap of strokes
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboEndCap_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboEndCap.SelectedValueChanged
        Dim EndCap As Drawing2D.LineCap = DirectCast(System.Enum.GetValues(GetType(Drawing2D.LineCap)).GetValue(cboEndCap.SelectedIndex), Drawing2D.LineCap)

        If Not Session.Stroke.Endcap = EndCap Then
            Session.Stroke.Endcap = EndCap
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the start cap of the strokes.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboStartCap_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStartCap.SelectedValueChanged
        Dim Startcap As Drawing2D.LineCap = DirectCast(System.Enum.GetValues(GetType(Drawing2D.LineCap)).GetValue(cboStartCap.SelectedIndex), Drawing2D.LineCap)
        If Not Session.Stroke.Startcap = Startcap Then
            Session.Stroke.Startcap = Startcap
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changed in the GDI+ Architect custom nud for width changes
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub nudWidth_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nudWidth.TextChanged, nudWidth.ValueChanged
        If nudWidth.Text.Length > 0 Then
            If IsNumeric(nudWidth.Text) Then
                Try
                    Single.Parse(nudWidth.Text)
                    Dim fWidth As Single = CSng(nudWidth.Text)
                    If Not Session.Stroke.Width = fWidth Then
                        Session.Stroke.Width = fWidth
                    End If

                Catch

                End Try
            End If
        End If
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the dash offset nud
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub nudDashOffset_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nudDashOffset.TextChanged, nudDashOffset.ValueChanged
        If nudDashOffset.Text.Length > 0 Then
            Try
                Dim fDashOffset As Single = CSng(nudDashOffset.Text)
                If Not Session.Stroke.DashOffset = fDashOffset Then
                    Session.Stroke.DashOffset = fDashOffset
                End If

            Catch ex As Exception
                MsgBox("Invalid width value")
            End Try
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Verifies a valid value is contained in the width nud if it looses focus and 
    ''' a valid value is not contained in the nud.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub nudWidth_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles nudWidth.LostFocus
        If nudWidth.Text.Length = 0 OrElse Not IsNumeric(nudWidth.Text) Then
            nudWidth.Text = "1.000"
        End If
    End Sub

#End Region
End Class
