''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : dgColorPicker
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Implements a custom color picker dialog
''' </summary>
''' <remarks>Credit to Ken Getz for this code!
''' </remarks>
''' -----------------------------------------------------------------------------
Friend Class dgColorPicker
    Inherits System.Windows.Forms.Form

    Private Enum ChangeStyle
        MouseMove
        RGB
        HSV
        None
    End Enum

    Private _ChangeType As ChangeStyle = ChangeStyle.None
    Private _ColorWheel As ColorWheel

    Private _LastMousePoint As Point


    Private _RGBHandler As ColorHandler.RGB
    Private _HSVHandler As ColorHandler.HSV
    Private _Alpha As Int32
    Private _Updating As Boolean = False


    Private Sub initializeColors(ByVal initColor As Color)
        _ChangeType = ChangeStyle.RGB
        _RGBHandler = New ColorHandler.RGB(initColor.R, initColor.G, initColor.B)
        _HSVHandler = ColorHandler.RGBtoHSV(_RGBHandler)
        Alpha = initColor.A

    End Sub
    Public Sub New(ByVal initColor As Color)
        MyBase.New()


        Trace.WriteLineIf(App.TraceOn, "ColorPicker.New")

        Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        Me.SetStyle(ControlStyles.UserPaint, True)
        Me.SetStyle(ControlStyles.DoubleBuffer, True)

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        pnlSelectedColor.Visible = False
        pnlBrightness.Visible = False
        pnlColor.Visible = False

        initializeColors(initColor)


        'Add any initialization after the InitializeComponent() call
        App.ToolTipManager.PopulatePopupTip(Tip, "ColorPicker", Me)

        ' Turn on double-buffering, so the form looks better. 
        ' These properties are set in design view, as well, but they
        ' have to be set to False in order for the Paint
        ' event to be able to display their contents.
        ' Never hurts to make sure they're invisible.

        ' Calculate the coordinates of the three
        ' required regions on the form.


        ' Create the new ColorWheel class, indicating
        ' the locations of the color wheel itself, the
        ' brightness area, and the position of the selected color.
        _ColorWheel = New ColorWheel(pnlColor.Bounds, _
        pnlBrightness.Bounds, _
        pnlSelectedColor.Bounds)

        _ColorWheel.Alpha = Alpha
        ' Set the RGB and HSV values 
        ' of the NumericUpDown controls.
        SetRGB(_RGBHandler)
        SetHSV(_HSVHandler)


        AddHandler _ColorWheel.ColorChanged, AddressOf ColorWheel_ColorChanged

    End Sub

#Region " Windows Form Designer generated code "


    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Private WithEvents nudRed As System.Windows.Forms.NumericUpDown
    Private WithEvents nudGreen As System.Windows.Forms.NumericUpDown
    Private WithEvents nudBlue As System.Windows.Forms.NumericUpDown
    Private WithEvents nudHue As System.Windows.Forms.NumericUpDown
    Private WithEvents nudSaturation As System.Windows.Forms.NumericUpDown
    Private WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents Label4 As System.Windows.Forms.Label
    Private WithEvents nudBrightness As System.Windows.Forms.NumericUpDown
    Private WithEvents Label5 As System.Windows.Forms.Label
    Private WithEvents Label6 As System.Windows.Forms.Label
    Private WithEvents Label7 As System.Windows.Forms.Label
    Private WithEvents pnlColor As System.Windows.Forms.Panel
    Private WithEvents pnlBrightness As System.Windows.Forms.Panel
    Private WithEvents pnlSelectedColor As System.Windows.Forms.Panel
    Private WithEvents btnCancel As System.Windows.Forms.Button
    Private WithEvents btnOK As System.Windows.Forms.Button
    Private WithEvents trkAlpha As System.Windows.Forms.TrackBar
    Private WithEvents Label8 As System.Windows.Forms.Label
    Private WithEvents lblAlpha As System.Windows.Forms.Label
    Private WithEvents Label9 As System.Windows.Forms.Label
    Private WithEvents nudAlpha As System.Windows.Forms.NumericUpDown
    Private WithEvents Tip As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.nudRed = New System.Windows.Forms.NumericUpDown
        Me.nudGreen = New System.Windows.Forms.NumericUpDown
        Me.nudBlue = New System.Windows.Forms.NumericUpDown
        Me.nudHue = New System.Windows.Forms.NumericUpDown
        Me.nudSaturation = New System.Windows.Forms.NumericUpDown
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.nudBrightness = New System.Windows.Forms.NumericUpDown
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.pnlColor = New System.Windows.Forms.Panel
        Me.pnlBrightness = New System.Windows.Forms.Panel
        Me.pnlSelectedColor = New System.Windows.Forms.Panel
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.trkAlpha = New System.Windows.Forms.TrackBar
        Me.Label8 = New System.Windows.Forms.Label
        Me.lblAlpha = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.nudAlpha = New System.Windows.Forms.NumericUpDown
        Me.Tip = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.nudRed, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudGreen, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudBlue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudHue, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudSaturation, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudBrightness, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.trkAlpha, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudAlpha, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'nudRed
        '
        Me.nudRed.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nudRed.Location = New System.Drawing.Point(96, 104)
        Me.nudRed.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.nudRed.Name = "nudRed"
        Me.nudRed.Size = New System.Drawing.Size(48, 22)
        Me.nudRed.TabIndex = 2
        '
        'nudGreen
        '
        Me.nudGreen.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nudGreen.Location = New System.Drawing.Point(96, 128)
        Me.nudGreen.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.nudGreen.Name = "nudGreen"
        Me.nudGreen.Size = New System.Drawing.Size(48, 22)
        Me.nudGreen.TabIndex = 3
        '
        'nudBlue
        '
        Me.nudBlue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nudBlue.Location = New System.Drawing.Point(96, 152)
        Me.nudBlue.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.nudBlue.Name = "nudBlue"
        Me.nudBlue.Size = New System.Drawing.Size(48, 22)
        Me.nudBlue.TabIndex = 4
        '
        'nudHue
        '
        Me.nudHue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nudHue.Location = New System.Drawing.Point(96, 16)
        Me.nudHue.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.nudHue.Name = "nudHue"
        Me.nudHue.Size = New System.Drawing.Size(48, 22)
        Me.nudHue.TabIndex = 5
        '
        'nudSaturation
        '
        Me.nudSaturation.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nudSaturation.Location = New System.Drawing.Point(96, 40)
        Me.nudSaturation.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.nudSaturation.Name = "nudSaturation"
        Me.nudSaturation.Size = New System.Drawing.Size(48, 22)
        Me.nudSaturation.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 104)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 23)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Red:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 128)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 23)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Green:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 152)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 23)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Blue:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(272, 208)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 24)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Color:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'nudBrightness
        '
        Me.nudBrightness.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nudBrightness.Location = New System.Drawing.Point(96, 64)
        Me.nudBrightness.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.nudBrightness.Name = "nudBrightness"
        Me.nudBrightness.Size = New System.Drawing.Size(48, 22)
        Me.nudBrightness.TabIndex = 11
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(16, 16)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 23)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Hue:"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(16, 40)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 23)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Saturation:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(16, 64)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(72, 23)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "Brightness:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'pnlColor
        '
        Me.pnlColor.Location = New System.Drawing.Point(200, 16)
        Me.pnlColor.Name = "pnlColor"
        Me.pnlColor.Size = New System.Drawing.Size(176, 176)
        Me.pnlColor.TabIndex = 15
        Me.pnlColor.Visible = False
        '
        'pnlBrightness
        '
        Me.pnlBrightness.Location = New System.Drawing.Point(408, 16)
        Me.pnlBrightness.Name = "pnlBrightness"
        Me.pnlBrightness.Size = New System.Drawing.Size(24, 176)
        Me.pnlBrightness.TabIndex = 16
        Me.pnlBrightness.Visible = False
        '
        'pnlSelectedColor
        '
        Me.pnlSelectedColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlSelectedColor.Location = New System.Drawing.Point(328, 208)
        Me.pnlSelectedColor.Name = "pnlSelectedColor"
        Me.pnlSelectedColor.Size = New System.Drawing.Size(48, 40)
        Me.pnlSelectedColor.TabIndex = 17
        Me.pnlSelectedColor.Visible = False
        '
        'btnOK
        '
        Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOK.Location = New System.Drawing.Point(256, 312)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.TabIndex = 18
        Me.btnOK.Text = "Ok"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(336, 312)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 19
        Me.btnCancel.Text = "Cancel"
        '
        'trkAlpha
        '
        Me.trkAlpha.Location = New System.Drawing.Point(16, 264)
        Me.trkAlpha.Maximum = 255
        Me.trkAlpha.Name = "trkAlpha"
        Me.trkAlpha.Size = New System.Drawing.Size(400, 42)
        Me.trkAlpha.TabIndex = 22
        Me.trkAlpha.TickFrequency = 5
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(16, 240)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(48, 23)
        Me.Label8.TabIndex = 23
        Me.Label8.Text = "Alpha:"
        '
        'lblAlpha
        '
        Me.lblAlpha.Location = New System.Drawing.Point(72, 240)
        Me.lblAlpha.Name = "lblAlpha"
        Me.lblAlpha.Size = New System.Drawing.Size(64, 23)
        Me.lblAlpha.TabIndex = 24
        Me.lblAlpha.Text = "lblAlpha"
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(16, 176)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(48, 23)
        Me.Label9.TabIndex = 26
        Me.Label9.Text = "Alpha:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'nudAlpha
        '
        Me.nudAlpha.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.nudAlpha.Location = New System.Drawing.Point(96, 176)
        Me.nudAlpha.Maximum = New Decimal(New Integer() {255, 0, 0, 0})
        Me.nudAlpha.Name = "nudAlpha"
        Me.nudAlpha.Size = New System.Drawing.Size(48, 22)
        Me.nudAlpha.TabIndex = 27
        '
        'dgColorPicker
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 16)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(442, 351)
        Me.Controls.Add(Me.nudRed)
        Me.Controls.Add(Me.nudGreen)
        Me.Controls.Add(Me.nudBlue)
        Me.Controls.Add(Me.nudHue)
        Me.Controls.Add(Me.nudSaturation)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.nudBrightness)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.pnlColor)
        Me.Controls.Add(Me.pnlBrightness)
        Me.Controls.Add(Me.pnlSelectedColor)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.trkAlpha)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.lblAlpha)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.nudAlpha)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "dgColorPicker"
        Me.Text = "Select Color"
        CType(Me.nudRed, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudGreen, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudBlue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudHue, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudSaturation, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudBrightness, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.trkAlpha, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudAlpha, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region




    Private Sub HandleMouse( _
     ByVal sender As Object, ByVal e As MouseEventArgs) _
     Handles MyBase.MouseMove, MyBase.MouseDown

        ' If you have the left mouse button down, 
        ' then update the selectedPoint value and 
        ' force a repaint of the color wheel.
        If e.Button = MouseButtons.Left Then
            _ChangeType = ChangeStyle.MouseMove
            _LastMousePoint = New Point(e.X, e.Y)
            Me.Invalidate()
        End If
    End Sub

    Private Sub frmMain_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseUp
        _ColorWheel.SetMouseUp()
        _ChangeType = ChangeStyle.None
    End Sub

    Private Sub HandleRGBChange( _
     ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles nudRed.ValueChanged, nudBlue.ValueChanged, nudGreen.ValueChanged

        ' If the R, G, or B values change, use this 
        ' code to update the HSV values and invalidate
        ' the color wheel (so it updates the pointers).
        ' Check the isInUpdate flag to avoid recursive events
        ' when you update the NumericUpdownControls.
        If Not _Updating Then
            _ChangeType = ChangeStyle.RGB
            _RGBHandler = New ColorHandler.RGB(CInt(nudRed.Value), _
             CInt(nudGreen.Value), CInt(nudBlue.Value))
            SetHSV(ColorHandler.RGBtoHSV(_RGBHandler))
            Me.Invalidate()
        End If
    End Sub

    Private Sub HandleHSVChange( _
     ByVal sender As System.Object, ByVal e As System.EventArgs) _
     Handles nudHue.ValueChanged, nudSaturation.ValueChanged, _
     nudBrightness.ValueChanged
        ' If the H, S, or V values change, use this 
        ' code to update the RGB values and invalidate
        ' the color wheel (so it updates the pointers).
        ' Check the isInUpdate flag to avoid recursive events
        ' when you update the NumericUpdownControls.
        If Not _Updating Then
            _ChangeType = ChangeStyle.HSV
            _HSVHandler = New ColorHandler.HSV(CInt(nudHue.Value), _
             CInt(nudSaturation.Value), CInt(nudBrightness.Value))
            SetRGB(ColorHandler.HSVtoRGB(_HSVHandler))
            Me.Invalidate()
        End If
    End Sub

    Private Sub SetRGB(ByVal RGB As ColorHandler.RGB)
        ' Update the RGB values on the form, but don't trigger
        ' the ValueChanged event of the form. The isInUpdate
        ' variable ensures that the event procedures
        ' exit without doing anything.
        _Updating = True
        RefreshNUD(nudRed, RGB.Red)
        RefreshNUD(nudBlue, RGB.Blue)
        RefreshNUD(nudGreen, RGB.Green)
        _Updating = False
    End Sub

    Private Sub SetHSV(ByVal HSV As ColorHandler.HSV)
        ' Update the HSV values on the form, but don't trigger
        ' the ValueChanged event of the form. The isInUpdate
        ' variable ensures that the event procedures
        ' exit without doing anything.
        _Updating = True
        RefreshNUD(nudHue, HSV.Hue)
        RefreshNUD(nudSaturation, HSV.Saturation)
        RefreshNUD(nudBrightness, HSV.Value)
        _Updating = False
    End Sub

    Private Sub HandleTextChanged(ByVal sender As Object, ByVal e As System.EventArgs) _
     Handles nudRed.TextChanged, nudBlue.TextChanged, nudGreen.TextChanged, _
     nudHue.TextChanged, nudSaturation.TextChanged, nudBrightness.TextChanged
        ' This step works around a bug -- unless you actively
        ' retrieve the Value, the min and max settings for the 
        ' control aren't honored when you type text. This may
        ' be fixed in the 1.1 version, but in VS.NET 1.0, this 
        ' step is required.
        Dim nud As NumericUpDown = DirectCast(sender, NumericUpDown)

        If nud.Text.Length > 0 AndAlso IsNumeric(nud.Text) Then
            Dim x As Decimal = DirectCast(sender, NumericUpDown).Value
        End If

    End Sub

    Private Sub RefreshNUD(ByVal nud As NumericUpDown, ByVal newValue As Integer)
        ' Update the value of the NumericUpDown control, 
        ' if the value is different than the current value.
        ' Refresh the control, causing an immediate repaint 
        If Not nud.Value = newValue Then
            nud.Value = newValue
            nud.Refresh()
        End If
    End Sub

    Public ReadOnly Property Color() As Color
        ' Get or set the color to be
        ' displayed in the color wheel.
        Get
            Return _ColorWheel.Color
        End Get

        'Set(ByVal Value As Color)
        '    ' Indicate the color change type. Either RGB or HSV
        '    ' will cause the color wheel to update the position
        '    ' of the pointer.
        '    _ChangeType = ChangeStyle.RGB
        '    _RGBHandler = New ColorHandler.RGB(Value.R, Value.G, Value.B)
        '    _HSVHandler = ColorHandler.RGBtoHSV(_RGBHandler)
        '    Alpha = Value.A

        'End Set
    End Property

    Private Sub ColorWheel_ColorChanged( _
     ByVal sender As Object, ByVal e As ColorChangedEventArgs)
        SetRGB(e.RGB)
        SetHSV(e.HSV)
    End Sub

    Private Sub ColorChooser1_Paint(ByVal sender As Object, ByVal e As PaintEventArgs) _
    Handles MyBase.Paint
        ' Depending on the circumstances, force a repaint
        ' of the color wheel passing different information.
        Select Case _ChangeType
            Case ChangeStyle.HSV
                _ColorWheel.Draw(e.Graphics, _HSVHandler)
            Case ChangeStyle.MouseMove, ChangeStyle.None
                _ColorWheel.Draw(e.Graphics, _LastMousePoint)
            Case ChangeStyle.RGB
                _ColorWheel.Draw(e.Graphics, _RGBHandler)
        End Select
    End Sub

    Private Property Alpha() As Int32
        Get
            Return _Alpha
        End Get
        Set(ByVal Value As Int32)
            _Alpha = Value
            If Not _ColorWheel Is Nothing Then
                _ColorWheel.Alpha = Value
            End If
            lblAlpha.Text = FormatPercent(_Alpha / 255, 2)
            nudAlpha.Value = _Alpha
            trkAlpha.Value = _Alpha

        End Set
    End Property

    Private Sub trkAlpha_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles trkAlpha.Scroll
        If _Alpha <> trkAlpha.Value Then
            Me.Alpha = trkAlpha.Value
            Me.Invalidate()
        End If
    End Sub

    Private Sub nudAlpha_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles nudAlpha.ValueChanged, nudAlpha.TextChanged
        If nudAlpha.Text.Length > 0 AndAlso IsNumeric(nudAlpha.Value) AndAlso _Alpha <> nudAlpha.Value Then
            Alpha = CInt(nudAlpha.Value)
            Me.Invalidate()
        End If
    End Sub
End Class
