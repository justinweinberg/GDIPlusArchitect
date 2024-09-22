Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : dgStrokePicker
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides an interface to set stroke properties in response to a double click 
''' on the stroke color area of the toolbox.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class dgStrokePicker
    Inherits System.Windows.Forms.Form

#Region "Invoker"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Intiates stroke selection
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Sub GO()
        Dim fmStrokeProps As New dgStrokePicker

        Dim iResp As DialogResult = fmStrokeProps.ShowDialog()

    End Sub
#End Region

#Region "Local Fields"

    '''<summary>A temporary stroke</summary>
    Private _Stroke As GDIStroke = New GDIStroke(Session.Stroke)

#End Region


#Region "Constructors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a strok picker dialog
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        populate()

        'Add any initialization after the InitializeComponent() call
        App.ToolTipManager.PopulatePopupTip(Tip, "StrokeProperties", Me)
    End Sub
#End Region

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
                _Stroke.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub


#Region " Windows Form Designer generated code "





    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Private WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents pnSolidBrush As System.Windows.Forms.Panel
    Private WithEvents picPreview As System.Windows.Forms.PictureBox
    Private WithEvents Label5 As System.Windows.Forms.Label
    Private WithEvents btnOk As System.Windows.Forms.Button
    Private WithEvents btnCancel As System.Windows.Forms.Button
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents Label4 As System.Windows.Forms.Label
    Private WithEvents Label6 As System.Windows.Forms.Label
    Private WithEvents picColor As System.Windows.Forms.PictureBox
    Private WithEvents cboDash As System.Windows.Forms.ComboBox
    Private WithEvents cboStartCap As System.Windows.Forms.ComboBox
    Private WithEvents cboEndCap As System.Windows.Forms.ComboBox
    Private WithEvents txtWidth As System.Windows.Forms.TextBox
    Private WithEvents Tip As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Label1 = New System.Windows.Forms.Label
        Me.picPreview = New System.Windows.Forms.PictureBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.pnSolidBrush = New System.Windows.Forms.Panel
        Me.Label5 = New System.Windows.Forms.Label
        Me.picColor = New System.Windows.Forms.PictureBox
        Me.btnOk = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboDash = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cboStartCap = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cboEndCap = New System.Windows.Forms.ComboBox
        Me.txtWidth = New System.Windows.Forms.TextBox
        Me.Tip = New System.Windows.Forms.ToolTip(Me.components)
        Me.pnSolidBrush.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Stroke Width:"
        '
        'picPreview
        '
        Me.picPreview.BackColor = System.Drawing.Color.White
        Me.picPreview.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picPreview.Location = New System.Drawing.Point(288, 32)
        Me.picPreview.Name = "picPreview"
        Me.picPreview.Size = New System.Drawing.Size(120, 112)
        Me.picPreview.TabIndex = 3
        Me.picPreview.TabStop = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(288, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 16)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Preview"
        '
        'pnSolidBrush
        '
        Me.pnSolidBrush.Controls.Add(Me.Label5)
        Me.pnSolidBrush.Controls.Add(Me.picColor)
        Me.pnSolidBrush.Location = New System.Drawing.Point(16, 152)
        Me.pnSolidBrush.Name = "pnSolidBrush"
        Me.pnSolidBrush.Size = New System.Drawing.Size(248, 72)
        Me.pnSolidBrush.TabIndex = 6
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 23)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Color:"
        '
        'picColor
        '
        Me.picColor.Location = New System.Drawing.Point(96, 8)
        Me.picColor.Name = "picColor"
        Me.picColor.Size = New System.Drawing.Size(72, 50)
        Me.picColor.TabIndex = 9
        Me.picColor.TabStop = False
        '
        'btnOk
        '
        Me.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnOk.Location = New System.Drawing.Point(248, 240)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.TabIndex = 7
        Me.btnOk.Text = "Ok"
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.SystemColors.Control
        Me.btnCancel.CausesValidation = False
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(328, 240)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "Cancel"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 16)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Dash Style:"
        '
        'cboDash
        '
        Me.cboDash.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboDash.Location = New System.Drawing.Point(112, 40)
        Me.cboDash.Name = "cboDash"
        Me.cboDash.Size = New System.Drawing.Size(152, 21)
        Me.cboDash.TabIndex = 9
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 72)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 16)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Start Cap:"
        '
        'cboStartCap
        '
        Me.cboStartCap.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboStartCap.Location = New System.Drawing.Point(112, 72)
        Me.cboStartCap.Name = "cboStartCap"
        Me.cboStartCap.Size = New System.Drawing.Size(152, 21)
        Me.cboStartCap.TabIndex = 11
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 104)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 16)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "End Cap:"
        '
        'cboEndCap
        '
        Me.cboEndCap.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboEndCap.Location = New System.Drawing.Point(112, 104)
        Me.cboEndCap.Name = "cboEndCap"
        Me.cboEndCap.Size = New System.Drawing.Size(152, 21)
        Me.cboEndCap.TabIndex = 13
        '
        'txtWidth
        '
        Me.txtWidth.AcceptsReturn = True
        Me.txtWidth.Location = New System.Drawing.Point(112, 8)
        Me.txtWidth.Name = "txtWidth"
        Me.txtWidth.Size = New System.Drawing.Size(40, 20)
        Me.txtWidth.TabIndex = 15
        Me.txtWidth.Text = ""
        '
        'dgStrokePicker
        '
        Me.AcceptButton = Me.btnOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(416, 269)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.picPreview)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.pnSolidBrush)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cboDash)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cboStartCap)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.cboEndCap)
        Me.Controls.Add(Me.txtWidth)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "dgStrokePicker"
        Me.ShowInTaskbar = False
        Me.Text = "Stroke Properties"
        Me.pnSolidBrush.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Helper Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populates the stroke dialog with current settings from the application wide 
    ''' stroke.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub populate()
        Trace.WriteLineIf(App.TraceOn, "Strokedialog.Populate")

        txtWidth.Text = _Stroke.Width.ToString

        With cboDash
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(Drawing2D.DashStyle)))

            .EndUpdate()
            .SelectedItem = System.Enum.GetName(GetType(Drawing2D.DashStyle), _Stroke.DashStyle)

        End With

        With cboStartCap
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(Drawing2D.LineCap)))
            .EndUpdate()
            .SelectedItem = System.Enum.GetName(GetType(Drawing2D.LineCap), _Stroke.Startcap)

        End With


        With cboEndCap
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(Drawing2D.LineCap)))

            .EndUpdate()
            .SelectedItem = System.Enum.GetName(GetType(Drawing2D.LineCap), _Stroke.Endcap)

        End With

        picColor.BackColor = _Stroke.Color

    End Sub

#End Region

#Region "Event Handlers"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Accepts the stroke dialog, assigning the current stroke to
    ''' the application wide session object.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Session.Stroke = _Stroke
        Me.DialogResult = DialogResult.OK
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders a preview of the stroke to the picPreview picture box.
    ''' the application wide session object.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub picPreview_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles picPreview.Paint
        Dim g As Graphics = e.Graphics
        g.DrawLine(Me._Stroke.Pen, 20, 20, picPreview.Width - 20, picPreview.Height - 20)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to a click on the color dialog to change the current color of the stroke.
    ''' the application wide session object.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' ----------------------------------------------------------------------------- 
    Private Sub picColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picColor.Click

        Dim tempColor As Color = Utility.PickColor(_Stroke.Color)
        If Not Color.op_Equality(tempColor, Color.Empty) Then
            picColor.BackColor = tempColor
            _Stroke.Color = tempColor
            picPreview.Invalidate()
        End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to a change in the dash style
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------     
    Private Sub cboDash_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDash.SelectedValueChanged
        _Stroke.DashStyle = DirectCast(System.Enum.GetValues(GetType(Drawing2D.DashStyle)).GetValue(cboDash.SelectedIndex), Drawing2D.DashStyle)
        picPreview.Invalidate()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Verifies the width entry is numeric and greater than 0
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------     
    Private Sub txtWidth_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtWidth.Validating
        If txtWidth.Text.Length > 0 AndAlso IsNumeric(txtWidth.Text) Then
            If CSng(txtWidth.Text) > 0 Then
                e.Cancel = False
                picPreview.Invalidate()
            Else
                e.Cancel = True
                txtWidth.SelectAll()
            End If

        Else
            e.Cancel = True
            txtWidth.SelectAll()
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to a change in end cap style
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------     
    Private Sub cboEndCap_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboEndCap.SelectedValueChanged
        _Stroke.Endcap = DirectCast(System.Enum.GetValues(GetType(Drawing2D.LineCap)).GetValue(cboEndCap.SelectedIndex), Drawing2D.LineCap)
        picPreview.Invalidate()
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to a change in start cap style
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------       
    Private Sub cboStartCap_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboStartCap.SelectedValueChanged
        _Stroke.Startcap = DirectCast(System.Enum.GetValues(GetType(Drawing2D.LineCap)).GetValue(cboStartCap.SelectedIndex), Drawing2D.LineCap)
        picPreview.Invalidate()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Cancels the stroke picker
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------     
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to a change in the width of the stoke
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------     
    Private Sub txtWidth_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWidth.Validated
        _Stroke.Width = CSng(txtWidth.Text)
        picPreview.Invalidate()
    End Sub

#End Region


End Class
