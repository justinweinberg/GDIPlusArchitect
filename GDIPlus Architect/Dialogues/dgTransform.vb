
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : dgTransform
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a dialog for transforming the size of objects
''' </summary>
''' <remarks>
''' This dialog could be expanded to perform all sorts of transformations to 
''' selected objects such as skew, etc.
''' </remarks>
''' -----------------------------------------------------------------------------
Friend Class dgTransform
    Inherits System.Windows.Forms.Form

#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the Tranform dialog
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        App.ToolTipManager.PopulatePopupTip(Tip, "Transform", Me)
    End Sub


#End Region

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
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents pnScale As System.Windows.Forms.Panel
    Private WithEvents txtScaleWidth As System.Windows.Forms.TextBox
    Private WithEvents txtScaleHeight As System.Windows.Forms.TextBox
    Private WithEvents chkScaleProportions As System.Windows.Forms.CheckBox
    Private WithEvents btnCancel As System.Windows.Forms.Button
    Private WithEvents btnOk As System.Windows.Forms.Button
    Private WithEvents Label4 As System.Windows.Forms.Label
    Private WithEvents Label5 As System.Windows.Forms.Label
    Private WithEvents btnHelp As System.Windows.Forms.Button
    Private WithEvents Tip As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.pnScale = New System.Windows.Forms.Panel
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtScaleHeight = New System.Windows.Forms.TextBox
        Me.chkScaleProportions = New System.Windows.Forms.CheckBox
        Me.txtScaleWidth = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOk = New System.Windows.Forms.Button
        Me.btnHelp = New System.Windows.Forms.Button
        Me.Tip = New System.Windows.Forms.ToolTip(Me.components)
        Me.pnScale.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnScale
        '
        Me.pnScale.Controls.Add(Me.Label5)
        Me.pnScale.Controls.Add(Me.Label4)
        Me.pnScale.Controls.Add(Me.txtScaleHeight)
        Me.pnScale.Controls.Add(Me.chkScaleProportions)
        Me.pnScale.Controls.Add(Me.txtScaleWidth)
        Me.pnScale.Controls.Add(Me.Label3)
        Me.pnScale.Controls.Add(Me.Label2)
        Me.pnScale.Location = New System.Drawing.Point(24, 40)
        Me.pnScale.Name = "pnScale"
        Me.pnScale.Size = New System.Drawing.Size(312, 144)
        Me.pnScale.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(120, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 23)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "%"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(120, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 23)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "%"
        '
        'txtScaleHeight
        '
        Me.txtScaleHeight.Location = New System.Drawing.Point(72, 48)
        Me.txtScaleHeight.MaxLength = 4
        Me.txtScaleHeight.Name = "txtScaleHeight"
        Me.txtScaleHeight.Size = New System.Drawing.Size(40, 20)
        Me.txtScaleHeight.TabIndex = 4
        Me.txtScaleHeight.Text = "100"
        '
        'chkScaleProportions
        '
        Me.chkScaleProportions.Checked = True
        Me.chkScaleProportions.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkScaleProportions.Location = New System.Drawing.Point(24, 96)
        Me.chkScaleProportions.Name = "chkScaleProportions"
        Me.chkScaleProportions.Size = New System.Drawing.Size(144, 24)
        Me.chkScaleProportions.TabIndex = 3
        Me.chkScaleProportions.Text = "Lock proportions"
        '
        'txtScaleWidth
        '
        Me.txtScaleWidth.Location = New System.Drawing.Point(72, 16)
        Me.txtScaleWidth.MaxLength = 4
        Me.txtScaleWidth.Name = "txtScaleWidth"
        Me.txtScaleWidth.Size = New System.Drawing.Size(40, 20)
        Me.txtScaleWidth.TabIndex = 2
        Me.txtScaleWidth.Text = "100"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(48, 23)
        Me.Label3.TabIndex = 1
        Me.Label3.Text = "Height:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 23)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Width:"
        '
        'btnCancel
        '
        Me.btnCancel.CausesValidation = False
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(264, 208)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 13
        Me.btnCancel.Text = "Cancel"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(184, 208)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.TabIndex = 12
        Me.btnOk.Text = "Ok"
        '
        'btnHelp
        '
        Me.btnHelp.Location = New System.Drawing.Point(32, 208)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.TabIndex = 14
        Me.btnHelp.Text = "Help"
        '
        'dgTransform
        '
        Me.AcceptButton = Me.btnOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(368, 245)
        Me.Controls.Add(Me.pnScale)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnHelp)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "dgTransform"
        Me.ShowInTaskbar = False
        Me.Text = "dgTransform"
        Me.pnScale.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Helper Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Applies the current transformation to the set of selected objects
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub applytransform()
        Trace.WriteLineIf(App.TraceOn, "Transform.Apply")

        If IsNumeric(txtScaleHeight.Text) AndAlso IsNumeric(txtScaleWidth.Text) Then
            Dim fWidthTransform As Single = CSng(txtScaleWidth.Text) / 100
            Dim fHeightTransform As Single = CSng(txtScaleHeight.Text) / 100

            If fWidthTransform > 0 AndAlso fHeightTransform > 0 Then
                MDIMain.ActiveDocument.Selected.Transform(fWidthTransform, fHeightTransform)
            End If
        End If

    End Sub

#End Region

#Region "Event Handlers"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Applies the transformation and closes the dialog
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        applytransform()
        Me.DialogResult = DialogResult.OK
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles typing in the scale height box
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub txtScaleHeight_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtScaleHeight.TextChanged
        If chkScaleProportions.Checked Then
            txtScaleWidth.Text = txtScaleHeight.Text
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles typing in the scale width box
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub txtScaleWidth_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtScaleWidth.TextChanged
        If chkScaleProportions.Checked Then
            txtScaleHeight.Text = txtScaleWidth.Text
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Cancels the dialog.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invokes transformation help
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHelp.Click
        App.HelpManager.InvokeHelpContents("PopTransform.htm")
    End Sub

#End Region

   
End Class
