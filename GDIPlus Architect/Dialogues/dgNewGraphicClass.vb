
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : dgNewGraphicClass
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a dialog for creating a new instance of a graphicsclass style 
''' GDIDocument.  This dialog is used to get the user defined 
''' width and height and set the class's back color property.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class dgNewGraphicClass
    Inherits System.Windows.Forms.Form

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a the dgNewGraphicsClass form.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        App.ToolTipManager.PopulatePopupTip(Tip, "ControlProperties", Me)
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
    Private WithEvents txtPixelsX As System.Windows.Forms.TextBox
    Private WithEvents txtPixelsY As System.Windows.Forms.TextBox
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents Label4 As System.Windows.Forms.Label
    Private WithEvents picBackColor As System.Windows.Forms.PictureBox
    Private WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents btnOk As System.Windows.Forms.Button
    Private WithEvents btnCancel As System.Windows.Forms.Button
    Private WithEvents btnHelp As System.Windows.Forms.Button
    Private WithEvents Tip As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.txtPixelsX = New System.Windows.Forms.TextBox
        Me.txtPixelsY = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnOk = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.picBackColor = New System.Windows.Forms.PictureBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.btnHelp = New System.Windows.Forms.Button
        Me.Tip = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'txtPixelsX
        '
        Me.txtPixelsX.Location = New System.Drawing.Point(96, 8)
        Me.txtPixelsX.Name = "txtPixelsX"
        Me.txtPixelsX.Size = New System.Drawing.Size(40, 20)
        Me.txtPixelsX.TabIndex = 0
        Me.txtPixelsX.Text = "300"
        '
        'txtPixelsY
        '
        Me.txtPixelsY.Location = New System.Drawing.Point(96, 32)
        Me.txtPixelsY.Name = "txtPixelsY"
        Me.txtPixelsY.Size = New System.Drawing.Size(40, 20)
        Me.txtPixelsY.TabIndex = 1
        Me.txtPixelsY.Text = "300"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Width:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Height:"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(112, 128)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.TabIndex = 4
        Me.btnOk.Text = "Ok"
        '
        'btnCancel
        '
        Me.btnCancel.CausesValidation = False
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(192, 128)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        '
        'picBackColor
        '
        Me.picBackColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picBackColor.Location = New System.Drawing.Point(96, 64)
        Me.picBackColor.Name = "picBackColor"
        Me.picBackColor.Size = New System.Drawing.Size(48, 40)
        Me.picBackColor.TabIndex = 9
        Me.picBackColor.TabStop = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 16)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Back Color:"
        '
        'btnHelp
        '
        Me.btnHelp.Location = New System.Drawing.Point(8, 128)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.TabIndex = 11
        Me.btnHelp.Text = "Help"
        '
        'dgNewGraphicClass
        '
        Me.AcceptButton = Me.btnOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(272, 157)
        Me.Controls.Add(Me.txtPixelsX)
        Me.Controls.Add(Me.txtPixelsY)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.picBackColor)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.btnHelp)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "dgNewGraphicClass"
        Me.ShowInTaskbar = False
        Me.Text = "Graphics Class Options"
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Property Accessors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the size specified by the user for the new graphics class
    ''' </summary>
    ''' <value>A Size structure containing the size to use for the GDIDocument.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ControlSize() As Size
        Get
            Return New Size(CInt(txtPixelsX.Text), CInt(txtPixelsY.Text))
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the back color specified by the user for the new graphics class
    ''' </summary>
    ''' <value>The back color to assign to the GDIDocument.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property SelectedBackColor() As Color
        Get
            Return picBackColor.BackColor
        End Get
    End Property
#End Region

#Region "Event Handlers"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Assures that the value specified in sText is specified, 
    ''' numeric, and greater than zero
    ''' </summary>
    ''' <param name="sText">The string to check</param>
    ''' <returns>A boolean indicating if the pixel entry is valid or not</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Justin]	1/11/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function validPixels(ByVal sText As String) As Boolean
        If sText.Length > 0 Then
            If IsNumeric(sText) Then
                Dim pxx As Int32 = CInt(CSng(sText))
                If pxx > 0 Then
                    sText = pxx.ToString
                    Return True
                Else
                    Return False
                End If

            Else
                Return False
            End If

        Else
            Return False
        End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Validates the x pixel text box.  
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
        Private Sub txtPixelsX_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtPixelsX.Validating
            Try
                e.Cancel = not validPixels(txtPixelsX.Text) 
            Catch ex As Exception
                e.Cancel = True
                txtPixelsX.SelectAll()
            End Try
        End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Validates the Y pixel text box. 
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub txtPixelsY_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtPixelsY.Validating
        Try

            e.Cancel = Not validPixels(txtPixelsY.Text)

        Catch ex As Exception
            e.Cancel = True
            txtPixelsY.SelectAll()
        End Try
    End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Pops up a dialog to select the back color for the GDIDocument.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub picBackColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picBackColor.Click
        Dim dgColor As New ColorDialog
        Dim dgResult As DialogResult = dgColor.ShowDialog()

        If dgResult = DialogResult.OK Then
            picBackColor.BackColor = dgColor.Color
        End If

        dgColor.Dispose()
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invokes help for dgNewGraphicsClass.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHelp.Click
        App.HelpManager.InvokeHelpContents("ClassOverview.htm")
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Accepts the dialog
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Me.DialogResult = DialogResult.OK
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Cancels the dialog.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub
#End Region


End Class
