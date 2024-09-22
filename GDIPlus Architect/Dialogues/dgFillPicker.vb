Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : dgFillPicker
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' dialog for when the user double clicks on the fill box (as opposed to uses the 
''' fill panel.)
''' </summary>
''' <remarks>The first thing the dialog has to establish is what the current fill is.
''' It then has the job of letting the user select from various types of 
''' fills and setting properties on them.
''' </remarks>
''' -----------------------------------------------------------------------------
Friend Class dgFillPicker
    Inherits System.Windows.Forms.Form


#Region "Invoker"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Initiates the  Fill Picker process
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Sub GO()
        Dim fp As New dgFillPicker
        fp.ShowDialog()
    End Sub

#End Region


#Region "Local Fields"
    '''<summary>Temporary solid fill</summary>
    Private _SolidFill As GDISolidFill = New GDISolidFill(Color.Black)
    '''<summary>Temporary hatch fill</summary>
    Private _hatchFill As GDIHatchFill = New GDIHatchFill(Color.White, Color.Black, Drawing2D.HatchStyle.BackwardDiagonal)
    '''<summary>Temporary gradient fill</summary>
    Private _GradientFill As GDIGradientFill = New GDIGradientFill(Color.Black, Color.White, GDIGradientFill.EnumGradientMode.ForwardDiagonal)
    '''<summary>Temporary texture fill</summary>
    Private _TextureFill As GDITexturedFill = New GDITexturedFill(App.Options.GetFirstTexture())

    '''<summary>Index of the currently selected fill type in the drop down of types.</summary>
    Private _SelectedFillIndex As Int32 = -1

#End Region


#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the fill properties dialog.  Populate the dialog 
    ''' with the current fill.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        mapCurrentFillToLocal()

        populateFillDropDown()

        App.ToolTipManager.PopulatePopupTip(Tip, "FillProperties", Me)

        'Add any initialization after the InitializeComponent() call

    End Sub


    Private Sub populateFillDropDown()

        With cboFills
            .BeginUpdate()
            .Items.Clear()
            .Items.Add(_SolidFill)
            .Items.Add(_hatchFill)
            .Items.Add(_GradientFill)
            .Items.Add(_TextureFill)
            .EndUpdate()
            .SelectedIndex = _SelectedFillIndex
        End With

        toggleFillPanels()

    End Sub

#End Region



#Region " Windows Form Designer generated code "



    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
                ' _GradientFill.Dispose()
                _hatchFill.Dispose()
                _SolidFill.Dispose()
                _GradientFill.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Private WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents Label4 As System.Windows.Forms.Label
    Private WithEvents cboGradient As System.Windows.Forms.ComboBox
    Private WithEvents cboFills As System.Windows.Forms.ComboBox
    Private WithEvents pnGradient As System.Windows.Forms.Panel
    Private WithEvents pnHatch As System.Windows.Forms.Panel
    Private WithEvents pnSolidBrush As System.Windows.Forms.Panel
    Private WithEvents picSolid As System.Windows.Forms.PictureBox
    Private WithEvents picHatchBack As System.Windows.Forms.PictureBox
    Private WithEvents picHatchFore As System.Windows.Forms.PictureBox
    Private WithEvents picPreview As System.Windows.Forms.PictureBox
    Private WithEvents picGrad2 As System.Windows.Forms.PictureBox
    Private WithEvents picGrad1 As System.Windows.Forms.PictureBox
    Private WithEvents cboHatchStyle As System.Windows.Forms.ComboBox
    Private WithEvents Label5 As System.Windows.Forms.Label
    Private WithEvents btnOk As System.Windows.Forms.Button
    Private WithEvents btnCancel As System.Windows.Forms.Button
    Private WithEvents pnTexture As System.Windows.Forms.Panel
    Private WithEvents Label7 As System.Windows.Forms.Label
    Private WithEvents lblTextures As System.Windows.Forms.Label
    Private WithEvents cboLibrary As System.Windows.Forms.ComboBox
    Private WithEvents txtTexture As System.Windows.Forms.TextBox
    Private WithEvents btnTextureFile As System.Windows.Forms.Button
    Private WithEvents pnGradCustom As System.Windows.Forms.Panel
    Private WithEvents Label6 As System.Windows.Forms.Label
    Private WithEvents nudAngle As System.Windows.Forms.NumericUpDown
    Private WithEvents Tip As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.cboFills = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.pnHatch = New System.Windows.Forms.Panel
        Me.picHatchBack = New System.Windows.Forms.PictureBox
        Me.picHatchFore = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboHatchStyle = New System.Windows.Forms.ComboBox
        Me.picPreview = New System.Windows.Forms.PictureBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.pnGradient = New System.Windows.Forms.Panel
        Me.pnGradCustom = New System.Windows.Forms.Panel
        Me.nudAngle = New System.Windows.Forms.NumericUpDown
        Me.Label6 = New System.Windows.Forms.Label
        Me.picGrad2 = New System.Windows.Forms.PictureBox
        Me.picGrad1 = New System.Windows.Forms.PictureBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.cboGradient = New System.Windows.Forms.ComboBox
        Me.pnSolidBrush = New System.Windows.Forms.Panel
        Me.Label5 = New System.Windows.Forms.Label
        Me.picSolid = New System.Windows.Forms.PictureBox
        Me.btnOk = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.pnTexture = New System.Windows.Forms.Panel
        Me.btnTextureFile = New System.Windows.Forms.Button
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtTexture = New System.Windows.Forms.TextBox
        Me.lblTextures = New System.Windows.Forms.Label
        Me.cboLibrary = New System.Windows.Forms.ComboBox
        Me.Tip = New System.Windows.Forms.ToolTip(Me.components)
        Me.pnHatch.SuspendLayout()
        Me.pnGradient.SuspendLayout()
        Me.pnGradCustom.SuspendLayout()
        CType(Me.nudAngle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnSolidBrush.SuspendLayout()
        Me.pnTexture.SuspendLayout()
        Me.SuspendLayout()
        '
        'cboFills
        '
        Me.cboFills.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFills.Location = New System.Drawing.Point(112, 8)
        Me.cboFills.Name = "cboFills"
        Me.cboFills.Size = New System.Drawing.Size(121, 21)
        Me.cboFills.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Fill Type:"
        '
        'pnHatch
        '
        Me.pnHatch.Controls.Add(Me.picHatchBack)
        Me.pnHatch.Controls.Add(Me.picHatchFore)
        Me.pnHatch.Controls.Add(Me.Label2)
        Me.pnHatch.Controls.Add(Me.cboHatchStyle)
        Me.pnHatch.Location = New System.Drawing.Point(16, 32)
        Me.pnHatch.Name = "pnHatch"
        Me.pnHatch.Size = New System.Drawing.Size(264, 208)
        Me.pnHatch.TabIndex = 2
        '
        'picHatchBack
        '
        Me.picHatchBack.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picHatchBack.Location = New System.Drawing.Point(168, 48)
        Me.picHatchBack.Name = "picHatchBack"
        Me.picHatchBack.Size = New System.Drawing.Size(72, 50)
        Me.picHatchBack.TabIndex = 10
        Me.picHatchBack.TabStop = False
        '
        'picHatchFore
        '
        Me.picHatchFore.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picHatchFore.Location = New System.Drawing.Point(16, 48)
        Me.picHatchFore.Name = "picHatchFore"
        Me.picHatchFore.Size = New System.Drawing.Size(72, 50)
        Me.picHatchFore.TabIndex = 9
        Me.picHatchFore.TabStop = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 16)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Hatch Pattern:"
        '
        'cboHatchStyle
        '
        Me.cboHatchStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboHatchStyle.Location = New System.Drawing.Point(96, 8)
        Me.cboHatchStyle.Name = "cboHatchStyle"
        Me.cboHatchStyle.Size = New System.Drawing.Size(144, 21)
        Me.cboHatchStyle.TabIndex = 4
        '
        'picPreview
        '
        Me.picPreview.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picPreview.Location = New System.Drawing.Point(312, 40)
        Me.picPreview.Name = "picPreview"
        Me.picPreview.Size = New System.Drawing.Size(75, 75)
        Me.picPreview.TabIndex = 3
        Me.picPreview.TabStop = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(312, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 16)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Preview"
        '
        'pnGradient
        '
        Me.pnGradient.Controls.Add(Me.pnGradCustom)
        Me.pnGradient.Controls.Add(Me.picGrad2)
        Me.pnGradient.Controls.Add(Me.picGrad1)
        Me.pnGradient.Controls.Add(Me.Label4)
        Me.pnGradient.Controls.Add(Me.cboGradient)
        Me.pnGradient.Location = New System.Drawing.Point(16, 32)
        Me.pnGradient.Name = "pnGradient"
        Me.pnGradient.Size = New System.Drawing.Size(264, 208)
        Me.pnGradient.TabIndex = 5
        '
        'pnGradCustom
        '
        Me.pnGradCustom.Controls.Add(Me.nudAngle)
        Me.pnGradCustom.Controls.Add(Me.Label6)
        Me.pnGradCustom.Location = New System.Drawing.Point(8, 120)
        Me.pnGradCustom.Name = "pnGradCustom"
        Me.pnGradCustom.Size = New System.Drawing.Size(144, 48)
        Me.pnGradCustom.TabIndex = 11
        Me.pnGradCustom.Visible = False
        '
        'nudAngle
        '
        Me.nudAngle.Location = New System.Drawing.Point(88, 8)
        Me.nudAngle.Maximum = New Decimal(New Integer() {360, 0, 0, 0})
        Me.nudAngle.Name = "nudAngle"
        Me.nudAngle.Size = New System.Drawing.Size(48, 20)
        Me.nudAngle.TabIndex = 13
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(40, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(40, 23)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Angle:"
        '
        'picGrad2
        '
        Me.picGrad2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picGrad2.Location = New System.Drawing.Point(168, 48)
        Me.picGrad2.Name = "picGrad2"
        Me.picGrad2.Size = New System.Drawing.Size(72, 50)
        Me.picGrad2.TabIndex = 8
        Me.picGrad2.TabStop = False
        '
        'picGrad1
        '
        Me.picGrad1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picGrad1.Location = New System.Drawing.Point(8, 48)
        Me.picGrad1.Name = "picGrad1"
        Me.picGrad1.Size = New System.Drawing.Size(72, 50)
        Me.picGrad1.TabIndex = 7
        Me.picGrad1.TabStop = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 16)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Gradient Type:"
        '
        'cboGradient
        '
        Me.cboGradient.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGradient.Location = New System.Drawing.Point(120, 16)
        Me.cboGradient.Name = "cboGradient"
        Me.cboGradient.Size = New System.Drawing.Size(121, 21)
        Me.cboGradient.TabIndex = 4
        '
        'pnSolidBrush
        '
        Me.pnSolidBrush.Controls.Add(Me.Label5)
        Me.pnSolidBrush.Controls.Add(Me.picSolid)
        Me.pnSolidBrush.Location = New System.Drawing.Point(16, 32)
        Me.pnSolidBrush.Name = "pnSolidBrush"
        Me.pnSolidBrush.Size = New System.Drawing.Size(264, 208)
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
        'picSolid
        '
        Me.picSolid.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picSolid.Location = New System.Drawing.Point(96, 8)
        Me.picSolid.Name = "picSolid"
        Me.picSolid.Size = New System.Drawing.Size(72, 50)
        Me.picSolid.TabIndex = 9
        Me.picSolid.TabStop = False
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(240, 256)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.TabIndex = 7
        Me.btnOk.Text = "Ok"
        '
        'btnCancel
        '
        Me.btnCancel.CausesValidation = False
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(320, 256)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 8
        Me.btnCancel.Text = "Cancel"
        '
        'pnTexture
        '
        Me.pnTexture.Controls.Add(Me.btnTextureFile)
        Me.pnTexture.Controls.Add(Me.Label7)
        Me.pnTexture.Controls.Add(Me.txtTexture)
        Me.pnTexture.Controls.Add(Me.lblTextures)
        Me.pnTexture.Controls.Add(Me.cboLibrary)
        Me.pnTexture.Location = New System.Drawing.Point(16, 32)
        Me.pnTexture.Name = "pnTexture"
        Me.pnTexture.Size = New System.Drawing.Size(264, 208)
        Me.pnTexture.TabIndex = 10
        '
        'btnTextureFile
        '
        Me.btnTextureFile.Location = New System.Drawing.Point(224, 56)
        Me.btnTextureFile.Name = "btnTextureFile"
        Me.btnTextureFile.Size = New System.Drawing.Size(24, 23)
        Me.btnTextureFile.TabIndex = 8
        Me.btnTextureFile.Text = "..."
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 40)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(56, 16)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "Source:"
        '
        'txtTexture
        '
        Me.txtTexture.Location = New System.Drawing.Point(8, 56)
        Me.txtTexture.Name = "txtTexture"
        Me.txtTexture.Size = New System.Drawing.Size(208, 20)
        Me.txtTexture.TabIndex = 6
        Me.txtTexture.Text = ""
        '
        'lblTextures
        '
        Me.lblTextures.Location = New System.Drawing.Point(8, 8)
        Me.lblTextures.Name = "lblTextures"
        Me.lblTextures.Size = New System.Drawing.Size(56, 16)
        Me.lblTextures.TabIndex = 5
        Me.lblTextures.Text = "Library:"
        '
        'cboLibrary
        '
        Me.cboLibrary.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLibrary.Location = New System.Drawing.Point(80, 8)
        Me.cboLibrary.Name = "cboLibrary"
        Me.cboLibrary.Size = New System.Drawing.Size(168, 21)
        Me.cboLibrary.TabIndex = 4
        '
        'dgFillPicker
        '
        Me.AcceptButton = Me.btnOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(400, 285)
        Me.Controls.Add(Me.cboFills)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.pnHatch)
        Me.Controls.Add(Me.picPreview)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.pnGradient)
        Me.Controls.Add(Me.pnSolidBrush)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.pnTexture)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "dgFillPicker"
        Me.ShowInTaskbar = False
        Me.Text = "Fill Properties"
        Me.pnHatch.ResumeLayout(False)
        Me.pnGradient.ResumeLayout(False)
        Me.pnGradCustom.ResumeLayout(False)
        CType(Me.nudAngle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnSolidBrush.ResumeLayout(False)
        Me.pnTexture.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Property Accessors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the selected fill.  Used for getting the out fill when this dialog 
    ''' closes.
    ''' </summary>
    ''' <returns>The final populated selected fill</returns>
    ''' -----------------------------------------------------------------------------
    Private ReadOnly Property SelectedFill() As GDIFill
        Get
            Return DirectCast(cboFills.SelectedItem, GDIFill)
        End Get
    End Property

#End Region

#Region "Helper Methods"





    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the respective local fill to the properties of the application wide 
    ''' current fill.
    ''' </summary>
     ''' -----------------------------------------------------------------------------
    Private Sub mapCurrentFillToLocal()
        Trace.WriteLineIf(App.TraceOn, "Filldialog.Populate")

        Dim f As GDIFill = Session.Fill

        If TypeOf f Is GDISolidFill Then
            _SolidFill.Dispose()
            _SolidFill = DirectCast(f.Clone(), GDISolidFill)
            _SelectedFillIndex = 0

        ElseIf TypeOf f Is GDIHatchFill Then
            _hatchFill.Dispose()
            _hatchFill = DirectCast(f.Clone(), GDIHatchFill)
            _SelectedFillIndex = 1
      
        ElseIf TypeOf f Is GDIGradientFill Then
            Dim oldGrad As GDIGradientFill = DirectCast(f, GDIGradientFill)
            _GradientFill.Dispose()
            _GradientFill = New GDIGradientFill(oldGrad.Color1, oldGrad.Color2, oldGrad.GradientMode)
            _SelectedFillIndex = 2

        ElseIf TypeOf f Is GDITexturedFill Then
            _TextureFill.Dispose()
            _TextureFill = DirectCast(f.Clone(), GDITexturedFill)
            _SelectedFillIndex = 3
        End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Shows the solid fill options panel
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub showSolidPanel()
        Trace.WriteLineIf(App.TraceOn, "Filldialog.showSolid")

        pnGradient.Visible = False
        pnHatch.Visible = False
        pnSolidBrush.Visible = True
        pnTexture.Visible = False
        picSolid.BackColor = _SolidFill.Color
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Shows the hatch fill options panel
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub showHatchPanel()
        Trace.WriteLineIf(App.TraceOn, "Filldialog.showHatch")

        pnGradient.Visible = False
        pnHatch.Visible = True
        pnSolidBrush.Visible = False
        pnTexture.Visible = False
        Dim x As Drawing2D.HatchStyle

        With cboHatchStyle
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(Drawing2D.HatchStyle)))
            .SelectedItem = _hatchFill.HatchStyle.ToString
            .EndUpdate()
            picHatchFore.BackColor = _hatchFill.HatchForeColor
            picHatchBack.BackColor = _hatchFill.HatchBackColor
        End With

     End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Shows the texture fill options panel
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub showTexturePanel()
        Trace.WriteLineIf(App.TraceOn, "Filldialog.showTexture")
        pnGradient.Visible = False
        pnHatch.Visible = False
        pnSolidBrush.Visible = False
        pnTexture.Visible = True

        With cboLibrary
            .BeginUpdate()
            .Items.Clear()
            For Each sImage As String In System.IO.Directory.GetFiles(App.Options.TexturePath)
                Dim fInfo As New System.IO.FileInfo(sImage)
                If fInfo.Extension = ".gif" OrElse fInfo.Extension = ".jpg" OrElse fInfo.Extension = ".jpeg" OrElse fInfo.Extension = ".png" OrElse fInfo.Extension = ".bmp" Then
                    cboLibrary.Items.Add(fInfo.Name)
                    If _TextureFill.ImageSource = fInfo.FullName Then
                        cboLibrary.Text = fInfo.Name
                    End If
                End If

            Next

            .EndUpdate()
        End With

     End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Shows the gradient fill options panel
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub showGradientPanel()
        Trace.WriteLineIf(App.TraceOn, "Filldialog.showGradient")
        pnGradient.Visible = True
        pnHatch.Visible = False
        pnSolidBrush.Visible = False
        pnTexture.Visible = False
        pnTexture.Visible = False
        With cboGradient
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(GDIGradientFill.EnumGradientMode)))

            .SelectedItem = _GradientFill.GradientMode.ToString
            .EndUpdate()
            picGrad1.BackColor = _GradientFill.Color1
            picGrad2.BackColor = _GradientFill.Color2

        End With

     End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles changes in the fill drop down, showing the appropriate panel for 
    ''' the cboFills entry
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub toggleFillPanels()
        If cboFills.SelectedItem Is _SolidFill Then
            showSolidPanel()
        ElseIf cboFills.SelectedItem Is _hatchFill Then
            showHatchPanel()
        ElseIf cboFills.SelectedItem Is _GradientFill Then
            showGradientPanel()
        Else
            showTexturePanel()
        End If


    End Sub
#End Region


#Region "Event Handlers"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles changes to the fill type drop down.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboFills_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFills.SelectedValueChanged
        toggleFillPanels()
        picPreview.Invalidate()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Cancels the fill picker.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles clicks on the OK button.  Sets the application wide GDIFill.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Trace.WriteLineIf(App.TraceOn, "Filldialog.Ok")
        If Not Me.SelectedFill Is Nothing Then

            Session.Fill = Me.SelectedFill
            Me.DialogResult = DialogResult.OK

        Else
            MsgBox("You currently have an invalid fill.  Please check your settings and try again.")
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders a preview of the current fill to the picPreview picture box.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub picPreview_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles picPreview.Paint
 
        e.Graphics.FillRectangle(Me.SelectedFill.Brush, 0, 0, picPreview.Width, picPreview.Height)
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles clicks on the first color gradient fill dialog box. 
    ''' Prompts for a color.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub picGrad1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picGrad1.Click
        Dim tempcol As Color = Utility.PickColor(_GradientFill.Color1)

        If Not Color.op_Equality(tempcol, Color.Empty) Then
            _GradientFill.Color1 = tempcol
            picGrad1.BackColor = tempcol
            picGrad1.Invalidate()
        End If

        picPreview.Invalidate()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a click on the second gradient fill color box.
    '''  Prompts for a color.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub picGrad2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picGrad2.Click
        Dim tempcol As Color = Utility.PickColor(_GradientFill.Color2)

        If Not Color.op_Equality(tempcol, Color.Empty) Then
            picGrad2.BackColor = tempcol
            _GradientFill.Color2 = tempcol
            picGrad2.Invalidate()
        End If

        picPreview.Invalidate()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a click on the hatch back color box. Prompts for a color.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub picHatchBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picHatchBack.Click
        Dim tempcol As Color = Utility.PickColor(_hatchFill.HatchBackColor)

        If Not Color.op_Equality(tempcol, Color.Empty) Then
            picHatchBack.BackColor = tempcol
            _hatchFill.HatchBackColor = tempcol
            picHatchBack.Invalidate()
        End If

        picPreview.Invalidate()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles clicks on the hatch fore color dialog box. Prompts for a color.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub picHatchFore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picHatchFore.Click
        Dim tempcol As Color = Utility.PickColor(_hatchFill.HatchForeColor)

        If Not Color.op_Equality(tempcol, Color.Empty) Then

            picHatchFore.BackColor = tempcol
            _hatchFill.HatchForeColor = tempcol
            picHatchFore.Invalidate()
        End If

        picPreview.Invalidate()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' handles changes in the hatch style drop down box.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboHatchStyle_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboHatchStyle.SelectedValueChanged
        _hatchFill.HatchStyle = DirectCast(System.Enum.GetValues(GetType(Drawing2D.HatchStyle)).GetValue(cboHatchStyle.SelectedIndex), Drawing2D.HatchStyle)

        picPreview.Invalidate()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles texture selection via a click on the texture ellipses button.
    ''' </summary>
    ''' <remarks>
    ''' This method prompts for a texture source, and, if valid, assigns the 
    ''' texture to the texture fill.
    ''' </remarks>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnTexture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTextureFile.Click
        'Dim dgopen As New file
        Dim dgopen As New OpenFileDialog
        dgopen.Filter = GDIImage.CONST_IMG_FILTER
        If _TextureFill.ImageSource <> String.Empty Then
            dgopen.InitialDirectory = _TextureFill.ImageSource
            dgopen.Title = "Texture Selection"
            dgopen.CheckPathExists = True
            dgopen.CheckFileExists = True
        End If

        Dim iresp As DialogResult = dgopen.ShowDialog()
        If iresp = DialogResult.OK Then
            Try
                Dim fInfo As System.IO.FileInfo = New System.IO.FileInfo(dgopen.FileName)
                _TextureFill.ImageSource = fInfo.FullName
                txtTexture.Text = fInfo.FullName()

            Catch ex As Exception
                Trace.WriteLineIf(App.TraceOn, "Texture.Exception" & ex.Message)
                MsgBox("Invalid file source")
            End Try
        End If

        picPreview.Invalidate()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles changes in the gradient drop down indicating the type of gradient 
    ''' to render.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboGradient_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGradient.SelectedValueChanged
        _GradientFill.GradientMode = DirectCast(System.Enum.GetValues(GetType(GDIGradientFill.EnumGradientMode)).GetValue(cboGradient.SelectedIndex), GDIGradientFill.EnumGradientMode)

        If _GradientFill.GradientMode = GDIGradientFill.EnumGradientMode.Custom Then
            nudAngle.Text = _GradientFill.Angle.ToString
            pnGradCustom.Visible = True

        Else
            pnGradCustom.Visible = False

        End If

        picPreview.Invalidate()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles clicks on the solid color box. Prompts for a color.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub picSolid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picSolid.Click

        Dim tempcol As Color = Utility.PickColor(_SolidFill.Color)

        If Not Color.op_Equality(tempcol, Color.Empty) Then
            _SolidFill.Color = tempcol
            picSolid.BackColor = tempcol
            picSolid.Invalidate()
        End If


        picPreview.Invalidate()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the selected texture from the texture libary drop down.
    ''' When this value changes, it attempts to make the image from the drop down 
    ''' the texture for the GDITexture brush.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboLibrary_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLibrary.SelectedValueChanged
        If cboLibrary.Text.Length > 0 Then
            Try
                Dim fInfo As New System.IO.FileInfo(App.Options.TexturePath & "\" & cboLibrary.Text)
                If fInfo.Exists Then
                    _TextureFill.ImageSource = fInfo.FullName
                    txtTexture.Text = fInfo.FullName
                    picPreview.Invalidate()
                End If
            Catch ex As Exception
                MsgBox("Unable to load this texture")
            End Try
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles changes in the gradient angle nud.  Sets the custom angle property for 
    ''' gradient fills.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub nudAngle_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nudAngle.TextChanged, nudAngle.ValueChanged
        If nudAngle.Text.Length > 0 Then
            _GradientFill.Angle = Single.Parse(nudAngle.Text)
            picPreview.Invalidate()
        End If
    End Sub

#End Region


End Class
