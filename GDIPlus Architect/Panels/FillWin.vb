Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : FillWin
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides an interface for setting the current fill, applying fills to selected objects,
''' and seeing the properties of fills.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class FillWin
    Inherits System.Windows.Forms.UserControl

#Region "Local Fields"

    Private _CBFillChanged As New Session.FillChanged(AddressOf OnFillChanged)

    '''<summary>Holds a solid fill.</summary>
    Private _TempSolidFill As GDISolidFill
    '''<summary>Holds a hatch fill</summary>
    Private _TempHatchFill As GDIHatchFill
    '''<summary>Holds a gradient fill</summary>
    Private _TempGradient As GDIGradientFill
    '''<summary>Holds a texture fill</summary>
    Private _TempTexture As GDITexturedFill


    '''<summary>Indicates whether a fill update is occurring and 
    ''' subsequent updates should be temporarily ignored</summary>
    Private _Updating As Boolean = False

#End Region


#Region "Constructors"


    Private Sub OnFillChanged(ByVal s As Object, ByVal e As EventArgs)
        Me.refreshPanel()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new intance of the Fill Window.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()


        _TempSolidFill = New GDISolidFill(Color.Black)
        _TempHatchFill = New GDIHatchFill(Color.White, Color.Black, Drawing2D.HatchStyle.BackwardDiagonal)
        _TempGradient = New GDIGradientFill(Color.Black, Color.White, GDIGradientFill.EnumGradientMode.ForwardDiagonal)
        _TempTexture = New GDITexturedFill(App.Options.GetFirstTexture())



        'This call is required by the Windows Form Designer.
        InitializeComponent()

        With cboFills
            .BeginUpdate()
            .Items.Clear()
            .Items.Add(New NVPair(0, "Solid"))
            .Items.Add(New NVPair(1, "Hatch"))
            .Items.Add(New NVPair(2, "Gradient"))
            .Items.Add(New NVPair(3, "Texture"))
            .EndUpdate()
        End With

        refreshPanel()

        GDIObjects.Session.setFillChangedCallBack(_CBFillChanged)
    End Sub


#End Region

#Region "Refresh"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The current fill can potentially be set from a variety of soures.  This 
    ''' includes the dropper tool, dgFillPicker, and the FillWin itself.  The current fill
    ''' is stored in Session.  This method updates the FillWin to reflect the 
    ''' state of that fill, regardless of how it got set.
    ''' </summary>
    ''' <remarks>Notice that this method short circuits with the _Updating argument
    ''' </remarks>
        ''' -----------------------------------------------------------------------------
    Private Sub refreshPanel()
        _Updating = True

        Dim fill As GDIFill = Session.Fill

        'Recreate the temp solid fill and show the solid panel
        If TypeOf fill Is GDISolidFill Then
            _TempSolidFill.Dispose()
            _TempSolidFill = Nothing
            _TempSolidFill = DirectCast(fill.Clone(), GDISolidFill)
            cboFills.SelectedIndex = 0
            showSolid()

            'Recreate temp hatch fill and show the hatch panel
        ElseIf TypeOf fill Is GDIHatchFill Then
            _TempHatchFill.Dispose()
            _TempHatchFill = Nothing
            _TempHatchFill = DirectCast(fill.Clone(), GDIHatchFill)
            cboFills.SelectedIndex = 1
            showHatch()

            'Recreate the temp gradient fill and show the gradient panel
        ElseIf TypeOf fill Is GDIGradientFill Then

            _TempGradient.Dispose()
            _TempGradient = Nothing
            _TempGradient = DirectCast(fill.Clone(), GDIGradientFill)


            cboFills.SelectedIndex = 2
            showGradient()

            'Recreate the temp texture fill and show the texture panel
        ElseIf TypeOf fill Is GDITexturedFill Then
            _TempTexture.Dispose()
            _TempTexture = Nothing
            _TempTexture = DirectCast(fill.Clone(), GDITexturedFill)
            cboFills.SelectedIndex = 3

            showTexture()

        End If

        _Updating = False

    End Sub


#End Region

#Region " Windows Form Designer generated code "





    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            _TempSolidFill.Dispose()
            _TempHatchFill.Dispose()
            _TempGradient.Dispose()
            _TempTexture.Dispose()

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
    Private WithEvents pnTexture As System.Windows.Forms.Panel
    Private WithEvents Label7 As System.Windows.Forms.Label
    Private WithEvents txtTexture As System.Windows.Forms.TextBox
    Private WithEvents lblTextures As System.Windows.Forms.Label
    Private WithEvents cboLibrary As System.Windows.Forms.ComboBox
    Private WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents cboFills As System.Windows.Forms.ComboBox
    Private WithEvents pnHatch As System.Windows.Forms.Panel
    Private WithEvents picHatchBack As System.Windows.Forms.PictureBox
    Private WithEvents picHatchFore As System.Windows.Forms.PictureBox
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents cboHatchStyle As System.Windows.Forms.ComboBox
    Private WithEvents pnGradient As System.Windows.Forms.Panel
    Private WithEvents pnGradCustom As System.Windows.Forms.Panel
    Private WithEvents Label6 As System.Windows.Forms.Label
    Private WithEvents txtCustomGradientAngle As System.Windows.Forms.TextBox
    Private WithEvents picGrad2 As System.Windows.Forms.PictureBox
    Private WithEvents Label4 As System.Windows.Forms.Label
    Private WithEvents cboGradient As System.Windows.Forms.ComboBox
    Private WithEvents pnSolidBrush As System.Windows.Forms.Panel
    Private WithEvents picSolid As System.Windows.Forms.PictureBox
    Private WithEvents picGrad1 As System.Windows.Forms.PictureBox
    Private WithEvents btnTexture As System.Windows.Forms.Button
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents cboWrapMode As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnTexture = New System.Windows.Forms.Panel
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboWrapMode = New System.Windows.Forms.ComboBox
        Me.btnTexture = New System.Windows.Forms.Button
        Me.txtTexture = New System.Windows.Forms.TextBox
        Me.cboLibrary = New System.Windows.Forms.ComboBox
        Me.lblTextures = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboFills = New System.Windows.Forms.ComboBox
        Me.pnHatch = New System.Windows.Forms.Panel
        Me.picHatchBack = New System.Windows.Forms.PictureBox
        Me.picHatchFore = New System.Windows.Forms.PictureBox
        Me.cboHatchStyle = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.pnGradient = New System.Windows.Forms.Panel
        Me.pnGradCustom = New System.Windows.Forms.Panel
        Me.txtCustomGradientAngle = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.picGrad2 = New System.Windows.Forms.PictureBox
        Me.picGrad1 = New System.Windows.Forms.PictureBox
        Me.cboGradient = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.pnSolidBrush = New System.Windows.Forms.Panel
        Me.picSolid = New System.Windows.Forms.PictureBox
        Me.pnTexture.SuspendLayout()
        Me.pnHatch.SuspendLayout()
        Me.pnGradient.SuspendLayout()
        Me.pnGradCustom.SuspendLayout()
        Me.pnSolidBrush.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnTexture
        '
        Me.pnTexture.Controls.Add(Me.Label3)
        Me.pnTexture.Controls.Add(Me.cboWrapMode)
        Me.pnTexture.Controls.Add(Me.btnTexture)
        Me.pnTexture.Controls.Add(Me.txtTexture)
        Me.pnTexture.Controls.Add(Me.cboLibrary)
        Me.pnTexture.Controls.Add(Me.lblTextures)
        Me.pnTexture.Controls.Add(Me.Label7)
        Me.pnTexture.Location = New System.Drawing.Point(8, 40)
        Me.pnTexture.Name = "pnTexture"
        Me.pnTexture.Size = New System.Drawing.Size(128, 128)
        Me.pnTexture.TabIndex = 21
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(0, 88)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 16)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Wrap Mode"
        '
        'cboWrapMode
        '
        Me.cboWrapMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboWrapMode.Location = New System.Drawing.Point(0, 104)
        Me.cboWrapMode.Name = "cboWrapMode"
        Me.cboWrapMode.Size = New System.Drawing.Size(121, 21)
        Me.cboWrapMode.TabIndex = 23
        '
        'btnTexture
        '
        Me.btnTexture.Location = New System.Drawing.Point(96, 51)
        Me.btnTexture.Name = "btnTexture"
        Me.btnTexture.Size = New System.Drawing.Size(32, 23)
        Me.btnTexture.TabIndex = 8
        Me.btnTexture.Text = "..."
        '
        'txtTexture
        '
        Me.txtTexture.Location = New System.Drawing.Point(0, 52)
        Me.txtTexture.Name = "txtTexture"
        Me.txtTexture.Size = New System.Drawing.Size(96, 20)
        Me.txtTexture.TabIndex = 6
        Me.txtTexture.Text = ""
        '
        'cboLibrary
        '
        Me.cboLibrary.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLibrary.Location = New System.Drawing.Point(0, 13)
        Me.cboLibrary.Name = "cboLibrary"
        Me.cboLibrary.Size = New System.Drawing.Size(120, 21)
        Me.cboLibrary.TabIndex = 4
        '
        'lblTextures
        '
        Me.lblTextures.Location = New System.Drawing.Point(0, 0)
        Me.lblTextures.Name = "lblTextures"
        Me.lblTextures.Size = New System.Drawing.Size(56, 16)
        Me.lblTextures.TabIndex = 5
        Me.lblTextures.Text = "Library"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(0, 38)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(56, 16)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "Source"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 16)
        Me.Label1.TabIndex = 13
        Me.Label1.Text = "Fill Type"
        '
        'cboFills
        '
        Me.cboFills.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFills.Location = New System.Drawing.Point(8, 13)
        Me.cboFills.Name = "cboFills"
        Me.cboFills.Size = New System.Drawing.Size(96, 21)
        Me.cboFills.TabIndex = 12
        '
        'pnHatch
        '
        Me.pnHatch.Controls.Add(Me.picHatchBack)
        Me.pnHatch.Controls.Add(Me.picHatchFore)
        Me.pnHatch.Controls.Add(Me.cboHatchStyle)
        Me.pnHatch.Controls.Add(Me.Label2)
        Me.pnHatch.Location = New System.Drawing.Point(8, 40)
        Me.pnHatch.Name = "pnHatch"
        Me.pnHatch.Size = New System.Drawing.Size(128, 104)
        Me.pnHatch.TabIndex = 14
        '
        'picHatchBack
        '
        Me.picHatchBack.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picHatchBack.Location = New System.Drawing.Point(96, 40)
        Me.picHatchBack.Name = "picHatchBack"
        Me.picHatchBack.Size = New System.Drawing.Size(30, 30)
        Me.picHatchBack.TabIndex = 10
        Me.picHatchBack.TabStop = False
        '
        'picHatchFore
        '
        Me.picHatchFore.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picHatchFore.Location = New System.Drawing.Point(0, 40)
        Me.picHatchFore.Name = "picHatchFore"
        Me.picHatchFore.Size = New System.Drawing.Size(30, 30)
        Me.picHatchFore.TabIndex = 9
        Me.picHatchFore.TabStop = False
        '
        'cboHatchStyle
        '
        Me.cboHatchStyle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboHatchStyle.Location = New System.Drawing.Point(0, 14)
        Me.cboHatchStyle.Name = "cboHatchStyle"
        Me.cboHatchStyle.Size = New System.Drawing.Size(128, 21)
        Me.cboHatchStyle.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(0, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 16)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Hatch Pattern"
        '
        'pnGradient
        '
        Me.pnGradient.Controls.Add(Me.pnGradCustom)
        Me.pnGradient.Controls.Add(Me.picGrad2)
        Me.pnGradient.Controls.Add(Me.picGrad1)
        Me.pnGradient.Controls.Add(Me.cboGradient)
        Me.pnGradient.Controls.Add(Me.Label4)
        Me.pnGradient.Location = New System.Drawing.Point(8, 40)
        Me.pnGradient.Name = "pnGradient"
        Me.pnGradient.Size = New System.Drawing.Size(128, 106)
        Me.pnGradient.TabIndex = 17
        '
        'pnGradCustom
        '
        Me.pnGradCustom.Controls.Add(Me.txtCustomGradientAngle)
        Me.pnGradCustom.Controls.Add(Me.Label6)
        Me.pnGradCustom.Location = New System.Drawing.Point(0, 71)
        Me.pnGradCustom.Name = "pnGradCustom"
        Me.pnGradCustom.Size = New System.Drawing.Size(128, 34)
        Me.pnGradCustom.TabIndex = 11
        Me.pnGradCustom.Visible = False
        '
        'txtCustomGradientAngle
        '
        Me.txtCustomGradientAngle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtCustomGradientAngle.Location = New System.Drawing.Point(0, 13)
        Me.txtCustomGradientAngle.Name = "txtCustomGradientAngle"
        Me.txtCustomGradientAngle.Size = New System.Drawing.Size(40, 20)
        Me.txtCustomGradientAngle.TabIndex = 11
        Me.txtCustomGradientAngle.Text = ""
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(0, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(56, 23)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Angle"
        '
        'picGrad2
        '
        Me.picGrad2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picGrad2.Location = New System.Drawing.Point(80, 40)
        Me.picGrad2.Name = "picGrad2"
        Me.picGrad2.Size = New System.Drawing.Size(30, 30)
        Me.picGrad2.TabIndex = 8
        Me.picGrad2.TabStop = False
        '
        'picGrad1
        '
        Me.picGrad1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picGrad1.Location = New System.Drawing.Point(0, 40)
        Me.picGrad1.Name = "picGrad1"
        Me.picGrad1.Size = New System.Drawing.Size(30, 30)
        Me.picGrad1.TabIndex = 7
        Me.picGrad1.TabStop = False
        '
        'cboGradient
        '
        Me.cboGradient.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboGradient.Location = New System.Drawing.Point(0, 16)
        Me.cboGradient.Name = "cboGradient"
        Me.cboGradient.Size = New System.Drawing.Size(112, 21)
        Me.cboGradient.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(0, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 16)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Gradient"
        '
        'pnSolidBrush
        '
        Me.pnSolidBrush.Controls.Add(Me.picSolid)
        Me.pnSolidBrush.Location = New System.Drawing.Point(8, 40)
        Me.pnSolidBrush.Name = "pnSolidBrush"
        Me.pnSolidBrush.Size = New System.Drawing.Size(128, 104)
        Me.pnSolidBrush.TabIndex = 18
        '
        'picSolid
        '
        Me.picSolid.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picSolid.Location = New System.Drawing.Point(0, 0)
        Me.picSolid.Name = "picSolid"
        Me.picSolid.Size = New System.Drawing.Size(30, 30)
        Me.picSolid.TabIndex = 9
        Me.picSolid.TabStop = False
        '
        'FillWin
        '
        Me.Controls.Add(Me.pnTexture)
        Me.Controls.Add(Me.cboFills)
        Me.Controls.Add(Me.pnHatch)
        Me.Controls.Add(Me.pnGradient)
        Me.Controls.Add(Me.pnSolidBrush)
        Me.Controls.Add(Me.Label1)
        Me.Name = "FillWin"
        Me.Size = New System.Drawing.Size(150, 170)
        Me.pnTexture.ResumeLayout(False)
        Me.pnHatch.ResumeLayout(False)
        Me.pnGradient.ResumeLayout(False)
        Me.pnGradCustom.ResumeLayout(False)
        Me.pnSolidBrush.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Helper Methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in cboFills.SelectedIndex, indicating that a new fill 
    ''' has been chosen.
    ''' </summary>
    ''' <remarks>A more elegant solution may have used an apply button to note that 
    ''' the user wants to change the fill rather than browse the options available.
    ''' </remarks>
        ''' -----------------------------------------------------------------------------
    Private Sub updateFill()
        Trace.WriteLineIf(App.TraceOn, "FillWin.updateFill")

        If _Updating = False Then
            Select Case cboFills.SelectedIndex
                Case 0
                    Session.Fill = _TempSolidFill
                Case 1
                    Session.Fill = _TempHatchFill
                Case 2
                    Session.Fill = _TempGradient
                Case 3
                    Session.Fill = _TempTexture
            End Select
        End If
    End Sub

#End Region


#Region "Show / Hide"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Moves the solid panel to the front and sets its values appropriately
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub showSolid()
        Trace.WriteLineIf(App.TraceOn, "FillWin.showSolid")

        pnGradient.Visible = False
        pnHatch.Visible = False
        pnSolidBrush.Visible = True
        pnTexture.Visible = False
        picSolid.BackColor = _TempSolidFill.Color
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Moves the hatch panel to the front and sets its values appropriately
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub showHatch()
        Trace.WriteLineIf(App.TraceOn, "FillWin.ShowHatch")

        pnGradient.Visible = False
        pnHatch.Visible = True
        pnSolidBrush.Visible = False
        pnTexture.Visible = False

        With cboHatchStyle
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(Drawing2D.HatchStyle)))
            .SelectedItem = _TempHatchFill.HatchStyle.ToString
            .EndUpdate()
        End With

        picHatchFore.BackColor = _TempHatchFill.HatchForeColor
        picHatchBack.BackColor = _TempHatchFill.HatchBackColor

    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Moves the texture panel to the front and sets its values appropriately
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub showTexture()
        Trace.WriteLineIf(App.TraceOn, "FillWin.showTexture")

        pnGradient.Visible = False
        pnHatch.Visible = False
        pnSolidBrush.Visible = False
        pnTexture.Visible = True

        With cboLibrary
            .BeginUpdate()
            .Items.Clear()
            Try
                For Each sImage As String In System.IO.Directory.GetFiles(App.Options.TexturePath)
                    Dim fInfo As New System.IO.FileInfo(sImage)
                    If fInfo.Extension = ".gif" OrElse fInfo.Extension = ".jpg" OrElse fInfo.Extension = ".jpeg" OrElse fInfo.Extension = ".png" OrElse fInfo.Extension = ".bmp" Then
                        cboLibrary.Items.Add(fInfo.Name)
                        If _TempTexture.ImageSource = fInfo.FullName Then
                            cboLibrary.Text = fInfo.Name
                        End If
                    End If

                Next
            Catch ex As Exception
                MsgBox("Invalid texture library")
            End Try
            .EndUpdate()
        End With


        With cboWrapMode
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(Drawing2D.WrapMode)))
            .SelectedItem = _TempTexture.WrapMode.ToString
            .EndUpdate()
        End With

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Moves the gradient panel to the front and sets its values appropriately
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub showGradient()
        Trace.WriteLineIf(App.TraceOn, "FillWin.showGradient")

        pnGradient.Visible = True
        pnHatch.Visible = False
        pnSolidBrush.Visible = False
        pnTexture.Visible = False
        pnTexture.Visible = False

        With cboGradient
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(GDIGradientFill.EnumGradientMode)))
            .SelectedItem = _TempGradient.GradientMode.ToString
            .EndUpdate()
        End With

        picGrad1.BackColor = _TempGradient.Color1
        picGrad2.BackColor = _TempGradient.Color2
    End Sub

#End Region





#Region "Event Handlers"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to clicks on the first gradient color box which initiates color selection
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub picGrad1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picGrad1.Click
        Dim tempcol As Color = Utility.PickColor(_TempGradient.Color1)

        If Not Color.op_Equality(tempcol, Color.Empty) Then
            _TempGradient.Color1 = tempcol
            picGrad1.BackColor = tempcol
            picGrad1.Invalidate()
        End If

        updateFill()

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to clicks on the second gradient color box which initiates color selection
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub picGrad2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picGrad2.Click
        Dim tempcol As Color = Utility.PickColor(_TempGradient.Color2)

        If Not Color.op_Equality(tempcol, Color.Empty) Then
            picGrad2.BackColor = tempcol
            _TempGradient.Color2 = tempcol
            picGrad2.Invalidate()
        End If

        updateFill()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to clicks on the back color hatch box which initiates color selection
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub picHatchBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picHatchBack.Click
        Dim tempcol As Color = Utility.PickColor(_TempHatchFill.HatchBackColor)

        If Not Color.op_Equality(tempcol, Color.Empty) Then
            picHatchBack.BackColor = tempcol
            _TempHatchFill.HatchBackColor = tempcol
            picHatchBack.Invalidate()
        End If

        updateFill()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to clicks on the fore color hatch box which initiates color selection
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub picHatchFore_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picHatchFore.Click
        Dim tempcol As Color = Utility.PickColor(_TempHatchFill.HatchForeColor)

        If Not Color.op_Equality(tempcol, Color.Empty) Then

            picHatchFore.BackColor = tempcol
            _TempHatchFill.HatchForeColor = tempcol
            picHatchFore.Invalidate()
        End If

        updateFill()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the hatch style drop down.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboHatchStyle_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboHatchStyle.SelectedValueChanged
        _TempHatchFill.HatchStyle = DirectCast(System.Enum.GetValues(GetType(Drawing2D.HatchStyle)).GetValue(cboHatchStyle.SelectedIndex), Drawing2D.HatchStyle)

        updateFill()


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Requests a texture source from the user 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnTexture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTexture.Click
        'Dim dgopen As New file
        Dim dgopen As New OpenFileDialog
        dgopen.Filter = GDIImage.CONST_IMG_FILTER
        If _TempTexture.ImageSource <> String.Empty Then
            dgopen.InitialDirectory = _TempTexture.ImageSource
            dgopen.Title = "Texture Selection"
            dgopen.CheckPathExists = True
            dgopen.CheckFileExists = True
        End If

        Dim iresp As DialogResult = dgopen.ShowDialog()
        If iresp = DialogResult.OK Then
            Try
                Dim fInfo As System.IO.FileInfo = New System.IO.FileInfo(dgopen.FileName)
                _TempTexture.ImageSource = fInfo.FullName
                txtTexture.Text = fInfo.FullName()

            Catch ex As Exception
                Trace.WriteLineIf(App.TraceOn, "TextureLoad.Exception" & ex.Message)

                MsgBox("Invalid file source")
            End Try
        End If
        updateFill()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Changes the current gradient mode in response to a change in cboGradient.  If the 
    ''' custom mode is chosen, displays the custom mode panel as well.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboGradient_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboGradient.SelectedValueChanged
        _TempGradient.GradientMode = DirectCast(System.Enum.GetValues(GetType(GDIGradientFill.EnumGradientMode)).GetValue(cboGradient.SelectedIndex), GDIGradientFill.EnumGradientMode)

        If _TempGradient.GradientMode = GDIGradientFill.EnumGradientMode.Custom Then
            txtCustomGradientAngle.Text = _TempGradient.Angle.ToString

            pnGradCustom.Visible = True
        Else
            pnGradCustom.Visible = False

        End If

        updateFill()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Validates a custom gradient angle
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub txtCustomGradientAngle_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCustomGradientAngle.Validating

        Try
            Single.Parse(txtCustomGradientAngle.Text)
            e.Cancel = False
        Catch ex As Exception

            MsgBox("Invalid angle value")
            e.Cancel = True
            txtCustomGradientAngle.SelectAll()
        End Try


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the value on the custom gradient if valid
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub txtCustomGradientAngle_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCustomGradientAngle.Validated
        _TempGradient.Angle = Single.Parse(txtCustomGradientAngle.Text)

        updateFill()
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the custom gradient angle on a hard return
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub txtCustomGradientAngle_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCustomGradientAngle.KeyDown
        If e.KeyCode = Keys.Enter OrElse e.KeyCode = Keys.Return Then
            Try
                _TempGradient.Angle = Single.Parse(txtCustomGradientAngle.Text)
            Catch ex As Exception
                MsgBox("Invalid angle value")
                txtCustomGradientAngle.SelectAll()
            End Try
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Brings up the color selector when the solid preview is clicked
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub picSolid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picSolid.Click

        Dim tempcol As Color = Utility.PickColor(_TempSolidFill.Color)

        If Not Color.op_Equality(tempcol, Color.Empty) Then
            _TempSolidFill.Color = tempcol
            picSolid.BackColor = tempcol
            picSolid.Invalidate()
        End If

        updateFill()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles changes to the selected texture in the texture library drop down.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboLibrary_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboLibrary.SelectedValueChanged
        If cboLibrary.Text.Length > 0 Then
            Try
                Dim fInfo As New System.IO.FileInfo(App.Options.TexturePath & "\" & cboLibrary.Text)
                If fInfo.Exists Then
                    txtTexture.Text = fInfo.FullName
                    _TempTexture.ImageSource = fInfo.FullName
                    updateFill()
                End If
            Catch ex As Exception
                MsgBox("Unable to load this texture")
            End Try
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles changes to the texture wrap mode caused by changes in cboWrapMode
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboWrapMode_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboWrapMode.SelectedValueChanged
        _TempTexture.WrapMode = DirectCast(System.Enum.GetValues(GetType(Drawing2D.WrapMode)).GetValue(cboWrapMode.SelectedIndex), Drawing2D.WrapMode)

        updateFill()

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Update the current fill and displayed fill panel in response to a change in cboFills.
    ''' Note that the updateFill method short circuits itself to avoid infinite loops.  In other 
    ''' words, we want the "user" selected values to call this, but not the programmatic changes.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboFills_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFills.SelectedIndexChanged
        updateFill()
    End Sub


#End Region
End Class
