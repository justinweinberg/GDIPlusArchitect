Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : ArrangeWin
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a user interface to perform various alignment operations to selected object(s).
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class ArrangeWin
    Inherits System.Windows.Forms.UserControl

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new ArrangeWin.  Sets up the image links to buttons and gets the 
    ''' initial values of align to canvas or align to margin.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        initImages()

        'Add any initialization after the InitializeComponent() call

        pshCanvas.Checked = App.AlignManager.AlignMode = EnumAlignMode.eCanvas
        pshMargins.Checked = App.AlignManager.AlignMode = EnumAlignMode.eMargins

    End Sub

#End Region

#Region " Windows Form Designer generated code "


    'UserControl overrides dispose to clean up the component list.
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
    Private WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents Label4 As System.Windows.Forms.Label
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents btnAlignLeft As GDIPlusArchitect.RollButton
    Private WithEvents btnAlignMiddle As GDIPlusArchitect.RollButton
    Private WithEvents btnAlignRight As GDIPlusArchitect.RollButton
    Private WithEvents btnAlignBottom As GDIPlusArchitect.RollButton
    Private WithEvents btnAlignCenter As GDIPlusArchitect.RollButton
    Private WithEvents btnAlignTop As GDIPlusArchitect.RollButton
    Private WithEvents btnSameSizeHeight As GDIPlusArchitect.RollButton
    Private WithEvents btnSameSizeWidth As GDIPlusArchitect.RollButton
    Private WithEvents btnHorizSpaceEqual As GDIPlusArchitect.RollButton
    Private WithEvents btnVertSpaceEqual As GDIPlusArchitect.RollButton
    Private WithEvents btnBringToFront As GDIPlusArchitect.RollButton
    Private WithEvents btnBringForward As GDIPlusArchitect.RollButton
    Private WithEvents btnSendBackward As GDIPlusArchitect.RollButton
    Private WithEvents btnSendToBack As GDIPlusArchitect.RollButton
    Private WithEvents btnSameSizeBoth As GDIPlusArchitect.RollButton
    Private WithEvents pshMargins As System.Windows.Forms.RadioButton
    Private WithEvents pshCanvas As System.Windows.Forms.RadioButton
    Private WithEvents Label5 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.pshMargins = New System.Windows.Forms.RadioButton
        Me.pshCanvas = New System.Windows.Forms.RadioButton
        Me.Label2 = New System.Windows.Forms.Label
        Me.btnAlignLeft = New GDIPlusArchitect.RollButton
        Me.btnAlignMiddle = New GDIPlusArchitect.RollButton
        Me.btnAlignRight = New GDIPlusArchitect.RollButton
        Me.btnAlignBottom = New GDIPlusArchitect.RollButton
        Me.btnAlignCenter = New GDIPlusArchitect.RollButton
        Me.btnAlignTop = New GDIPlusArchitect.RollButton
        Me.btnSameSizeHeight = New GDIPlusArchitect.RollButton
        Me.btnSameSizeWidth = New GDIPlusArchitect.RollButton
        Me.btnHorizSpaceEqual = New GDIPlusArchitect.RollButton
        Me.btnVertSpaceEqual = New GDIPlusArchitect.RollButton
        Me.btnBringToFront = New GDIPlusArchitect.RollButton
        Me.btnBringForward = New GDIPlusArchitect.RollButton
        Me.btnSendBackward = New GDIPlusArchitect.RollButton
        Me.btnSendToBack = New GDIPlusArchitect.RollButton
        Me.btnSameSizeBoth = New GDIPlusArchitect.RollButton
        Me.Label5 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(10, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(42, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Align:"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(10, 58)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 17)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Same Size:"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(80, 57)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(50, 17)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "Spacing:"
        '
        'pshMargins
        '
        Me.pshMargins.Appearance = System.Windows.Forms.Appearance.Button
        Me.pshMargins.AutoCheck = False
        Me.pshMargins.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pshMargins.Location = New System.Drawing.Point(62, 2)
        Me.pshMargins.Name = "pshMargins"
        Me.pshMargins.Size = New System.Drawing.Size(46, 20)
        Me.pshMargins.TabIndex = 28
        Me.pshMargins.Text = "Margins"
        Me.pshMargins.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pshCanvas
        '
        Me.pshCanvas.Appearance = System.Windows.Forms.Appearance.Button
        Me.pshCanvas.AutoCheck = False
        Me.pshCanvas.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pshCanvas.Location = New System.Drawing.Point(107, 2)
        Me.pshCanvas.Name = "pshCanvas"
        Me.pshCanvas.Size = New System.Drawing.Size(45, 20)
        Me.pshCanvas.TabIndex = 29
        Me.pshCanvas.Text = "Canvas"
        Me.pshCanvas.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(10, 96)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 12)
        Me.Label2.TabIndex = 49
        Me.Label2.Text = "Order:"
        '
        'btnAlignLeft
        '
        Me.btnAlignLeft.BackColor = System.Drawing.SystemColors.Control
        Me.btnAlignLeft.Enabled = False
        Me.btnAlignLeft.ImageIndex = 0
        Me.btnAlignLeft.ImageList = Nothing
        Me.btnAlignLeft.Location = New System.Drawing.Point(10, 33)
        Me.btnAlignLeft.Name = "btnAlignLeft"
        Me.btnAlignLeft.Size = New System.Drawing.Size(20, 20)
        Me.btnAlignLeft.TabIndex = 50
        '
        'btnAlignMiddle
        '
        Me.btnAlignMiddle.BackColor = System.Drawing.SystemColors.Control
        Me.btnAlignMiddle.Enabled = False
        Me.btnAlignMiddle.ImageIndex = 0
        Me.btnAlignMiddle.ImageList = Nothing
        Me.btnAlignMiddle.Location = New System.Drawing.Point(30, 33)
        Me.btnAlignMiddle.Name = "btnAlignMiddle"
        Me.btnAlignMiddle.Size = New System.Drawing.Size(20, 20)
        Me.btnAlignMiddle.TabIndex = 51
        '
        'btnAlignRight
        '
        Me.btnAlignRight.BackColor = System.Drawing.SystemColors.Control
        Me.btnAlignRight.Enabled = False
        Me.btnAlignRight.ImageIndex = 0
        Me.btnAlignRight.ImageList = Nothing
        Me.btnAlignRight.Location = New System.Drawing.Point(50, 33)
        Me.btnAlignRight.Name = "btnAlignRight"
        Me.btnAlignRight.Size = New System.Drawing.Size(20, 20)
        Me.btnAlignRight.TabIndex = 52
        '
        'btnAlignBottom
        '
        Me.btnAlignBottom.BackColor = System.Drawing.SystemColors.Control
        Me.btnAlignBottom.Enabled = False
        Me.btnAlignBottom.ImageIndex = 0
        Me.btnAlignBottom.ImageList = Nothing
        Me.btnAlignBottom.Location = New System.Drawing.Point(120, 33)
        Me.btnAlignBottom.Name = "btnAlignBottom"
        Me.btnAlignBottom.Size = New System.Drawing.Size(20, 20)
        Me.btnAlignBottom.TabIndex = 55
        '
        'btnAlignCenter
        '
        Me.btnAlignCenter.BackColor = System.Drawing.SystemColors.Control
        Me.btnAlignCenter.Enabled = False
        Me.btnAlignCenter.ImageIndex = 0
        Me.btnAlignCenter.ImageList = Nothing
        Me.btnAlignCenter.Location = New System.Drawing.Point(100, 33)
        Me.btnAlignCenter.Name = "btnAlignCenter"
        Me.btnAlignCenter.Size = New System.Drawing.Size(20, 20)
        Me.btnAlignCenter.TabIndex = 54
        '
        'btnAlignTop
        '
        Me.btnAlignTop.BackColor = System.Drawing.SystemColors.Control
        Me.btnAlignTop.Enabled = False
        Me.btnAlignTop.ImageIndex = 0
        Me.btnAlignTop.ImageList = Nothing
        Me.btnAlignTop.Location = New System.Drawing.Point(80, 33)
        Me.btnAlignTop.Name = "btnAlignTop"
        Me.btnAlignTop.Size = New System.Drawing.Size(20, 20)
        Me.btnAlignTop.TabIndex = 53
        '
        'btnSameSizeHeight
        '
        Me.btnSameSizeHeight.BackColor = System.Drawing.SystemColors.Control
        Me.btnSameSizeHeight.Enabled = False
        Me.btnSameSizeHeight.ImageIndex = 0
        Me.btnSameSizeHeight.ImageList = Nothing
        Me.btnSameSizeHeight.Location = New System.Drawing.Point(30, 72)
        Me.btnSameSizeHeight.Name = "btnSameSizeHeight"
        Me.btnSameSizeHeight.Size = New System.Drawing.Size(20, 20)
        Me.btnSameSizeHeight.TabIndex = 57
        '
        'btnSameSizeWidth
        '
        Me.btnSameSizeWidth.BackColor = System.Drawing.SystemColors.Control
        Me.btnSameSizeWidth.Enabled = False
        Me.btnSameSizeWidth.ImageIndex = 0
        Me.btnSameSizeWidth.ImageList = Nothing
        Me.btnSameSizeWidth.Location = New System.Drawing.Point(10, 72)
        Me.btnSameSizeWidth.Name = "btnSameSizeWidth"
        Me.btnSameSizeWidth.Size = New System.Drawing.Size(20, 20)
        Me.btnSameSizeWidth.TabIndex = 56
        '
        'btnHorizSpaceEqual
        '
        Me.btnHorizSpaceEqual.BackColor = System.Drawing.SystemColors.Control
        Me.btnHorizSpaceEqual.Enabled = False
        Me.btnHorizSpaceEqual.ImageIndex = 0
        Me.btnHorizSpaceEqual.ImageList = Nothing
        Me.btnHorizSpaceEqual.Location = New System.Drawing.Point(100, 72)
        Me.btnHorizSpaceEqual.Name = "btnHorizSpaceEqual"
        Me.btnHorizSpaceEqual.Size = New System.Drawing.Size(20, 20)
        Me.btnHorizSpaceEqual.TabIndex = 59
        '
        'btnVertSpaceEqual
        '
        Me.btnVertSpaceEqual.BackColor = System.Drawing.SystemColors.Control
        Me.btnVertSpaceEqual.Enabled = False
        Me.btnVertSpaceEqual.ImageIndex = 0
        Me.btnVertSpaceEqual.ImageList = Nothing
        Me.btnVertSpaceEqual.Location = New System.Drawing.Point(80, 72)
        Me.btnVertSpaceEqual.Name = "btnVertSpaceEqual"
        Me.btnVertSpaceEqual.Size = New System.Drawing.Size(20, 20)
        Me.btnVertSpaceEqual.TabIndex = 58
        '
        'btnBringToFront
        '
        Me.btnBringToFront.BackColor = System.Drawing.SystemColors.Control
        Me.btnBringToFront.Enabled = False
        Me.btnBringToFront.ImageIndex = 0
        Me.btnBringToFront.ImageList = Nothing
        Me.btnBringToFront.Location = New System.Drawing.Point(80, 110)
        Me.btnBringToFront.Name = "btnBringToFront"
        Me.btnBringToFront.Size = New System.Drawing.Size(20, 20)
        Me.btnBringToFront.TabIndex = 63
        '
        'btnBringForward
        '
        Me.btnBringForward.BackColor = System.Drawing.SystemColors.Control
        Me.btnBringForward.Enabled = False
        Me.btnBringForward.ImageIndex = 0
        Me.btnBringForward.ImageList = Nothing
        Me.btnBringForward.Location = New System.Drawing.Point(10, 110)
        Me.btnBringForward.Name = "btnBringForward"
        Me.btnBringForward.Size = New System.Drawing.Size(20, 20)
        Me.btnBringForward.TabIndex = 62
        '
        'btnSendBackward
        '
        Me.btnSendBackward.BackColor = System.Drawing.SystemColors.Control
        Me.btnSendBackward.Enabled = False
        Me.btnSendBackward.ImageIndex = 0
        Me.btnSendBackward.ImageList = Nothing
        Me.btnSendBackward.Location = New System.Drawing.Point(30, 110)
        Me.btnSendBackward.Name = "btnSendBackward"
        Me.btnSendBackward.Size = New System.Drawing.Size(20, 20)
        Me.btnSendBackward.TabIndex = 61
        '
        'btnSendToBack
        '
        Me.btnSendToBack.BackColor = System.Drawing.SystemColors.Control
        Me.btnSendToBack.Enabled = False
        Me.btnSendToBack.ImageIndex = 0
        Me.btnSendToBack.ImageList = Nothing
        Me.btnSendToBack.Location = New System.Drawing.Point(100, 110)
        Me.btnSendToBack.Name = "btnSendToBack"
        Me.btnSendToBack.Size = New System.Drawing.Size(20, 20)
        Me.btnSendToBack.TabIndex = 60
        '
        'btnSameSizeBoth
        '
        Me.btnSameSizeBoth.BackColor = System.Drawing.SystemColors.Control
        Me.btnSameSizeBoth.Enabled = False
        Me.btnSameSizeBoth.ImageIndex = 0
        Me.btnSameSizeBoth.ImageList = Nothing
        Me.btnSameSizeBoth.Location = New System.Drawing.Point(50, 72)
        Me.btnSameSizeBoth.Name = "btnSameSizeBoth"
        Me.btnSameSizeBoth.Size = New System.Drawing.Size(20, 20)
        Me.btnSameSizeBoth.TabIndex = 64
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(10, 1)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(49, 17)
        Me.Label5.TabIndex = 65
        Me.Label5.Text = "Arrange to:"
        '
        'ArrangeWin
        '
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.pshMargins)
        Me.Controls.Add(Me.pshCanvas)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnAlignLeft)
        Me.Controls.Add(Me.btnAlignMiddle)
        Me.Controls.Add(Me.btnAlignRight)
        Me.Controls.Add(Me.btnAlignBottom)
        Me.Controls.Add(Me.btnAlignCenter)
        Me.Controls.Add(Me.btnAlignTop)
        Me.Controls.Add(Me.btnSameSizeHeight)
        Me.Controls.Add(Me.btnSameSizeWidth)
        Me.Controls.Add(Me.btnHorizSpaceEqual)
        Me.Controls.Add(Me.btnVertSpaceEqual)
        Me.Controls.Add(Me.btnBringToFront)
        Me.Controls.Add(Me.btnBringForward)
        Me.Controls.Add(Me.btnSendBackward)
        Me.Controls.Add(Me.btnSendToBack)
        Me.Controls.Add(Me.btnSameSizeBoth)
        Me.Controls.Add(Me.Label5)
        Me.Name = "ArrangeWin"
        Me.Size = New System.Drawing.Size(155, 145)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Refresh"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Enables and disabled appropriately based upon how many objects are currently 
    ''' selected in a document.
    ''' </summary>
    ''' <param name="doc">The document which is currently selected.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub refreshPanel()
        Trace.WriteLineIf(App.TraceOn, "ArrangeWin.Refresh")

        Dim bHaveInstance As Boolean = Not MDIMain.ActiveDocument Is Nothing
        Dim bHaveSelection As Boolean
        Dim bTwoOrMoreSelected As Boolean

        If bHaveInstance Then
            bHaveSelection = MDIMain.ActiveDocument.Selected.Count > 0
            bTwoOrMoreSelected = MDIMain.ActiveDocument.Selected.Count > 1
        Else
            bHaveSelection = False
            bTwoOrMoreSelected = False
        End If

        If bTwoOrMoreSelected OrElse _
        (bHaveSelection AndAlso _
        (Not App.AlignManager.AlignMode = EnumAlignMode.eNormal)) Then
            enableAlignCommands()
        Else
            disableAlignCommands()
        End If

        btnSendBackward.Enabled = bHaveSelection
        btnBringForward.Enabled = bHaveSelection
        btnSendToBack.Enabled = bHaveSelection
        btnBringToFront.Enabled = bHaveSelection


        refreshAlignTo()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Refreshes the align to canvas and align to margins button based upon the current 
    ''' alignmode in the Session object.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub refreshAlignTo()

        pshCanvas.Checked = App.AlignManager.AlignMode = EnumAlignMode.eCanvas
        pshMargins.Checked = App.AlignManager.AlignMode = EnumAlignMode.eMargins

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disables all align commands.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub disableAlignCommands()
        btnAlignTop.Enabled = False
        btnAlignCenter.Enabled = False
        btnAlignBottom.Enabled = False

        btnAlignLeft.Enabled = False
        btnAlignMiddle.Enabled = False
        btnAlignRight.Enabled = False

        btnSameSizeHeight.Enabled = False
        btnSameSizeWidth.Enabled = False

        btnVertSpaceEqual.Enabled = False
        btnHorizSpaceEqual.Enabled = False
        btnSameSizeBoth.Enabled = False

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Enables all align commands
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub enableAlignCommands()
        btnAlignTop.Enabled = True
        btnAlignCenter.Enabled = True
        btnAlignBottom.Enabled = True

        btnAlignLeft.Enabled = True
        btnAlignMiddle.Enabled = True
        btnAlignRight.Enabled = True

        btnSameSizeHeight.Enabled = True
        btnSameSizeWidth.Enabled = True

        btnVertSpaceEqual.Enabled = True
        btnHorizSpaceEqual.Enabled = True
        btnSameSizeBoth.Enabled = True

    End Sub

#End Region

#Region "Image setup"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the image indexes and image lists of the buttons on the alignment panel 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub initImages()

        btnAlignTop.ImageIndex = IconManager.EnumIcons.AlignTop
        btnAlignCenter.ImageIndex = IconManager.EnumIcons.AlignCenter
        btnAlignBottom.ImageIndex = IconManager.EnumIcons.AlignBottom

        btnAlignLeft.ImageIndex = IconManager.EnumIcons.AlignLeft
        btnAlignMiddle.ImageIndex = IconManager.EnumIcons.AlignMiddle
        btnAlignRight.ImageIndex = IconManager.EnumIcons.AlignRight

        btnSameSizeHeight.ImageIndex = IconManager.EnumIcons.SameSizeHeight
        btnSameSizeWidth.ImageIndex = IconManager.EnumIcons.SameSizeWidth

        btnVertSpaceEqual.ImageIndex = IconManager.EnumIcons.VertSpaceEqual
        btnHorizSpaceEqual.ImageIndex = IconManager.EnumIcons.HorizSpaceEqual
        btnSameSizeBoth.ImageIndex = IconManager.EnumIcons.SameSizeBoth
        btnSendBackward.ImageIndex = IconManager.EnumIcons.SendBackward
        btnBringForward.ImageIndex = IconManager.EnumIcons.BringForward
        btnSendToBack.ImageIndex = IconManager.EnumIcons.SendToBack
        btnBringToFront.ImageIndex = IconManager.EnumIcons.BringToFront

        btnAlignTop.ImageList = App.IconManager.IconImageList
        btnAlignCenter.ImageList = App.IconManager.IconImageList
        btnAlignBottom.ImageList = App.IconManager.IconImageList
        btnAlignLeft.ImageList = App.IconManager.IconImageList
        btnAlignMiddle.ImageList = App.IconManager.IconImageList
        btnAlignRight.ImageList = App.IconManager.IconImageList
        btnSameSizeHeight.ImageList = App.IconManager.IconImageList
        btnSameSizeWidth.ImageList = App.IconManager.IconImageList
        btnSameSizeBoth.ImageList = App.IconManager.IconImageList

        btnVertSpaceEqual.ImageList = App.IconManager.IconImageList
        btnHorizSpaceEqual.ImageList = App.IconManager.IconImageList
        btnSendBackward.ImageList = App.IconManager.IconImageList
        btnBringForward.ImageList = App.IconManager.IconImageList
        btnSendToBack.ImageList = App.IconManager.IconImageList
        btnBringToFront.ImageList = App.IconManager.IconImageList

    End Sub


#End Region

#Region "Event Handlers"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to clicks on the Canvas button. Toggles margins if selected and sets the 
    ''' session align variable appropriately.  Notifies the GDIMenu of changes in the AlignTo
    ''' so it can update its user interface.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub pshCanvas_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pshCanvas.Click
        pshCanvas.Checked = Not pshCanvas.Checked
        If pshCanvas.Checked Then
            App.AlignManager.AlignMode = EnumAlignMode.eCanvas
            pshMargins.Checked = False
        Else
            If pshMargins.Checked = False Then
                App.AlignManager.AlignMode = EnumAlignMode.eNormal
            End If
        End If

        App.MDIMain.RefreshMenu()


    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to clicks on the Margin button. Toggles canvas if selected and sets the 
    ''' session align variable appropriately.  Notifies the GDIMenu of changes in the AlignTo
    ''' so it can update its user interface.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub pshMargins_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pshMargins.Click
        pshMargins.Checked = Not pshMargins.Checked
        If pshMargins.Checked Then
            App.AlignManager.AlignMode = EnumAlignMode.eMargins
            pshCanvas.Checked = False
        Else
            If pshCanvas.Checked = False Then
                App.AlignManager.AlignMode = EnumAlignMode.eNormal
            End If
        End If
        App.MDIMain.RefreshMenu()


    End Sub

    'Invokes a same size both
    Private Sub btnSameSizeBoth_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnSameSizeBoth.ButtonClick
        App.AlignManager.SameSizeBoth()
    End Sub

    'Invokes a same size height
    Private Sub btnSameSizeHeight_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnSameSizeHeight.ButtonClick
        App.AlignManager.SameSizeHeight()
    End Sub

    'Invokes a same size width
    Private Sub btnSameSizeWidth_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnSameSizeWidth.ButtonClick
        App.AlignManager.SameSizeWidth()
    End Sub

    'Invokes equal horizontal spacing 
    Private Sub btnHorizSpaceEqual_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnHorizSpaceEqual.ButtonClick
        App.AlignManager.distributewidths()

    End Sub


    'Invokes equal vertical spacing 
    Private Sub btnVertSpaceEqual_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnVertSpaceEqual.ButtonClick
        App.AlignManager.distributeheights()
    End Sub

    'Invokes align right
    Private Sub btnAlignRight_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnAlignRight.ButtonClick
             App.AlignManager.AlignRight()
     End Sub


    'Invokes align middle
    Private Sub btnAlignMiddle_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnAlignMiddle.ButtonClick
        App.AlignManager.alignvertcenter()

    End Sub



    'Invokes align left
    Private Sub btnAlignLeft_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnAlignLeft.ButtonClick
        App.AlignManager.AlignLeft()

    End Sub


    'Invokes align top
    Private Sub btnAlignTop_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnAlignTop.ButtonClick
        App.AlignManager.AlignTop()
    End Sub


    'Invokes align center
    Private Sub btnAlignCenter_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnAlignCenter.ButtonClick
        App.AlignManager.AlignHorizCenter()
    End Sub

    'Invokes align bottom
    Private Sub btnAlignBottom_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnAlignBottom.ButtonClick
        App.AlignManager.AlignBottom()

    End Sub

    'Invokes bring forward 
    Private Sub btnBringForward_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnBringForward.ButtonClick
        App.AlignManager.BringForward()
    End Sub

    'Invokes send backward 
    Private Sub btnSendBackward_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnSendBackward.ButtonClick
        App.AlignManager.sendbackward()

    End Sub

    'Invokes a bring to front 
    Private Sub btnBringToFront_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnBringToFront.ButtonClick
        App.AlignManager.bringtofront()

    End Sub


    'Invokes a send to back
    Private Sub btnSendToBack_ButtonClick(ByVal s As Object, ByVal e As System.EventArgs) Handles btnSendToBack.ButtonClick
        App.AlignManager.sendtoback()

    End Sub
#End Region
End Class
