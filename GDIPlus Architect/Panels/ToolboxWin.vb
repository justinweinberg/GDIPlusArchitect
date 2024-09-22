Imports GDIObjects



''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : ToolboxWin
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A control to hold the tool window which is composed of a collection of ToolButtons
''' and color pick boxes.  The tool window is used to select the current application wide 
''' tool that is used on the surface.  The actual management of tools is located 
''' in the ToolManager class.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class ToolboxWin
    Inherits System.Windows.Forms.UserControl


    Private _CBStrokeChanged As New GDIObjects.Session.StrokeChanged(AddressOf OnStrokeChanged)
    Private _CBFillChanged As New GDIObjects.Session.FillChanged(AddressOf OnFillChanged)

#Region "Type Declarations"


    'Default button height for tool buttons
    Private Const CONST_BUTTON_HEIGHT As Int32 = 26
    'The height of the panel that holds the tool buttons
    Private Const CONST_LIST_PANEL_HEIGHT As Int32 = 312
    'The amount to scroll in response to a scroll up or down on the tool window.
    Private Const CONST_LIST_SCROLL_INCREMENT As Int32 = 26
#End Region

    Private Sub OnFillChanged(ByVal s As Object, ByVal e As EventArgs)
        picFillColor.Invalidate()
    End Sub

    Private Sub OnStrokeChanged(ByVal s As Object, ByVal e As EventArgs)
        picStrokeColor.Invalidate()
    End Sub

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the Toolbox window.  Refreshes the colors in the 
    ''' preview boxes and updates the images displayed in the tool buttons.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        pnToolBar.Top = 0

        '    setstyle(ControlStyles.DoubleBuffer, True)
        picStrokeColor.Invalidate()
        picFillColor.Invalidate()



        initImages()

        GDIObjects.Session.setStrokeChangedCallBack(_CBStrokeChanged)
        GDIObjects.Session.setFillChangedCallBack(_CBFillChanged)

    End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Assigns images to the tool buttons in the toolbox window.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub initImages()
        btnCircle.ImageIndex = IconManager.EnumIcons.circle
        btnCircle.Caption = "Ellipse"
        btnDropper.ImageIndex = IconManager.EnumIcons.drop
        btnDropper.Caption = "Dropper"
        btnField.ImageIndex = IconManager.EnumIcons.field
        btnFill.ImageIndex = IconManager.EnumIcons.bucket
        btnPlaceImage.ImageIndex = IconManager.EnumIcons.import
        btnLine.ImageIndex = IconManager.EnumIcons.line
        btnMagnify.ImageIndex = IconManager.EnumIcons.magnify
        btnPen.ImageIndex = IconManager.EnumIcons.pen

        btnPointer.ImageIndex = IconManager.EnumIcons.pointer
        btnSquare.ImageIndex = IconManager.EnumIcons.rect
        btnText.ImageIndex = IconManager.EnumIcons.alpha
        btnHand.ImageIndex = IconManager.EnumIcons.hand
        btnScrollDown.ImageIndex = IconManager.EnumIcons.DownButton
        btnScrollUp.ImageIndex = IconManager.EnumIcons.UpButton


        btnCircle.ImageList = App.IconManager.IconImageList
        btnDropper.ImageList = App.IconManager.IconImageList
        btnField.ImageList = App.IconManager.IconImageList
        btnFill.ImageList = App.IconManager.IconImageList
        btnPlaceImage.ImageList = App.IconManager.IconImageList
        btnLine.ImageList = App.IconManager.IconImageList
        btnMagnify.ImageList = App.IconManager.IconImageList
        btnPen.ImageList = App.IconManager.IconImageList
        btnPointer.ImageList = App.IconManager.IconImageList
        btnSquare.ImageList = App.IconManager.IconImageList
        btnText.ImageList = App.IconManager.IconImageList
        btnHand.ImageList = App.IconManager.IconImageList
        btnScrollDown.ImageList = App.IconManager.IconImageList
        btnScrollUp.ImageList = App.IconManager.IconImageList

    End Sub



#End Region


#Region " Windows Form Designer generated code "





    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Private WithEvents lblFill As System.Windows.Forms.Label
    Private WithEvents lblStroke As System.Windows.Forms.Label
    Private WithEvents picFillColor As System.Windows.Forms.PictureBox
    Private WithEvents picStrokeColor As System.Windows.Forms.PictureBox
    Private WithEvents btnScrollDown As GDIPlusArchitect.ScrollButton
    Private WithEvents btnScrollUp As GDIPlusArchitect.ScrollButton
    Private WithEvents pnToolBar As System.Windows.Forms.Panel
    Private WithEvents btnPointer As GDIPlusArchitect.ToolButton
    Private WithEvents btnHand As GDIPlusArchitect.ToolButton
    Private WithEvents btnText As GDIPlusArchitect.ToolButton
    Private WithEvents btnSquare As GDIPlusArchitect.ToolButton
    Private WithEvents btnCircle As GDIPlusArchitect.ToolButton
    Private WithEvents btnLine As GDIPlusArchitect.ToolButton
    Private WithEvents btnField As GDIPlusArchitect.ToolButton
    Private WithEvents btnPen As GDIPlusArchitect.ToolButton
    Private WithEvents btnPlaceImage As GDIPlusArchitect.ToolButton
    Private WithEvents btnFill As GDIPlusArchitect.ToolButton
    Private WithEvents btnDropper As GDIPlusArchitect.ToolButton
    Private WithEvents btnMagnify As GDIPlusArchitect.ToolButton
    Private WithEvents pnContainer As System.Windows.Forms.Panel
    Private WithEvents pnScrollArea As System.Windows.Forms.Panel

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblFill = New System.Windows.Forms.Label
        Me.lblStroke = New System.Windows.Forms.Label
        Me.picFillColor = New System.Windows.Forms.PictureBox
        Me.picStrokeColor = New System.Windows.Forms.PictureBox
        Me.pnContainer = New System.Windows.Forms.Panel
        Me.pnScrollArea = New System.Windows.Forms.Panel
        Me.pnToolBar = New System.Windows.Forms.Panel
        Me.btnPointer = New GDIPlusArchitect.ToolButton
        Me.btnHand = New GDIPlusArchitect.ToolButton
        Me.btnText = New GDIPlusArchitect.ToolButton
        Me.btnSquare = New GDIPlusArchitect.ToolButton
        Me.btnCircle = New GDIPlusArchitect.ToolButton
        Me.btnLine = New GDIPlusArchitect.ToolButton
        Me.btnField = New GDIPlusArchitect.ToolButton
        Me.btnPen = New GDIPlusArchitect.ToolButton
        Me.btnPlaceImage = New GDIPlusArchitect.ToolButton
        Me.btnFill = New GDIPlusArchitect.ToolButton
        Me.btnDropper = New GDIPlusArchitect.ToolButton
        Me.btnMagnify = New GDIPlusArchitect.ToolButton
        Me.btnScrollDown = New GDIPlusArchitect.ScrollButton
        Me.btnScrollUp = New GDIPlusArchitect.ScrollButton
        Me.pnContainer.SuspendLayout()
        Me.pnScrollArea.SuspendLayout()
        Me.pnToolBar.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblFill
        '
        Me.lblFill.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblFill.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.lblFill.Location = New System.Drawing.Point(8, 424)
        Me.lblFill.Name = "lblFill"
        Me.lblFill.Size = New System.Drawing.Size(24, 16)
        Me.lblFill.TabIndex = 5
        Me.lblFill.Text = "Fill"
        '
        'lblStroke
        '
        Me.lblStroke.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblStroke.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.lblStroke.Location = New System.Drawing.Point(8, 472)
        Me.lblStroke.Name = "lblStroke"
        Me.lblStroke.Size = New System.Drawing.Size(40, 16)
        Me.lblStroke.TabIndex = 6
        Me.lblStroke.Text = "Stroke"
        '
        'picFillColor
        '
        Me.picFillColor.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picFillColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picFillColor.Location = New System.Drawing.Point(8, 440)
        Me.picFillColor.Name = "picFillColor"
        Me.picFillColor.Size = New System.Drawing.Size(32, 24)
        Me.picFillColor.TabIndex = 21
        Me.picFillColor.TabStop = False
        '
        'picStrokeColor
        '
        Me.picStrokeColor.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.picStrokeColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picStrokeColor.Location = New System.Drawing.Point(8, 488)
        Me.picStrokeColor.Name = "picStrokeColor"
        Me.picStrokeColor.Size = New System.Drawing.Size(32, 24)
        Me.picStrokeColor.TabIndex = 22
        Me.picStrokeColor.TabStop = False
        '
        'pnContainer
        '
        Me.pnContainer.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnContainer.Controls.Add(Me.pnScrollArea)
        Me.pnContainer.Controls.Add(Me.btnScrollDown)
        Me.pnContainer.Controls.Add(Me.btnScrollUp)
        Me.pnContainer.Location = New System.Drawing.Point(0, 0)
        Me.pnContainer.Name = "pnContainer"
        Me.pnContainer.Size = New System.Drawing.Size(112, 424)
        Me.pnContainer.TabIndex = 41
        '
        'pnScrollArea
        '
        Me.pnScrollArea.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnScrollArea.Controls.Add(Me.pnToolBar)
        Me.pnScrollArea.Location = New System.Drawing.Point(0, 26)
        Me.pnScrollArea.Name = "pnScrollArea"
        Me.pnScrollArea.Size = New System.Drawing.Size(112, 310)
        Me.pnScrollArea.TabIndex = 41
        '
        'pnToolBar
        '
        Me.pnToolBar.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnToolBar.Controls.Add(Me.btnPointer)
        Me.pnToolBar.Controls.Add(Me.btnHand)
        Me.pnToolBar.Controls.Add(Me.btnText)
        Me.pnToolBar.Controls.Add(Me.btnSquare)
        Me.pnToolBar.Controls.Add(Me.btnCircle)
        Me.pnToolBar.Controls.Add(Me.btnLine)
        Me.pnToolBar.Controls.Add(Me.btnField)
        Me.pnToolBar.Controls.Add(Me.btnPen)
        Me.pnToolBar.Controls.Add(Me.btnPlaceImage)
        Me.pnToolBar.Controls.Add(Me.btnFill)
        Me.pnToolBar.Controls.Add(Me.btnDropper)
        Me.pnToolBar.Controls.Add(Me.btnMagnify)
        Me.pnToolBar.Location = New System.Drawing.Point(0, 0)
        Me.pnToolBar.Name = "pnToolBar"
        Me.pnToolBar.Size = New System.Drawing.Size(112, 312)
        Me.pnToolBar.TabIndex = 39
        '
        'btnPointer
        '
        Me.btnPointer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPointer.Caption = "Pointer"
        Me.btnPointer.ImageIndex = 0
        Me.btnPointer.ImageList = Nothing
        Me.btnPointer.Location = New System.Drawing.Point(0, 0)
        Me.btnPointer.Name = "btnPointer"
        Me.btnPointer.Selected = False
        Me.btnPointer.Size = New System.Drawing.Size(100, 26)
        Me.btnPointer.TabIndex = 27
        '
        'btnHand
        '
        Me.btnHand.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnHand.Caption = "Hand"
        Me.btnHand.ImageIndex = 0
        Me.btnHand.ImageList = Nothing
        Me.btnHand.Location = New System.Drawing.Point(0, 26)
        Me.btnHand.Name = "btnHand"
        Me.btnHand.Selected = False
        Me.btnHand.Size = New System.Drawing.Size(100, 26)
        Me.btnHand.TabIndex = 37
        '
        'btnText
        '
        Me.btnText.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnText.Caption = "Text"
        Me.btnText.ImageIndex = 0
        Me.btnText.ImageList = Nothing
        Me.btnText.Location = New System.Drawing.Point(0, 52)
        Me.btnText.Name = "btnText"
        Me.btnText.Selected = False
        Me.btnText.Size = New System.Drawing.Size(100, 26)
        Me.btnText.TabIndex = 26
        '
        'btnSquare
        '
        Me.btnSquare.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSquare.Caption = "Rectangle"
        Me.btnSquare.ImageIndex = 0
        Me.btnSquare.ImageList = Nothing
        Me.btnSquare.Location = New System.Drawing.Point(0, 78)
        Me.btnSquare.Name = "btnSquare"
        Me.btnSquare.Selected = False
        Me.btnSquare.Size = New System.Drawing.Size(100, 26)
        Me.btnSquare.TabIndex = 28
        '
        'btnCircle
        '
        Me.btnCircle.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCircle.Caption = "Ellipse"
        Me.btnCircle.ImageIndex = 0
        Me.btnCircle.ImageList = Nothing
        Me.btnCircle.Location = New System.Drawing.Point(0, 104)
        Me.btnCircle.Name = "btnCircle"
        Me.btnCircle.Selected = False
        Me.btnCircle.Size = New System.Drawing.Size(100, 26)
        Me.btnCircle.TabIndex = 36
        '
        'btnLine
        '
        Me.btnLine.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLine.Caption = "Line"
        Me.btnLine.ImageIndex = 0
        Me.btnLine.ImageList = Nothing
        Me.btnLine.Location = New System.Drawing.Point(0, 130)
        Me.btnLine.Name = "btnLine"
        Me.btnLine.Selected = False
        Me.btnLine.Size = New System.Drawing.Size(100, 26)
        Me.btnLine.TabIndex = 29
        '
        'btnField
        '
        Me.btnField.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnField.Caption = "Field"
        Me.btnField.ImageIndex = 0
        Me.btnField.ImageList = Nothing
        Me.btnField.Location = New System.Drawing.Point(0, 156)
        Me.btnField.Name = "btnField"
        Me.btnField.Selected = False
        Me.btnField.Size = New System.Drawing.Size(102, 26)
        Me.btnField.TabIndex = 35
        '
        'btnPen
        '
        Me.btnPen.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPen.Caption = "Pen"
        Me.btnPen.ImageIndex = 0
        Me.btnPen.ImageList = Nothing
        Me.btnPen.Location = New System.Drawing.Point(0, 182)
        Me.btnPen.Name = "btnPen"
        Me.btnPen.Selected = False
        Me.btnPen.Size = New System.Drawing.Size(100, 26)
        Me.btnPen.TabIndex = 34
        '
        'btnPlaceImage
        '
        Me.btnPlaceImage.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPlaceImage.Caption = "Image"
        Me.btnPlaceImage.ImageIndex = 0
        Me.btnPlaceImage.ImageList = Nothing
        Me.btnPlaceImage.Location = New System.Drawing.Point(0, 208)
        Me.btnPlaceImage.Name = "btnPlaceImage"
        Me.btnPlaceImage.Selected = False
        Me.btnPlaceImage.Size = New System.Drawing.Size(100, 26)
        Me.btnPlaceImage.TabIndex = 30
        '
        'btnFill
        '
        Me.btnFill.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFill.Caption = "Fill"
        Me.btnFill.ImageIndex = 0
        Me.btnFill.ImageList = Nothing
        Me.btnFill.Location = New System.Drawing.Point(0, 234)
        Me.btnFill.Name = "btnFill"
        Me.btnFill.Selected = False
        Me.btnFill.Size = New System.Drawing.Size(102, 26)
        Me.btnFill.TabIndex = 33
        '
        'btnDropper
        '
        Me.btnDropper.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDropper.Caption = "Dropper"
        Me.btnDropper.ImageIndex = 0
        Me.btnDropper.ImageList = Nothing
        Me.btnDropper.Location = New System.Drawing.Point(0, 260)
        Me.btnDropper.Name = "btnDropper"
        Me.btnDropper.Selected = False
        Me.btnDropper.Size = New System.Drawing.Size(100, 26)
        Me.btnDropper.TabIndex = 31
        '
        'btnMagnify
        '
        Me.btnMagnify.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnMagnify.Caption = "Magnify"
        Me.btnMagnify.ImageIndex = 0
        Me.btnMagnify.ImageList = Nothing
        Me.btnMagnify.Location = New System.Drawing.Point(0, 286)
        Me.btnMagnify.Name = "btnMagnify"
        Me.btnMagnify.Selected = False
        Me.btnMagnify.Size = New System.Drawing.Size(100, 26)
        Me.btnMagnify.TabIndex = 32
        '
        'btnScrollDown
        '
        Me.btnScrollDown.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnScrollDown.ImageIndex = 0
        Me.btnScrollDown.ImageList = Nothing
        Me.btnScrollDown.Location = New System.Drawing.Point(72, 394)
        Me.btnScrollDown.Name = "btnScrollDown"
        Me.btnScrollDown.Size = New System.Drawing.Size(26, 26)
        Me.btnScrollDown.TabIndex = 40
        '
        'btnScrollUp
        '
        Me.btnScrollUp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnScrollUp.ImageIndex = 0
        Me.btnScrollUp.ImageList = Nothing
        Me.btnScrollUp.Location = New System.Drawing.Point(74, 0)
        Me.btnScrollUp.Name = "btnScrollUp"
        Me.btnScrollUp.Size = New System.Drawing.Size(26, 26)
        Me.btnScrollUp.TabIndex = 39
        '
        'ToolboxWin
        '
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.Controls.Add(Me.lblFill)
        Me.Controls.Add(Me.lblStroke)
        Me.Controls.Add(Me.picFillColor)
        Me.Controls.Add(Me.picStrokeColor)
        Me.Controls.Add(Me.pnContainer)
        Me.Name = "ToolboxWin"
        Me.Size = New System.Drawing.Size(112, 528)
        Me.pnContainer.ResumeLayout(False)
        Me.pnScrollArea.ResumeLayout(False)
        Me.pnToolBar.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region



#Region "Refresh"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Update the tool button display and current mode based upon an explicit 
    ''' mode setting.  There are cases when other parts of the application, most 
    ''' notably the current drawing surface nede to change the selected tool such 
    ''' as when the user presses "ESC" during an operation to abort it.
    ''' </summary>
    ''' <param name="curMode"></param>
    ''' -----------------------------------------------------------------------------
    Public Sub refreshPanel()
        picFillColor.Invalidate()

        Select Case App.ToolManager.GlobalMode


            Case EnumTools.eCircle
                btnCircle.Selected = True
                ToggleToolMode(btnCircle)
            Case EnumTools.eDropper
                btnDropper.Selected = True
                ToggleToolMode(btnDropper)
            Case EnumTools.eField
                btnField.Selected = True
                ToggleToolMode(btnField)
            Case EnumTools.eFill
                btnFill.Selected = True
                ToggleToolMode(btnFill)
            Case EnumTools.eHand
                btnHand.Selected = True
                ToggleToolMode(btnHand)
            Case EnumTools.eLine
                btnLine.Selected = True
                ToggleToolMode(btnLine)
            Case EnumTools.eMagnify
                btnMagnify.Selected = True
                ToggleToolMode(btnMagnify)
            Case EnumTools.ePen
                btnPen.Selected = True
                ToggleToolMode(btnPen)
            Case EnumTools.ePlacing
                btnPlaceImage.Selected = True
                ToggleToolMode(btnPlaceImage)
            Case EnumTools.ePointer
                btnPointer.Selected = True
                ToggleToolMode(btnPointer)
            Case EnumTools.eSquare
                btnSquare.Selected = True
                ToggleToolMode(btnSquare)
            Case EnumTools.eText
                btnText.Selected = True
                ToggleToolMode(btnText)

        End Select
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Switches tool modes based upon which button was clicked.  Updates the 
    ''' application wide cursor to show the current most cursor.
    ''' </summary>
    ''' <param name="btn">The tool button clicked</param>
    ''' -----------------------------------------------------------------------------
    Protected Sub ToggleToolMode(ByVal btn As ToolButton)
        Trace.WriteLineIf(App.TraceOn, "Toolbox.Toggle.Button")
        Togglebuttons(btn)

        Select Case App.ToolManager.GlobalMode

            Case EnumTools.ePointer
                App.ToolManager.ToolCursor = Cursors.Arrow
            Case EnumTools.eText
                App.ToolManager.ToolCursor = Cursors.IBeam
            Case EnumTools.eMagnify
                App.ToolManager.ToolCursor = App.IconManager.CursorMagnify
            Case EnumTools.eDropper
                App.ToolManager.ToolCursor = App.IconManager.CursorDropper
            Case EnumTools.eFill
                App.ToolManager.ToolCursor = App.IconManager.CursorBucket
            Case EnumTools.eHand
                App.ToolManager.ToolCursor = App.IconManager.CursorHand
            Case Else
                App.ToolManager.ToolCursor = Cursors.Cross
        End Select
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Toggles the state of all buttons based upon a selected button.  Each button is 
    ''' expected to mark itself as unselected if it doesn't equal the incoming button 
    ''' reference
    ''' </summary>
    ''' <param name="s">The button that is selected</param>
    ''' <remarks>
    ''' This could also have been implement as:
    '''     For i As Int32 = 0 To pnToolBar.Controls.Count - 1
    '''          If TypeOf pnToolBar.Controls(i) Is ToolButton Then
    '''             Dim btn As ToolButton = DirectCast(pnToolBar.Controls(i), ToolButton)
    '''             btn.ToggleSelection(selButton)
    '''          End If
    '''      Next
    ''' 
    ''' But referencing the buttons explicitly seems more performant.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Sub Togglebuttons(ByVal selButton As ToolButton)
        btnCircle.ToggleSelection(selButton)
        btnDropper.ToggleSelection(selButton)
        btnField.ToggleSelection(selButton)
        btnFill.ToggleSelection(selButton)
        btnPlaceImage.ToggleSelection(selButton)
        btnLine.ToggleSelection(selButton)
        btnMagnify.ToggleSelection(selButton)
        btnPen.ToggleSelection(selButton)
        btnPointer.ToggleSelection(selButton)
        btnSquare.ToggleSelection(selButton)
        btnText.ToggleSelection(selButton)
        btnHand.ToggleSelection(selButton)
    End Sub

#End Region


#Region "Event Handlers"




    'Sets the current tool mode to the pointer tool
    Private Sub btnPointer_Selected(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPointer.ButtonSelected
        App.ToolManager.GlobalMode = EnumTools.ePointer
        ToggleToolMode(btnPointer)
    End Sub


    'Sets the current tool mode to the ellipse tool 
    Private Sub btnCircle_Selected(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCircle.ButtonSelected
        App.ToolManager.GlobalMode = EnumTools.eCircle
        ToggleToolMode(btnCircle)
    End Sub



    'Sets the current tool mode to the rectangle tool
    Private Sub btnSquare_Selected(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSquare.ButtonSelected
        App.ToolManager.GlobalMode = EnumTools.eSquare
        ToggleToolMode(btnSquare)
    End Sub

    'Initiates fill selection
    Private Sub picFillColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picFillColor.Click, lblFill.Click
        dgFillPicker.GO()
    End Sub

    'Initiates stroke selection
    Private Sub picStrokeColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picStrokeColor.Click, lblStroke.Click
        dgStrokePicker.GO()
    End Sub

    'Sets the current mode to the text tool
    Private Sub btnText_Selected(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnText.ButtonSelected
        App.ToolManager.GlobalMode = EnumTools.eText
        ToggleToolMode(btnText)
    End Sub

    'Sets the current mode to the line tool
    Private Sub btnLine_Selected(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLine.ButtonSelected
        App.ToolManager.GlobalMode = EnumTools.eLine
        ToggleToolMode(btnLine)

    End Sub

    'Sets the current mode to the pen tool
    Private Sub btnPen_Selected(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPen.ButtonSelected
        App.ToolManager.GlobalMode = EnumTools.ePen
        ToggleToolMode(btnPen)
    End Sub

    'Sets the current mode to the fill tool
    Private Sub btnFill_Selected(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFill.ButtonSelected
        App.ToolManager.GlobalMode = EnumTools.eFill
        ToggleToolMode(btnFill)
    End Sub

    'Sets the current mode to the dropper tool
    Private Sub Dropper_Selected(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDropper.ButtonSelected
        App.ToolManager.GlobalMode = EnumTools.eDropper
        ToggleToolMode(btnDropper)
    End Sub

    'Sets the current mode to the field tool 
    Private Sub btnField_Selected(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnField.ButtonSelected
        App.ToolManager.GlobalMode = EnumTools.eField
        ToggleToolMode(btnField)
    End Sub

    'Sets the mode to the magnify tool mode
    Private Sub btnMagnify_Selected(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMagnify.ButtonSelected
        App.ToolManager.GlobalMode = EnumTools.eMagnify
        ToggleToolMode(btnMagnify)
    End Sub


    'responds to changed in the size of the panel
    Private Sub pnScrollArea_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pnContainer.Resize
        If Me.Visible Then
            pnToolBar.Top = 0
            CheckScrollButtons()
        End If
    End Sub

    'Sets the tool mode to hand mode (for moving the surface around)
    Private Sub btnHand_Selected(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHand.ButtonSelected
        App.ToolManager.GlobalMode = EnumTools.eHand
        ToggleToolMode(btnHand)
    End Sub




#End Region


#Region "Stroke/Fill Paint"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the current fill to the picture box that displays the current fill.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub picFillColor_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles picFillColor.Paint
        Dim g As Graphics = e.Graphics

        If TypeOf Session.Fill Is GDIGradientFill Then
            Dim gFill As GDIGradientFill = DirectCast(Session.Fill, GDIGradientFill)
            Dim tempbrush As Drawing2D.LinearGradientBrush

            If gFill.GradientMode = GDIGradientFill.EnumGradientMode.Custom Then
                tempbrush = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, picFillColor.Width, picFillColor.Height), gFill.Color1, gFill.Color2, gFill.Angle)
            Else
                tempbrush = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, picFillColor.Width, picFillColor.Height), gFill.Color1, gFill.Color2, CType(gFill.GradientMode, Drawing2D.LinearGradientMode))
            End If

            g.FillRectangle(tempbrush, 0, 0, picFillColor.Width, picFillColor.Height)
            tempbrush.Dispose()
            gFill = Nothing
        Else
            g.FillRectangle(Session.Fill.Brush, 0, 0, picFillColor.Width, picFillColor.Height)
        End If


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the color of the current stroke to the stroke color box.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub picStrokeColor_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles picStrokeColor.Paint
        Dim g As Graphics = e.Graphics
        Dim b As Brush = New SolidBrush(Session.Stroke.Color)
        g.FillRectangle(b, 0, 0, picFillColor.Width, picFillColor.Height)
    End Sub


#End Region


#Region "Scroll Related Members"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Checks which scroll buttons are necessary when visibility changes.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub ToolBoxWin_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.VisibleChanged
        If Me.Visible Then
            CheckScrollButtons()
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Scrolls the tool window up as necessary and checks the scroll buttons.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub scrollUp()
        pnToolBar.Top += CONST_LIST_SCROLL_INCREMENT
        CheckScrollButtons()
    End Sub

    'Scrolls the tool window up 
    Private Sub btnScrollUp_Scroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnScrollUp.Scroll
        scrollUp()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Scrolls the tool window down if necessary and then checks which buttons should 
    ''' be enabled
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub scrollDown()
        pnToolBar.Top -= CONST_LIST_SCROLL_INCREMENT
        CheckScrollButtons()
    End Sub

    'Scrolls down
    Private Sub btnScrollDown_Scroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnScrollDown.Scroll
        scrollDown()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to wheel, scrolling the tool window view if appropriate.
    ''' </summary>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseWheel(ByVal e As System.Windows.Forms.MouseEventArgs)
        If e.Delta > 0 Then
            If btnScrollUp.Enabled = True Then
                scrollUp()
            End If
        Else
            If btnScrollDown.Enabled = True Then
                scrollDown()
            End If
        End If

        MyBase.OnMouseWheel(e)

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines which scroll buttons should be enabled after a scroll.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub CheckScrollButtons()

        If pnToolBar.Top < 0 Then
            btnScrollUp.Enabled = True
        Else
            btnScrollUp.Enabled = False
        End If

        If pnScrollArea.Height < CONST_LIST_PANEL_HEIGHT + pnToolBar.Top Then
            btnScrollDown.Enabled = True
        Else
            btnScrollDown.Enabled = False
        End If
    End Sub

#End Region


#Region "Image Import"


    Private Sub btnPlaceImage_Selected(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlaceImage.ButtonSelected
        Dim imageSource As String = promptForImage()
        ToggleToolMode(btnPointer)

        If Not imageSource = String.Empty Then


            App.ToolManager.ToolCursor = Cursors.Cross
            App.ToolManager.ArmedImageSrc = imageSource
            App.ToolManager.GlobalMode = EnumTools.ePlacing
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Prompts for an image source using the OpenFileDialog.  Returns the image source 
    ''' selected by the user or an empty string if no source was selected
    ''' </summary>
    ''' <returns>An image source selected by the user.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function promptForImage() As String
        Dim dgImgOpen As New OpenFileDialog

        dgImgOpen.Filter = GDIImage.CONST_IMG_FILTER

        Dim iresp As DialogResult = dgImgOpen.ShowDialog()

        Try

            If iresp = DialogResult.OK Then
                Return dgImgOpen.FileName
            Else
                Return String.Empty
            End If
        Catch ex As Exception
            MsgBox("Error accessing resource")
            Return String.Empty
        Finally
            dgImgOpen.Dispose()
        End Try

    End Function
#End Region



    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            GDIObjects.Session.removeStrokeCallback(_CBStrokeChanged)
            GDIObjects.Session.removeFillCallback(_CBFillChanged)
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class
