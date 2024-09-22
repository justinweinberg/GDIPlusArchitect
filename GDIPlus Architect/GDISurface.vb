Imports GDIObjects
Imports System.Drawing
Imports System.Drawing.Drawing2D



''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : GDISurface
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' The most important class in the GDIPlus Architect portion of the solution.  This class
''' is responsible for rendering a GDIDocument, responding to application wide state 
''' changes and redrawing documents, handling the current tool settings, and responding to 
''' mouse actions appropriately. 
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class GDISurface
    Inherits System.Windows.Forms.UserControl

#Region "Document specific callback declarations"
    Private _CBFillChanged As New GDIObjects.Session.FillChanged(AddressOf OnFillChanged)
    Private _CBStrokeChanged As New GDIObjects.Session.StrokeChanged(AddressOf OnStrokeChanged)
#End Region


#Region "Type Declarations"

    '''<summary>Width of the minor grid pen</summary>
    Private Const CONST_MINORGRID_WIDTH As Int32 = 1
    '''<summary>Width of the grid pen</summary>
    Private Const CONST_GRID_WIDTH As Int32 = 1
    '''<summary>Width of the margins</summary>
    Private Const CONST_MARGIN_WIDTH As Int32 = 1


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Denotes what should happen when the user scrolls the wheel.
    ''' </summary>
    ''' <remarks>Every time the user clicks the middle mouse button, the wheel action
    ''' toggles between zoom and scroll.  This enumeration tracks the current state 
    ''' of what action the wheel should take.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Enum EnumWheelAction As Integer
        ''' <summary>The document will zoom in and out on wheel actions</summary>
        eZoomDocument = 0
        ''' <summary>The document will scroll on wheel actions</summary>
        eScrollDocument = 1
    End Enum

#End Region


#Region "Event Declarations"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Raised when the current zoom of the document changes
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Event ZoomChanged(ByVal s As Object, ByVal e As EventArgs)

#End Region

#Region "Local Fields"



    '''<summary>The current zoom</summary>
    Private _Zoom As Single = 1.0!

    '''<summary>X Dots per inch.  Established from the System.Drawing.Graphics 
    ''' context passed on paint.</summary>
    Private _XDPI As Single = 96.0!

    '''<summary>Y Dots per inch.  Established from the System.Drawing.Graphics
    '''  context passed on paint.</summary>
    Private _YDPI As Single = 96.0!


    '''<summary>GDIContextMenu object to display upon a Right click on
    '''  the surface</summary>
    Private _ContextMenu As SurfaceContextMenu

    '''<summary>The mouse button being held down</summary>
    Private _ButtonDown As MouseButtons = MouseButtons.None


    '''<summary>The object on the surface that the mouse is hovering over,
    '''  if any</summary>
    Private _MouseOverObject As GDIObject = Nothing




    '''<summary>A reference to the document that this surface is displaying</summary>
    Private _Document As GDIDocument

    '''<summary>The action that should be taken upon moving the mouse wheel. 
    ''' The default is to zoom, but upon a middle mouse click the action toggles 
    ''' to a vertical scroll.</summary>
    Private _WheelAction As EnumWheelAction = EnumWheelAction.eZoomDocument


    '''<summary>For hand tool operations.  Denotes where a hand operation
    '''  began and the last point recorded for the hand tool</summary>
    Private _HandToolPoint As Point




    '''<summary>Last known mouse cursor position.  This is the cursor position
    '''  without any transforms</summary>
    Private _SurfacePoint As Point



    '''<summary>The mouse coordinates expressed in the object's coordinate 
    ''' space as a Point structure The pointF variant is used primarily for 
    ''' hit testing where a number of division operations are performed.</summary>
    Private _ObjectPointF As PointF
    '''<summary>The mouse coordinates expressed in the object's coordinate 
    ''' space as a Point structure</summary>
    Private _ObjectPoint As Point

    '''<summary>Point in object coordinate space snapped to grid, margins, 
    ''' and/or guides depending on settings under options.</summary>
    Private _ObjectPointSnapped As Point


    '''<summary>Whether the shift key is held down or not </summary>
    Private _ShiftDown As Boolean = False

    '''<summary>Whether the control key is held down or not </summary>
    Private _ControlDown As Boolean = False


#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a GDISurface.  Sets up the properties typically 
    ''' set in InitializeComponent and sets various control styles (DoubleBuffer, 
    ''' UserPaint, ResizeRedraw)
    ''' </summary>
    ''' <param name="doc">The document being rendered to hte surface.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal doc As GDIDocument)
        MyBase.New()


        Trace.WriteLineIf(App.TraceOn, "GDISurface.New")

        SetStyle(ControlStyles.AllPaintingInWmPaint, True)
        SetStyle(ControlStyles.DoubleBuffer, True)
        SetStyle(ControlStyles.UserPaint, True)
        SetStyle(ControlStyles.ResizeRedraw, True)

        _Document = doc

        _ContextMenu = New SurfaceContextMenu

        With Me
            .AutoScroll = True
            .AutoScrollMinSize = New System.Drawing.Size(384, 528)
            .BackColor = System.Drawing.SystemColors.Window
            .Dock = System.Windows.Forms.DockStyle.Fill
            .Location = New System.Drawing.Point(0, 0)
            .Name = "GDISurface"
            .Size = New System.Drawing.Size(400, 541)
            .TabIndex = 0
        End With
 
        GDIObjects.Session.setFillChangedCallBack(_CBFillChanged)
        GDIObjects.Session.setStrokeChangedCallBack(_CBStrokeChanged)



        Me.Name = "GraphicsSurface"


    End Sub
#End Region

#Region "Property Accessors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the GDIDocument that is rendered on the surface.
    ''' </summary>
    ''' <value>A GDIDocument to render to the surface.</value>
    ''' -----------------------------------------------------------------------------
    Public Property Document() As GDIDocument
        Get
            Return _Document
        End Get
        Set(ByVal Value As GDIDocument)
            _Document = Value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the Zoom Factor of the window.  This is how much the view is zoomed.
    ''' </summary>
    ''' <value>A single (float) indicating the percent zoomed in or out.</value>
    ''' <remarks>Raises the Zoomchanged event on a property Set.</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property Zoom() As Single
        Get
            Return _Zoom
        End Get
        Set(ByVal Value As Single)
            If Value > 0.05 Then
                _Zoom = Value
                RaiseEvent ZoomChanged(Me, EventArgs.Empty)
            End If

            Invalidate()
        End Set
    End Property


#End Region



#Region "Delegate Handlers"



    Protected Sub OnFillChanged(ByVal s As Object, ByVal e As EventArgs)
        _Document.Selected.setFill(Session.Fill)
        Invalidate()
    End Sub

    Protected Sub OnStrokeChanged(ByVal s As Object, ByVal e As EventArgs)
        _Document.Selected.setStroke(Session.Stroke)
        Invalidate()
    End Sub

#End Region

#Region "Guide Manipulation"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Checks whether a the mouse is over a guide or not
    ''' </summary>
    ''' <param name="ptObject"></param>
    ''' <param name="btn"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub hitTestGuides()
        For Each guide As GDIGuide In _Document.Guides
            If guide.HitTest(_ObjectPoint) Then
                App.ToolManager.SelectedGuide = guide
                Invalidate()
                Return
            End If

        Next

    End Sub


    Public Sub InsertHGuide()
        _Document.InsertGuide(GDIGuide.EnumGuideDirection.eHoriz, _ObjectPoint.Y)
        Invalidate()
    End Sub


    Public Sub InsertVGuide()
        _Document.InsertGuide(GDIGuide.EnumGuideDirection.eVert, _ObjectPoint.X)
        Invalidate()
    End Sub

#End Region


#Region "Cleanup"



    'UserControl1 overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then


            GDIObjects.Session.removeFillCallback(_CBFillChanged)
            GDIObjects.Session.removeStrokeCallback(_CBStrokeChanged)
            _CBFillChanged = Nothing
            _CBStrokeChanged = Nothing
            _ContextMenu = Nothing
            Me._Document.Dispose()
            Me._Document = Nothing

        End If

        MyBase.Dispose(disposing)
    End Sub
#End Region

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the surface's cursor based upon a hit handle.
    ''' </summary>
    ''' <param name="objectRotation">Amount of rotation of the hit object</param>
    ''' <param name="eHitHandle">The handle on the object that was hit.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub updateCursor(ByVal objectRotation As Single, ByVal eHitHandle As EnumDragHandles)

        If objectRotation > 0 Then
            Me.Cursor = Cursors.Cross
        Else
            Select Case eHitHandle
                Case EnumDragHandles.eBottomleft
                    Me.Cursor = Cursors.SizeNESW

                Case EnumDragHandles.eBottomRight
                    Me.Cursor = Cursors.SizeNWSE

                Case EnumDragHandles.eTopLeft
                    Me.Cursor = Cursors.SizeNWSE

                Case EnumDragHandles.eTopRight
                    Me.Cursor = Cursors.SizeNESW

                Case EnumDragHandles.eBottom
                    Me.Cursor = Cursors.SizeNS

                Case EnumDragHandles.eLeft
                    Me.Cursor = Cursors.SizeWE

                Case EnumDragHandles.eTop
                    Me.Cursor = Cursors.SizeNS

                Case EnumDragHandles.eRight
                    Me.Cursor = Cursors.SizeWE

                Case Else
                    Me.Cursor = Cursors.Cross

            End Select

        End If

    End Sub

#Region "Hit Testing"




    Private Sub checkDragState()
        If _ButtonDown = MouseButtons.Left Then
            If _Document.Selected.Count > 0 Then

                 App.ToolManager.ActiveToolMode = App.ToolManager.TM.eDragPending
             Else

                App.ToolManager.ActiveToolMode = App.ToolManager.TM.eBoundPending

            End If
        End If

    End Sub


    Private Sub hitTestObjects()

        'Get an object as the specified point 
        Dim hitObject As GDIObject = _
        _Document.CurrentPage.GDIObjects.FindObjectAtPoint(_ObjectPointF, _Zoom)

        'Let the selected set react to the hit object.  The reason for this is that GDI+ Architect 
        'allows for multiple selected objects and depending on the state of the shift key, the selected 
        'set will either include or exclude objects.
        _Document.Selected.handleNewSelection(hitObject, _ShiftDown)
        checkDragState()

        Invalidate()

    End Sub


#End Region


    Private Sub handleRightClick()
        Trace.WriteLineIf(App.TraceOn, "GDISurface.rightclick")

        If App.ToolManager.ToolInUse = False Then
            hitTestObjects()

            hitTestGuides()

            _ContextMenu.show(Me.PointToScreen(_SurfacePoint))

        End If
    End Sub



    Private Sub handleDropper()
        Dim SelObj As GDIObject = _Document.CurrentPage.GDIObjects.FindObjectAtPoint(_ObjectPointF, _Zoom)

        If Not SelObj Is Nothing Then
            If TypeOf SelObj Is GDIFilledShape Then

                Session.Fill = DirectCast(SelObj, GDIFilledShape).Fill

            ElseIf TypeOf SelObj Is GDIImage Then
                Dim tempcolor As Color = DirectCast(SelObj, GDIImage).ColorAtPoint(_ObjectPoint)
                Session.Fill = New GDISolidFill(tempcolor)

            End If


        End If


    End Sub







    Private Sub handlePointer()
        If App.ToolManager.ActiveToolMode = App.ToolManager.TM.eClickInTextBox Then
            If Not App.ToolManager.ClickinTextbox(_ObjectPointF, _Zoom) Then
                App.ToolManager.EndTool(_ShiftDown)
                App.ToolManager.ActiveToolMode = App.ToolManager.TM.eNone
            End If
        Else

            If Not App.ToolManager.SelectedDragHandle = EnumDragHandles.eNone Then
                App.ToolManager.ActiveToolMode = App.ToolManager.TM.eDraggingHandle
                App.ToolManager.BeginDragHandle(_MouseOverObject, App.ToolManager.SelectedDragHandle, _ObjectPointSnapped)
            Else
                hitTestObjects()
            End If
        End If
    End Sub




  

    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        updatePoints(e.X, e.Y)

        _ButtonDown = e.Button

        Select Case e.Button
            Case MouseButtons.Left
                handleToolDown()
            Case MouseButtons.Middle
                handleMiddleClick()
            Case MouseButtons.Right
                handleRightClick()
        End Select

        MyBase.OnMouseDown(e)
    End Sub

    Private Sub handleToolDown()

        App.ToolManager.handleToolDown(_ObjectPointSnapped, _ObjectPointF, _Zoom, _ShiftDown)


        Select Case App.ToolManager.GlobalMode

            Case EnumTools.eMagnify
                magnifySelection()
                Invalidate()

            Case EnumTools.eHand
                App.ToolManager.ActiveToolMode = App.ToolManager.TM.eHanding
                _HandToolPoint = _SurfacePoint

            Case EnumTools.eFill
                hitTestObjects()
                _Document.Selected.setFill(Session.Fill)
                Invalidate()

            Case EnumTools.eDropper
                handleDropper()

            Case EnumTools.ePointer
                handlePointer()
        End Select
    End Sub

    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        Trace.WriteLineIf(App.TraceOn, "GDISurface.mouseup")

        updatePoints(e.X, e.Y)
        _ButtonDown = MouseButtons.None

        App.ToolManager.OnToolUp(_ShiftDown, _ObjectPoint)

        Invalidate()


        MyBase.OnMouseUp(e)
    End Sub


#Region "Zoom and Mouse Wheel"

    Private Sub magnifySelection()
        Trace.WriteLineIf(App.TraceOn, "GDISurface.MagnifySelection")

        Dim fZoomlast As Single = _Zoom
        'Converts mGDIDOCParent.zoom to a percentage of 100
        Dim fzoomperc As Single = _Zoom * 100

        'for each detent detected, mGDIDOCParent.zoom a factor of 2 
        If fzoomperc = 75 Then
            If _ControlDown Then
                fzoomperc = 50.0!
            Else
                fzoomperc = 100.0!
            End If
        ElseIf fzoomperc = 100 Then
            If _ControlDown Then
                fzoomperc = 75.0!
            Else
                fzoomperc = 150.0!

            End If
        ElseIf fzoomperc < 100 Then
            If _ControlDown Then
                fzoomperc = CSng(fzoomperc / 1.5)
            Else
                fzoomperc = CSng(fzoomperc * 1.5)
            End If
        Else
            If _ControlDown Then
                fzoomperc = fzoomperc - 50
            Else
                fzoomperc = fzoomperc + 50
            End If
        End If

        '1) Define current center of new view
        '2) Find the point difference between the center view and the clicked point
        '3) set autoscroll based on the difference 

        '1) xcenter = autoscroll.x + clientrectangle.width / 2 
        '2) xcenter - 
        Me.Zoom = fzoomperc / 100


        Dim ptTranslated As Point = New Point(Math.Abs(AutoScrollPosition.X), Math.Abs(AutoScrollPosition.Y))


        Dim xChange As Single = _SurfacePoint.X - ClientRectangle.Width \ 2
        Dim yChange As Single = _SurfacePoint.Y - ClientRectangle.Height \ 2

        Dim zcurrent As Single = _Zoom

        Dim ptOffset As PointF = New PointF(ptTranslated.X + xChange, ptTranslated.Y + yChange)
        'update autoscroll to reflect the center view
        '  AutoScrollPosition = New Point(CInt(ptTranslated.X + (xChange * zRatio)), CInt(ptTranslated.Y + (yChange * zRatio)))
        Dim ptEquiv As Point = New Point(CInt((ptOffset.X * zcurrent) / fZoomlast), CInt((ptOffset.Y * zcurrent) / fZoomlast))

        AutoScrollPosition = New Point(ptEquiv.X, ptEquiv.Y)

        Invalidate()
    End Sub



    Private Sub handleMiddleClick()
        Trace.WriteLineIf(App.TraceOn, "GDISurface.handlemiddleclick")

        Select Case _WheelAction
            Case EnumWheelAction.eZoomDocument
                _WheelAction = EnumWheelAction.eScrollDocument
            Case EnumWheelAction.eScrollDocument
                _WheelAction = EnumWheelAction.eZoomDocument
        End Select
    End Sub

    Protected Overrides Sub OnMouseWheel(ByVal e As System.Windows.Forms.MouseEventArgs)
        Trace.WriteLineIf(App.TraceOn, "GDISurface.mousewheel")

        If Not e.Delta = 0 Then

            Dim detents As Int32 = Math.Abs(e.Delta \ 120)

            If _WheelAction = EnumWheelAction.eZoomDocument Then
                'negative is a "mouse up" - mGDIDOCParent.zoom in 
                Dim bZoomOut As Boolean = e.Delta < 0
                Dim fzoomperc As Single = _Zoom * 100

                'for each detent detected, mGDIDOCParent.zoom a factor of 2 
                For i As Int32 = 1 To detents
                    If fzoomperc > 50 AndAlso fzoomperc <= 75 Then
                        If bZoomOut Then
                            fzoomperc = 50.0!
                        Else
                            fzoomperc = 100.0!
                        End If
                    ElseIf (fzoomperc > 75 AndAlso fzoomperc <= 100) OrElse (fzoomperc >= 100 AndAlso fzoomperc < 150) Then
                        If bZoomOut Then
                            fzoomperc = 75.0!
                        Else
                            fzoomperc = 150.0!

                        End If
                    ElseIf fzoomperc < 100 Then
                        If bZoomOut Then
                            fzoomperc = CSng(fzoomperc / 1.5)
                        Else
                            fzoomperc = CSng(fzoomperc * 1.5)
                        End If
                    Else
                        If bZoomOut Then
                            fzoomperc = fzoomperc - 50
                        Else
                            fzoomperc = fzoomperc + 50
                        End If
                    End If
                Next

                'He converts back to a numeric value here 
                Me.Zoom = fzoomperc / 100

            Else
                MyBase.OnMouseWheel(e)
            End If

        End If

        Invalidate()

    End Sub
#End Region


#Region "Painting"




    Private Sub cacheDPI(ByVal g As Graphics)
        _XDPI = g.DpiX
        _YDPI = g.DpiY
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets up the graphics mode for painting to the surface.  Sets the smoothingmode 
    ''' and texthint based off of he document settings and then trasnforms the 
    ''' document to match the viewport of the surface.
    ''' </summary>
    ''' <param name="g">A s</param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Justin]	1/10/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Sub setGraphicsMode(ByVal g As Graphics)

        With g
            .SmoothingMode = _Document.SmoothingMode
            .TextRenderingHint = _Document.TextRenderingHint

            .TranslateTransform(Me.AutoScrollPosition.X, Me.AutoScrollPosition.Y)
        End With
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Paints the background area of the surface.
    ''' </summary>
    ''' <param name="g">Graphics context to draw against</param>
    ''' <param name="rectPageBounds">The printable area (the page bounds)</param>
    ''' -----------------------------------------------------------------------------
    Private Sub paintBackGround(ByVal g As Graphics, ByVal rectPageBounds As Rectangle)
        Dim bBackbrush As Brush = New SolidBrush(_Document.BackColor)

        'Clear the entire surface with the non print area color 
        g.Clear(App.Options.ColorNonPrintArea)

        g.FillRectangle(bBackbrush, rectPageBounds)



        bBackbrush.Dispose()

    End Sub




    Private Sub paintMinorGrid(ByVal g As Graphics, ByVal rectpageBounds As Rectangle)
        'Converts the grid line size from 100ths of an inch, to inches, to pixels, and multiplies by the mGDIDOCParent.zoom
        Dim iXStep As Single = App.Options.MinorGridSize * _Zoom
        Dim iYStep As Single = App.Options.MinorGridSize * _Zoom

        Dim pMinorGrid As New Pen(App.Options.ColorMinorGrid, CONST_MINORGRID_WIDTH)

        For i As Single = 0 To rectpageBounds.Height Step iYStep
            g.DrawLine(pMinorGrid, 0, i, rectpageBounds.Width, i)
        Next

        For i As Single = 0 To rectpageBounds.Width Step iXStep
            g.DrawLine(pMinorGrid, i, 0, i, rectpageBounds.Height)
        Next

        pMinorGrid.Dispose()

    End Sub


    Private Sub paintGrid(ByVal g As Graphics, ByVal rectPageBounds As Rectangle)
        'Converts the grid line size from 100ths of an inch, to inches, to pixels, and multiplies by the mGDIDOCParent.zoom
        Dim iXStep As Single = App.Options.MajorGridSize * _Zoom
        Dim iYStep As Single = App.Options.MajorGridSize * _Zoom

        Dim pGrid As New Pen(App.Options.ColorMajorGrid, CONST_GRID_WIDTH)

        For i As Single = 0 To rectPageBounds.Height Step iYStep
            g.DrawLine(pGrid, 0, i, rectPageBounds.Width, i)
        Next

        For i As Single = 0 To rectPageBounds.Width Step iXStep
            g.DrawLine(pGrid, i, 0, i, rectPageBounds.Height)
        Next

        pGrid.Dispose()

    End Sub

    Private Sub paintMargins(ByVal g As Graphics)
        Dim rectMargins As Rectangle = RectToObjectCoordinates(_Document.PrintableArea)
        Dim pMargin As New Pen(App.Options.ColorMargin, CONST_MARGIN_WIDTH)

        pMargin.DashStyle = Drawing2D.DashStyle.Dash
        g.DrawRectangle(pMargin, rectMargins)

        pMargin.Dispose()

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders margins, grids and guides as appropriate based on application settings
    ''' </summary>
    ''' <param name="g">Graphics context to draw against</param>
    ''' <param name="rectPageBounds">The entire bounds of the page in object coordinates.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub paintPageDecorations(ByVal g As Graphics, ByVal rectPageBounds As Rectangle)
        'Draw the minor grid
        If App.Options.ShowMinorGrid Then
            paintMinorGrid(g, rectPageBounds)
        End If

        'Draw the major grid
        If App.Options.ShowGrid Then
            paintGrid(g, rectPageBounds)
        End If


        'Draw the margins for PrintDocument type documents
        If App.Options.ShowMargins AndAlso _Document.ExportSettings.DocumentType = EnumDocumentTypes.ePrintDocument Then
            paintMargins(g)
        End If

        'Draw Guides
        If App.Options.ShowGuides Then
            _Document.DrawGuides(g, _Zoom)
        End If


    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders selection rectangles and mouse over highlights as appropriate based 
    ''' on user settings, selected objects in the document, and the mouse state.
    ''' </summary>
    ''' <param name="g">Graphics context to draw against</param
    ''' -----------------------------------------------------------------------------
    Private Sub paintHighlights(ByVal g As Graphics)
        If Not App.ToolManager.ActiveToolMode = App.ToolManager.TM.eDragging Then

            _Document.Selected.DrawSelected(g, _Zoom)

            If Not _MouseOverObject Is Nothing AndAlso _
            _Document.Selected.Contains(_MouseOverObject) = False _
            AndAlso App.Options.ShowMouseOverHandles Then

                _Document.CurrentPage.GDIObjects.highlightObject(g, _MouseOverObject, _Zoom, App.Options.ColorMouseOver)
            End If

        End If
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        Dim g As Graphics = e.Graphics

        'Converts the Instance size in 100ths of an inch to inches, then to pixels, and multiplies by the  zoom factor
        Dim rectPageBounds As Rectangle = RectToObjectCoordinates(_Document.RectPageSize)

        'Setup scrolling
        If Not AutoScrollMinSize.Height = rectPageBounds.Height AndAlso _
        Not AutoScrollMinSize.Width = rectPageBounds.Width Then

            AutoScrollMinSize = New Size(rectPageBounds.Width, rectPageBounds.Height)
        End If

        'Set the graphic mode.
        setGraphicsMode(g)

        'Get the DPI from the graphics context and save it for later.
        cacheDPI(g)

        'Fill the background of the surface (render the print area and non print area)
        paintBackGround(g, rectPageBounds)

        'Draws margins, guides, and grids onto the surface
        paintPageDecorations(g, rectPageBounds)


        'Ask each object on the current page to draw itself to the surface
        With _Document.CurrentPage.GDIObjects
            .DrawObjects(g, _Zoom, EnumDrawMode.eNormal)
        End With

        'Let the toolmanager do any tool specific drawing.
        App.ToolManager.draw(g, _Zoom)

        'Draws highlights and selection rectangles
        paintHighlights(g)


        'Call the base paint method
        MyBase.OnPaint(e)
    End Sub

#End Region

#Region "Hand tool"




    Private Sub updatehand(ByVal surfacePoint As Point)
        Trace.WriteLineIf(App.TraceOn, "GDISurface.updateHand")


        Dim ChangeX As Single = surfacePoint.X - _HandToolPoint.X
        Dim ChangeY As Single = surfacePoint.Y - _HandToolPoint.Y

        Dim ptTranslated As Point = New Point(Math.Abs(AutoScrollPosition.X), Math.Abs(AutoScrollPosition.Y))
        Me.AutoScrollPosition = New Point(ptTranslated.X - CInt(ChangeX * _Zoom), ptTranslated.Y - CInt(ChangeY * _Zoom))

        _HandToolPoint.X = surfacePoint.X
        _HandToolPoint.Y = surfacePoint.Y
    End Sub
#End Region

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the mouse move event for when the mouse moves over the surface.
    ''' </summary>
    ''' <param name="e">Standard EventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)

        updatePoints(e.X, e.Y)

        Select Case App.ToolManager.ActiveToolMode

            Case App.ToolManager.TM.eHanding
                updatehand(_SurfacePoint)

            Case App.ToolManager.TM.eTextBegin
                App.ToolManager.BeginDefineTextbox(_ObjectPointSnapped)
                App.ToolManager.ActiveToolMode = App.ToolManager.TM.eTextDefinebox

            Case App.ToolManager.TM.eTextDefinebox
                App.ToolManager.UpdateTool(_ObjectPointSnapped, MouseButtons.Left, _ShiftDown, _Zoom)
                Invalidate()

            Case App.ToolManager.TM.eClickInTextBox
                If e.Button = MouseButtons.Left Then
                    App.ToolManager.AttemptTextSelection(_ObjectPointF)
                End If

            Case App.ToolManager.TM.eMovingGuide
                If App.ToolManager.SelectedGuide.Direction = GDIGuide.EnumGuideDirection.eHoriz Then
                    App.ToolManager.SelectedGuide.XY = _ObjectPoint.Y
                Else
                    App.ToolManager.SelectedGuide.XY = _ObjectPoint.X
                End If
                Invalidate()

            Case App.ToolManager.TM.eDragPending
                App.ToolManager.BeginDragging(_ObjectPoint)

            Case App.ToolManager.TM.eBoundPending
                App.ToolManager.BeginBounding(_ObjectPoint)



            Case App.ToolManager.TM.eFielding, App.ToolManager.TM.ePlacing, _
            App.ToolManager.TM.eLining, App.ToolManager.TM.eDrawing, App.ToolManager.TM.ePenning, _
            App.ToolManager.TM.eDraggingHandle

                App.ToolManager.UpdateTool(_ObjectPointSnapped, _ButtonDown, _ShiftDown, _Zoom)
                Invalidate()

            Case App.ToolManager.TM.eDragging, App.ToolManager.TM.eBounding
                App.ToolManager.UpdateTool(_ObjectPoint, _ButtonDown, _ShiftDown, _Zoom)
                Invalidate()

            Case Else

                If App.ToolManager.GlobalMode = EnumTools.ePointer Then


                    Dim SelObj As GDIObject = _Document.CurrentPage.GDIObjects.FindObjectAtPoint(_ObjectPointF, _Zoom)

                    If Not SelObj Is Nothing Then

                        _MouseOverObject = SelObj

                        Dim eHitHandle As EnumDragHandles = _
                        SelObj.HitTestHandles(_ObjectPointF, _Zoom)

                        If Not eHitHandle = EnumDragHandles.eNone Then

                            updateCursor(SelObj.Rotation, eHitHandle)
                            App.ToolManager.SelectedDragHandle = eHitHandle

                        Else
                            App.ToolManager.SelectedDragHandle = EnumDragHandles.eNone
                            Me.Cursor = Cursors.Default
                        End If


                    Else
                        _MouseOverObject = Nothing
                        App.ToolManager.SelectedDragHandle = EnumDragHandles.eNone
                        Me.Cursor = Cursors.Default


                    End If
                    Me.Invalidate()
                End If

        End Select

        MyBase.OnMouseMove(e)
    End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Lets the tool manager handle lost focus on the surface
    ''' </summary>
    ''' <param name="e">Standard EventArgs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub GDISurface_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.LostFocus
        Trace.WriteLineIf(App.TraceOn, "GDISurface.Unfocus")

        _ShiftDown = False
        _ControlDown = False

        App.ToolManager.handleLostFocus()


    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the current cursor based on the tool in use
    ''' </summary>
    ''' <param name="e">Standard EventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseEnter(ByVal e As System.EventArgs)
        Me.Cursor = App.ToolManager.ToolCursor
        MyBase.OnMouseEnter(e)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Lets the tool manager handle a double click on the surface
    ''' </summary>
    ''' <param name="e">Standard EventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        App.ToolManager.handleDoubleClick(_ObjectPoint, _ObjectPointF)
    End Sub



#Region "Coordinate Translations Methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Takes mouse coordinates and creates a series of points used in different 
    ''' places within the application.
    ''' </summary>
    ''' <param name="x">X coordinate of the mouse position</param>
    ''' <param name="y">Y coordinate of the mouse position</param>
    ''' -----------------------------------------------------------------------------
    Private Sub updatePoints(ByVal x As Int32, ByVal y As Int32)

        'The surface point is the mouse coordinates
        _SurfacePoint = New Point(x, y)

        'The object point is the mouse coordinate translates to the object coordinate space
        _ObjectPoint = surfaceToObjCoord(_SurfacePoint)
        'Same as the object point, but a pointF
        _ObjectPointF = surfaceToObjCoord(New PointF(x, y))
        'Object point snapped to margins, grid, etc based on snap settings
        _ObjectPointSnapped = snappedObjectPoint()
    End Sub


     Private Function RectToObjectCoordinates(ByVal rect As Rectangle) As Rectangle
        'Converts to inches from inches * 100

        Dim x As Int32 = CInt((rect.X / _XDPI) * _XDPI * _Zoom)
        Dim y As Int32 = CInt((rect.Y / _YDPI) * _YDPI * _Zoom)
        Dim width As Int32 = CInt((rect.Width / _XDPI) * _XDPI * _Zoom)
        Dim height As Int32 = CInt((rect.Height / _YDPI) * _YDPI * _Zoom)

        'Multiple all units by the dots per inch.  This converts the rectangle from inches to pixels

        'Additionally, multiply the rectangle by the mGDIDOCParent.zoom factor.  if it's < 1, it will shrink our rectangle.  
        'If it Then 's > 1 it will expand it giving it a mGDIDOCParent.zoomed in look
        Return New Rectangle(x, y, width, height)
    End Function

     Public Overloads Function surfaceToObjCoord(ByVal ptIn As PointF) As PointF
        Dim pt As PointF
        pt.X = ((ptIn.X - Me.AutoScrollPosition.X) / _Zoom)
        pt.Y = ((ptIn.Y - Me.AutoScrollPosition.Y) / _Zoom)
        Return pt
    End Function

    Public Overloads Function surfaceToObjCoord(ByVal ptIn As Point) As Point
        Dim pt As Point

        pt.X = CInt((ptIn.X - Me.AutoScrollPosition.X) / _Zoom)
        pt.Y = CInt((ptIn.Y - Me.AutoScrollPosition.Y) / _Zoom)
        Return pt

    End Function




#End Region


#Region "Snapping"


    Public Function snapPoint(ByVal ptInitial As Point, ByRef bXSnapped As Boolean, ByRef bYSnapped As Boolean) As Point
        Dim ptGridMatch As Point
        Dim ptOut As PointF
        Dim iXStep As Single
        Dim iYStep As Single
        ptOut.X = ptInitial.X
        ptOut.Y = ptInitial.Y

        bYSnapped = False
        bXSnapped = False


        If App.Options.SnapToMinorGrid Then
            iXStep = XPixels(App.Options.MinorGridSize / _XDPI)
            iYStep = YPixels(App.Options.MinorGridSize / _YDPI)

            Dim ticksX As Int32 = CInt(ptInitial.X / iXStep)
            Dim ticksY As Int32 = CInt(ptInitial.Y / iYStep)

            ptGridMatch.X = CInt((ticksX) * iXStep)
            ptGridMatch.Y = CInt(ticksY * iYStep)

            If Math.Abs(ptGridMatch.X - ptInitial.X) <= App.Options.GridElasticity Then
                ptOut.X = ptGridMatch.X
                bXSnapped = True
            End If

            If Math.Abs(ptGridMatch.Y - ptInitial.Y) <= App.Options.GridElasticity Then
                ptOut.Y = ptGridMatch.Y
                bYSnapped = True
            End If
        End If


        If App.Options.SnapToMajorGrid Then
            iXStep = XPixels(App.Options.MajorGridSize / _XDPI)
            iYStep = YPixels(App.Options.MajorGridSize / _YDPI)

            Dim ticksX As Int32 = CInt(ptInitial.X / iXStep)
            Dim ticksY As Int32 = CInt(ptInitial.Y / iYStep)

            ptGridMatch.X = CInt((ticksX) * iXStep)
            ptGridMatch.Y = CInt(ticksY * iYStep)

            If Math.Abs(ptGridMatch.X - ptInitial.X) <= App.Options.GridElasticity Then
                ptOut.X = ptGridMatch.X
                bXSnapped = True
            End If

            If Math.Abs(ptGridMatch.Y - ptInitial.Y) <= App.Options.GridElasticity Then
                ptOut.Y = ptGridMatch.Y
                bYSnapped = True
            End If
        End If


        'Snap to margins
        If App.Options.SnapToMargins AndAlso _Document.ExportSettings.DocumentType = EnumDocumentTypes.ePrintDocument Then
            Dim rectMargins As Rectangle = _Document.PrintableArea


            If Math.Abs(rectMargins.X - ptInitial.X) <= App.Options.GuideElasticity Then
                ptOut.X = rectMargins.X
                bXSnapped = True
            End If

            If Math.Abs(rectMargins.Y - ptInitial.Y) <= App.Options.GuideElasticity Then
                ptOut.Y = rectMargins.Y
                bYSnapped = True
            End If


            If Math.Abs(rectMargins.Right - ptInitial.X) <= App.Options.GuideElasticity Then
                ptOut.X = rectMargins.Right
                bXSnapped = True
            End If


            If Math.Abs(rectMargins.Bottom - ptInitial.Y) <= App.Options.GuideElasticity Then
                ptOut.Y = rectMargins.Bottom
                bYSnapped = True
            End If
        End If



        'Snap to guides last - overrides everything else  
        If App.Options.SnapToGuides AndAlso MDIMain.ActiveDocument.Guides.Count > 0 Then

            For i As Int32 = 0 To MDIMain.ActiveDocument.Guides.Count - 1
                Dim gd As GDIGuide = DirectCast(MDIMain.ActiveDocument.Guides(i), GDIGuide)

                Select Case gd.Direction
                    Case GDIGuide.EnumGuideDirection.eHoriz
                        If Math.Abs(gd.XY - ptInitial.Y) <= App.Options.GuideElasticity Then
                            ptOut.Y = gd.XY
                            bYSnapped = True
                        End If
                    Case GDIGuide.EnumGuideDirection.eVert
                        If Math.Abs(gd.XY - ptInitial.X) <= App.Options.GuideElasticity Then
                            ptOut.X = gd.XY
                            bXSnapped = True
                        End If
                End Select
            Next
        End If


        Return Point.Round(ptOut)


    End Function



    'Convert from 100ths of an inch to pixels
    Private Function XPixels(ByVal value As Double) As Single
        Return CSng(value * _XDPI)
    End Function

    'Convert from 100ths of an inch to pixels
    Private Function YPixels(ByVal value As Double) As Single
        Return CSng(value * _YDPI)
    End Function

    Private Function snappedObjectPoint() As Point
        Dim ptGridMatch As Point
        Dim ptOut As Point

        Dim iXStep As Single
        Dim iYStep As Single

        ptOut.X = CInt(_ObjectPoint.X)
        ptOut.Y = CInt(_ObjectPoint.Y)


        If App.Options.SnapToMinorGrid Then
            iXStep = XPixels(App.Options.MinorGridSize / _XDPI)
            iYStep = YPixels(App.Options.MinorGridSize / _YDPI)

            Dim ticksX As Int32 = CInt(_ObjectPoint.X / iXStep)
            Dim ticksY As Int32 = CInt(_ObjectPoint.Y / iYStep)

            ptGridMatch.X = CInt((ticksX) * iXStep)
            ptGridMatch.Y = CInt(ticksY * iYStep)

            If Math.Abs(ptGridMatch.X - _ObjectPoint.X) <= App.Options.GridElasticity Then
                ptOut.X = ptGridMatch.X
            End If

            If Math.Abs(ptGridMatch.Y - _ObjectPoint.Y) <= App.Options.GridElasticity Then
                ptOut.Y = ptGridMatch.Y
            End If
        End If


        If App.Options.SnapToMajorGrid Then
            iXStep = XPixels(App.Options.MajorGridSize / _XDPI)
            iYStep = YPixels(App.Options.MajorGridSize / _YDPI)

            Dim ticksX As Int32 = CInt(_ObjectPoint.X / iXStep)
            Dim ticksY As Int32 = CInt(_ObjectPoint.Y / iYStep)

            ptGridMatch.X = CInt(ticksX * iXStep)
            ptGridMatch.Y = CInt(ticksY * iYStep)

            If Math.Abs(ptGridMatch.X - _ObjectPoint.X) <= App.Options.GridElasticity Then
                ptOut.X = ptGridMatch.X
            End If

            If Math.Abs(ptGridMatch.Y - _ObjectPoint.Y) <= App.Options.GridElasticity Then
                ptOut.Y = ptGridMatch.Y
            End If
        End If


        'Snap to margins
        If App.Options.SnapToMargins AndAlso _Document.ExportSettings.DocumentType = EnumDocumentTypes.ePrintDocument Then
            Dim rectMargins As Rectangle = _Document.PrintableArea


            If Math.Abs(rectMargins.X - _ObjectPoint.X) <= App.Options.GuideElasticity Then
                ptOut.X = rectMargins.X
            End If

            If Math.Abs(rectMargins.Y - _ObjectPoint.Y) <= App.Options.GuideElasticity Then
                ptOut.Y = rectMargins.Y
            End If


            If Math.Abs(rectMargins.Right - _ObjectPoint.X) <= App.Options.GuideElasticity Then
                ptOut.X = rectMargins.Right
            End If


            If Math.Abs(rectMargins.Bottom - _ObjectPoint.Y) <= App.Options.GuideElasticity Then
                ptOut.Y = rectMargins.Bottom
            End If
        End If



        'Snap to guides last - overrides everything else  
        If App.Options.SnapToGuides AndAlso MDIMain.ActiveDocument.Guides.Count > 0 Then

            For i As Int32 = 0 To MDIMain.ActiveDocument.Guides.Count - 1
                Dim gd As GDIGuide = DirectCast(MDIMain.ActiveDocument.Guides(i), GDIGuide)

                Select Case gd.Direction
                    Case GDIGuide.EnumGuideDirection.eHoriz
                        If Math.Abs(gd.XY - _ObjectPoint.Y) <= App.Options.GuideElasticity Then
                            ptOut.Y = gd.XY
                        End If
                    Case GDIGuide.EnumGuideDirection.eVert
                        If Math.Abs(gd.XY - _ObjectPoint.X) <= App.Options.GuideElasticity Then
                            ptOut.X = gd.XY
                        End If
                End Select
            Next
        End If


        Return ptOut


    End Function

#End Region




#Region "key press related methods and events"


    Const WM_KEYDOWN As Int32 = &H100
    Const VK_LEFT As Int32 = &H25
    Const VK_UP As Int32 = &H26
    Const VK_RIGHT As Int32 = &H27
    Const VK_DOWN As Int32 = &H28
    Const VK_ENTER As Int32 = &HD



    Private Sub handleCommandKey(ByRef e As System.Windows.Forms.KeyEventArgs)
        If Not App.ToolManager.ToolInUse() Then
            Select Case e.KeyCode
                Case Keys.Delete, Keys.Clear, Keys.Back
                    _Document.Selected.DeleteAll()

                Case Keys.Left
                    If e.Control Then
                        MDIMain.ActiveDocument.Selected.Nudge(-App.Options.LargeNudge, 0)
                    Else
                        MDIMain.ActiveDocument.Selected.Nudge(-App.Options.SmallNudge, 0)
                    End If
                    e.Handled = True

                Case Keys.Right

                    If e.Control Then
                        MDIMain.ActiveDocument.Selected.Nudge(App.Options.LargeNudge, 0)
                    Else
                        MDIMain.ActiveDocument.Selected.Nudge(App.Options.SmallNudge, 0)
                    End If
                    e.Handled = True

                Case Keys.Up
                    If e.Control Then
                        MDIMain.ActiveDocument.Selected.Nudge(0, -App.Options.LargeNudge)
                    Else
                        MDIMain.ActiveDocument.Selected.Nudge(0, -App.Options.SmallNudge)
                    End If
                    e.Handled = True

                Case Keys.Down
                    If e.Control Then
                        MDIMain.ActiveDocument.Selected.Nudge(0, App.Options.LargeNudge)
                    Else
                        MDIMain.ActiveDocument.Selected.Nudge(0, App.Options.SmallNudge)
                    End If
                    e.Handled = True

                Case Else

                    HandleToolChange(e.KeyCode)

            End Select

            Me.Cursor = App.ToolManager.ToolCursor
        End If
    End Sub


    Private Sub HandleToolChange(ByVal key As System.Windows.Forms.Keys)


        Select Case key
            Case Keys.O
                App.ToolManager.GlobalMode = EnumTools.eCircle

            Case Keys.I
                App.ToolManager.GlobalMode = EnumTools.eDropper
            Case Keys.F
                App.ToolManager.GlobalMode = EnumTools.eField
            Case Keys.B
                App.ToolManager.GlobalMode = EnumTools.eFill
            Case Keys.H
                App.ToolManager.GlobalMode = EnumTools.eHand
            Case Keys.N
                App.ToolManager.GlobalMode = EnumTools.eLine
            Case Keys.M
                App.ToolManager.GlobalMode = EnumTools.eMagnify
            Case Keys.P
                App.ToolManager.GlobalMode = EnumTools.ePen
            Case Keys.V
                App.ToolManager.GlobalMode = EnumTools.ePointer
            Case Keys.R
                App.ToolManager.GlobalMode = EnumTools.eSquare
            Case Keys.T
                App.ToolManager.GlobalMode = EnumTools.eText
        End Select

    End Sub

    Private Sub handleEscape()
        Select Case App.ToolManager.ActiveToolMode
            Case App.ToolManager.TM.ePenning, App.ToolManager.TM.eClickInTextBox
                App.ToolManager.EndTool(_ShiftDown)
                App.ToolManager.ActiveToolMode = App.ToolManager.TM.eNone
            Case Else
                _Document.Selected.DeselectAll()
        End Select
    End Sub


    Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not e.KeyChar = Chr(8) AndAlso App.ToolManager.ActiveToolMode = App.ToolManager.TM.eClickInTextBox Then
            handleTextkey(e)
        End If
        MyBase.OnKeyPress(e)
    End Sub
    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)

        Select Case App.ToolManager.ActiveToolMode

            Case App.ToolManager.TM.eClickInTextBox
                Select Case e.KeyCode

                    Case Keys.Return, Keys.Delete, Keys.Back, Keys.Left, Keys.Right, Keys.Down, Keys.Up
                        handetextkeydown(e)
                        e.Handled = True
                        Return
                    Case Keys.Escape
                        handleEscape()
                End Select

            Case Else

                Select Case e.KeyCode
                    Case Keys.Escape
                        handleEscape()

                    Case Keys.ShiftKey
                        _ShiftDown = True

                    Case Keys.ControlKey
                        _ControlDown = True
                    Case Else
                        handleCommandKey(e)
                End Select
        End Select


    End Sub


    Private Sub handetextkeydown(ByVal e As KeyEventArgs)

        App.ToolManager.handleTextKeyDown(e)

        Invalidate()
        e.Handled = True


    End Sub


    Private Sub handleTextkey(ByVal e As KeyPressEventArgs)

        App.ToolManager.handleTextKeyPress(e)

        Invalidate()
        e.Handled = True


    End Sub


    Protected Overrides Sub OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case e.KeyCode
            Case Keys.ControlKey
                _ControlDown = False
                e.Handled = True
            Case Keys.ShiftKey
                _ShiftDown = False
                e.Handled = True
        End Select
        MyBase.OnKeyUp(e)
    End Sub


    Protected Overrides Function IsInputKey(ByVal keyData As System.Windows.Forms.Keys) As Boolean

        Dim bInputKey As Boolean = True

        Select Case keyData

            Case Keys.Delete
                Return True
            Case Keys.Left
                Return True
            Case Keys.Right
                Return True
            Case Keys.Down
                Return True
            Case Keys.Up
                Return True
            Case Else
                bInputKey = MyBase.IsInputKey(keyData)
        End Select

    End Function


#End Region



End Class