Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : ToolManager
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a wrapper for all tool related functionality.  The ToolManager is 
''' responsible for maintaining the currently selected tool, the in use state of 
''' tools, and providing methods to work with tools.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class ToolManager

#Region "Type Declarations"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The in use tool mode, or active tool mode.  The active tool mode is what 
    ''' is actually occurring on a drawing surface, which is different than the 
    ''' currently selected tool in the tool window.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Enum TM
        ''' <summary>No tool is actively in use</summary>
        eNone = 0
        ''' <summary>The user is drawing a line</summary>
        eLining
        ''' <summary>The user has clicked with the text tool on an empty space, but 
        ''' has not decided what to do yet.  If the user moves the mouse, they will 
        ''' begin to draw a bounded text box.  Alternatively, if they release the mouse up,
        ''' they can begin typing out a text box and created a non bounded box.</summary>
        eTextBegin
        ''' <summary>The user has clicked inside of a text box while 
        ''' using the pointer, invoking a text edit.</summary>
        eClickInTextBox
        ''' <summary>The user is drawing a bounding box after a TextBegin operation.  They 
        ''' are in the process of creating a fixed width text box.</summary>
        eTextDefinebox
        ''' <summary>The user is placing an image on the surface</summary>
        ePlacing
        ''' <summary>The user is drawing a simple shape (ellipse, rectangle)</summary>
        eDrawing
        ''' <summary>The user is drawing with the pen tool</summary>
        ePenning
        ''' <summary>The user is drawing a a GDIField</summary>
        eFielding
        ''' <summary>The user has taken the first step in a drag operation by selecting 
        ''' an object or series of objects and has held the mouse down.  This is before 
        ''' the actual drag operation begins and the user has not released the mouse.  If 
        ''' the user continues to hold down the mouse in this state, this will convert to 
        ''' an eDragging state.</summary>
        eDragPending
        ''' <summary>The user has clicked on an empty area of the surface, signalling that 
        ''' they are interested in beginning a bound operation.  If the user continues 
        ''' to hold the mouse drown, this converts to an eBounding state.</summary>
        eBoundPending
        ''' <summary>The user is dragging an object or objects along the surface.</summary>
        eDragging
        ''' <summary>The user is drawing a bounding box to select objects on the surface.</summary>
        eBounding
        ''' <summary>The user is dragging a drag handle.  Before this becomes valid it is 
        ''' assumed that the user has already selected a drag handle.</summary>
        eDraggingHandle
        ''' <summary>The user is moving a guide along the surface.</summary>
        eMovingGuide
        ''' <summary>The user is changing the viewed position with the hand tool.</summary>
        eHanding
    End Enum

#End Region



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Lets the ToolManager respond to tool up commands on the drawing surface.
    ''' Depending on the activetool mode, different actions are taken.
    ''' </summary>
    ''' <param name="shiftDown">Whether shift is held down when the tool up occurs</param>
    ''' <param name="ptObject">The up point in object coordinate space</param>
    ''' -----------------------------------------------------------------------------
    Public Sub OnToolUp(ByVal shiftDown As Boolean, ByVal ptObject As Point)

        'When the tool is up, there isn't a selected drag handle.
        _SelectedDragHandle = EnumDragHandles.eNone

        Select Case _ActiveToolMode
            'For pending tools that never fired and the hand tool, mark the activetoolmode
            'as none.
            Case TM.eDragPending, TM.eBoundPending, TM.eHanding
                _ActiveToolMode = TM.eNone

                'For tools that finish on a mouse up, call EndTool and set the tool mode 
                'to none.
            Case TM.eDragging, TM.eBounding, TM.eLining, _
                 TM.eDrawing, TM.eFielding, TM.eDraggingHandle

                EndTool(shiftDown)
                _ActiveToolMode = TM.eNone

                'The place tool is the same as the above case, but the cursor has to be 
                'reset to a pointer.  With the other tools, the cursor should stay the 
                'same
            Case TM.ePlacing
                EndTool(shiftDown)
                _ActiveToolMode = TM.eNone

                GlobalMode = EnumTools.ePointer

                'The next two cases are for the more complicated text tool functionality.

                'TM.eTextDefinebox is the mode when the user has dragged a text box area
                'along the surface.  In other words, the user moused down with the text 
                'tool and moved the tool.

                'TM.eTextBegin is the mode that occurs when the user has placed the 
                'mouse down with the text tool and immediately brought it back up. 
                'If this occurs, it's assumed the user is going to start typing out 
                'text.

            Case TM.eTextDefinebox
                'When the tool up happens when drawing a textbox, place the text box
                'into edit mode and let the user start typing.

                HandleDrawnTextBox(shiftDown)

            Case TM.eTextBegin
                'When the text tool comes up without moving, immediately start a text 
                'flow and place the user in text edit mode.
                _ActiveToolMode = TM.eClickInTextBox
                _CurrentGDITool = New TextTool(ptObject)
                markEditingText()


        End Select

    End Sub

 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends a Text bounds box draw operation and places the user into text edit mode.
    ''' </summary>
    ''' <param name="bShiftDown">Whether shift is held down or not</param>
    ''' -----------------------------------------------------------------------------
    Private Sub HandleDrawnTextBox(ByVal bShiftDown As Boolean)
        Dim rect As Rectangle = DirectCast(_CurrentGDITool, TextDragBox).Bounds

        _CurrentGDITool.EndTool(bShiftDown)
        _CurrentGDITool.Dispose()
        _CurrentGDITool = Nothing


        If rect.Width > 7 AndAlso rect.Height > 7 Then
            BeginFixedText(rect)
        Else
            Dim pt As New Point(rect.X, rect.Y)
            _CurrentGDITool = New TextTool(pt)
            markEditingText()

        End If


    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the tool portion of lost focus on a surface.
    ''' </summary>
    ''' <remarks>With most tools, we aren't interested if they are drawn off the surface - 
    ''' this is marked as a tool that never actualized.  However, the text tool and the 
    ''' pen tool are tools that accept multiple clicks and exist regardless of the 
    ''' specific mouse state.  When the user looses focus on a surface using these tools,
    ''' we want to record an end tool for these two tools.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub handleLostFocus()

        Select Case _ActiveToolMode

            Case TM.ePenning, TM.eClickInTextBox
                Me.EndTool(False)
                _ActiveToolMode = TM.eNone

        End Select
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the tool relevant portion of double clicks on a surface.
    ''' </summary>
    ''' <param name="ptObject">The click point in object coordinates</param>
    ''' <param name="ptObjectF">The click point in precise non rounded object coordinates</param>
    ''' <remarks>The only tool currently interested in double clicks is the text tool.
    ''' If a double click occurs on a GDIText box, the user wishes to select it and put 
    ''' it in text edit mode, and the cursor should be at the point it was double clicked.
    ''' -----------------------------------------------------------------------------
    Public Sub handleDoubleClick(ByVal ptObject As Point, ByVal ptObjectF As PointF)

        If TypeOf App.MDIMain.ActiveDocument.Selected.LastSelected Is GDIText Then
            'Double click occurred on a text box - start text editing.

            Dim tempgdi As GDIText = DirectCast(App.MDIMain.ActiveDocument.Selected.LastSelected, GDIText)

            _CurrentGDITool = New TextTool(ptObject, tempgdi.CharIndexAtPoint(ptObjectF), tempgdi)

            markEditingText()

            _ActiveToolMode = TM.eClickInTextBox

        End If

    End Sub

#Region "Local Fields"

    '''<summary>For image import, the path to an image which is being imported to the current surface</summary>
    Private _armedImageSrc As String

    '''<summary>Whether a GDIText field on the surface is currently being edited. 
    ''' This is used to denote if cut/copy, etc. should pass through to the text field. </summary>
    Private _TextEditing As Boolean = False


    ''' <summary> The currently selected Guide, if any. </summary>
    Private _SelectedGuide As GDIGuide = Nothing


    '''<summary>The selected drag handle or eNone if no drag handle is selected</summary>
    Private _SelectedDragHandle As EnumDragHandles = EnumDragHandles.eNone



    '''<summary>The selected tool in the tool box - the applicaiton wide current tool</summary>
    Private _GlobalToolmode As EnumTools = EnumTools.ePointer

    '''<summary>The current tool mode.  This is what is
    '''  actually happening on the drawing surface, and does not map one to one 
    ''' to the global tool mode.</summary>
    Private _ActiveToolMode As TM

    '''<summary>The current GDITool being used, if any, or nothing if no tool is in use</summary>
    Private _CurrentGDITool As GDITool = Nothing

    '''<summary>The application wide  tool cursor</summary>
    Private _ToolCursor As Cursor = Cursors.Default

#End Region



#Region "text Tool Related Members"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Tries to select text in the current TextTool based on a hit point on the surface
    ''' </summary>
    ''' <param name="ptobject">Down point where character selection should occur </param>
    ''' -----------------------------------------------------------------------------
    Public Sub AttemptTextSelection(ByVal ptobject As PointF)
        DirectCast(_CurrentGDITool, TextTool).AttemptSelection(ptobject)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Lets the texttool respond to clicking inside the text box.  Returns a boolean 
    ''' indicating if the click was within the box.  The surface uses this to determine 
    ''' if it should stop text editing mode (users expect a click outside a box to stop 
    ''' current editing)
    ''' </summary>
    ''' <param name="pt">The point to check if in box.</param>
    ''' <returns>A boolean indicating if the hit was in the box or not</returns>
    ''' -----------------------------------------------------------------------------
    Public Function ClickinTextbox(ByVal pt As PointF, ByVal zoom As Single) As Boolean
        Dim texttool As TextTool = DirectCast(_CurrentGDITool, TextTool)
        Dim bInBox As Boolean = texttool.PointIsInBox(pt, zoom)
        If bInBox = False Then
            Return False
        Else
            texttool.hittestCharPoint(pt)
            Return True
        End If
    End Function



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Allows the text tool to handle a key down event.
    ''' </summary>
    ''' <param name="e">The key pressed as a KeyEventArgs enumeration</param>
    '''-----------------------------------------------------------------------------
    Public Sub handleTextKeyDown(ByVal e As KeyEventArgs)
        DirectCast(_CurrentGDITool, TextTool).handleKeyDown(e)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sends a delete notification to the current TextTool
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub notifyTextDelete()
        Me.handleTextKeyDown(New KeyEventArgs(Keys.Delete))
        MDIMain.RefreshActiveDocument()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Allows the text tool to handle a key press event
    ''' </summary>
    ''' <param name="e">The key pressed as a KeyEventArgs enumeration</param>
    ''' -----------------------------------------------------------------------------
    Public Sub handleTextKeyPress(ByVal e As KeyPressEventArgs)
        DirectCast(_CurrentGDITool, TextTool).handleKeyPress(e)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Notifies the text tool of a copy command
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub notifyTextCopy()
        DirectCast(_CurrentGDITool, TextTool).DoCopy()
        MDIMain.RefreshActiveDocument()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Notifies the text tool of an undo command
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub notifyTextUndo()
        DirectCast(_CurrentGDITool, TextTool).DoUndo()
        MDIMain.RefreshActiveDocument()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Notifies the text tool of a redo command
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub notifyTextRedo()
        DirectCast(_CurrentGDITool, TextTool).DoRedo()
        MDIMain.RefreshActiveDocument()

    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Notifies the text tool of a paste command
    ''' </summary>
    ''' <param name="sText">The text to paste</param>
    ''' -----------------------------------------------------------------------------
    Public Sub notifyTextPaste(ByVal sText As String)
        DirectCast(_CurrentGDITool, TextTool).DoPaste(sText)
        MDIMain.RefreshActiveDocument()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Notifies the text tool a cut operation is occurring.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub NotifyTextCut()
        DirectCast(_CurrentGDITool, TextTool).DoCut()
        MDIMain.RefreshActiveDocument()

    End Sub
#End Region


#Region "Begin Tools"

    Public Sub BeginFixedText(ByVal rect As Rectangle)
        Trace.WriteLineIf(App.TraceOn, "GDISurface.BeginFixedText")

        _CurrentGDITool = New TextTool(rect)
        markEditingText()

        Me.ActiveToolMode = TM.eClickInTextBox
        MDIMain.RefreshActiveDocument()

    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Begins a drag handle operation.  Creates a draghandle tool.
    ''' </summary>
    ''' <param name="obj">Object to resize</param>
    ''' <param name="initHandle">The handle being resized</param>
    ''' <param name="ptInitial">The point where the handle was positioned at.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub BeginDragHandle(ByVal obj As GDIObject, ByVal initHandle As EnumDragHandles, ByVal ptInitial As Point)
        _CurrentGDITool = New DragHandleTool(obj, initHandle, ptInitial)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Begins a drag operation.  Creates a dragtool.
    ''' </summary>
    ''' <param name="ptSnapped">Point where dragging began, snapped appropriately.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub BeginDragging(ByVal ptObject As Point)
        _ActiveToolMode = TM.eDragging
        _CurrentGDITool = New DragTool(ptObject)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Begins a bounding box operation.  Creates a boundingtool.
    ''' </summary>
    ''' <param name="ptSnapped">Point where bounding began, snapped appropriately.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub BeginBounding(ByVal ptObject As Point)
        _ActiveToolMode = App.ToolManager.TM.eBounding
        _CurrentGDITool = New BoundingTool(ptObject)
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Starts an image place operation.  Creates a new place tool.
    ''' </summary>
    ''' <param name="ptsnapped">point to begin placing the image at, snapped appropriately</param>
    ''' <param name="sImageSource">A full path to the image resource being placed.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub BeginImagePlace(ByVal ptsnapped As Point, ByVal sImageSource As String)
        _CurrentGDITool = New PlaceTool(ptsnapped)
        Try
            DirectCast(_CurrentGDITool, PlaceTool).ImageSource = sImageSource
        Catch ex As System.ArgumentException
            _CurrentGDITool.Dispose()
            _CurrentGDITool = Nothing
            Throw New System.ArgumentException("Invalid bitmap file")
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Starts creating a field tool with variable width.
    ''' </summary>
    ''' <param name="ptSnapped">Top left point of the new field, snapped appropriately</param>
    ''' -----------------------------------------------------------------------------
    Public Sub BeginField(ByVal ptSnapped As Point)
        _ActiveToolMode = App.ToolManager.TM.eFielding
        _CurrentGDITool = New FieldTool(ptSnapped)
    End Sub
 


    Public Sub handlePen(ByVal shiftDown As Boolean, ByVal ptSnapped As Point)

        If _ActiveToolMode = TM.ePenning Then
            Dim pathclosed As Boolean = AddPenPoint()

            If pathclosed Then
                _ActiveToolMode = TM.eNone
                EndTool(shiftDown)
            End If

        Else
            _ActiveToolMode = TM.ePenning
            App.MDIMain.ActiveDocument.Selected.DeselectAll()
            BeginPen(ptSnapped)
        End If
    End Sub



    Friend Sub handleToolDown(ByVal pointSnapped As Point, ByVal objectPointF As PointF, ByVal zoom As Single, ByVal shiftdown As Boolean)
        Trace.WriteLineIf(App.TraceOn, "GDISurface.handleleftclick")

        If _ActiveToolMode = TM.eMovingGuide Then
            _SelectedGuide = Nothing
            _ActiveToolMode = TM.eNone
        End If

        If _ActiveToolMode = TM.eClickInTextBox Then
            If Not ClickinTextbox(objectPointF, zoom) Then
                EndTool(shiftdown)
                _ActiveToolMode = TM.eNone
            End If
        End If

        Select Case GlobalMode



            Case EnumTools.eText
                HandleNewText()

            Case EnumTools.eLine
                BeginLine(pointSnapped)


            Case EnumTools.ePlacing
                BeginPlace(pointSnapped)

            Case EnumTools.ePen
                handlePen(shiftdown, pointSnapped)

            Case EnumTools.eField
                BeginField(pointSnapped)

            Case EnumTools.eCircle, EnumTools.eSquare
                BeginShape(pointSnapped)

        End Select

    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Starts creating a line on a surfcae.  Creates a new LineTool.
    ''' </summary>
    ''' <param name="ptorigin"></param>
    ''' -----------------------------------------------------------------------------
    Public Sub BeginLine(ByVal ptorigin As Point)
        _ActiveToolMode = TM.eLining
        _CurrentGDITool = New LineTool(ptorigin)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Starts text edit mode for a GDIText field on the surface.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub markEditingText()
        Trace.WriteLineIf(App.TraceOn, "GDIMenu.BeginTextEditMode")
        _TextEditing = True
        App.MDIMain.RefreshMenu()
    End Sub


    Public Sub HandleNewText()
        If _ActiveToolMode = TM.eNone Then
            _ActiveToolMode = TM.eTextBegin
        End If
    End Sub


    Public Sub BeginPlace(ByVal ptSnapped As Point)

        _ActiveToolMode = TM.ePlacing

        Try
            BeginImagePlace(ptSnapped, _armedImageSrc)

        Catch ex As System.ArgumentException
            _ActiveToolMode = App.ToolManager.TM.eNone
            _GlobalToolmode = EnumTools.ePointer
            MsgBox("Invalid Bitmap file")
        End Try

    End Sub






    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Starts a draw operation.  Creates a DrawingTool instance.
    ''' </summary>
    ''' <param name="ptSnapped">Point where drawing should begin, snapped appropriately.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub BeginShape(ByVal ptSnapped As Point)
        _ActiveToolMode = App.ToolManager.TM.eDrawing
        _CurrentGDITool = New ShapeTool(ptSnapped)
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Begins a pen tool operation.  Creates a PenTool object.
    ''' </summary>
    ''' <param name="ptSnapped">Point where the pen should begin, snapped appropriately.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub BeginPen(ByVal ptSnapped As Point)
        _CurrentGDITool = New PenTool(ptSnapped)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Begins a fixed width text box operation.  The TextDragBox tool is used when the
    ''' user is defining a fixed width text box.
    ''' </summary>
    ''' <param name="ptSnapped">Top left corner to begin creating the textbox at, snapped appropriately.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub BeginDefineTextbox(ByVal ptSnapped As Point)
        _CurrentGDITool = New TextDragBox(ptSnapped)
    End Sub



#End Region


#Region "Pen Specific Functionality"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a new pen point to a current pen tool.
    ''' </summary>
    ''' <returns>A boolean value indicating if the pen tool has closed its current 
    ''' path.  If it has, this becomes a closed path object.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function AddPenPoint() As Boolean
        Return DirectCast(_CurrentGDITool, PenTool).AddPenPoint()
    End Function

#End Region

#Region "End Tools"





    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Generic end tool operation.  Used for most tools to indicate the surface is 
    ''' done with the tool.  
    ''' </summary>
    ''' <param name="bShiftDown">Whether shift is held as the tool ends.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub EndTool(ByVal bShiftDown As Boolean)
        'This just saves a couple overrides
        _TextEditing = False
        App.MDIMain.RefreshMenu()


        _CurrentGDITool.EndTool(bShiftDown)
        _CurrentGDITool.Dispose()

        _CurrentGDITool = Nothing
    End Sub


#End Region

#Region "Update Tools"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Lets a tool update themselves in response to changes in mouse positions and 
    ''' buttons
    ''' </summary>
    ''' <param name="pt">the latest mouse point</param>
    ''' <param name="btn">the mouse button being held</param>
    ''' -----------------------------------------------------------------------------
    Public Sub UpdateTool(ByVal pt As Point, ByVal btn As MouseButtons, ByVal bShiftDown As Boolean, ByVal zoom As Single)
        _CurrentGDITool.UpdateTool(pt, btn, bShiftDown, zoom)
    End Sub

#End Region

#Region "Tool drawing"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws the current tool to a graphics context
    ''' </summary>
    ''' <param name="g">Graphics context to draw against.</param>
    ''' <param name="fScale">Current zoom factor of the surface</param>
    ''' <remarks>
    '''</remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub draw(ByVal g As Graphics, ByVal fScale As Single)
        If Not _CurrentGDITool Is Nothing Then
            _CurrentGDITool.draw(g, fScale)
        End If

    End Sub


#End Region

#Region "Property Accessors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns whether the tool manager is in text edit mode or not.  This allows 
    ''' menus and panels to check if their cut/copy/undo, etc should apply to text 
    ''' or to objects.
    ''' </summary>
    ''' <value>A boolean indicating if the tool manager is in text edit mode</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property InTextEdit() As Boolean
        Get
            Return _TextEditing
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets a path to the armed image source.  An armed image source is a 
    ''' path to an image that the user is placing.
    ''' </summary>
    ''' <value>A full path to an image source.</value>
    ''' -----------------------------------------------------------------------------
    Public Property ArmedImageSrc() As String
        Get
            Return _armedImageSrc
        End Get
        Set(ByVal Value As String)
            _armedImageSrc = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or Sets the activetoolmode.  For more information, see the TM
    ''' enumeration.
    ''' </summary>
    ''' <value>The active tool mode.</value>
    ''' -----------------------------------------------------------------------------
    Public Property ActiveToolMode() As TM
        Get
            Return _ActiveToolMode
        End Get
        Set(ByVal Value As TM)
            _ActiveToolMode = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the application wide selected guide.
    ''' </summary>
    ''' <value>A GDIGuide</value>
    ''' -----------------------------------------------------------------------------
    Public Property SelectedGuide() As GDIGuide
        Get
            Return _SelectedGuide
        End Get
        Set(ByVal Value As GDIGuide)
            _SelectedGuide = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the application wide selected drag handle.
    ''' </summary>
    ''' <value>A GDIGuide</value>
    ''' -----------------------------------------------------------------------------
    Public Property SelectedDragHandle() As EnumDragHandles
        Get
            Return _SelectedDragHandle
        End Get
        Set(ByVal Value As EnumDragHandles)
            _SelectedDragHandle = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the appropriate cursor for the current tool mode.
    ''' </summary>
    ''' <value></value>
    ''' -----------------------------------------------------------------------------
    Public Property ToolCursor() As Cursor
        Get
            Return _ToolCursor
        End Get
        Set(ByVal Value As Cursor)
            _ToolCursor = Value
        End Set
    End Property




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the current toolbox toolmode (the selected tool in the toolbox)
    ''' </summary>
    ''' <value>An enumTools value indicating the currently selected tool.</value>
    ''' -----------------------------------------------------------------------------
    Public Property GlobalMode() As EnumTools
        Get
            Return _GlobalToolmode
        End Get
        Set(ByVal Value As EnumTools)
            _GlobalToolmode = Value
            App.PanelManager.OnToolsChanged()
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a boolean indicating if the tool mode has an active mode or not.
    ''' </summary>
    ''' <value>True if an active tool is in use, false otherwise.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ToolInUse() As Boolean
        Get
            Return Not _ActiveToolMode = TM.eNone
        End Get
    End Property


#End Region

End Class
