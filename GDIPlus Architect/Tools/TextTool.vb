Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : TextTool
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Tool used to manipulate text on the drawing surface by typing into the text.
''' </summary>
''' <remarks>This tool is solely used to allow user to type onto the text objcets 
''' on the surface.  The TextDragBox tool is used to position text within fixed 
''' bounded boxes
''' </remarks>
''' -----------------------------------------------------------------------------
Friend Class TextTool
    Inherits GDITool

#Region "Local Fields"


    '''<summary>Whether to draw the cursor or not</summary>
    Private _DrawCursor As Boolean = True



    '''<summary>Temporary object created for text operations.  This will 
    ''' become an actual member of the page depending on what happens 
    ''' during text operations.</summary>
    Private _GDITextObject As GDIText

    '''<summary>When lines of text are selected (highlighed), 
    ''' The position where the cursor lies.</summary>
    Private _TextIndexCurrent As Int32

    '''<summary>When lines of text are selected (highlighed),
    '''the position of the first selected character.</summary>
    Private _TextIndexStart As Int32 = -1


    '''<summary>The rectangle to draw the text cursor at.</summary>
    Private _CursorRect As RectangleF

    '''<summary>Whether the text object being manipulated (or created) is in 
    ''' a fixed box or a variable mode.  Fixed implies that the text wraps 
    ''' around the edges of the box.</summary>
    Private _FixedLayout As Boolean = False




    '''<summary>A brush used to paint selected text.</summary>
    Private _HiglightBrush As New SolidBrush(Color.LightGreen)


    '''<summary>Whether the highlight rectangle needs to be updated or not</summary>
    Private _DrawHighlight As Boolean = False


    '''<summary>Stores the rotation, if any on the text object.  
    ''' We don't want users to have to try and read rotated text, so this value 
    ''' stored the last rotation value while the text is being manipulated</summary>
    Private _Rotation As Single = 0



    '''<summary>Boolean indicating whether an existing GDIText object is being
    '''  edited or a new object is being created.</summary>
    Private _Editing As Boolean = False


    '''<summary>A private undo and redo stack used when text is
    '''  copied, cut, and paste on the surface during the edit
    '''  operation</summary>
    Private _UndoStack As New Stack(20)
    Private _RedoStack As New Stack(20)


    '''<summary>Series of rectangles that represent the highlighted characters in the text object during manipulation</summary>
    Private _SelectedCharacters() As RectangleF

#End Region

#Region "Constructors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new text tool instance for the case when the user has dragged a 
    ''' rectangle.  In this case, the layout is fixed.
    ''' </summary>
    ''' <param name="rect">The fixed layout rect</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal rect As Rectangle)
        MyBase.New(New Point(rect.X, rect.Y))
        'Set up the local text object appropriately for when the user has drawn a text 
        'area rectangle
        _GDITextObject = New GDIText(" ", App.Options.Font, rect, True)
        _FixedLayout = True

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new text tool for when the user has clicked with the text tool. 
    ''' In this case the text object created has a variable length layout.
    ''' </summary>
    ''' <param name="pt"></param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal pt As Point)
        MyBase.New(pt)
        'Set up the local text object appropriately for when the user clicks to start typing 
        'text
        _GDITextObject = New GDIText(" ", App.Options.Font, _
        New Rectangle(pt.X, pt.Y, 0, App.Options.Font.Height), False)

        _FixedLayout = False

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new text tool for when the user is editing an existing object.
    ''' </summary>
    ''' <param name="pt">Origin point of the text</param>
    ''' <param name="miCharIndex">The char index "hit" by the tool when text editing began.</param>
    ''' <param name="editingObject">The GDIText object being edited by the tool</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal pt As Point, ByVal miCharIndex As Int32, ByVal editingObject As GDIText)
        MyBase.New(pt)
        'If editing, get the rotation
        _Rotation = editingObject.Rotation
        editingObject.Rotation = 0

        'Assign the tool's text object to the object being edited
        _GDITextObject = editingObject
        'Add a space trailer.  To see why this is done, try doing it without this.
        _GDITextObject.Text &= " "

        _FixedLayout = editingObject.Wrap
        _TextIndexCurrent = miCharIndex
        _Editing = True

    End Sub


#End Region

#Region "Hit testing and selection"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a boolean indicating whether a point is within the bounds of 
    ''' the text box.
    ''' </summary>
    ''' <param name="pt">The point to check.</param>
    ''' <returns>A boolean indicating if the text was within the box.</returns>
    ''' -----------------------------------------------------------------------------
    Friend Function PointIsInBox(ByVal pt As PointF, ByVal zoom As Single) As Boolean
        _TextIndexStart = -1

        Return _GDITextObject.HitTest(pt, zoom)

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Attempts to set the cursor appropriately based upon a click point within the 
    ''' text box.
    ''' </summary>
    ''' <param name="pt"></param>
    ''' -----------------------------------------------------------------------------
    Public Sub hittestCharPoint(ByVal pt As PointF)
        _TextIndexCurrent = _GDITextObject.CharIndexAtPoint(pt)
        _DrawCursor = True

        MDIMain.RefreshActiveDocument()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Tries to select the text at a point.  This is used when the user clicks on a 
    ''' character while in text edit mode.
    ''' </summary>
    ''' <param name="pt">The point at which the mouse fell translated into the object's 
    ''' coordinates.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub AttemptSelection(ByVal pt As PointF)
        Dim miSelIndex As Int32
        miSelIndex = _GDITextObject.CharIndexAtPoint(pt)
        _DrawCursor = True

        If Not _TextIndexStart = miSelIndex Then
            _TextIndexStart = miSelIndex
            _DrawHighlight = True
        End If

        MDIMain.RefreshActiveDocument()

    End Sub

#End Region

#Region "Local event handlers"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a local paste into the text box
    ''' </summary>
    ''' <param name="sText">The text to psate.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub DoPaste(ByVal sText As String)
        If _TextIndexStart > -1 Then
            deleteSelection()
        End If

        If _TextIndexCurrent < 0 Then
            _TextIndexCurrent = 0
        End If

        _UndoStack.Push(_GDITextObject.Text)

        _GDITextObject.Text = _GDITextObject.Text.Insert(_TextIndexCurrent, sText)

        _TextIndexCurrent += sText.Length

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Undoes a local text operation.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub DoUndo()
        If Not _UndoStack.Count = 0 Then
            Dim sUndoText As String = CStr(_UndoStack.Pop())

            checkIndexes(sUndoText)
            _RedoStack.Push(_GDITextObject.Text)
            _GDITextObject.Text = sUndoText
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Redoes as local text operation
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub DoRedo()
        If Not _RedoStack.Count = 0 Then
            Dim sRedoText As String = _RedoStack.Pop().ToString

            checkIndexes(sRedoText)
            _UndoStack.Push(_GDITextObject.Text)
            _GDITextObject.Text = sRedoText
        End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a local copy on the text object
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub DoCopy()

        If _TextIndexCurrent < 0 Then
            _TextIndexCurrent = 0
        End If

        copyselection()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a local cut on the text object
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub DoCut()
        If _TextIndexStart > -1 Then
            cutselection()
        End If
    End Sub
#End Region

#Region "misc Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Verifies that a valid text position is still selected after an undo or a redo
    ''' </summary>
    ''' <param name="stext"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub checkIndexes(ByVal stext As String)

        _TextIndexStart = -1

        If _TextIndexCurrent < 0 Then
            _TextIndexCurrent = 0

        End If
        If _TextIndexCurrent > stext.Length - 1 Then
            _TextIndexCurrent = stext.Length - 1
        End If

    End Sub
#End Region

#Region "Base Class Implementors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Not used with the text tool.
    ''' </summary>
    ''' <param name="ptPoint"></param>
    ''' <param name="btn"></param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub UpdateTool(ByVal ptPoint As System.Drawing.Point, ByVal btn As MouseButtons, ByVal bShiftDown As Boolean, ByVal zoom As Single)

    End Sub






    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends a text tool operation and updates an existing GDIText object or creates a 
    ''' new object as appropriate.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub EndTool(ByVal bShiftDown As Boolean)
        _GDITextObject.Text = _GDITextObject.Text.Remove(_GDITextObject.Length - 1, 1)
        If _GDITextObject.Text.Length > 0 Then

            If _Editing = False Then
                Dim x As New GDIText(_GDITextObject.Text, _GDITextObject.Font, _GDITextObject.Bounds, _FixedLayout)

                x.Fill = Session.Fill
                x.Stroke = Session.Stroke

                MDIMain.ActiveDocument.AddObjectToPage(x, "Created text")
 
          

            Else

                _GDITextObject.Rotation = _Rotation
                MDIMain.ActiveDocument.recordHistory("Updated text")
            End If
        End If

        If Not _PenOutline Is Nothing Then
            _PenOutline.Dispose()
        End If

    End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the text tool to the current surface.
    ''' </summary>
    ''' <param name="g">A graphics context to draw the text tool to.</param>
    ''' <param name="fScale">Current zoom factor of the drawing surface</param>
    ''' -----------------------------------------------------------------------------
    Friend Overloads Overrides Sub draw(ByVal g As System.Drawing.Graphics, ByVal fScale As Single)
        _PenOutline.Width = 1 / fScale
        Dim mpcursor As New Pen(Color.Black, 2 / fScale)

        App.GraphicsManager.BeginScaledView(g)


        If Not _TextIndexStart = -1 AndAlso _DrawHighlight Then
            _SelectedCharacters = _GDITextObject.getSelectedCharRanges(_TextIndexCurrent, _TextIndexStart)
            g.FillRectangles(_HiglightBrush, _SelectedCharacters)

        ElseIf _DrawCursor AndAlso _TextIndexCurrent > -1 Then
            _CursorRect = _GDITextObject.getRectangleForCharPos(_TextIndexCurrent)
            g.DrawLine(mpcursor, _CursorRect.X, _CursorRect.Top, _CursorRect.X, _CursorRect.Bottom)
        End If


        _GDITextObject.RenderTool(g)


        App.GraphicsManager.EndScaledView(g)


        mpcursor.Dispose()


    End Sub


    Friend Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then

            If Not _HiglightBrush Is Nothing Then
                _HiglightBrush.Dispose()
            End If
        End If

        MyBase.Dispose(disposing)
    End Sub

#End Region

#Region "Key press handlers"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Allows the text tool to handle a key press event, updating the contents of the 
    ''' text box with the new character
    ''' </summary>
    ''' <param name="e">The key pressed</param>
    ''' -----------------------------------------------------------------------------
    Public Sub handleKeyPress(ByVal e As KeyPressEventArgs)

        If _TextIndexStart > -1 Then
            deleteSelection()
        End If

        If _TextIndexCurrent < 0 Then
            _TextIndexCurrent = 0
        End If
        _GDITextObject.Text = _GDITextObject.Text.Insert(_TextIndexCurrent, e.KeyChar.ToString)

        _TextIndexCurrent += 1
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a key down event inside the text box.
    ''' </summary>
    ''' <param name="e">A KeyEventArgs matching the keypressed.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub handleKeyDown(ByVal e As KeyEventArgs)
        Select Case e.KeyData

            Case Keys.Return
                handleNewLine()

            Case Keys.Delete
                handleDelete()

            Case Keys.Back
                handleBackSpace()

            Case Keys.Left
                handleLeft()

            Case Keys.Right
                handleRight()

            Case Keys.Down
                handleDown()

            Case Keys.Up
                handleUp()

        End Select

        _DrawCursor = True

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a new line key press inside a text box.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub handleNewLine()
        If _TextIndexStart > -1 Then
            deleteSelection()
        End If

        If _TextIndexCurrent < _GDITextObject.Length Then
            _GDITextObject.Text = _GDITextObject.Text.Insert(_TextIndexCurrent, ControlChars.CrLf)
        Else
            _GDITextObject.Text &= ControlChars.CrLf
        End If

        _TextIndexCurrent += 2


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a backspace operation inside the text box
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub handleBackSpace()
        If _TextIndexStart > -1 Then
            deleteSelection()

        ElseIf _TextIndexCurrent > 0 Then
            _GDITextObject.Text = _GDITextObject.Text.Remove(_TextIndexCurrent - 1, 1)
            _TextIndexCurrent = Math.Max(0, _TextIndexCurrent - 1)
        End If


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a delete operation inside the text box
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub handleDelete()
        If _TextIndexStart > -1 Then
            deleteSelection()
        ElseIf _TextIndexCurrent < _GDITextObject.Length - 1 Then
            _GDITextObject.Text = _GDITextObject.Text.Remove(_TextIndexCurrent, 1)
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a right arrow key  press inside the text box
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub handleRight()
        _DrawCursor = True
        _TextIndexCurrent = Math.Min(_GDITextObject.Length - 1, _TextIndexCurrent + 1)
        _TextIndexStart = -1
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a left arrow key  press inside the text  box
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub handleLeft()

        _TextIndexCurrent = Math.Max(0, _TextIndexCurrent - 1)
        _TextIndexStart = -1
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a down arrow key press inside the text box
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub handleDown()
        _TextIndexStart = -1
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles an up arrow key inside the text box
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub handleUp()
        _TextIndexStart = -1
    End Sub

#End Region


#Region "Selected Text Manipulation"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deletes the currently highlighted (selected) text inside the text box.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub deleteSelection()
        If _TextIndexCurrent > _TextIndexStart Then
            Dim temppos As Int32 = _TextIndexCurrent
            _TextIndexCurrent = _TextIndexStart
            _TextIndexStart = temppos
        End If
        _UndoStack.Push(_GDITextObject.Text)

        _GDITextObject.Text = _GDITextObject.Text.Remove(_TextIndexCurrent, _TextIndexStart - _TextIndexCurrent)
        ' _TextIndexCurrent = miSelStartPos
        _TextIndexStart = -1
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a copy on the currently selected text in the text box.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub copyselection()
        If _TextIndexCurrent > _TextIndexStart Then
            Dim temppos As Int32 = _TextIndexCurrent
            _TextIndexCurrent = _TextIndexStart
            _TextIndexStart = temppos
        End If

        If _TextIndexCurrent > -1 Then
            Dim sCopyText As String = _GDITextObject.Text.Substring(_TextIndexCurrent, _TextIndexStart - _TextIndexCurrent)

            System.Windows.Forms.Clipboard.SetDataObject(sCopyText, True)
        End If
        _TextIndexStart = -1
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a cut on the currently selected text in the text box.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub cutselection()
        If _TextIndexCurrent > _TextIndexStart Then
            Dim temppos As Int32 = _TextIndexCurrent
            _TextIndexCurrent = _TextIndexStart
            _TextIndexStart = temppos
        End If
        If _TextIndexCurrent > -1 Then
            Dim sCutText As String = _GDITextObject.Text.Substring(_TextIndexCurrent, _TextIndexStart - _TextIndexCurrent)
            _UndoStack.Push(_GDITextObject.Text)

            _GDITextObject.Text = _GDITextObject.Text.Remove(_TextIndexCurrent, _TextIndexStart - _TextIndexCurrent)

            System.Windows.Forms.Clipboard.SetDataObject(sCutText, True)
        End If

        _TextIndexStart = -1

    End Sub

#End Region
End Class
