
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : TextDragBox
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Tool for creating rectangular fixed width text areas.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class TextDragBox
    Inherits GDITool


#Region "Local Fields"
    '''<summary>Current bounds of the text area</summary>
    Private _Bounds As Rectangle

    '''<summary>The last point recorded within our rectangle</summary>
    Private _PTEnd As Point

#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a TextDragbox, beginning at the point specified in 
    ''' ptObject
    ''' </summary>
    ''' <param name="ptObject">The point to begin creating text at.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal ptObject As Point)
        MyBase.New(ptObject)
        _Bounds = New Rectangle(ptObject, New Size(0, 0))
        _PTEnd = ptObject
    End Sub

#End Region

#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the current bounds of the TextDragBox tool.
    ''' </summary>
    ''' <value>A rectangle containing the current bounds of the tool</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Bounds() As Rectangle
        Get
            Return _Bounds
        End Get
    End Property

#End Region

#Region "Base Class Implementors"
 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the TextDragBox tool based upon the last mouse point and the buttons 
    ''' being held
    ''' </summary>
    ''' <param name="ptPoint"></param>
    ''' <param name="btn"></param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub UpdateTool(ByVal ptPoint As System.Drawing.Point, ByVal btn As MouseButtons, ByVal bShiftDown As Boolean, ByVal zoom As Single)
        _PTEnd = MyBase.getEndPoint(ptPoint, bShiftDown)
        _Bounds = getDrawingRect(ptPoint, bShiftDown)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the TextDrag tool to the current graphics surface.
    ''' </summary>
    ''' <param name="g">Graphics context to draw against</param>
    ''' <param name="fScale">Current zoom factor of the surface</param>
    ''' -----------------------------------------------------------------------------
    Friend Overloads Overrides Sub Draw(ByVal g As System.Drawing.Graphics, ByVal fScale As Single)
        _PenOutline.Width = 1 / fScale

        App.GraphicsManager.BeginScaledView(g)

        g.DrawRectangle(_PenOutline, _Bounds.X, _Bounds.Y, _Bounds.Width, _Bounds.Height)

        App.GraphicsManager.EndScaledView(g)

    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends the TextDragTool operation
    ''' </summary>
    ''' <param name="bShiftDown">Whether shift is being held down or not</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub EndTool(ByVal bShiftDown As Boolean)
        App.ToolManager.BeginFixedText(_Bounds)
    End Sub
#End Region




End Class
 