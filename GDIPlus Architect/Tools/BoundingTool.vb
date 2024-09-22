
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : BoundingTool
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Responsible for providing bounding box functionality
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class BoundingTool
    Inherits GDITool

#Region "Local Fields"



    '''<summary>Rectangle representing the bounding box</summary>
    Private _Bounds As Rectangle
#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a bounding box tool
    ''' </summary>
    ''' <param name="ptPoint">The point at which bounding began on the surface</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal ptPoint As Point)
        MyBase.New(ptPoint)
        _Bounds = New Rectangle(ptPoint, New Size(0, 0))
    End Sub
#End Region

#Region "Base Class Implementors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws the bounding tool to the graphics surface
    ''' </summary>
    ''' <param name="g">Graphics context to draw against</param>
    ''' <param name="fScale">Current zoom factor of the surface</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub draw(ByVal g As Graphics, ByVal fScale As Single)
        _PenOutline.Width = 1 / fScale
        Dim gCt As Drawing2D.GraphicsContainer = g.BeginContainer
        g.ScaleTransform(fScale, fScale)
        g.DrawRectangle(_PenOutline, _Bounds)
        g.EndContainer(gCt)
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends the bounding tool operation
    ''' </summary>
    ''' <param name="bShiftDown">Whether shift is held down or not</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub EndTool(ByVal bShiftDown As Boolean)
        MDIMain.ActiveDocument.Selected.handleSelection(_Bounds, bShiftDown)
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the bounding tool with the latest information from the client surface
    ''' </summary>
    ''' <param name="ptPoint">Last recorded mouse point </param>
    ''' <param name="btn">Which buttons are held down</param>
    ''' <param name="bShiftDown">Whether shift is being held down</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub UpdateTool(ByVal ptPoint As System.Drawing.Point, ByVal btn As MouseButtons, ByVal bShiftDown As Boolean, ByVal Zoom As Single)
        Dim temporigin As New Point
        Dim tempextent As New Size

        temporigin.X = Math.Min(_PTOrigin.X, ptPoint.X)
        temporigin.Y = Math.Min(_PTOrigin.Y, ptPoint.Y)

        tempextent.Width = Math.Abs(ptPoint.X - _PTOrigin.X)
        tempextent.Height = Math.Abs(ptPoint.Y - _PTOrigin.Y)

        _Bounds = New Rectangle(temporigin, tempextent)
    End Sub


#End Region
End Class
