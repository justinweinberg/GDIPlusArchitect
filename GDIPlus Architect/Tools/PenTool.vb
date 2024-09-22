Imports System.Drawing.Drawing2D
Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : PenTool
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Wraps pen tool related functionality
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class PenTool
    Inherits GDITool

#Region "Local Fields"

    '''<summary>A GDIOpenPath that is being created with each pen click.  Used to perform 
    ''' hit testing and render the tool to the surface.</summary>
    Private _GDIOpenPath As GDIOpenPath

    '''<summary>Whether the path has been closed or not.  This occurs when the user clicks 
    ''' the pen tool and the current point is in range of the start point.</summary>
    Private _PathClosed As Boolean = False

    '''<summary>Whether a curve segment is being drawn or not</summary>
    Private _Curving As Boolean = False
    '''<summary>Whether the current pen position is within range of the start point of the path.
    ''' if the user clicks when this is true, the path will close itself</summary>
    Private _InCloseRange As Boolean = False
#End Region


#Region "Constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the path tool.
    ''' </summary>
    ''' <param name="ptSnapped"></param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal ptSnapped As Point)
        MyBase.New(ptSnapped)
        _GDIOpenPath = New GDIOpenPath(_PTOrigin)
        _PTLastPoint = _PTOrigin
    End Sub
#End Region


#Region "Public Methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a pen point to the path being created.  Returns whether the path has 
    ''' been closed or not.
    ''' </summary>
    ''' <returns>A boolean indicating if the previously open path has been closed 
    ''' by the user.</returns>
    ''' <remarks>As the user brings the mouse into the area of the initial point, 
    ''' if they add a pen point the previously open path closes.  This method returns 
    ''' a value indicating this has occurred. 
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Function AddPenPoint() As Boolean

        If _GDIOpenPath.Points.Length > 1 AndAlso _InCloseRange Then
            _PathClosed = True
            Return _PathClosed
        End If

        _GDIOpenPath.addLineSegment(_PTLastPoint)

        Return False
    End Function

#End Region


#Region "Helper Methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Continues a curvature segment with the last recorded point
    ''' </summary>
    ''' <param name="ptSnapped">Point to add to the curved segment.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub UpdateCurve(ByVal ptSnapped As Point)
        _GDIOpenPath.UpdateCurveSegment(New PointF(ptSnapped.X, ptSnapped.Y))
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Begins a curvature segment
    ''' </summary>
    ''' <param name="ptSnapped">Point to begin curving at, snapped appropriately</param>
    ''' -----------------------------------------------------------------------------
    Private Sub beginCurve(ByVal ptSnapped As Point)
        _GDIOpenPath.BeginCurveSegment(New PointF(ptSnapped.X, ptSnapped.Y))
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a boolean indicating if the pen tool is within "closeable range"
    ''' As the user moves the mouse toward the initial pen point, a circle is drawn 
    ''' indicating the path is closable.  This method determines if the pen tool
    ''' is within range.
    ''' </summary>
    ''' <param name="ptSnapped">A point snapped to grid appropriately</param>
    ''' <returns>True if the path is in close range, false otherwise.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function CanClosePath(ByVal ptSnapped As Point, ByVal zoom As Single) As Boolean
        If _GDIOpenPath.HitTestHandles(Point.op_Implicit(ptSnapped), zoom) = _
        EnumDragHandles.ePointHandle AndAlso _
        _GDIOpenPath.Points.Length > 1 Then
            If PointF.op_Equality(_GDIOpenPath.LastHitHandle(), _GDIOpenPath.Points(0)) Then
                Return True
            End If
        End If
    End Function

#End Region

#Region "Base Clase Overrides"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends the pen tool operation
    ''' </summary>
    ''' <param name="bShiftDown">Whether shift is being held down or not</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub EndTool(ByVal bShiftDown As Boolean)
        If _PathClosed Then

            MDIMain.ActiveDocument.AddObjectToPage(_GDIOpenPath.GetClosedPath(), "Pen Tool")
        ElseIf _GDIOpenPath.Points.Length > 1 Then
            MDIMain.ActiveDocument.AddObjectToPage(_GDIOpenPath, "Pen Tool")
        End If

        _PenOutline.Dispose()


        _GDIOpenPath = Nothing
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the pen tool based upon the last known mouse point and what buttons are 
    ''' being held down.
    ''' </summary>
    ''' <param name="ptPoint">The last recorded mouse point</param>
    ''' <param name="btn">Which mouse button, if any, is being held down</param>
    ''' <param name="bShiftDown">Whether shift is held down or not</param>
    ''' <remarks>Whereas most tools the mouse is held down through 
    ''' the operation, when users use a pen tool, they expect to add point to point 
    ''' What this is saying is if a point is added and dragged with the most down,
    ''' the user intends to create a curvature point.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub UpdateTool(ByVal ptPoint As System.Drawing.Point, ByVal btn As MouseButtons, ByVal bShiftDown As Boolean, ByVal zoom As Single)


        'This is a bit confusing.  W
        Select Case btn
            Case MouseButtons.Left
                If Not _Curving Then
                    _Curving = True
                    beginCurve(ptPoint)
                Else
                    UpdateCurve(ptPoint)
                End If

            Case Else
                If CanClosePath(ptPoint, zoom) Then
                    _InCloseRange = True
                Else
                    _InCloseRange = False
                End If

                _Curving = False

        End Select

        If bShiftDown Then
            Dim ptlast As Point = Point.Round(_GDIOpenPath.Points(_GDIOpenPath.Points.Length - 1))
            _PTLastPoint = Utility.ShiftedDownPoint(ptlast, ptPoint)

        Else
            _PTLastPoint = ptPoint
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the pen tool to the current surface
    ''' </summary>
    ''' <param name="g">Graphics context to draw against </param>
    ''' <param name="fScale">Current zoom factor of the surface</param>
    ''' -----------------------------------------------------------------------------
    Friend Overloads Overrides Sub Draw(ByVal g As System.Drawing.Graphics, ByVal fScale As Single)
        _PenOutline.Width = 1 / fScale

 
            App.GraphicsManager.BeginScaledView(g)


            g.DrawLine(_PenOutline, _GDIOpenPath.Points(_GDIOpenPath.Points.Length - 1), Point.op_Implicit(_PTLastPoint))

            If _InCloseRange Then
                g.DrawEllipse(_PenOutline, New RectangleF(_GDIOpenPath.Points(0).X - 5 / fScale, _GDIOpenPath.Points(0).Y - 5 / fScale, 10 / fScale, 10 / fScale))
            End If

            _GDIOpenPath.RenderTool(g)

            App.GraphicsManager.EndScaledView(g)
     End Sub


#End Region

End Class
