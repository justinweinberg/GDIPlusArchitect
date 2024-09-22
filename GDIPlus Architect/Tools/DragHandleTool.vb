Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : DragHandleTool
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Response for handling manipulation of drag handles for the selected object
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class DragHandleTool
    Inherits GDITool

#Region "Type Declaratioins"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The current direction the user is dragging in.  Used fo when shift is held to 
    ''' perform proportional drag handle resizing.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Enum EnumDragDirection
        eUpDown
        eLeftRight
    End Enum
#End Region


#Region "Local Fields"


    '''<summary>The object whose drag handles are being manipulated</summary>
    Private _SelectedObject As GDIObject

    '''<summary>The drag handle of _SelectedObject being manipulated</summary>
    Private _SelectedHandle As EnumDragHandles

    '''<summary>Whether shift was held down the last time the tool was updated. 
    '''to understand how this is used requires a deeper look at the code. </summary>
    Private _ShiftDownLastUpdate As Boolean = False

    '''<summary>Direction that a drag with shift down (proportional) is being dragged in.</summary>
    Private _ProportionalDrag As EnumDragDirection

    '''<summary>The rectangle of the drag handle itself</summary>
    Private _DragHandleRectangle As Rectangle
    '''<summary>The original bounds _SelectedObject prior to manipulating handles</summary>
    Private _OriginalBounds As Rectangle

    '''<summary>Amount of rotation on the object whose drag handles are being manipulated</summary>
    Private _Rotation As Single = 0

#End Region


#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the DragHandle tool given an initial object, a 
    ''' drag handle point, and which handle is being manipulated
    ''' </summary>
    ''' <param name="gdiobj">The object being manipulated</param>
    ''' <param name="initHandle">The type of handle being manipulated</param>
    ''' <param name="ptOrigin">The initial point dragging began at.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal gdiobj As GDIObject, ByVal initHandle As EnumDragHandles, _
                   ByVal ptOrigin As Point)

        MyBase.New(ptOrigin)

        _SelectedObject = gdiobj
        _DragHandleRectangle = _SelectedObject.Bounds
        _OriginalBounds = _SelectedObject.Bounds
        _SelectedHandle = initHandle

        _Rotation = _SelectedObject.Rotation

        If _SelectedHandle = EnumDragHandles.ePointHandle Then

            If TypeOf _SelectedObject Is GDIClosedPath Then

                Dim cp As GDIClosedPath = DirectCast(_SelectedObject, GDIClosedPath)
                Dim iPT As Int32 = cp.Points.IndexOf(cp.Points, cp.LastHitHandle)
                If iPT > 0 Then
                    ipt -= 1
                End If
                _PTOrigin = Point.Round(cp.Points(ipt))

            ElseIf TypeOf _SelectedObject Is GDILine Then

                Dim op As GDIOpenPath = DirectCast(_SelectedObject, GDIOpenPath)
                Dim iPT As Int32 = op.Points.IndexOf(op.Points, op.LastHitHandle)
                If iPT = 0 Then
                    iPT = 1
                Else
                    iPT = 0
                End If

                _PTOrigin = Point.Round(op.Points(ipt))
            ElseIf TypeOf _SelectedObject Is GDIOpenPath Then

                Dim op As GDIOpenPath = DirectCast(_SelectedObject, GDIOpenPath)
                Dim iPT As Int32 = op.Points.IndexOf(op.Points, op.LastHitHandle)
                If iPT > 0 Then
                    ipt -= 1
                End If

                _PTOrigin = Point.Round(op.Points(ipt))
            End If



        End If

        _PTLastPoint = ptOrigin

    End Sub

#End Region


#Region "Point Manipulation Methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets a top handle point appropriately.
    ''' </summary>
    ''' <param name="pt">the last mouse point </param>
    ''' <param name="bshiftdown">Whether shift is being held down</param>
    ''' <returns>The "out" drag handle.  If the user "flips" the drag handles over the 
    ''' surface, the drag handle becomes its mirror opposite.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function setTop(ByVal pt As Point, ByVal bshiftdown As Boolean) As EnumDragHandles

        Dim temprect As New Rectangle(_DragHandleRectangle.X, pt.Y, _DragHandleRectangle.Width, _DragHandleRectangle.Height + (_DragHandleRectangle.Y - pt.Y))

        '  If temprect.Height < 0 Then
        '  Return EnumDragHandles.eTopRight
        If temprect.Height < 0 Then
            Return EnumDragHandles.eBottom
        Else
            _DragHandleRectangle = temprect

            Return EnumDragHandles.eTop
        End If
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets a bottom handle point appropriately
    ''' </summary>
    ''' <param name="pt">the last mouse point </param>
    ''' <param name="bshiftdown">Whether shift is being held down</param>
    ''' <returns>The "out" drag handle.  If the user "flips" the drag handles over the 
    ''' surface, the drag handle becomes its mirror opposite.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function setBottom(ByVal pt As Point, ByVal bshiftdown As Boolean) As EnumDragHandles
        Dim temprect As New Rectangle(_DragHandleRectangle.X, _DragHandleRectangle.Y, _DragHandleRectangle.Width, pt.Y - _DragHandleRectangle.Y)

        '  If temprect.Height < 0 Then
        '  Return EnumDragHandles.eTopRight
        If temprect.Height < 0 Then
            Return EnumDragHandles.eTop
        Else
            _DragHandleRectangle = temprect

            Return EnumDragHandles.eBottom
        End If
    End Function



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets a left drag handle appropriately
    ''' </summary>
    ''' <param name="pt">the last mouse point </param>
    ''' <param name="bshiftdown">Whether shift is being held down</param>
    ''' <returns>The "out" drag handle.  If the user "flips" the drag handles over the 
    ''' surface, the drag handle becomes its mirror opposite.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function setLeft(ByVal pt As Point, ByVal bshiftdown As Boolean) As EnumDragHandles

        Dim temprect As New Rectangle(pt.X, _OriginalBounds.Y, _OriginalBounds.Width + (_OriginalBounds.X - pt.X), _OriginalBounds.Height)

        If temprect.Width < 0 Then
            Return EnumDragHandles.eRight
        Else
            _DragHandleRectangle = temprect

            Return EnumDragHandles.eLeft
        End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets a right drag handle appropriately
    ''' </summary>
    ''' <param name="pt">the last mouse point </param>
    ''' <param name="bshiftdown">Whether shift is being held down</param>
    ''' <returns>The "out" drag handle.  If the user "flips" the drag handles over the 
    ''' surface, the drag handle becomes its mirror opposite.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function setRight(ByVal pt As Point, ByVal bshiftdown As Boolean) As EnumDragHandles
        Dim temprect As New Rectangle(_DragHandleRectangle.X, _DragHandleRectangle.Y, pt.X - _DragHandleRectangle.X, _DragHandleRectangle.Height)

        '  If temprect.Height < 0 Then
        '  Return EnumDragHandles.eTopRight
        If temprect.Width < 0 Then
            Return EnumDragHandles.eLeft
        Else
            _DragHandleRectangle = temprect

            Return EnumDragHandles.eRight
        End If
    End Function


    'Is my mouse point on the "bottom edge" of the resizing rect or along the "right edge" of the resizing rect?

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a drag handle operation on a bottom right corner drag handle
    ''' </summary>
    ''' <param name="pt">the last mouse point </param>
    ''' <param name="bshiftdown">Whether shift is being held down</param>
    ''' <returns>The "out" drag handle.  If the user "flips" the drag handles over the 
    ''' surface, the drag handle becomes its mirror opposite.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function setBottomRight(ByVal pt As Point, ByVal bshiftdown As Boolean) As EnumDragHandles
        Dim temprect As Rectangle

        If bshiftdown AndAlso _DragHandleRectangle.Width > 1 AndAlso _DragHandleRectangle.Height > 1 Then
            If _ProportionalDrag = EnumDragDirection.eLeftRight Then
                Dim propHeight As Double = (_SelectedObject.Bounds.Height * (pt.X - _DragHandleRectangle.X)) / _SelectedObject.Bounds.Width
                temprect = New Rectangle(_DragHandleRectangle.X, _DragHandleRectangle.Y, pt.X - _DragHandleRectangle.X, CInt(propHeight))
            Else
                Dim propWidth As Double = (_SelectedObject.Bounds.Width * (pt.Y - _DragHandleRectangle.Y)) / _SelectedObject.Bounds.Height
                temprect = New Rectangle(_DragHandleRectangle.X, _DragHandleRectangle.Y, CInt(propWidth), pt.Y - _DragHandleRectangle.Y)
            End If
        Else
            temprect = New Rectangle(_DragHandleRectangle.X, _DragHandleRectangle.Y, pt.X - _DragHandleRectangle.X, pt.Y - _DragHandleRectangle.Y)
        End If

        If temprect.Height < 0 Then
            Return EnumDragHandles.eTopRight
        ElseIf temprect.Width < 0 Then
            Return EnumDragHandles.eBottomleft
        Else
            _DragHandleRectangle = temprect

            Return EnumDragHandles.eBottomRight
        End If
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a drag operation on a top left corner drag handle
    ''' </summary>
    ''' <param name="pt">the last mouse point </param>
    ''' <param name="bshiftdown">Whether shift is being held down</param>
    ''' <returns>The "out" drag handle.  If the user "flips" the drag handles over the 
    ''' surface, the drag handle becomes its mirror opposite.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function setTopLeft(ByVal pt As Point, ByVal bshiftdown As Boolean) As EnumDragHandles

        Dim temprect As Rectangle

        Dim scaledRect As Rectangle

        If bshiftdown AndAlso _DragHandleRectangle.Width > 1 AndAlso _DragHandleRectangle.Height > 1 Then
            If _ProportionalDrag = EnumDragDirection.eLeftRight Then
                '(height_old * width_new) / width_old


                Dim width_new As Int32 = _DragHandleRectangle.Width + (_DragHandleRectangle.X - pt.X)
                Dim height_new As Int32 = (_SelectedObject.Bounds.Height * width_new) \ _SelectedObject.Bounds.Width

                'height_new = (height_old * width_new) / width_old
                scaledRect.Width = width_new
                scaledRect.Height = height_new
                scaledRect.Y = _DragHandleRectangle.Bottom - scaledRect.Height
                scaledRect.X = _DragHandleRectangle.Right - scaledRect.Width
                temprect = scaledRect

            Else
                Dim height_new As Int32 = _DragHandleRectangle.Height + (_DragHandleRectangle.Y - pt.Y)
                Dim width_new As Int32 = (_SelectedObject.Bounds.Width * height_new) \ _SelectedObject.Bounds.Height

                scaledRect.Width = width_new
                scaledRect.Height = height_new
                scaledRect.Y = _DragHandleRectangle.Bottom - scaledRect.Height
                scaledRect.X = _DragHandleRectangle.Right - scaledRect.Width
                temprect = scaledRect
            End If

        Else
            temprect = New Rectangle(pt.X, pt.Y, _DragHandleRectangle.Width + (_DragHandleRectangle.X - pt.X), _DragHandleRectangle.Height + (_DragHandleRectangle.Y - pt.Y))
        End If



        If temprect.Height < 0 Then
            Return EnumDragHandles.eBottomleft
        ElseIf temprect.Width < 0 Then
            Return EnumDragHandles.eTopRight
        Else
            _DragHandleRectangle = temprect
            Return EnumDragHandles.eTopLeft
        End If
    End Function




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a drag handle on a top right corner handle
    ''' </summary>
    ''' <param name="pt">the last mouse point </param>
    ''' <param name="bshiftdown">Whether shift is being held down</param>
    ''' <returns>The "out" drag handle.  If the user "flips" the drag handles over the 
    ''' surface, the drag handle becomes its mirror opposite.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function setTopRight(ByVal pt As Point, ByVal bshiftdown As Boolean) As EnumDragHandles
        Dim temprect As Rectangle

        Dim scaledRect As Rectangle

        If bshiftdown AndAlso _DragHandleRectangle.Width > 1 AndAlso _DragHandleRectangle.Height > 1 Then
            If _ProportionalDrag = EnumDragDirection.eLeftRight Then
                '(height_old * width_new) / width_old


                Dim width_new As Int32 = pt.X - _DragHandleRectangle.X
                Dim height_new As Int32 = (_SelectedObject.Bounds.Height * width_new) \ _SelectedObject.Bounds.Width

                'height_new = (height_old * width_new) / width_old
                scaledRect.Width = width_new
                scaledRect.Height = height_new
                scaledRect.Y = _DragHandleRectangle.Bottom - scaledRect.Height
                temprect = New Rectangle(_DragHandleRectangle.X, scaledRect.Y, scaledRect.Width, scaledRect.Height)

            Else
                Dim height_new As Int32 = _DragHandleRectangle.Height + (_DragHandleRectangle.Y - pt.Y)
                Dim width_new As Int32 = (_SelectedObject.Bounds.Width * height_new) \ _SelectedObject.Bounds.Height

                scaledRect.Width = width_new
                scaledRect.Height = height_new
                'the x value should be determined by the new width based on the original right...
                scaledRect.Y = _DragHandleRectangle.Bottom - scaledRect.Height
                temprect = New Rectangle(_DragHandleRectangle.X, scaledRect.Y, scaledRect.Width, scaledRect.Height)
            End If

        Else
            temprect = New Rectangle(_DragHandleRectangle.X, pt.Y, pt.X - _DragHandleRectangle.X, _DragHandleRectangle.Height + (_DragHandleRectangle.Y - pt.Y))
        End If

        If temprect.Height < 0 Then
            Return EnumDragHandles.eBottomRight
        ElseIf temprect.Width < 0 Then
            Return EnumDragHandles.eTopLeft
        Else
            _DragHandleRectangle = temprect
            Return EnumDragHandles.eTopRight
        End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a drag operation on a bottom left corner drag handle
    ''' </summary>
    ''' <param name="pt">the last mouse point </param>
    ''' <param name="bshiftdown">Whether shift is being held down</param>
    ''' <returns>The "out" drag handle.  If the user "flips" the drag handles over the 
    ''' surface, the drag handle becomes its mirror opposite.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function setBottomLeft(ByVal pt As Point, ByVal bshiftdown As Boolean) As EnumDragHandles
        Dim temprect As Rectangle

        Dim scaledRect As Rectangle = New Rectangle(_DragHandleRectangle.X, _DragHandleRectangle.Y, _DragHandleRectangle.Width, _DragHandleRectangle.Height)

        If bshiftdown AndAlso _DragHandleRectangle.Width > 1 AndAlso _DragHandleRectangle.Height > 1 Then
            If _ProportionalDrag = EnumDragDirection.eLeftRight Then
                '(height_old * width_new) / width_old


                Dim width_new As Int32 = _DragHandleRectangle.Width + (_DragHandleRectangle.X - pt.X)
                Dim height_new As Int32 = (_SelectedObject.Bounds.Height * width_new) \ _SelectedObject.Bounds.Width

                'height_new = (height_old * width_new) / width_old
                scaledRect.Width = width_new
                scaledRect.Height = height_new
                scaledRect.X = _DragHandleRectangle.Right - scaledRect.Width
                temprect = New Rectangle(scaledRect.X, _DragHandleRectangle.Y, scaledRect.Width, scaledRect.Height)

            Else
                Dim height_new As Int32 = pt.Y - _DragHandleRectangle.Y
                Dim width_new As Int32 = (_SelectedObject.Bounds.Width * height_new) \ _SelectedObject.Bounds.Height

                scaledRect.Width = width_new
                scaledRect.Height = height_new
                'the x value should be determined by the new width based on the original right...
                scaledRect.X = _DragHandleRectangle.Right - scaledRect.Width
                temprect = New Rectangle(scaledRect.X, _DragHandleRectangle.Y, scaledRect.Width, scaledRect.Height)
            End If

        Else
            temprect = New Rectangle(pt.X, _DragHandleRectangle.Y, _DragHandleRectangle.Width + (_DragHandleRectangle.X - pt.X), pt.Y - _DragHandleRectangle.Y)
        End If


        If temprect.Height < 0 Then
            Return EnumDragHandles.eTopLeft

        ElseIf temprect.Width < 0 Then
            Return EnumDragHandles.eBottomRight

        Else
            _DragHandleRectangle = temprect
            Return EnumDragHandles.eBottomleft
        End If


    End Function


#End Region


#Region "Base Class Implementors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws the drag handle tool to the surface
    ''' </summary>
    ''' <param name="g">Graphics context to draw against</param>
    ''' <param name="fScale">Zoom factor of the current surface.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overloads Overrides Sub draw(ByVal g As System.Drawing.Graphics, ByVal fScale As Single)
        _PenOutline.Width = 1 / fScale

        App.GraphicsManager.BeginScaledView(g)

        Dim gContainer As Drawing2D.GraphicsContainer
        Dim mtx As Drawing2D.Matrix

        If _Rotation > 0 Then
            gContainer = g.BeginContainer()
            mtx = g.Transform()
            mtx.RotateAt(_Rotation, _SelectedObject.RotationPoint)
            g.Transform = mtx
        End If

        If Not _SelectedHandle = EnumDragHandles.ePointHandle Then
            g.DrawRectangle(_PenOutline, _DragHandleRectangle)
        End If

        'Reset the graphics object to its original value
        If _Rotation > 0 Then
            g.EndContainer(gContainer)
            mtx.Dispose()
            mtx = Nothing
            gContainer = Nothing
        End If

        App.GraphicsManager.EndScaledView(g)
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends the draghandle tool operation
    ''' </summary>
    ''' <param name="bShiftDown">Whether shift is being held down or not</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub EndTool(ByVal bShiftDown As Boolean)
        If Not Point.op_Equality(_PTOrigin, _PTLastPoint) Then
            If Not _SelectedHandle = EnumDragHandles.ePointHandle Then
                If _Rotation = 0 Then
                    _SelectedObject.Bounds = _DragHandleRectangle
                Else
                    _SelectedObject.Bounds = Rectangle.Round(translateRotatedDragRect(_DragHandleRectangle))
                End If

            End If

            MDIMain.ActiveDocument.recordHistory("Dragged handle")
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the drag handle tool based upon the last mouse point
    ''' </summary>
    ''' <param name="ptPoint">Point, snapped appropriately</param>
    ''' <param name="btn">Button held down</param>
    ''' <param name="bShiftDown">Whether shift is down or not</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub UpdateTool(ByVal ptPoint As System.Drawing.Point, ByVal btn As MouseButtons, ByVal bShiftDown As Boolean, ByVal zoom As Single)
        If Not _SelectedObject Is Nothing Then


            If bShiftDown <> _ShiftDownLastUpdate Then
                _DragHandleRectangle = _SelectedObject.Bounds
                getdirection(ptPoint)
                _ShiftDownLastUpdate = bShiftDown
            End If

            If _Rotation > 0 Then
                ptPoint = MapPoint(ptPoint)
            End If

            Select Case _SelectedHandle
                Case EnumDragHandles.eLeft

                    _SelectedHandle = setLeft(ptPoint, bShiftDown)

                Case EnumDragHandles.eTopLeft
                    _SelectedHandle = setTopLeft(ptPoint, bShiftDown)

                Case EnumDragHandles.eTop
                    _SelectedHandle = setTop(ptPoint, bShiftDown)

                Case EnumDragHandles.eTopRight
                    _SelectedHandle = setTopRight(ptPoint, bShiftDown)

                Case EnumDragHandles.eRight
                    _SelectedHandle = setRight(ptPoint, bShiftDown)

                Case EnumDragHandles.eBottomRight
                    _SelectedHandle = setBottomRight(ptPoint, bShiftDown)

                Case EnumDragHandles.eBottom
                    _SelectedHandle = setBottom(ptPoint, bShiftDown)

                Case EnumDragHandles.eBottomleft
                    _SelectedHandle = setBottomLeft(ptPoint, bShiftDown)

                Case EnumDragHandles.ePointHandle
                    handlePointSet(ptPoint, bShiftDown)

            End Select

        End If

        _PTLastPoint = ptPoint
    End Sub

#End Region


#Region "Local Members"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Maps the mouse position from rotated coordinate space to normal space.
    ''' </summary>
    ''' <param name="pt">Point to map from rotated coordinated space to normal space</param>
    ''' <returns>The point in normal coordinate space.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function MapPoint(ByVal pt As Point) As Point


        Dim mtx As New Drawing2D.Matrix

        mtx.RotateAt(-_Rotation, _SelectedObject.RotationPoint)

        Dim transform() As Point = {pt}
        mtx.TransformPoints(transform)

        Return transform(0)
        mtx.Dispose()

    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the current proportional drag direction based upon the last recorded point.
    ''' Used for when dragging handles with shift held.
    ''' </summary>
    ''' <param name="ptsnapped">Point snapped appropriately</param>
    ''' -----------------------------------------------------------------------------
    Private Sub getdirection(ByVal ptsnapped As Point)

        If Math.Abs(ptsnapped.X - _PTLastPoint.X) > Math.Abs(ptsnapped.Y - _PTLastPoint.Y) Then
            _ProportionalDrag = EnumDragDirection.eLeftRight
        Else
            _ProportionalDrag = EnumDragDirection.eUpDown
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Translates a drag rectangle into standard coordinates.  Used to assign 
    ''' the graphical bounds to the actual, unrotated object bounds.
    ''' </summary>
    ''' <param name="rect">Rectangle to translate </param>
    ''' <returns></returns>
    ''' -----------------------------------------------------------------------------
    Private Function translateRotatedDragRect(ByVal rect As Rectangle) As RectangleF
        '1) Apply the current rotation point to the drag rectangle's edges
        Dim mtx As New System.Drawing.Drawing2D.Matrix
        Dim gPath As Drawing2D.GraphicsPath = New Drawing2D.GraphicsPath

        mtx.RotateAt(_Rotation, _SelectedObject.RotationPoint, Drawing2D.MatrixOrder.Append)

        Dim pts() As PointF = {New PointF(rect.Left, rect.Top), New PointF(rect.Right, rect.Top), New PointF(rect.Right, rect.Bottom), New PointF(rect.X, rect.Bottom)}

        mtx.TransformPoints(pts)

        'Create a temporary path to get the bounds
        gPath.AddRectangle(rect)
        gPath.Transform(mtx)

        Dim rectBounds As RectangleF = gPath.GetBounds()

        'Get the center of the graphic path bounds
        Dim newCenter As PointF = New PointF(rectBounds.X + rectBounds.Width / 2, rectBounds.Y + rectBounds.Height / 2)
        mtx.Reset()

        'Unrotate, so to speak, around the new center point
        mtx.RotateAt(360 - _Rotation, newCenter, Drawing2D.MatrixOrder.Append)
        mtx.TransformPoints(pts)

        mtx.Dispose()

        gPath.Dispose()


        'Return a rectangle based on the unrotated points
        Return New RectangleF(pts(0).X, pts(0).Y, rect.Width, rect.Height)



    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles drag handles for when the drag handle type of point is "Point" or "Curve"
    ''' In these cases, its not really a drag handle, but a point move operation.
    ''' </summary>
    ''' <param name="ptPoint">Point to update the point to</param>
    ''' <param name="bShiftDown">Whether shift is being held down or not.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub handlePointSet(ByVal ptPoint As Point, ByVal bShiftDown As Boolean)
        Dim ptShifted As Point = ptPoint

        If bShiftDown Then
            ptShifted = Utility.ShiftedDownPoint(_PTOrigin, ptPoint)
        End If

        _SelectedObject.handlePointSet(ptShifted)

    End Sub

#End Region

End Class
