
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : DragTool
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Responsible for drag operations, where an object or objects are dragged to 
''' another position on the canvas.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class DragTool
    Inherits GDITool

#Region "Local Fields"
    '''<summary>Latest bounding rectangle</summary>
    Private _Bounds As Rectangle

#End Region

#Region "Constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new DragTool based on an origin pointof the drag operation
    ''' </summary>
    ''' <param name="ptSnapped"></param>
    ''' -----------------------------------------------------------------------------
    Public sub New(ByVal ptSnapped As Point)
        MyBase.New(ptSnapped)

        _PTOrigin = ptSnapped
        _PTLastPoint = ptSnapped

        MDIMain.ActiveDocument.Selected.initDragPoint(_PTOrigin)
    End Sub


#End Region

#Region "Base Class Implementors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the dragtool to a surface
    ''' </summary>
    ''' <param name="g">Graphics context to draw against</param>
    ''' <param name="fScale">Current Zoom factor of the surface</param>
    ''' -----------------------------------------------------------------------------
    Friend Overloads Overrides Sub Draw(ByVal g As System.Drawing.Graphics, ByVal fScale As Single)
        _PenOutline.Width = 1 / fScale

        Dim gCt As Drawing2D.GraphicsContainer = g.BeginContainer
        g.ScaleTransform(fScale, fScale)

        g.DrawRectangle(_PenOutline, _Bounds)
        g.EndContainer(gCt)

        gCt = Nothing

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends a drag tool operation
    ''' </summary>
    ''' <param name="bShiftDown">Whether shift is held down or not</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub EndTool(ByVal bShiftDown As Boolean)
        If Not Point.op_Equality(_PTOrigin, _PTLastPoint) Then
            MDIMain.ActiveDocument.Selected.NotifyDragEnd()
            MDIMain.ActiveDocument.recordHistory("Dragged objects")
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the drag tool based upon the latest mouse point 
    ''' </summary>
    ''' <param name="ptPoint">Mouse point, snapped appropriately</param>
    ''' <param name="btn">Which button is being held down</param>
    ''' <param name="bShiftDown">Whether shift is held down or not</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub UpdateTool(ByVal ptPoint As System.Drawing.Point, ByVal btn As MouseButtons, ByVal bShiftDown As Boolean, ByVal zoom As Single)
        _PTLastPoint = ptPoint



        MDIMain.ActiveDocument.Selected.updateDragRect(ptPoint)

        Dim temprect As Rectangle = MDIMain.ActiveDocument.Selected.CurrentBounds
        Dim arrPoints() As Point = New Point() {New Point(temprect.X, temprect.Y), New Point(temprect.Right, temprect.Y), New Point(temprect.Right, temprect.Bottom), New Point(temprect.X, temprect.Bottom)}

        Dim arrPointsSnap(3) As Point

        Dim bsnappedx As Boolean = False
        Dim bsnappedy As Boolean = False

        For i As Int32 = 0 To 3
            arrPointsSnap(i) = App.MDIMain.snapPoint(arrPoints(i), bsnappedx, bsnappedy)

            If bsnappedx = False Then
                arrPointsSnap(i).X = 100000
            End If

            If bsnappedy = False Then
                arrPointsSnap(i).Y = 100000
            End If
        Next


        Dim iClosestX As Int32 = -1
        Dim iClosestY As Int32 = -1

        Dim iClosestXVal As Int32 = 90000
        Dim iClosestYVal As Int32 = 90000


        For i As Int32 = 0 To 3
            If Math.Abs(arrPointsSnap(i).X - arrPoints(i).X) < iClosestXVal Then
                iClosestXVal = arrPointsSnap(i).X - arrPoints(i).X
                iClosestX = i
            End If

            If Math.Abs(arrPointsSnap(i).Y - arrPoints(i).Y) < iClosestYVal Then
                iClosestYVal = arrPointsSnap(i).Y - arrPoints(i).Y
                iClosestY = i
            End If
        Next

        If iClosestX > -1 Then
            ptPoint.X += (arrPointsSnap(iClosestX).X - arrPoints(iClosestX).X)
        End If

        If iClosestY > -1 Then
            ptPoint.Y += (arrPointsSnap(iClosestY).Y - arrPoints(iClosestY).Y)
        End If



        MDIMain.ActiveDocument.Selected.updateDragRect(ptPoint)
        _Bounds = MDIMain.ActiveDocument.Selected.CurrentBounds
        MDIMain.ActiveDocument.Selected.UpdateDraggedPositions(ptPoint)

    End Sub
#End Region

End Class
