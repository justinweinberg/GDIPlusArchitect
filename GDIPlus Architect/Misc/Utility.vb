
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : Formulas
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a wrapper for complex mathematical formulas used in multiple places
''' within the application.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class Utility


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invokes the color picker
    ''' </summary>
    ''' <param name="initColor">Initial color to set the picker to.</param>
    ''' <returns>The selected color or an empty color if the user elected ot 
    ''' cancel the dialog.</returns>
    ''' -----------------------------------------------------------------------------
    Public Shared Function PickColor(ByVal initColor As Color) As Color

        Dim outcolor As Color = Color.Empty

        If App.Options.UseCustomColorPicker Then
            Dim picker As New dgColorPicker(initColor)
            '            picker.Color = initColor

            Dim iresp As DialogResult = picker.ShowDialog()

            If iresp = DialogResult.OK Then
                outcolor = picker.Color
            End If
        Else
            Dim dgcolor As New ColorDialog
            Dim iresp As DialogResult = dgcolor.ShowDialog()
            If iresp = DialogResult.OK Then
                outcolor = dgcolor.Color
            End If
        End If

        Return outcolor
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the distance between two points.
    ''' </summary>
    ''' <param name="pt1">The first point</param>
    ''' <param name="pt2">The second point</param>
    ''' <returns>The distance between pt1 and pt2</returns>
    ''' <remarks> The distance between two points is: 
    '''  sqrt( (x1 - x2)^2  +  (y1 - y2)^2  )
    ''' </remarks>
        ''' -----------------------------------------------------------------------------
    Private Shared Function DistanceBetweenPoints(ByVal pt1 As Point, ByVal pt2 As Point) As Double
        Return Math.Sqrt((pt1.X - pt2.X) ^ 2 + (pt1.Y - pt2.Y) ^ 2)
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the point that is appropriate when Shift is being held down while 
    ''' manipulating a point to point type of shape.  
    ''' 
    ''' The user expects when shift is being held that 45 degree lines will be drawn.
    ''' 
    ''' This functionality determines to which 45 degree angle the point should snap.
    ''' </summary>
    ''' <param name="ptorigin">the origin point from the shift</param>
    ''' <param name="ptsnapped">The mouse position with a snapped to grid operation 
    ''' already performed</param>
    ''' <returns>A point appropriate for a shift held operation</returns>
    ''' -----------------------------------------------------------------------------
    Public Shared Function ShiftedDownPoint(ByVal ptorigin As Point, ByVal ptsnapped As Point) As Point
        Dim radius As Int32 = CInt(Utility.DistanceBetweenPoints(ptorigin, ptsnapped))

        Dim ptTop As Point = New Point(ptorigin.X, ptorigin.Y - radius)
        Dim ptBottom As Point = New Point(ptorigin.X, ptorigin.Y + radius)
        Dim ptLeft As Point = New Point(ptorigin.X - radius, ptorigin.Y)
        Dim ptRight As Point = New Point(ptorigin.X + radius, ptorigin.Y)

        Dim triside As Int32 = CInt(radius / 1.4142)

        Dim ptdiagTopRight As Point = New Point(ptorigin.X + triside, ptorigin.Y - triside)
        Dim ptdiagTopLeft As Point = New Point(ptorigin.X - triside, ptorigin.Y - triside)
        Dim ptdiagBottomleft As Point = New Point(ptorigin.X - triside, ptorigin.Y + triside)
        Dim ptdiagBottomRight As Point = New Point(ptorigin.X + triside, ptorigin.Y + triside)



        Dim pts() As Point = New Point() {ptTop, ptLeft, ptRight, ptBottom, ptdiagTopRight, ptdiagTopLeft, ptdiagBottomleft, ptdiagBottomRight}


        Dim ptClosest As Point
        'assign an initial value 
        Dim leastDistance As Double = DistanceBetweenPoints(ptsnapped, pts(0))
        ptClosest = pts(0)


        For i As Int32 = 0 To pts.Length - 1
            Dim templeastDistance As Double = DistanceBetweenPoints(ptsnapped, pts(i))
            If templeastDistance < leastDistance Then
                ptClosest = pts(i)
                leastDistance = templeastDistance
            End If
        Next

        Return ptClosest




    End Function

End Class
