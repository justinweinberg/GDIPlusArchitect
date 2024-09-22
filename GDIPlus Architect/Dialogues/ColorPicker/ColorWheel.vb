Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' ''' Class	 : ColorWheel
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Implements a color wheel chooser for custom color picking.
''' </summary>
''' <remarks>Credit to Ken Getz for this code!
''' </remarks>
''' -----------------------------------------------------------------------------
Friend Class ColorWheel
    Implements IDisposable

#Region "Type Declarations"


    ' Keep track of the current mouse state. 
    Public Enum MouseState
        MouseUp
        ClickOnColor
        DragInColor
        ClickOnBrightness
        DragInBrightness
        ClickOutsideRegion
        DragOutsideRegion
    End Enum

    ' The code needs to convert back and forth between 
    ' degrees and radians. There are 2*PI radians in a 
    ' full circle, and 360 degrees. This constant allows
    ' you to convert back and forth.
    Private Const CONST_DEGREES_PER_RADIAN As Double = 180.0 / Math.PI

    ' COLOR_COUNT represents the number of distinct colors
    ' used to create the circular gradient. Its value 
    ' is somewhat arbitrary -- change this to 6, for 
    ' example, to see what happens. 1536 (6 * 256) seems 
    ' a good compromise -- it's enought to get a full 
    ' range of colors, but it doesn't overwhelm the processor
    ' attempting to generate the image. The color wheel
    ' contains 6 sections, and each section displays 
    ' 256 colors. Seems like a reasonable compromise.
    Private Const CONST_COLOR_COUNT As Int32 = 6 * 256

#End Region

#Region "Event Declarations"
    Public Event ColorChanged(ByVal sender As Object, ByVal e As ColorChangedEventArgs)
#End Region


    ' These resources should be disposed
    ' of when you're done with them.
    Private _colorRegion As Region
    Private _brightnessRegion As Region

    'A bitmap containing the wheel image bitmap.  This allows the bitmap to only be created 
    'once and then stored here, saving quite a bit of work
    Private _WheelBitmap As Bitmap

    Private _currentState As MouseState = MouseState.MouseUp

    Private _CenterPoint As Point
    Private _WheelRadius As Int32


    Private _WheelRectangle As Rectangle

    'Rectangle in which the brightness selection is drawn
    Private _BrightnessRectangle As Rectangle


    'Rectangle into which the selected color is painted 
    Private _SelectedColorRectangle As Rectangle
    Private _BrightnessX As Int32
    Private _BrightnessScaling As Double


    'The selected color from the wheel + the brightness rectangle 
    Private _SelectedColor As Color = Color.White
    'The selected color at full brightness
    Private _FullColor As Color = Color.White


    'Red Blue Green
    Private _RGB As ColorHandler.RGB
    Private _HSV As ColorHandler.HSV

    ' Locations for the two "pointers" on the form.
    Private _colorPoint As Point
    Private _BrightnessPoint As Point

    Private _Alpha As Int32

    Private _Brightness As Int32
    Private _BrightnessMin As Int32
    Private _BrightnessMax As Int32

    Protected Overridable Sub OnColorChanged( _
    ByVal RGB As ColorHandler.RGB, ByVal HSV As ColorHandler.HSV)
        Dim e As New ColorChangedEventArgs(RGB, HSV)
        RaiseEvent ColorChanged(Me, e)
    End Sub

    Private Sub Dispose() Implements IDisposable.Dispose
        ' Dispose of graphic resources
        If Not _WheelBitmap Is Nothing Then
            _WheelBitmap.Dispose()
        End If
        If Not _colorRegion Is Nothing Then
            _colorRegion.Dispose()
        End If
        If Not _brightnessRegion Is Nothing Then
            _brightnessRegion.Dispose()
        End If

    End Sub

    Public Sub New( _
    ByVal colorRectangle As Rectangle, _
    ByVal brightnessRectangle As Rectangle, _
    ByVal selectedColorRectangle As Rectangle)
        Trace.WriteLineIf(App.TraceOn, "ColorWheel.New")

        ' Caller must provide locations for color wheel
        ' (colorRectangle), brightness "strip" (brightnessRectangle)
        ' and location to display selected color (selectedColorRectangle).

        Dim path As GraphicsPath

        Try
            ' Store away locations for later use. 
            _WheelRectangle = colorRectangle
            _BrightnessRectangle = brightnessRectangle
            _SelectedColorRectangle = selectedColorRectangle

            ' Calculate the center of the circle.
            ' Start with the location, then offset
            ' the point by the radius.
            ' Use the smaller of the width and height of
            ' the colorRectangle value.
            _WheelRadius = Math.Min( _
            colorRectangle.Width, colorRectangle.Height) \ 2
            _CenterPoint = colorRectangle.Location
            _CenterPoint.Offset(_WheelRadius, _WheelRadius)

            ' Start the pointer in the center.
            _colorPoint = _CenterPoint

            ' Create a region corresponding to the color circle.
            ' Code uses this later to determine if a specified
            ' point is within the region, using the IsVisible 
            ' method.
            path = New GraphicsPath
            path.AddEllipse(colorRectangle)
            _colorRegion = New Region(path)

            ' Set the range for the brightness selector.
            _BrightnessMin = _BrightnessRectangle.Top
            _BrightnessMax = _BrightnessRectangle.Bottom

            ' Create a region corresponding to the
            ' brightness rectangle, with a little extra 
            ' "breathing room". 
            With brightnessRectangle
                path = New GraphicsPath
                path.AddRectangle( _
                New Rectangle(.Left, .Top - 10, _
                .Width + 10, .Height + 20))
            End With
            ' Create region corresponding to brightness
            ' rectangle. Later code uses this to 
            ' determine if a specified point is within
            ' the region, using the IsVisible method.
            _brightnessRegion = New Region(path)

            ' Set the location for the brightness indicator "marker".
            ' Also calculate the scaling factor, scaling the height
            ' to be between 0 and 255. 
            _BrightnessX = brightnessRectangle.Left + _
            brightnessRectangle.Width
            _BrightnessScaling = 255 / _
            (_BrightnessMax - _BrightnessMin)

            ' Calculate the location of the brightness
            ' pointer. Assume it's at the highest position.
            _BrightnessPoint = New Point(_BrightnessX, _BrightnessMax)

            ' Create the bitmap that contains the circular gradient.
            _WheelBitmap = CreateWheelBitmap()

        Finally
            If Not path Is Nothing Then
                path.Dispose()
            End If
        End Try
    End Sub

    Public Sub SetMouseUp()
        ' Indicate that the user has
        ' released the mouse.
        _currentState = MouseState.MouseUp
    End Sub

    Public Overloads Sub Draw(ByVal g As Graphics, ByVal HSV As ColorHandler.HSV)
        ' Given HSV values, update the screen.
        _HSV = HSV
        CalcCoordsAndUpdate(_HSV)
        UpdateDisplay(g)
    End Sub

    Public Overloads Sub Draw(ByVal g As Graphics, ByVal RGB As ColorHandler.RGB)
        ' Given RGB values, calculate HSV and then update the screen.
        _HSV = ColorHandler.RGBtoHSV(RGB)
        CalcCoordsAndUpdate(_HSV)
        UpdateDisplay(g)
    End Sub

    Public Overloads Sub Draw(ByVal g As Graphics, ByVal mousePoint As Point)
        ' You've moved the mouse. 
        ' Now update the screen to match.

        Dim distance As Double
        Dim degrees As Integer
        Dim delta As Point
        Dim newColorPoint As Point
        Dim newBrightnessPoint As Point
        Dim newPoint As Point

        ' Keep track of the previous color pointer point, 
        ' so you can put the mouse there in case the 
        ' user has clicked outside the circle.
        newColorPoint = _colorPoint
        newBrightnessPoint = _BrightnessPoint


        If _currentState = MouseState.MouseUp Then
            If Not mousePoint.IsEmpty Then
                If _colorRegion.IsVisible(mousePoint) Then
                    ' Is the mouse point within the color circle?
                    ' If so, you just clicked on the color wheel.
                    _currentState = MouseState.ClickOnColor
                ElseIf _brightnessRegion.IsVisible(mousePoint) Then
                    ' Is the mouse point within the brightness area?
                    ' You clicked on the brightness area.
                    _currentState = MouseState.ClickOnBrightness
                Else
                    ' Clicked outside the color and the brightness
                    ' regions. In that case, just put the 
                    ' pointers back where they were.
                    _currentState = MouseState.ClickOutsideRegion
                End If
            End If
        End If

        Select Case _currentState
            Case MouseState.ClickOnBrightness, MouseState.DragInBrightness
                ' Calculate new color information
                ' based on the brightness, which may have changed.
                newPoint = mousePoint
                If newPoint.Y < _BrightnessMin Then
                    newPoint.Y = _BrightnessMin
                ElseIf newPoint.Y > _BrightnessMax Then
                    newPoint.Y = _BrightnessMax
                End If
                newBrightnessPoint = New Point(_BrightnessX, newPoint.Y)
                _Brightness = CInt((_BrightnessMax - newPoint.Y) * _BrightnessScaling)
                _HSV.Value = _Brightness
                _RGB = ColorHandler.HSVtoRGB(_HSV)

            Case MouseState.ClickOnColor, MouseState.DragInColor
                ' Calculate new color information
                ' based on selected color, which may have changed.
                newColorPoint = mousePoint

                ' Calculate x and y distance from the center,
                ' and then calculate the angle corresponding to the
                ' new location.
                delta = New Point( _
                mousePoint.X - _CenterPoint.X, _
                mousePoint.Y - _CenterPoint.Y)
                degrees = CalcDegrees(delta)

                ' Calculate distance from the center to the new point 
                ' as a fraction of the radius. Use your old friend, 
                ' the Pythagorean theorem, to calculate this value.
                distance = Math.Sqrt( _
                delta.X * delta.X + delta.Y * delta.Y) / _WheelRadius

                If _currentState = MouseState.DragInColor Then
                    If distance > 1 Then
                        ' Mouse is down, and outside the circle, but you 
                        ' were previously dragging in the color circle. 
                        ' What to do?
                        ' In that case, move the point to the edge of the 
                        ' circle at the correct angle.
                        distance = 1
                        newColorPoint = PolarToRect(degrees, _WheelRadius, _CenterPoint)
                    End If
                End If

                ' Calculate the new HSV and RGB values.
                _HSV.Hue = CInt(degrees * 255 / 360)
                _HSV.Saturation = CInt(distance * 255)
                _HSV.Value = _Brightness
                _RGB = ColorHandler.HSVtoRGB(_HSV)
                _FullColor = ColorHandler.HSVtoColor( _
                _HSV.Hue, _HSV.Saturation, 255)
        End Select

        _SelectedColor = ColorHandler.HSVtoColor(_HSV)

        _SelectedColor = Color.FromArgb(_Alpha, _SelectedColor.R, _SelectedColor.G, _SelectedColor.B)
        ' Raise an event back to the parent form,
        ' so the form can update any UI it's using 
        ' to display selected color values.
        OnColorChanged(_RGB, _HSV)

        ' On the way out, set the new state.
        Select Case _currentState
            Case MouseState.ClickOnBrightness
                _currentState = MouseState.DragInBrightness
            Case MouseState.ClickOnColor
                _currentState = MouseState.DragInColor
            Case MouseState.ClickOutsideRegion
                _currentState = MouseState.DragOutsideRegion
        End Select

        ' Store away the current points for next time.
        _colorPoint = newColorPoint
        _BrightnessPoint = newBrightnessPoint

        ' Draw the gradients and points. 
        UpdateDisplay(g)
    End Sub

    Private Function CalcBrightnessPoint(ByVal brightness As Integer) As Point
        ' Take the value for brightness (0 to 255), scale to the 
        ' scaling used in the brightness bar, then add the value 
        ' to the bottom of the bar. Return the correct point at which 
        ' to display the brightness pointer.
        Return New Point( _
        _BrightnessX, _
        CInt(_BrightnessMax - brightness / _BrightnessScaling))
    End Function

    Private Sub UpdateDisplay(ByVal g As Graphics)
        ' Update the gradients, and place the 
        ' pointers correctly based on colors and 
        ' brightness.

        ' Draw the "selected color" rectangle.
        Dim selectedBrush As New SolidBrush(_SelectedColor)


        ' Draw the color wheel into the wheel rectangle
        g.DrawImage(_WheelBitmap, _WheelRectangle)

        g.FillRectangle(selectedBrush, _SelectedColorRectangle)

        ' Draw the "brightness" rectangle.
        DrawBrightnessRect(g, _FullColor)

        ' Draw the two pointers.
        DrawColorPointer(g, _colorPoint)
        DrawBrightnessPointer(g, _BrightnessPoint)

        selectedBrush.Dispose()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the brightness selecter portion of the color picker
    ''' </summary>
    ''' <param name="g">Graphics context to draw against</param>
    ''' <param name="TopColor"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub DrawBrightnessRect(ByVal g As Graphics, ByVal TopColor As Color)

        Dim lgb As New LinearGradientBrush( _
        _BrightnessRectangle, TopColor, Color.Black, LinearGradientMode.Vertical)

        g.FillRectangle(lgb, _BrightnessRectangle)



    End Sub


    Private Sub CalcCoordsAndUpdate(ByVal HSV As ColorHandler.HSV)
        ' Convert color to real-world coordinates and then calculate
        ' the various points. HSV.Hue represents the degrees (0 to 360), 
        ' HSV.Saturation represents the radius. 
        ' This procedure doesn't draw anything--it simply 
        ' updates class-level variables. The UpdateDisplay
        ' procedure uses these values to update the screen.

        ' Given the angle (HSV.Hue), and distance from 
        ' the center (HSV.Saturation), and the center, 
        ' calculate the point corresponding to 
        ' the selected color, on the color wheel.
        _colorPoint = PolarToRect( _
        HSV.Hue / 255 * 360, _
        HSV.Saturation / 255 * _WheelRadius, _
        _CenterPoint)

        ' Given the brightness (HSV.Value), calculate the 
        ' point corresponding to the brightness indicator.
        _BrightnessPoint = CalcBrightnessPoint(HSV.Value)

        ' Store information about the selected color.
        _Brightness = HSV.Value
        _SelectedColor = ColorHandler.HSVtoColor(HSV)

        _RGB = ColorHandler.HSVtoRGB(HSV)

        _SelectedColor = Color.FromArgb(_Alpha, _SelectedColor.R, _SelectedColor.G, _SelectedColor.B)

        ' The full color is the same as HSV, except that the 
        ' brightness is set to full (255). This is the top-most
        ' color in the brightness gradient.
        _FullColor = ColorHandler.HSVtoColor(HSV.Hue, HSV.Saturation, 255)
    End Sub


    Private Function CalcDegrees(ByVal pt As Point) As Integer
        Dim degrees As Int32

        If pt.X = 0 Then
            ' The point is on the y-axis. Determine whether 
            ' it's above or below the x-axis, and return the 
            ' corresponding angle. Note that the orientation of the
            ' y-coordinate is backwards. That is, A positive Y value 
            ' indicates a point BELOW the x-axis.
            If pt.Y > 0 Then
                degrees = 270
            Else
                degrees = 90
            End If
        Else
            ' This value needs to be multiplied
            ' by -1 because the y-coordinate
            ' is opposite from the normal direction here.
            ' That is, a y-coordinate that's "higher" on 
            ' the form has a lower y-value, in this coordinate
            ' system. So everything's off by a factor of -1 when
            ' performing the ratio calculations.
            degrees = CInt(-Math.Atan(pt.Y / pt.X) * CONST_DEGREES_PER_RADIAN)

            ' If the x-coordinate of the selected point
            ' is to the left of the center of the circle, you 
            ' need to add 180 degrees to the angle. ArcTan only
            ' gives you a value on the right-hand side of the circle.
            If pt.X < 0 Then
                degrees += 180
            End If

            ' Ensure that the return value is 
            ' between 0 and 360.
            degrees = (degrees + 360) Mod 360
        End If
        Return degrees
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Function CreateWheelBitmap() As Bitmap
        Dim newGraphics As Graphics
        Dim pgb As PathGradientBrush

        ' Create a new PathGradientBrush, supplying
        ' an array of points created by calling
        ' the GetPoints method.
        pgb = New PathGradientBrush( _
        GetPoints(_WheelRadius, New Point(_WheelRadius, _WheelRadius)))

        ' Set the various properties. Note the SurroundColors
        ' property, which contains an array of points, 
        ' in a one-to-one relationship with the points
        ' that created the gradient.
        With pgb
            .CenterColor = Color.White
            .CenterPoint = New PointF(_WheelRadius, _WheelRadius)
            .SurroundColors = GetColors()
        End With

        ' Create a new bitmap containing
        ' the color wheel gradient, so the 
        ' code only needs to do all this 
        ' work once. Later code uses the bitmap
        ' rather than recreating the gradient.
        Dim wheelBMP As New Bitmap( _
        _WheelRectangle.Width, _WheelRectangle.Height, _
        PixelFormat.Format32bppArgb)

        newGraphics = Graphics.FromImage(wheelBMP)
        newGraphics.FillEllipse(pgb, 0, 0, _
        wheelBMP.Width, _WheelRectangle.Height)


        pgb.Dispose()
        newGraphics.Dispose()

        Return wheelBMP

        Return wheelBMP
    End Function

    Private Function GetColors() As Color()
        ' Create an array of COLOR_COUNT
        ' colors, looping through all the 
        ' hues between 0 and 255, broken
        ' into COLOR_COUNT intervals. HSV is
        ' particularly well-suited for this, 
        ' because the only value that changes
        ' as you create colors is the Hue.
        Dim Colors(CONST_COLOR_COUNT - 1) As Color

        For i As Int32 = 0 To CONST_COLOR_COUNT - 1
            Colors(i) = ColorHandler.HSVtoColor(i * 255 \ CONST_COLOR_COUNT, 255, 255)
        Next

        Return Colors
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="radius"></param>
    ''' <param name="centerPoint"></param>
    ''' <returns></returns>
    ''' -----------------------------------------------------------------------------
    Private Function GetPoints(ByVal radius As Double, ByVal centerPoint As Point) As Point()

        ' Generate the array of points that describe
        ' the locations of the COLOR_COUNT colors to be 
        ' displayed on the color wheel.
        Dim Points(CONST_COLOR_COUNT - 1) As Point

        For i As Int32 = 0 To CONST_COLOR_COUNT - 1
            Points(i) = PolarToRect(i * 360 / CONST_COLOR_COUNT, radius, centerPoint)
        Next

        Return Points

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts a coordinate in polar space to rectangular space
    ''' </summary>
    ''' <param name="degrees"></param>
    ''' <param name="radius"></param>
    ''' <param name="centerPoint"></param>
    ''' <returns></returns>
    ''' -----------------------------------------------------------------------------
    Private Function PolarToRect(ByVal degrees As Double, ByVal radius As Double, ByVal center As Point) As Point

        ' Given the center of a circle and its radius, along
        ' with the angle corresponding to the point, 
        ' find the coordinates.  In other words, convert 
        ' from polar to rectangular coordinates.
        Dim radians As Double = degrees / CONST_DEGREES_PER_RADIAN

        Return New Point(center.X + CInt(Math.Floor(radius * Math.Cos(radians))), _
        center.Y - CInt(Math.Floor(radius * Math.Sin(radians))))
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws the selected color square on the color wheel
    ''' </summary>
    ''' <param name="g">Graphics context to draw against </param>
    ''' <param name="pt">Top left point of the rectangle</param>
    ''' -----------------------------------------------------------------------------
    Private Sub DrawColorPointer(ByVal g As Graphics, ByVal pt As Point)
        Const SIZE As Integer = 6

        g.DrawRectangle(Pens.Black, pt.X - SIZE \ 2, pt.Y - SIZE \ 2, SIZE, SIZE)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws the triangular brightness pointer 
    ''' </summary>
    ''' <param name="g">graphics context to draw against</param>
    ''' <param name="pt">Top left point of the triangle</param>
    ''' -----------------------------------------------------------------------------
    Private Sub DrawBrightnessPointer(ByVal g As Graphics, ByVal pt As Point)
        Const HEIGHT As Integer = 10
        Const WIDTH As Integer = 7

        Dim Points() As Point = {pt, New Point(pt.X + WIDTH, pt.Y + HEIGHT \ 2), _
        New Point(pt.X + WIDTH, pt.Y - HEIGHT \ 2)}

        g.FillPolygon(Brushes.Black, Points)
    End Sub



#Region "Property Accessors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the alpha value of the color 
    ''' </summary>
    ''' <value>Ain int32 between 0 and 255 representing the alpha value</value>
    ''' -----------------------------------------------------------------------------
    Public Property Alpha() As Int32
        Get
            Return _Alpha
        End Get
        Set(ByVal Value As Int32)
            _Alpha = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the selected color
    ''' </summary>
    ''' <value>The selected color</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Color() As Color
        Get
            Return _SelectedColor
        End Get
    End Property
#End Region


End Class

