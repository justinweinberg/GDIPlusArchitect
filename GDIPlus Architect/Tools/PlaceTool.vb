Imports GDIObjects

Friend Class PlaceTool
    Inherits GDITool

#Region "Local Fields"

    '''<summary>Bounds of the image to place</summary>
    Private _Bounds As Rectangle

    '''<summary>Last point recorded thus far for the mouse</summary>
    Private _PTEnd As Point

    '''<summary>The image to place</summary>
    Private _Image As Bitmap
    '''<summary>Path to the image to place</summary>
    Private _ImageSource As String

#End Region

#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the PlaceTool given an origin to begin creating the 
    ''' place rectangle at.
    ''' </summary>
    ''' <param name="pt">the point to begin placing at.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal pt As Point)
        MyBase.New(pt)
        _PTEnd = pt
    End Sub

#End Region

#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Loads an image from a path
    ''' </summary>
    ''' <param name="src">Source to load the image from.</param>
    ''' -----------------------------------------------------------------------------
    Public Property ImageSource() As String
        Get
            Return _ImageSource
        End Get
        Set(ByVal Value As String)
            Try
                _Image = New Bitmap(Value)
                _ImageSource = Value

            Catch ex As Exception
                _Image = Nothing

                Throw New System.ArgumentException("Invalid Bitmap file")
            End Try
        End Set
    End Property



#End Region

#Region "Base Class Implementors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the PlaceTool 
    ''' </summary>
    ''' <param name="ptPoint">Last mouse point, snapped appropriately</param>
    ''' <param name="btn">Mouse button being held</param>
    ''' <param name="bShiftDown">Whether the shift key is down or not.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub UpdateTool(ByVal ptPoint As System.Drawing.Point, ByVal btn As MouseButtons, ByVal bShiftDown As Boolean, ByVal Zoom As Single)
        _PTEnd = MyBase.getEndPoint(ptPoint, bShiftDown)
        _Bounds = MyBase.getDrawingRect(ptPoint, bShiftDown)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the PlaceTool to the current surface
    ''' </summary>
    ''' <param name="g">Graphics context to draw against</param>
    ''' <param name="fScale">Zoom factor of the surface</param>
    ''' -----------------------------------------------------------------------------
    Friend Overloads Overrides Sub draw(ByVal g As System.Drawing.Graphics, ByVal fScale As Single)
        App.GraphicsManager.BeginScaledView(g)

        g.DrawImage(_Image, _Bounds)

        App.GraphicsManager.EndScaledView(g)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends the PlaceTool operation.
    ''' </summary>
    ''' <param name="bShiftDown">Whether shift is held or not</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub EndTool(ByVal bShiftDown As Boolean)
        If Point.op_Inequality(_PTEnd, _PTOrigin) Then

            If _Bounds.Width > 0 AndAlso _Bounds.Height > 0 Then
                placeImage(_Bounds)
            End If
        Else

            placeImage()
        End If
    End Sub

#End Region

#Region "Helper Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Places an image with its original scale (no bounding rect)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub placeImage()
        Dim newimage As New GDILinkedImage(96, 96, CInt(_PTOrigin.X), CInt(_PTOrigin.Y), _ImageSource)
        MDIMain.ActiveDocument.AddObjectToPage(newimage, "Placed Image")
     End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Places an image within a specific rectangular bounds
    ''' </summary>
    ''' <param name="rect"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub placeImage(ByVal rect As Rectangle)
        Dim newimage As New GDILinkedImage(96, 96, _Bounds, _ImageSource)
        MDIMain.ActiveDocument.AddObjectToPage(newimage, "Placed Image")
     End Sub

#End Region


#Region "Disposal and Cleanup"



    Friend Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not _Image Is Nothing Then
                _Image.Dispose()
                _Image = Nothing
            End If

        End If

        MyBase.Dispose(disposing)
    End Sub

#End Region

End Class
