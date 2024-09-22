

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIImage
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Virtual class that fillable shapes (shapes with a fill property such as 
''' text, rectangles, etc.) inherit from.  Extracts common functionality for 
''' fillable shapes into this virtual class.
''' </summary>
'''  -----------------------------------------------------------------------------
<Serializable()> _
   Public MustInherit Class GDIImage
    Inherits GDIObject


#Region "Type Declarations"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A filter string of valid image types to use on the GDI+ Architect surface
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Const CONST_IMG_FILTER As String = "Bitmap Files *.bmp;*.gif;*.jpeg;*.jpg;*.png;|*.bmp;*.gif;*.jpeg;*.jpg;*.png;| All files (*.*)|*.*"

#End Region


#Region "Nonserialized Fields"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A temporary, nonserialized bitmap that is used to draw the bitmap to the surface.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
Protected _Image As Bitmap

#End Region

#Region "Local Fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' X DPI of the image
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _XDPI As Int32
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Y DPI of the image.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _YDPI As Int32



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Boolean indicating whether the image has been manually sized or if the image 
    ''' should use its default sizes.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _Sized As Boolean = False
#End Region
 
#Region "Type Selector"



        ''' -----------------------------------------------------------------------------
        ''' Project	 : GDIObjects
        ''' Class	 : GDIImage.ImageFileChooser
        ''' 
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Provides a simple dialogue for selecting a file.
        ''' </summary>
        ''' -----------------------------------------------------------------------------
    Public Class ImageFileChooser
        Inherits System.Windows.Forms.Design.FileNameEditor

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Sets up the Imagefile chooser dialogue with the filter and default extension
        ''' </summary>
        ''' <param name="openFileDialog">The dialogue to initialize.</param>
        ''' -----------------------------------------------------------------------------
        Protected Overrides Sub InitializeDialog(ByVal openFileDialog As System.Windows.Forms.OpenFileDialog)
        openFileDialog.Filter = CONST_IMG_FILTER
        openFileDialog.DefaultExt = "bmp"
    End Sub
    End Class
#End Region

#Region "Misc Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if a file at a given full path is a bitmap or not.
    ''' </summary>
    ''' <param name="sPath">The path to the file.</param>
    ''' <returns>A Boolean indicating if the file is a bitmap or not.</returns>
    ''' -----------------------------------------------------------------------------
    Public Shared Function IsBitmap(ByVal sPath As String) As Boolean
        Dim x As Bitmap

        Try
            x = New Bitmap(sPath)
            Return True
        Catch ex As Exception
            Return False
        Finally
            If Not x Is Nothing Then
                x.Dispose()
            End If
        End Try

    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the color at a specific point in an image.
    ''' </summary>
    ''' <param name="ptObject">A point translated to the object's coordinate space</param>
    ''' <returns>The color at the point</returns>
    ''' -----------------------------------------------------------------------------
    Public Function ColorAtPoint(ByVal ptObject As Point) As Color
        Trace.WriteLineIf(Session.Tracing, "Image.ColorAtPoint: " & ptObject.ToString)


 
        Try
 

            Dim ptTranslated As PointF
            '_Bounds = New Rectangle(_Bounds.X, _Bounds.Y, CInt(fWidth), CInt(fHeight))

            'Translate the point to the image's coordinate space
            If Me.Bounds.Contains(ptObject) Then

                ptTranslated.X = ptObject.X - _Bounds.X
                ptTranslated.Y = ptObject.Y - _Bounds.Y

                ptTranslated.X = (ptTranslated.X * _Image.Width) / Bounds.Width
                ptTranslated.Y = (ptTranslated.Y * _Image.Height) / Bounds.Height


                Return _Image.GetPixel(CInt(ptTranslated.X), CInt(ptTranslated.Y))
            Else
                Return Color.White
            End If
        Catch ex As Exception
            Return Color.White
        Finally
 
        End Try

     End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Resets an image to its original size.   
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub RevertSize()
        _Sized = False
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sizes an image automatically based upon its DPI, resolution, and size.
    ''' </summary>
    ''' <param name="img">The image to size.</param>
    ''' -----------------------------------------------------------------------------
    Protected Sub autoSize(ByVal img As Image)
        Trace.WriteLineIf(Session.Tracing, "Image.autosize")
        'Pixels per inch = width / reso of image
        'So if the image is 500 pix and a reso of 100 then this is 5 px / inch.
        'Multiply by  px / inch to get the new pixel size

        Dim fWidth As Single = (img.Width / img.HorizontalResolution) * Me._XDPI
        Dim fHeight As Single = (img.Height / img.VerticalResolution) * Me._YDPI

        _Bounds = New Rectangle(_Bounds.X, _Bounds.Y, CInt(fWidth), CInt(fHeight))

    End Sub

#End Region




#Region "Drawing related functionality"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws the image to the surface.
    ''' </summary>
    ''' <param name="g">A graphics context to draw against.</param>
    ''' <param name="eDrawMode">The current draw mode (to the surface, to a graphic, to a printer)</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub Draw(ByVal g As System.Drawing.Graphics, ByVal eDrawMode As EnumDrawMode)

 
        If Not _Sized Then
            autoSize(_Image)
            _Sized = True
        End If

        MyBase.BeginDraw(g)

        g.DrawImage(_Image, Bounds)


        MyBase.EndDraw(g)

    End Sub


#End Region

#Region "Cleanup and Disposal"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of a GDIImage.  Specifically releases the _Image member.
    ''' </summary>
    ''' <param name="disposing">Whether unmanaged resources are being exposed or not.</param>
    ''' -----------------------------------------------------------------------------
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
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
