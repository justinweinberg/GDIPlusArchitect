Imports System.Drawing.Imaging
Imports System.Threading
Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : Exp
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Serves as a base class for exporting images from GDI+ Architect documents.
''' </summary>
''' -----------------------------------------------------------------------------
Friend MustInherit Class Exp

#Region "Local Fields"



    '''<summary>Friendly name of the export type (GIF, Metafile, etc)</summary>
    Protected _Name As String = String.Empty
    '''<summary>MIMEType of the export as a string </summary>
    Protected _MimeType As String = String.Empty
    '''<summary>A file filter that matches the save as options for the type.</summary>
     Protected _FileFilter As String = String.Empty

#End Region


#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructor for a new Exp object
    ''' </summary>
    ''' <param name="name">Name of the type of exported document (Metafile, PNG, etc)</param>
    ''' <param name="MimeType">MimeType appropriate to the export type.</param>
    ''' <param name="FileFilter">Filter used in the save as dialog to filter to appropriate 
    ''' types.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal name As String, ByVal MimeType As String, ByVal FileFilter As String)
        _Name = name
        _MimeType = MimeType
        _FileFilter = FileFilter
    End Sub

#End Region


#Region "Required Implementation"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Requires implementation by inheritors.  Given a document and a page, exports the 
    ''' page to the image type of the inheritor to the path specified in the sFullpath 
    ''' argument.
    ''' </summary>
    ''' <param name="doc">The document whose page is being exported</param>
    ''' <param name="pg">The page to export</param>
    ''' <param name="sFullpath">The target path to export to.</param>
    ''' -----------------------------------------------------------------------------
    Public MustOverride Sub Export(ByVal doc As GDIDocument, ByVal pg As GDIPage, ByVal sFullpath As String)

#End Region


#Region "Helper Functions"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the  encoder that matches the Exp class MimeType.
    ''' </summary>
    ''' <returns>An ImageCodedInfo object that can be used to encode the image.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function GetEncoderInfo() As ImageCodecInfo

        Dim encoders() As ImageCodecInfo

        encoders = ImageCodecInfo.GetImageEncoders()

        'Iterate over encoders, finding the appropriate mime type
        For i As Int32 = 0 To encoders.Length - 1
            If encoders(i).MimeType = _MimeType Then
                Return encoders(i)
            End If
        Next

        Return Nothing
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a 32 bit bitmap with the page to export rendered to its surface.
    ''' </summary>
    ''' <param name="doc">The document being exported </param>
    ''' <param name="pg">The page to export </param>
    ''' <returns>A bitmap with the page rendered to its surface</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function getBitmap(ByVal doc As GDIDocument, ByVal pg As GDIPage) As Bitmap
        Dim bmp As New Bitmap(doc.RectPageSize.Width, _
                              doc.RectPageSize.Height, _
                            Imaging.PixelFormat.Format32bppPArgb)

        Dim g As Graphics = Graphics.FromImage(bmp)

        'Set the back color, text hint, and smoothing mode.
        g.Clear(doc.BackColor)
        g.TextRenderingHint = doc.TextRenderingHint
        g.SmoothingMode = doc.SmoothingMode


        'Ask the page to draw itself as a bitmap
        pg.ToBitmap(g)
        g.Dispose()

        Return bmp

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a bitmap with the page to export rendered to its surface in the format 
    ''' specified by the format argument (PixelFormat).
    ''' </summary>
    ''' <param name="doc">The document being exported </param>
    ''' <param name="pg">The page to export </param>
    ''' <param name="format">format the bitmap is to be exported to.</param>
    ''' <returns>A bitmap with the page rendered to its surface in the format specified in the format argument.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function getBitmap(ByVal doc As GDIDocument, ByVal pg As GDIPage, ByVal format As Imaging.PixelFormat) As Bitmap

        'Create a new bitmap based off of the document's page size
        Dim bmp As New Bitmap(doc.RectPageSize.Width, doc.RectPageSize.Height, format)
        Dim g As Graphics = Graphics.FromImage(bmp)

        'Set the back color, text hint, and smoothing mode.
        g.Clear(doc.BackColor)
        g.TextRenderingHint = doc.TextRenderingHint
        g.SmoothingMode = doc.SmoothingMode

        'Ask the page to render itself as a bitmap
        pg.ToBitmap(g)

        g.Dispose()

        Return bmp

    End Function
#End Region

#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the file filter for the EXP type as a string.
    ''' </summary>
    ''' <value>The file filter as a string.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property FileFilter() As String
        Get
            Return _FileFilter
        End Get
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the friendly name of the export type as a string.
    ''' </summary>
    ''' <value>The name as a string.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Name() As String
        Get
            Return _Name
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the mime type of the export type as a string.
    ''' </summary>
    ''' <value>The file filter as a string.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property MimeType() As String
        Get
            Return _MimeType
        End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the name of the export as a string.
    ''' </summary>
    ''' <returns>The name of the export type</returns>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function ToString() As String
        Return _Name
    End Function
#End Region
End Class
