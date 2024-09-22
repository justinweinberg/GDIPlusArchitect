Imports System.Drawing.Imaging
Imports GDIObjects
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : ExpJPEG
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Responsible for exporting the current page as a JPEG.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class ExpJPEG
    Inherits ExpMillions

#Region "Local Fields"
    'The quality of the JPEG
    Private _Quality As Long = 80
#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructor for a new instance of the EXPJPG class
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New("JPEG", "image/jpeg", "jpg |*.jpg| All files |*.*")
    End Sub

#End Region

#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the quality of the exported JPEG
    ''' </summary>
    ''' <value></value>
    ''' -----------------------------------------------------------------------------
    Public Property Quality() As Long
        Get
            Return _Quality
        End Get
        Set(ByVal Value As Long)
            _Quality = Value
        End Set
    End Property

#End Region

#Region "Base Class Implementors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a JPEG file
    ''' </summary>
    ''' <param name="sFileName">The path to save the metafile to.</param>
    ''' <param name="doc">The parent document being exported.</param>
    ''' <param name="pg">The page being exported.</param>
    ''' -----------------------------------------------------------------------------
    Public Overrides Sub Export(ByVal doc As GDIDocument, ByVal pg As GDIPage, ByVal sFullpath As String)
        Dim myImageCodecInfo As ImageCodecInfo
        Dim myEncoder As Encoder
        Dim myEncoderParameter As EncoderParameter
        Dim myEncoderParameters As EncoderParameters

        Dim bmp As Bitmap = MyBase.getBitmap(doc, pg)


        '// Get an ImageCodecInfo object that represents the JPEG codec.
        myImageCodecInfo = GetEncoderInfo()

        Dim codecinfo As ImageCodecInfo() = myImageCodecInfo.GetImageEncoders
        '// Create an Encoder object based on the GUID
        '// for the Quality parameter category.
        myEncoder = Encoder.Quality

        '// Create an EncoderParameters object.
        '// An EncoderParameters object has an array of EncoderParameter
        '// objects. In this case, there is only one
        '// EncoderParameter object in the array.
        myEncoderParameters = New EncoderParameters(1)


        '// Save the bitmap as a JPEG file with quality level 50.
        myEncoderParameter = New EncoderParameter(myEncoder, Quality)
        myEncoderParameters.Param(0) = myEncoderParameter
        bmp.Save(sFullpath, myImageCodecInfo, myEncoderParameters)

        bmp.Dispose()
        myEncoderParameter.Dispose()
        myEncoderParameters.Dispose()
    End Sub


#End Region
End Class
