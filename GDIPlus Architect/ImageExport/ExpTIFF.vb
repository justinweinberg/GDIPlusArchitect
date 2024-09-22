Imports System.Drawing.Imaging
Imports GDIObjects
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : ExpTIFF
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Exports the current page to a TIFF
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class ExpTIFF
    Inherits ExpMillions

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the ExpTiff object
    ''' </summary>
    ''' <param name="name">Friendly name to give the instance.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal name As String)
        MyBase.New(name, "image/tiff", "tif |*.tif| All files |*.*")
    End Sub
#End Region

#Region "Base Class Implementors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exports the current page to a TIFF.
    ''' </summary>
    ''' <param name="doc">The current document </param>
    ''' <param name="pg">The page to export</param>
    ''' <param name="sFullpath">A full path to the place to save the TIFF.</param>
    ''' -----------------------------------------------------------------------------
    Public Overrides Sub Export(ByVal doc As GDIDocument, ByVal pg As GDIPage, ByVal sFullpath As String)
        Dim bmp As Bitmap

        'Depending on quality, ask the base class to create a BMP based on either 24 or 32alpha
        Select Case BitQuality
            Case ExpMillions.EnumBitQuality.eQuality24
                bmp = MyBase.getBitmap(doc, pg, PixelFormat.Format24bppRgb)
            Case ExpMillions.EnumBitQuality.eQuality32Alpha
                bmp = MyBase.getBitmap(doc, pg, PixelFormat.Format32bppArgb)
        End Select

        'Save the BMP as as TIFF
        bmp.Save(sFullpath, System.Drawing.Imaging.ImageFormat.Tiff)

        bmp.Dispose()

    End Sub

#End Region
End Class
