Imports System.Drawing.Imaging
Imports GDIObjects
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : ExpPNG
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Class that handles exporting PNGs
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class ExpPNG
    Inherits ExpMillions

#Region "Constructors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the ExpPNG class (for exporting PNGS)
    ''' </summary>
    ''' <param name="name">Friendly name to give the instance.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal name As String)
        MyBase.New(name, "image/png", "png |*.png| All files |*.*")
    End Sub


#End Region

#Region "Base Class Implementors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exports the current page to a PNG
    ''' </summary>
    ''' <param name="doc">The document to export </param>
    ''' <param name="pg">The page to export </param>
    ''' <param name="sFullpath">The path to save the PNG to.</param>
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

        'Save to the PNG format
        bmp.Save(sFullpath, System.Drawing.Imaging.ImageFormat.Png)

        bmp.Dispose()

    End Sub

#End Region
End Class
