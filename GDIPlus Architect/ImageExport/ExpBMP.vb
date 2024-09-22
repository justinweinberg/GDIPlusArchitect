Imports System.Drawing.Imaging
Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : ExpBMP
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Responsible for exporting graphical Bitmaps of pages
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class ExpBMP
    Inherits ExpMillions

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new ExpBMP.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New("Bitmap - 24 bit", "image/bitmap", "bmp |*.bmp| All files |*.*")
    End Sub

#End Region

#Region "Base Class Implementors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exports the current page to a bitmap
    ''' </summary>
    ''' <param name="doc">Document being exported</param>
    ''' <param name="pg">Page to export</param>
    ''' <param name="sFullpath">Path to save the page to.</param>
    ''' -----------------------------------------------------------------------------
    Public Overrides Sub Export(ByVal doc As GDIDocument, ByVal pg As GDIPage, ByVal sFullpath As String)
        Dim bmp As Bitmap


        'Based on the quality settings, ask the base EXP class to get a bitmap for the page.
        Select Case BitQuality
            Case ExpMillions.EnumBitQuality.eQuality24
                bmp = MyBase.getBitmap(doc, pg, PixelFormat.Format24bppRgb)
            Case ExpMillions.EnumBitQuality.eQuality32Alpha
                bmp = MyBase.getBitmap(doc, pg, PixelFormat.Format32bppArgb)
        End Select

        bmp.Save(sFullpath, System.Drawing.Imaging.ImageFormat.Bmp)


        bmp.Dispose()

    End Sub
#End Region

End Class
