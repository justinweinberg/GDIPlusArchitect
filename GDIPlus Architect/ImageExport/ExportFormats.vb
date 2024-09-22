Imports System.Drawing.Imaging
Imports GDIObjects
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : ExportFormats
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a constainer for the various types of graphic exports that GDI+ 
''' Architect supports.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class ExportFormats
    Inherits CollectionBase

#Region "Constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the ExportFormats class
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        populateformats()
    End Sub

#End Region


#Region "Helper methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populates the list with the various export formats supported by GDI+ Architet
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub populateformats()

        'Creates the formats
        Dim expJPEG As New ExpJPEG

        Dim expTIFF32 As New ExpTIFF("TIFF - 32 bit (alpha channel)")
        expTIFF32.BitQuality = ExpMillions.EnumBitQuality.eQuality32Alpha

        Dim expTIFF24 As New ExpTIFF("TIFF - 24 bit")
        expTIFF24.BitQuality = ExpMillions.EnumBitQuality.eQuality24

        Dim expmeta As New ExpMeta

        Dim expBMP24 As New ExpBMP
        expBMP24.BitQuality = ExpMillions.EnumBitQuality.eQuality24

        Dim expPNG24 As New ExpPNG("PNG - 24 bit")
        expPNG24.BitQuality = ExpMillions.EnumBitQuality.eQuality24

        Dim expPNG32 As New ExpPNG("PNG - 32 bit (alpha channel)")
        expPNG32.BitQuality = ExpMillions.EnumBitQuality.eQuality32Alpha

        'Add them to the collection
        With innerlist
            .Add(expJPEG)
            .Add(expPNG32)
            .Add(expPNG24)
            .Add(expBMP24)
            .Add(expTIFF32)
            .Add(expTIFF24)
            .Add(expmeta)
        End With


    End Sub

#End Region

#Region "Collection Implementation"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Property Accessor for the ExportFormats collection
    ''' </summary>
    ''' <param name="iIndex">Index to retrieve the export format at.</param>
    ''' <value>Returns an Exp inheritor</value>
    ''' -----------------------------------------------------------------------------
    Default Public ReadOnly Property Item(ByVal iIndex As Int32) As Exp
        Get
            Return DirectCast(innerlist(iIndex), Exp)
        End Get
    End Property
#End Region
End Class
