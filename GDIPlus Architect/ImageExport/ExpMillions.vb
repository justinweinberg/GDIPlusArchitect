Imports GDIObjects
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : ExpMillions
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Serves as a base class for exporting to graphics with many colors
''' </summary>
''' -----------------------------------------------------------------------------
Friend MustInherit Class ExpMillions
    Inherits Exp

#Region "Type Declarations"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Enumeration that represents the quality of the export
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Enum EnumBitQuality As Integer
        '''<summary>24 bit export</summary>
        eQuality24
        '''<summary>32 bit export (with alpha channel)</summary>
        eQuality32Alpha
    End Enum

#End Region

#Region "Local Fields"
    '''<summary>Field that holds the quality argument of the export</summary>
    Protected _BitQuality As EnumBitQuality = EnumBitQuality.eQuality24

#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new export for millions of color type graphic files.
    ''' </summary>
    ''' <param name="name">Friendly name of the export type</param>
    ''' <param name="mimetype">Mime type of the  format</param>
    ''' <param name="filter">String to use to filter file selection</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal name As String, ByVal mimetype As String, ByVal filter As String)
        MyBase.New(name, mimetype, filter)
    End Sub


#End Region

#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The Quality of the export format as specified by the user.
    ''' </summary>
    ''' <value>An EnumBitQuality that indicates the quality of the export.</value>
    ''' -----------------------------------------------------------------------------
    Public Property BitQuality() As EnumBitQuality
        Get
            Return _BitQuality
        End Get
        Set(ByVal Value As EnumBitQuality)
            _BitQuality = Value
        End Set
    End Property
#End Region

End Class
