

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : ColorChangedEventArgs
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides color changed event argument 
''' </summary>
''' <remarks>Credit to Ken Getz for this code!
''' </remarks>
''' -----------------------------------------------------------------------------
Friend Class ColorChangedEventArgs
    Inherits EventArgs

#Region "Local Fields"

    'The color's RGB value
    Private _RGB As ColorHandler.RGB
    'The color's HSV value
    Private _HSV As ColorHandler.HSV
#End Region



#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the ColorChangedEventArgs
    ''' </summary>
    ''' <param name="RGB">RGB value of the color</param>
    ''' <param name="HSV">HSV value of the color</param>
    ''' -----------------------------------------------------------------------------
    Sub New(ByVal RGB As ColorHandler.RGB, ByVal HSV As ColorHandler.HSV)
        _RGB = RGB
        _HSV = HSV
    End Sub

#End Region

#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the RGB value of the color
    ''' </summary>
    ''' <value></value>
    ''' -----------------------------------------------------------------------------
    ReadOnly Property RGB() As ColorHandler.RGB
        Get
            Return _RGB
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the HSV value of the color
    ''' </summary>
    ''' <value></value>
    ''' -----------------------------------------------------------------------------
    ReadOnly Property HSV() As ColorHandler.HSV
        Get
            Return _HSV
        End Get
    End Property

#End Region
End Class