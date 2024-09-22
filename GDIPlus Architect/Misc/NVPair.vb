
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : NVPair
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a simple wrapper for pairs of name/int32 value information (eg ID=1, name=foo)
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class NVPair

#Region "Local Fields"


    '''<summary>The ID of the item </summary>
    Private _ID As Int32 = -1
    '''<summary>The name of the item</summary>
    Private _Name As String = String.Empty

#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new NVPair instance
    ''' </summary>
    ''' <param name="Id">The ID (value) of the NVPair</param>
    ''' <param name="name">The name of the NVPair</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal Id As Int32, ByVal name As String)
        _ID = Id
        _Name = name
    End Sub

#End Region


#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the ID of the NVPair item.
    ''' </summary>
    ''' <value>The ID of the NVPair</value>
    ''' -----------------------------------------------------------------------------
    Public Property ID() As Int32
        Get
            Return _ID
        End Get
        Set(ByVal Value As Int32)
            _ID = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the name of the NVPair item
    ''' </summary>
    ''' <value>The name of the nvpair</value>
    ''' -----------------------------------------------------------------------------
    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal Value As String)
            _Name = Value
        End Set
    End Property

#End Region

#Region "Overrides"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a string representation of the NVPair (the name of the item)
    ''' </summary>
    ''' <returns></returns>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function toString() As String
        Return _Name
    End Function
#End Region
End Class
