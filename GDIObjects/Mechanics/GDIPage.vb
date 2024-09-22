''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIPage
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Represents an instance of a page inside a GDIDocument.  
''' While the graphics class type of  uses this, it is limited to a single page. 
''' The print document class contains multiple GDIPages, one per page on the surface.
''' </summary>
'''  -----------------------------------------------------------------------------
<Serializable()> _
Public Class GDIPage
    Implements IDisposable

#Region "Non serialized fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether this page has been disposed or not
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _Disposed As Boolean = False

#End Region


#Region "Local Fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The collection of GDIObject contained in this page (
    ''' Read GDIObjCol as GDIObject Collection)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _GDIObjects As GDIObjCol

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the page.  This is an arbitrary value.  If not set the page will 
    ''' render itself in lists as "Page X" where X is the page number.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Name As String = ""

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The page number of this page.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _PageNumber As Int32 = 0
#End Region

#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructor for a GDIPage
    ''' </summary>
    ''' <param name="iNum">The page number of the page.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal iNum As Int32)
        _PageNumber = iNum
        _GDIObjects = New GDIObjCol
        _Name = "Page " & CStr(Me.PageNum)
    End Sub

#End Region


#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the name of this page.  Solely used for a more friendly interface
    ''' in GDI+ Architect.
    ''' </summary>
    ''' <value>A string with the page name.</value>
    ''' -----------------------------------------------------------------------------
    Public Property Name() As String
        Get
            Return Me.ToString
        End Get
        Set(ByVal Value As String)
            _Name = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets this page's page number.
    ''' </summary>
    ''' <value>Integer value representing the page number</value>
    '''     ''' -----------------------------------------------------------------------------
    Public Property PageNum() As Int32
        Get
            Return _PageNumber
        End Get
        Set(ByVal Value As Int32)
            _PageNumber = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the collection of objects contained on this page
    ''' </summary>
    ''' <value>The collection of objects on the page, a GDIObjCol (GDI Object Collection) </value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property GDIObjects() As GDIObjects.GDIObjCol
        Get
            Return _GDIObjects
        End Get
    End Property

#End Region


#Region "Drawing related functions"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the page to a bitmap.  
    ''' All of the work for this is done inside the page's object collection (GDIObjCol) 
    ''' collection.
    ''' </summary>
    ''' <param name="gbmp">A graphics context to write the page to, created from a bitmap.</param>
    ''' <param name="doc">The parent GDIDocument this page belongs to.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub ToBitmap(ByVal g As Graphics)
        _GDIObjects.PrintObjects(g)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the page to a meta file.
    ''' All of the work for this is done inside the page's object collection (GDIObjCol) 
    ''' 
    ''' </summary>
    ''' <param name="gMeta">A graphics context to write the page to.  Created from a metafile.</param>
    ''' <param name="doc">The parent GDIDocument this page belongs to.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub DrawToMetafile(ByVal g As Graphics)
        _GDIObjects.PrintObjects(g)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the page to a print document.
    ''' All of the work for this is done inside the page's object collection (GDIObjCol) 
    ''' </summary>
    ''' <param name="g">A graphics context to write the page to. 
    ''' This is created from a print document.</param>
    ''' <param name="doc">The parent GDIDocument this page belongs to.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub Print(ByVal g As Graphics)
        _GDIObjects.PrintObjects(g)
    End Sub

#End Region

#Region "Misc Methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a string representation of the page.
    ''' </summary>
    ''' <returns>A string containing the name of the page</returns>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function ToString() As String

        Return _Name

    End Function

#End Region


#Region "Dispoal and Cleanup"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of a page.
    ''' </summary>
    ''' <param name="disposing">If unmanaged resources are being disposed or not</param>
    ''' -----------------------------------------------------------------------------
    Private Overloads Sub Dispose(ByVal disposing As Boolean)
        If Not _Disposed Then
            If disposing Then

                For Each gdiobj As GDIObject In Me.GDIObjects
                    gdiobj.Dispose()
                Next

            End If


        End If
        _Disposed = True
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Default dispose implementation
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Overloads Sub Dispose() Implements System.IDisposable.Dispose
        Me.Dispose(True)
    End Sub

#End Region
End Class
