
''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : Session
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Brokers application wide global settings relevant to the GDIObject project and 
''' the GDIPlus Architect project.
''' </summary>
''' -----------------------------------------------------------------------------
Public Class Session

#Region "Local Fields"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to write debug tracing or not 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _Tracing As Boolean = False


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sundry local settings
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _Settings As Settings


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Current in use application wide stroke
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _Stroke As GDIStroke

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Currently in use application wide Fill
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _Fill As GDIFill


    Private Shared _GraphicsManager As GraphicsManager


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Manager for the current document's properties and settings
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _DocumentManager As DocumentManager

#End Region


#Region "Constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new instance of the global manager
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Shared Sub New()
        _Settings = New Settings
        _Stroke = New GDIStroke
        _Fill = New GDISolidFill(Color.Black)
        _GraphicsManager = New GraphicsManager

        AddHandler _Stroke.StrokeUpdated, AddressOf Stroke_StrokeUpdated
        AddHandler _Fill.FillUpdated, AddressOf Fill_FillUpdated

        _DocumentManager = New DocumentManager

    End Sub

#End Region

#Region "Stroke and Fill Delegated and Events"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Delegate to notify interested parties that the session's stroke has been changed 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Delegate Sub StrokeChanged(ByVal s As Object, ByVal e As EventArgs)
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Event to notify listeners that the session's fill has been changed.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Delegate Sub FillChanged(ByVal s As Object, ByVal e As EventArgs)

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Local instance of the stroke changed delegate 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _StrokeChangedInvokeList As StrokeChanged
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Local instance of the fill changed delegate
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _FillChangedInvokeList As FillChanged


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a listener callback to the stroke changed event
    ''' </summary>
    ''' <param name="cb">The callback to invoke upon the stoke changing.</param>
    ''' -----------------------------------------------------------------------------
    Public Shared Sub setStrokeChangedCallBack(ByVal cb As StrokeChanged)
        If _StrokeChangedInvokeList Is Nothing Then
            _StrokeChangedInvokeList = cb
        Else
            _StrokeChangedInvokeList = CType(System.Delegate.Combine(_StrokeChangedInvokeList, cb), StrokeChanged)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a listener to the fill changed callback
    ''' </summary>
    ''' <param name="cb">The callback to invoke upon fill changing</param>
    ''' -----------------------------------------------------------------------------
    Public Shared Sub setFillChangedCallBack(ByVal cb As FillChanged)
        If _FillChangedInvokeList Is Nothing Then
            _FillChangedInvokeList = cb
        Else
            _FillChangedInvokeList = CType(System.Delegate.Combine(_FillChangedInvokeList, cb), FillChanged)
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invokes the strokechanged delegate.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared Sub InvokeStrokeCallbacks()
        _StrokeChangedInvokeList.Invoke(_Stroke, EventArgs.Empty)
    End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invokes the fillchanged delegate.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared Sub InvokeFillCallbacks()
        _FillChangedInvokeList.Invoke(_Fill, EventArgs.Empty)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Removes a listener from the stroke changed event
    ''' </summary>
    ''' <param name="cb">The callback to remove</param>
    ''' -----------------------------------------------------------------------------
    Public Shared Sub removeStrokeCallback(ByVal cb As StrokeChanged)
        If Not _StrokeChangedInvokeList Is Nothing Then
            _StrokeChangedInvokeList = CType(System.Delegate.Remove(_StrokeChangedInvokeList, cb), StrokeChanged)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Removes a listener from the fill changed delegate 
    ''' </summary>
    ''' <param name="cb">The callback to remove.</param>
    ''' -----------------------------------------------------------------------------
    Public Shared Sub removeFillCallback(ByVal cb As FillChanged)
        If Not _FillChangedInvokeList Is Nothing Then
            _FillChangedInvokeList = CType(System.Delegate.Remove(_FillChangedInvokeList, cb), FillChanged)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handler for the FillUpdated event.  Informs listeners of the fill update
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared Sub Fill_FillUpdated(ByVal s As Object, ByVal e As EventArgs)
        InvokeFillCallbacks()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handle for the Stroke updated event.  Informs listeners of the stroke update
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared Sub Stroke_StrokeUpdated(ByVal s As Object, ByVal e As EventArgs)
        InvokeStrokeCallbacks()
    End Sub

#End Region


#Region "Property Accessors"


 


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether debug trace is enabled or not.
    ''' </summary>
    ''' <value>A Boolean indicating if tracing is enabled or not.</value>
    ''' <remarks>This value is controlled from the user interface project.</remarks>
    ''' -----------------------------------------------------------------------------
    Public Shared Property Tracing() As Boolean
        Get
            Return _Tracing
        End Get
        Set(ByVal Value As Boolean)
            _Tracing = Value
        End Set

    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets various settings relevant to the entire GDIObject project.
    ''' </summary>
    ''' <value>A settings object which contains various settings needed in the project.</value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property Settings() As Settings
        Get
            Return _Settings
        End Get
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a reference to the current document manager
    ''' </summary>
    ''' <value>The DocumentManager used to provide context relevant to 
    ''' the currently selected document</value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property DocumentManager() As DocumentManager
        Get
            Return _DocumentManager
        End Get

    End Property


  
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a reference to the GraphicsManager.
    ''' </summary>
    ''' <value>A GraphicsManager used to manipulate graphics within the GDIObjects project</value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property GraphicsManager() As GraphicsManager
        Get
            Return _GraphicsManager
        End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the current GDIFill that new shapes should be filled with.  
    ''' See the GDIFill virtual class for more information about fills.
    ''' </summary>
    ''' <value>A GDIFill.</value>
    ''' <remarks>On set, this clones the incoming fill and assigns the clone to the _Fill
    ''' property.  It then reassigns the Fill_FillUpdated handler 
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Shared Property Fill() As GDIFill

        Get
            Return _Fill
        End Get
        Set(ByVal Value As GDIFill)
            Trace.WriteLineIf(Session.Tracing, "Session.fill.Set")

            If Not _Fill Is Nothing Then
                _Fill.Dispose()
                _Fill = Nothing
            End If
            _Fill = Value.Clone()

            AddHandler _Fill.FillUpdated, AddressOf Fill_FillUpdated


            InvokeFillCallbacks()
        End Set

    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the current stroke that new shapes will be stroked with.
    ''' </summary>
    ''' <value>The GDIStroke that new shapes should be stroke with</value>
    ''' <remarks>On set, this clones the incoming stroke and assigns the clone to the _Stroke
    ''' property.  It then reassigns the Stroke_StrokeUpdated handler 
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Shared Property Stroke() As GDIStroke
        Get
            Return _Stroke
        End Get
        Set(ByVal Value As GDIStroke)
            Trace.WriteLineIf(Session.Tracing, "Session.Stroke.Set")
            _Stroke.Dispose()
            _Stroke = Value.Clone()


            AddHandler _Stroke.StrokeUpdated, AddressOf Stroke_StrokeUpdated

            InvokeStrokeCallbacks()
        End Set

    End Property


#End Region


End Class
