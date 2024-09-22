Imports System.ComponentModel

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIFilledShape
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Virtual class that fillable shapes (shapes with a fill property such as 
''' text, rectangles, etc.) inherit from.  Extracts common functionality for 
''' fillable shapes into this virtual class.
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public MustInherit Class GDIFilledShape
    Inherits GDIShape


#Region "Local Fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The fill used to fill shapes.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _Fill As GDIFill


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to track the fill of this object or not.  
    ''' </summary>
    ''' <remarks>For gradient and texture fills, the way they are rendered can depends 
    ''' on a point or rectangle.  If TrackFill is true, these fills base their position 
    ''' off of their parent object position.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Protected _TrackFill As Boolean = True

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Boolean indicating whether to draw the fill of this shape or not.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _DrawFill As Boolean = True


#End Region



#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new instance of a GDIFilledShape.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.new()

        Trace.WriteLineIf(Session.Tracing, "FilledShape.New")

        _Fill = Session.Fill.clone()


        AddHandler _Fill.FillUpdated, AddressOf OnFillUpdated

    End Sub


#End Region

#Region "Delegates and Events"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Notifies the base GDIShape class that the current path is invalid and needs to be 
    ''' recreated.
    ''' </summary>
    ''' <remarks>Serialization requires this otherwise should be private method to be public.
    ''' Setting this to private causes binary serialization to throw an exception.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub OnFillUpdated(ByVal s As Object, ByVal e As EventArgs)
        _ResetPath = True
    End Sub
#End Region

#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Boolean indicating whether to stroke the shape or not.
    ''' </summary>
    ''' <value>A Boolean indicating whether to draw the stroke of the shape or not.</value>
    ''' <remarks>This override of the default functionality is used to disallow 
    ''' setting both drawfill and drawstroke to false.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    <Browsable(True), DefaultValue(True), Description("Whether or not to draw a pen stroke for object.")> _
     Public Overrides Property DrawStroke() As Boolean
        Get
            Return _DrawStroke
        End Get
        Set(ByVal Value As Boolean)
            Trace.WriteLineIf(Session.Tracing, "Shape.DrawStroke.Set: " & Value)

            If _DrawFill = True OrElse Value = True Then
                _DrawStroke = Value
            End If

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets a value indicating whether to draw the fill portion of the shape.
    ''' </summary>
    ''' <value>A Boolean indicating whether to fill the shape.</value>
    ''' -----------------------------------------------------------------------------
    <Browsable(True), DefaultValue(True), Description("Whether or not to fill the object.")> _
    Public Overridable Property DrawFill() As Boolean
        Get
            Return _DrawFill
        End Get
        Set(ByVal Value As Boolean)
            Trace.WriteLineIf(Session.Tracing, "Shape.DrawFill.Set: " & Value)

            If _DrawStroke = True OrElse Value = True Then
                _DrawFill = Value
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the bounds on the shape.
    ''' </summary>
    ''' <value>A rectangle that bounds the shape.</value>
    ''' <remarks>Notifies child fills that the parent shape's bounds have changed on a set.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Overrides Property Bounds() As System.Drawing.Rectangle
        Get
            Return _Bounds
        End Get
        Set(ByVal Value As System.Drawing.Rectangle)
            Trace.WriteLineIf(Session.Tracing, "FilledShape.Bounds.Set")
            MyBase.Bounds = Value
            Fill.OnParentUpdated(Me)
        End Set
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the GDIFilled shape's current fill
    ''' </summary>
    ''' <value>A GDIFill used to fill the shape.</value>
    ''' <remarks>Raises the onParentUpdated event upon Set.  Also notice that this property
    ''' clones the brush rather than doing a straight reference assignment upon Set.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    <BrowsableAttribute(True), Description("The fill (brush) used to paint the interior of this graphic")> _
    Public Property Fill() As GDIFill
        Get
            Return _Fill
        End Get
        Set(ByVal Value As GDIFill)
            Trace.WriteLineIf(Session.Tracing, "FilledShape.Fill.Set" & Value.ToString)
            If Not _Fill Is Nothing Then
                _Fill.Dispose()
            End If
            _Fill = Value.clone()
            _Fill.OnParentUpdated(Me)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to "track" the fill 
    ''' </summary>
    ''' <value>
    ''' A Boolean indicating whether tracking is enabled for this shape's fill or not.
    ''' </value>
    ''' <remarks>For some fills how it is rendered is determined by 
    ''' coordinates.  For example, a gradient Fill has a start and end point for the gradient 
    ''' If track fill is set to true, these coordinates are derived from this shape 
    ''' (the parent shape of the fill).  If it is set to false, the fill is left at the 
    ''' last point it was tracked.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------   
    <DefaultValue(True), Description("Whether this graphic's fill should align to the graphic's bounds  Applies only to gradient and texture fills.  Set this to false for various effects")> _
    Public Property TrackFill() As Boolean
        Get
            Return _TrackFill
        End Get
        Set(ByVal Value As Boolean)
            Trace.WriteLineIf(Session.Tracing, "FilledShape.TrackFill.Set" & Value)

            _TrackFill = Value
            Fill.OnParentUpdated(Me)
        End Set
    End Property


#End Region

#Region "Disposal and cleanup"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the fill, which in turn disposes of allocated brushes.
    ''' </summary>
    ''' <param name="disposing">Whether unmanaged resources are being disposed or not.</param>
    ''' -----------------------------------------------------------------------------
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            _Fill.Dispose()
        End If

        MyBase.Dispose(disposing)
    End Sub

#End Region

#Region "Serialization"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deserialized a Filled shape.
    ''' </summary>
    ''' <param name="doc">The parent document the filled shape belongs to.</param>
    ''' <param name="pg">The page the filled shape is on, if any.</param>
    ''' <returns>A Boolean indicating if deserialization was successful.</returns>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Function deserialize(ByVal doc As GDIDocument, ByVal pg As GDIPage) As Boolean
        Trace.WriteLineIf(Session.Tracing, "FilledShape.Deserialize ")

        Return Me.Fill.deserialize() And MyBase.deserialize(doc, pg)
    End Function


#End Region
End Class
