Imports System.ComponentModel
Imports System.CodeDom

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIShape
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Virtual class that all GDIShapes inherit from, Provides common 
''' functionality for all GDIShapes.
''' </summary>
'''  -----------------------------------------------------------------------------
<Serializable()> _
Public MustInherit Class GDIShape
    Inherits GDIObject

#Region "Non serialized fields"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Path used to render the shape to a drawing surface.  
    ''' </summary>
    ''' <remarks>Each GDIShape, regardless if 
    ''' there is a simpler render method, is rendered to a path and then the path is 
    ''' rendered to the drawing surface.  This allows for a single point of interaction for 
    ''' drawing code and as you look at inheritors, it will become clearer what the benefit 
    ''' of doing it this way is.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
Protected _Path As Drawing2D.GraphicsPath

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Indicates whether our _Path member is invalid and needs to be recreated.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
        Protected _ResetPath As Boolean = True
#End Region

#Region "Serialized fields"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The stroke used to paint the shape.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _Stroke As GDIStroke

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to draw the stroke on the shape or not.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _DrawStroke As Boolean = True
#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a GDIShape
    ''' </summary>
    ''' <remarks>Adds a handler to the OnStrokeUpdated event 
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.new()
        Trace.WriteLineIf(Session.Tracing, "Shape.New")
        _Stroke = New GDIStroke(Session.Stroke)

        AddHandler _Stroke.StrokeUpdated, AddressOf OnStrokeUpdated
    End Sub
#End Region


#Region "Delegates and Events"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to a stroke changed event by setting _ResetPath to true, signifying 
    ''' that the path is invalid and the shape needs to be redrawn.
    ''' </summary>
    ''' <remarks>This method would be private, but serialization requires delegate 
    ''' handlers such as this to be marked as public or an exception will be thrown
    ''' when deserializing.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub OnStrokeUpdated(ByVal s As Object, ByVal e As EventArgs)
        _ResetPath = True
    End Sub

#End Region


#Region "Requires Implementation"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Implementers must override this to create the path that will be drawn to surfaces.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected MustOverride Sub createPath()

#End Region

#Region "Property Accessors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the stroke used to draw the stroke portion of the shape (the outline)
    ''' </summary>
    ''' <value>A GDIStroke.</value>
    ''' -----------------------------------------------------------------------------
    <BrowsableAttribute(True), Description("The stroke (pen) used to outline of this graphic")> _
    Public Overridable Property Stroke() As GDIStroke
        Get
            Return _Stroke
        End Get

        Set(ByVal Value As GDIStroke)
            Trace.WriteLineIf(Session.Tracing, "Shape.Stroke.Set")

            _Stroke = Value.clone()
            _ResetPath = True
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the bounds of the shape.  The bounds is a rectangle containing the 
    ''' shape in its non rotated position.
    ''' </summary>
    ''' <value>A rectangle to make the bounds.</value>
    ''' <remarks>The GDIShape overrides the Bounds property of the GDIObject so that it 
    ''' can reset the path used to render the shape when the bounds is explicitly set.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Overrides Property Bounds() As System.Drawing.Rectangle
        Get

            Return _Bounds

        End Get
        Set(ByVal Value As System.Drawing.Rectangle)
            Trace.WriteLineIf(Session.Tracing, "Shape.Bounds.Set:" & Value.ToString)

            _Bounds = Value
            _ResetPath = True
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets a value indicating whether to draw the stroke of this shape or not.
    ''' </summary>
    ''' <value>A Boolean indicating whether to draw the stroke</value>
    ''' -----------------------------------------------------------------------------
    <Browsable(True), DefaultValue(True), Description("Whether or not to draw a pen stroke for object.")> _
    Public Overridable Property DrawStroke() As Boolean
        Get
            Return _DrawStroke
        End Get
        Set(ByVal Value As Boolean)
            Trace.WriteLineIf(Session.Tracing, "Shape.DrawStroke.Set: " & Value)

            _DrawStroke = Value

        End Set
    End Property



#End Region

#Region "base class Implementers"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deserializes a shape.  Recreates the nonserialized _Path property based on the 
    ''' shape's settings.
    ''' </summary>
    ''' <param name="doc">The parent document being deserialized.</param>
    ''' <param name="pg">The page this object exists on.</param>
    ''' <returns>True.  By default GDIShapes do not expect deserialization issues.</returns>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Function deserialize(ByVal doc As GDIDocument, ByVal pg As GDIPage) As Boolean
        Trace.WriteLineIf(Session.Tracing, "Shape.Deserializet")

        MyBase.deserialize(doc, pg)
        createPath()

        Return True

    End Function


#End Region


#Region "Drawing related methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Resets the path used to render the shape to a graphics surface.  Inheritors can 
    ''' override this method to do any custom operations they need to when the path 
    ''' is being reset.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected Overridable Sub resetPath()
        If _Path Is Nothing Then
            _Path = New Drawing2D.GraphicsPath
        Else
            _Path.Reset()
        End If
    End Sub


#End Region




#Region "Disposal and cleanup Cleanup"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the Shape.
    ''' </summary>
    ''' <param name="disposing">If managed resources are being disposed or not.</param>
    ''' -----------------------------------------------------------------------------
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            ' Release  managed resources  
            _Stroke.Dispose()
        End If

        MyBase.Dispose(disposing)

    End Sub

#End Region

End Class

