Imports System.ComponentModel
Imports System.CodeDom

#Region "Fill Property Browser"


''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : FillBrowser
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides an interface for selecting fills in a property grid.
''' </summary>
''' <remarks>For more information about how this works, see the framework's 
''' ExpandableObjectConverter class.
''' </remarks>
''' -----------------------------------------------------------------------------
Public Class FillBrowser
    Inherits ExpandableObjectConverter

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' This overload causes the property grid to always display "Fill Properties" regardless 
    ''' of what type of fill is being displayed.  Removing it will show the name of the type 
    ''' of fill.
    ''' </summary>
    ''' <param name="context">Type Descriptor context</param>
    ''' <param name="culture">Culture info</param>
    ''' <param name="value">The object being browsed</param>
    ''' <param name="destinationType">The output type expected. In the case of the property 
    ''' grid, this is a string.</param>
    ''' <returns>A string containing the words "Fill Properties" for the property grid.</returns>
    ''' <remarks>This could be expanded on by interrogating the incoming value property and 
    ''' building something more interesting in the return string.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Overloads Overrides Function ConvertTo(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object, ByVal destinationType As System.Type) As Object
        Return "Fill Properties"
    End Function
End Class

#End Region

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIFill
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Virtual class all GDI+ Architect fills inherit from.
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable(), TypeConverter(GetType(FillBrowser))> _
Public MustInherit Class GDIFill
    Implements IDisposable

#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructor for a new instance of a fill.  Sets the parent and consolidation settings.
    ''' </summary>
    ''' <remarks>Notice that this sets consolidate based on the session settings.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        _Consolidate = Session.Settings.ConsolidateFills
    End Sub

#End Region

#Region "Event Declarations"
    


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Raised when a property of a fill is changed.
    ''' </summary>
    ''' <remarks>The session listens for this property and multicasts to interested listeners.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Event FillUpdated(ByVal s As Object, ByVal e As EventArgs)

#End Region


#Region "Non Serialized Fields"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether this Fill has been disposed or not.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _Disposed As Boolean = False
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A System.Drawing.Brush used to paint the GDIFilledShape the GDIFill belongs to.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Protected _Brush As Brush

#End Region

#Region "Field Declarations"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to consolidate this fill with similar fills 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Consolidate As Boolean




#End Region


#Region "Must Implement Methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Required for inheritors to implement.  Updates the fill's current brush.
    ''' </summary>
    ''' <remarks>This method should probably have been named "UpdateBrush"
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Protected MustOverride Sub UpdateFill()

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Required for inheritors.  Deserializes a fill.
    ''' </summary>
    ''' <param name="parent">Parent shape being deserialized.</param>
    ''' <returns>True for the default case.  If inheritors anticipate potential problems 
    ''' deserializing a fill, they should interrogate their properties and return false 
    ''' if the fill cannot be deserialized.</returns>
    ''' -----------------------------------------------------------------------------
    Friend Overridable Function deserialize() As Boolean
        Return True
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Required for inheritors.  What to do when the fill's parent shape has been updated.
    ''' </summary>
    ''' <param name="obj">The parent shape which has changed.</param>
    ''' -----------------------------------------------------------------------------
    Public MustOverride Sub OnParentUpdated(ByVal obj As GDIFilledShape)


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts the fill to XML (for SVG export)
    ''' </summary>
    ''' <param name="xmlDoc">XMLDocument to append the XML to.</param>
    ''' <param name="defs">Definitions section of the SVG document.</param>
    ''' <param name="parent">Parent to which this fill belongs.</param>
    ''' -----------------------------------------------------------------------------
    Public MustOverride Sub toXML(ByVal xmlDoc As Xml.XmlDocument, ByVal defs As Xml.XmlNode, ByVal parent As Xml.XmlNode)

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits code for a consolidated object.
    ''' </summary>
    ''' <param name="sSharedName">The shared consolidated name to use</param>
    ''' <param name="declarations">CodeDOM Declaration section</param>
    ''' <param name="InitGraphics">the GDI+ Architect initGraphics method.</param>
    ''' <param name="RenderGDI">The GDI+ Architect renderGDI method.</param>
    ''' <param name="DisposeGDI">The GDI+ Architect disposeGDI method</param>
    ''' <param name="ExportSettings">The currently selected export settings</param>
    ''' <returns>A string containing the name of the fill created in code.</returns>
    ''' -----------------------------------------------------------------------------
    Friend MustOverride Function emit(ByVal sSharedName As String, _
    ByVal declarations As CodeDom.CodeTypeMemberCollection, _
    ByVal InitGraphics As CodeDom.CodeMemberMethod, _
    ByVal RenderGDI As CodeDom.CodeStatementCollection, _
    ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
    ByVal ExportSettings As ExportSettings) As String


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits code.  Inheritors are responsible for implementing this class.
    ''' </summary>
    ''' <param name="obj">The parent object whose fill is being emitted.</param>
    ''' <param name="declarations">CodeDOM declaration sections</param>
    ''' <param name="InitGraphics">The GDI+ Architect initializegrapics method</param>
    ''' <param name="RenderGDI">The GDI+ architect render GDI method</param>
    ''' <param name="DisposeGDI">The GDI+ architect dispose GDI method.</param>
    ''' <param name="ExportSettings">The export settings.</param>
    ''' <param name="Consolidated">The set of consolidated objects thus far.</param>
    ''' <returns>A string containing the name of the fill created in code.</returns>
    ''' -----------------------------------------------------------------------------
    Friend MustOverride Function emit(ByVal obj As GDIObject, _
    ByVal declarations As CodeDom.CodeTypeMemberCollection, _
    ByVal InitGraphics As CodeDom.CodeMemberMethod, _
    ByVal RenderGDI As CodeDom.CodeStatementCollection, _
    ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
    ByVal ExportSettings As ExportSettings, _
    ByVal Consolidated As ExportConsolidate) As String
#End Region

#Region "Delegates and Events"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Provides a way for inherited objects to raise the FillUpdated event as needed.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected Sub NotifyFillUpdated()
        RaiseEvent FillUpdated(Me, EventArgs.Empty)
    End Sub

#End Region


#Region "Code Generation Related Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if two fills are identical (for when consolidating fills).
    ''' </summary>
    ''' <param name="fill1">First fill to compare</param>
    ''' <param name="fill2">Second fill to compare.</param>
    ''' <returns>A Boolean indicating whether the two fills are equiv.</returns>
    ''' -----------------------------------------------------------------------------
    Public Shared Function ExportEquality(ByVal fill1 As GDIFill, ByVal fill2 As GDIFill) As Boolean
        Trace.WriteLineIf(Session.Tracing, "Fill.ExportEquality")

        If fill1.GetType().Name = fill2.GetType().Name Then
            If TypeOf fill1 Is GDIGradientFill Then
                Return GDIGradientFill.ExportEquality(DirectCast(fill1, GDIGradientFill), DirectCast(fill2, GDIGradientFill))
            ElseIf TypeOf fill1 Is GDISolidFill Then
                Return GDISolidFill.op_Equality(DirectCast(fill1, GDISolidFill), DirectCast(fill2, GDISolidFill))
            ElseIf TypeOf fill1 Is GDIHatchFill Then
                Return GDIHatchFill.op_Equality(DirectCast(fill1, GDIHatchFill), DirectCast(fill2, GDIHatchFill))
            ElseIf TypeOf fill1 Is GDITexturedFill Then
                Return GDITexturedFill.ExportEquality(DirectCast(fill1, GDITexturedFill), DirectCast(fill2, GDITexturedFill))

            End If
        Else
            Return False
        End If
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retrieves a color assignment statement.  This is used when emitting code.
    ''' </summary>
    ''' <param name="val">The color.</param>
    ''' <returns>A code expression that assigns the color to the appropriate value.</returns>
    ''' <remarks>This statement first attempts to used named colors if possible.
    ''' If this isn't possible, it uses the ARGB values instead.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Protected Function getColorAssignment(ByVal val As Color) As CodeExpression
        Trace.WriteLineIf(Session.Tracing, "Fill.getColorAssignment")

        Dim argbColorInvocation As New CodeDom.CodeMethodInvokeExpression
        Dim namedColorField As New CodeDom.CodeFieldReferenceExpression

        If val.IsNamedColor AndAlso val.A = 255 Then

            With namedColorField
                .TargetObject = New CodeTypeReferenceExpression(GetType(System.Drawing.Color))
                .FieldName = val.Name
                Return namedColorField
            End With

        Else

            With argbColorInvocation
                .Method.TargetObject = New CodeTypeReferenceExpression(GetType(System.Drawing.Color))
                .Method.MethodName = "FromArgb"
                .Parameters.Add(New CodePrimitiveExpression(val.A))
                .Parameters.Add(New CodePrimitiveExpression(val.R))
                .Parameters.Add(New CodePrimitiveExpression(val.G))
                .Parameters.Add(New CodePrimitiveExpression(val.B))
            End With

            Return argbColorInvocation

        End If

    End Function
#End Region

#Region "Property Accessors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to consolidate this fill with identical fills.
    ''' </summary>
    ''' <value>True if the fill should consolidate, false otherwise</value>
    ''' -----------------------------------------------------------------------------
    <Description("Whether on export to consolidate with identical fills.  See help for more details.")> _
    Public Property Consolidate() As Boolean
        Get
            Return _Consolidate
        End Get
        Set(ByVal Value As Boolean)
            _Consolidate = Value
            Trace.WriteLineIf(Session.Tracing, "Fill.Consolidate: " & Value.ToString)
        End Set
    End Property


#End Region

#Region "Cleanup and Disposal"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Calls the default dispose
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the fill, releasing the brush resource.
    ''' </summary>
    ''' <param name="disposing">Whether unmanaged resources are being disposed or not.</param>
    ''' -----------------------------------------------------------------------------
    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        ' Check to see if Dispose has already been called.
        If Not _Disposed Then

            If disposing Then
                ' Dispose managed resources
                If Not _Brush Is Nothing Then
                    _Brush.Dispose()
                End If
            End If
            ' Release unmanaged resources.    

        End If

        _Disposed = True
    End Sub



#End Region

#Region "Misc Methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Clones a fill of a specific object, returning a new fill.
    ''' </summary>
    ''' <returns>A new fill with identical properties to the original fill.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function Clone() As GDIFill
        If TypeOf Me Is GDISolidFill Then
            Return New GDISolidFill(DirectCast(Me, GDISolidFill))
        ElseIf TypeOf Me Is GDIHatchFill Then
            Return New GDIHatchFill(DirectCast(Me, GDIHatchFill))
        ElseIf TypeOf Me Is GDIGradientFill Then
            Return New GDIGradientFill(DirectCast(Me, GDIGradientFill))
        ElseIf TypeOf Me Is GDITexturedFill Then
            Return New GDITexturedFill(DirectCast(Me, GDITexturedFill))

        End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a System.Drawing brush capable of painting the current fill to a surface.
    ''' </summary>
    ''' <value>A Brush with which the a graphics context can be painted to render 
    ''' the fill.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Brush() As Brush
        Get
            If _Brush Is Nothing Then
                UpdateFill()
            End If

            Return _Brush

        End Get
    End Property


#End Region


End Class
