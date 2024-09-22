Imports System.ComponentModel
Imports System.CodeDom

#Region "Type Declarations"



''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : StrokeBrowser
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Class used to render strokes inside the property grid.  For more information 
''' on how the mechanics of this works, see the MSDN documentation on 
''' the ExpandableObjectConvertor class.
''' </summary>
''' -----------------------------------------------------------------------------
Public Class StrokeBrowser
    Inherits ExpandableObjectConverter



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Makes the property grid display "Stroke Properties" instead of "GDIStroke"
    ''' </summary>
    ''' <param name="context">ITypeDescriptorContext</param>
    ''' <param name="culture">current culture</param>
    ''' <param name="value">Object being described.</param>
    ''' <param name="destinationType">Expected output type of convertTo.  For the prop grid 
    ''' this is always a string.</param>
    ''' <returns>The string  "Stroke Properties"</returns>
    ''' -----------------------------------------------------------------------------
    Public Overloads Overrides Function ConvertTo(ByVal context As System.ComponentModel.ITypeDescriptorContext, ByVal culture As System.Globalization.CultureInfo, ByVal value As Object, ByVal destinationType As System.Type) As Object
        Return "Stroke Properties"
    End Function

End Class

#End Region
''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIStroke
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' The GDIStroke class is responsible for providing a stroke to strokable GDIObjects.  This 
''' maps to pen functionality in System.Drawing.   
''' </summary>
'''  -----------------------------------------------------------------------------
<Serializable(), TypeConverter(GetType(StrokeBrowser))> _
Public Class GDIStroke
    Implements IDisposable



#Region "Non persisted fields"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether this object has been disposed or not
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _Disposed As Boolean = False

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A pen used to draw the GDIStroke as needed
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _Pen As Pen
#End Region

#Region "Field Declarations"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to consolidate this stroke with similar strokes during code export.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Consolidate As Boolean

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Alignment property of the stroke.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Alignment As Drawing2D.PenAlignment = Drawing2D.PenAlignment.Center


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' DashStyle of the stroke
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _DashStyle As Drawing2D.DashStyle = Drawing2D.DashStyle.Solid

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Width of the stroke
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Width As Single = 1
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Dash offset of the stroke
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _DashOffset As Single = 0.0!

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Start cap of the stroke
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _StartCap As Drawing2D.LineCap = Drawing2D.LineCap.Flat
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' End cap of the stroke 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _EndCap As Drawing2D.LineCap = Drawing2D.LineCap.Flat
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Dash cap used on the stroke
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _DashCap As Drawing2D.DashCap = Drawing2D.DashCap.Flat

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Line join (see Drawing2D.LineJoin for more information about this property)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _LineJoin As Drawing2D.LineJoin = Drawing2D.LineJoin.Miter

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Miter limit of the stroke (see Drawing2D for information on miter limits)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _MiterLimit As Single = 10
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The color to draw this stroke with.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Color As Color = Color.Black

#End Region

#Region "Event Declarations"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Raised when a stroke is updated.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Event StrokeUpdated(ByVal s As Object, ByVal e As EventArgs)

#End Region

#Region "Constructors"

    
    Public Sub New()
        MyBase.New()
    End Sub

    '''' -----------------------------------------------------------------------------
    '''' <summary>
    '''' Constructor used for the "sessions" stroke - the stroke currently being used 
    '''' to stroke new objects.  All other strokes created in GDI+ Architect use the other 
    '''' constructor, which typically takes the stroke instantiated in the Session.
    '''' </summary>
    '''' <param name="parent">Parent shape this stroke will render.  In the one time this 
    '''' constructor is used, however, this value is Nothing (null) and 
    '''' could be removed.</param>
    '''' -----------------------------------------------------------------------------
    'Public Sub New(ByVal parent As GDIShape)
    '    _Consolidate = Session.Settings.ConsolidateStrokes
    'End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a stroke given another a parent shape and another stroke to base it on.
    ''' </summary>
    ''' <param name="parent">The parent shape used to stroke objects</param>
    ''' <param name="stroke">The stroke to base the new stroke off of.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal stroke As GDIStroke)

        _Alignment = stroke.Alignment
        _Color = stroke.Color
        _DashCap = stroke.DashCap
        _DashStyle = stroke.DashStyle
        _DashOffset = stroke.DashOffset
        _EndCap = stroke.Endcap
        _LineJoin = stroke.LineJoin
        _MiterLimit = stroke.MiterLimit
        _StartCap = stroke.Startcap()
        _Width = stroke.Width()
        _Consolidate = Session.Settings.ConsolidateFills
    End Sub

#End Region

#Region "Code Generation"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a CodeDOM usable color expression.  This is used during the code generation
    ''' options.
    ''' </summary>
    ''' <param name="val">The color to return a code representation of.</param>
    ''' <returns>A code expression representing the color.</returns>
    ''' <remarks>Notice that this code first attempts to use named colors if possible.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Function getColorAssignment(ByVal val As Color) As CodeExpression
        Dim outparamInvoke As New CodeDom.CodeMethodInvokeExpression
        Dim simpleOut As New CodeDom.CodeFieldReferenceExpression

        If val.IsNamedColor AndAlso val.A = 255 Then
            With simpleOut
                .TargetObject = New CodeTypeReferenceExpression(GetType(Drawing.Color))
                .FieldName = val.Name
                Return simpleOut
            End With

        Else
            With outparamInvoke
                .Method.TargetObject = New CodeTypeReferenceExpression(GetType(Drawing.Color))
                .Method.MethodName = "FromArgb"
                .Parameters.Add(New CodePrimitiveExpression(val.A))
                .Parameters.Add(New CodePrimitiveExpression(val.R))
                .Parameters.Add(New CodePrimitiveExpression(val.G))
                .Parameters.Add(New CodePrimitiveExpression(val.B))
            End With

            Return outparamInvoke

        End If


    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a stroke to code.
    ''' </summary>
    ''' <param name="parent">The parent object this stroke belongs to.</param>
    ''' <param name="Declarations">The declarations section of the CodeDOM.</param>
    ''' <param name="InitGraphics">The GDI+ Architect's InitGraphics method.</param>
    ''' <param name="RenderGDI">The GDI+ Architect RenderGDI method.</param>
    ''' <param name="DisposeGDI">The GDI+ Architect's dispose method</param>
    ''' <param name="ExportSettings">The currently selected export settings.</param>
    ''' <param name="Consolidated">The set of consolidated objects.</param>
    ''' <returns>The name of the stroke created in code.</returns>
    ''' -----------------------------------------------------------------------------
    Friend Overridable Function emit(ByVal parent As GDIObject, _
    ByVal Declarations As CodeDom.CodeTypeMemberCollection, _
    ByVal InitGraphics As CodeDom.CodeMemberMethod, _
    ByVal RenderGDI As CodeDom.CodeStatementCollection, _
    ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
    ByVal ExportSettings As ExportSettings, _
    ByVal Consolidated As ExportConsolidate) As String


        If Consolidated.hasStrokeMatch(Me) AndAlso Me.Consolidate Then
            Return Consolidated.getStrokeName(Me)
        Else
            Dim declareStroke As New CodeMemberField
            Dim createStatement As New CodeObjectCreateExpression
            Dim disposeStroke As New CodeDom.CodeMethodInvokeExpression

            'Create the initialization args and assignment  for stroke
            With createStatement
                .CreateType = New CodeTypeReference(GetType(Pen))
                .Parameters.Add(getColorAssignment(Me.Color))
                .Parameters.Add(New CodePrimitiveExpression(Me.Width))
            End With

            'Create the member variable for the stroke
            With declareStroke
                .InitExpression = createStatement
                .Name = parent.ExportName & "Pen"
                .Attributes = MemberAttributes.Private
                .Type = New CodeTypeReference(GetType(System.Drawing.Pen))
            End With

            If Not _Alignment = CType(0, Drawing2D.PenAlignment) Then
                Dim AssignAlignment As New CodeAssignStatement
                AssignAlignment.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "Alignment")
                AssignAlignment.Right = New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(Drawing2D.PenAlignment)), _Alignment.ToString)
                InitGraphics.Statements.Add(AssignAlignment)
            End If

            If Not _DashCap = CType(0, Drawing2D.DashCap) Then
                Dim AssignDashCap As New CodeAssignStatement
                AssignDashCap.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "DashCap")
                AssignDashCap.Right = New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(Drawing2D.DashCap)), _DashCap.ToString)
                InitGraphics.Statements.Add(AssignDashCap)
            End If

            If Not _DashOffset = 0.0! Then
                Dim AssignDashOffset As New CodeAssignStatement

                AssignDashOffset.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "DashOffset")
                AssignDashOffset.Right = New CodePrimitiveExpression(_DashOffset)
                InitGraphics.Statements.Add(AssignDashOffset)
            End If


            If Not _DashStyle = CType(0, Drawing2D.DashStyle) Then
                Dim AssignDashStyle As New CodeAssignStatement
                AssignDashStyle.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "DashStyle")
                AssignDashStyle.Right = New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(Drawing2D.DashStyle)), _DashStyle.ToString)
                InitGraphics.Statements.Add(AssignDashStyle)
            End If

            If Not _StartCap = CType(0, Drawing2D.LineCap) Then
                Dim AssignStartcap As New CodeAssignStatement
                AssignStartcap.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "StartCap")
                AssignStartcap.Right = New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(Drawing2D.LineCap)), _StartCap.ToString)
                InitGraphics.Statements.Add(AssignStartcap)
            End If

            If Not _EndCap = CType(0, Drawing2D.LineCap) Then
                Dim AssignEndcap As New CodeAssignStatement
                AssignEndcap.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "EndCap")
                AssignEndcap.Right = New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(Drawing2D.LineCap)), _EndCap.ToString)
                InitGraphics.Statements.Add(AssignEndcap)
            End If

            If Not _LineJoin = CType(0, Drawing2D.LineJoin) Then
                Dim AssignJoin As New CodeAssignStatement
                AssignJoin.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "LineJoin")
                AssignJoin.Right = New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(Drawing2D.LineJoin)), _LineJoin.ToString)
                InitGraphics.Statements.Add(AssignJoin)
            End If

            If Not _MiterLimit = 10.0! Then
                Dim MiterLimit As New CodeAssignStatement
                MiterLimit.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "MiterLimit")
                MiterLimit.Right = New CodePrimitiveExpression(_MiterLimit)
                InitGraphics.Statements.Add(MiterLimit)
            End If

            'create dispose call
            With disposeStroke
                .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, declareStroke.Name)
                .Method.MethodName = "Dispose"
            End With

            'add object disposal
            With DisposeGDI
                .Statements.Add(disposeStroke)
            End With

            With Declarations
                .Add(declareStroke)
            End With


            Return declareStroke.Name
        End If


    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts a stroke to an SVG XML representation.
    ''' </summary>
    ''' <param name="xmldoc">The SVG document containing the stroke.</param>
    ''' <param name="parent">The parent node to append the stroke to.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub toXML(ByVal xmldoc As Xml.XmlDocument, ByVal parent As Xml.XmlNode)
        Dim arg As String = String.Empty

        Dim attr As Xml.XmlAttribute


        If Me.Color.IsNamedColor Then
            arg = Me.Color.Name
        Else
            arg = String.Format("rgb({0},{1},{2})", Me.Color.R, Me.Color.G, Me.Color.B)
        End If

        attr = xmldoc.CreateAttribute("stroke")
        attr.Value = arg
        parent.Attributes.Append(attr)

        attr = xmldoc.CreateAttribute("stroke-width")
        attr.Value = String.Format("{0}", Me.Width.ToString)
        parent.Attributes.Append(attr)


        If Me.Color.A < 255 Then
            arg = FormatNumber(Color.A / 255, 3)
            attr = xmldoc.CreateAttribute("stroke-opacity")
            attr.Value = arg
            parent.Attributes.Append(attr)
        End If

        If Not _Alignment = CType(0, Drawing2D.PenAlignment) Then
            arg = System.Enum.GetName(GetType(Drawing2D.PenAlignment), _Alignment)
            attr = xmldoc.CreateAttribute("stroke-alignment")
            attr.Value = arg
            parent.Attributes.Append(attr)
        End If

        If Not _DashCap = CType(0, Drawing2D.DashCap) Then
            arg = System.Enum.GetName(GetType(Drawing2D.DashCap), _DashCap)
            attr = xmldoc.CreateAttribute("stroke-dashcap")
            attr.Value = arg
            parent.Attributes.Append(attr)
        End If

        If Not _DashOffset = 0.0! Then
            arg = _DashOffset.ToString
            attr = xmldoc.CreateAttribute("stroke-dashoffset")
            attr.Value = arg
            parent.Attributes.Append(attr)
        End If


        If Not _DashStyle = CType(0, Drawing2D.DashStyle) AndAlso Not _DashStyle = Drawing2D.DashStyle.Solid Then
            'stroke-dasharray' 
            Dim sDashArray As String

            Select Case _DashStyle
                Case Drawing2D.DashStyle.Dash
                    arg = 3 * Me.Width & " " & 3 * Me.Width
                Case Drawing2D.DashStyle.DashDot
                    arg = 3 * Me.Width & " " & Me.Width & " " & Me.Width & " " & Me.Width
                Case Drawing2D.DashStyle.DashDotDot
                    arg = 3 * Me.Width & " " & Me.Width & " " & Me.Width & " " & Me.Width & " " & Me.Width
                Case Drawing2D.DashStyle.Dot
                    arg = Me.Width & " " & Me.Width
            End Select

            attr = xmldoc.CreateAttribute("stroke-dasharray")
            attr.Value = arg
            parent.Attributes.Append(attr)

        End If

        If Not _StartCap = CType(0, Drawing2D.LineCap) Then
            arg = System.Enum.GetName(GetType(Drawing2D.LineCap), _StartCap)
            attr = xmldoc.CreateAttribute("stroke-StartCap")
            attr.Value = arg
            parent.Attributes.Append(attr)
        End If

        If Not _EndCap = CType(0, Drawing2D.LineCap) Then
            arg = System.Enum.GetName(GetType(Drawing2D.LineCap), _EndCap)
            attr = xmldoc.CreateAttribute("stroke-EndCap")
            attr.Value = arg
            parent.Attributes.Append(attr)
        End If

        If Not _LineJoin = CType(0, Drawing2D.LineJoin) Then

            Select Case _LineJoin
                Case Drawing2D.LineJoin.Bevel
                    arg = "bevel"
                Case Drawing2D.LineJoin.Miter
                    arg = "miter"
                Case Drawing2D.LineJoin.MiterClipped
                    arg = "miterclipped"
                Case Drawing2D.LineJoin.Round
                    arg = "round"
            End Select

            attr = xmldoc.CreateAttribute("stroke-linejoin")
            attr.Value = arg
            parent.Attributes.Append(attr)
        End If

        If Not _MiterLimit = 10.0! Then
            arg = _MiterLimit.ToString
            attr = xmldoc.CreateAttribute("stroke-miterlimit")
            attr.Value = arg
            parent.Attributes.Append(attr)
        End If


    End Sub



#End Region


#Region "public Methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a copy of a stroke
    ''' </summary>
    ''' <returns>A new copy of the current stroke.  A clone.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function Clone() As GDIStroke
        Return New GDIStroke(Me)
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if two strokes are equal (all properties the same).
    ''' </summary>
    ''' <param name="stroke1">The first stroke to compare</param>
    ''' <param name="stroke2">The second stroke to compare</param>
    ''' <returns>A Boolean indicating if the two strokes are equal.</returns>
    ''' -----------------------------------------------------------------------------
    Public Shared Function op_Equality(ByVal stroke1 As GDIStroke, ByVal stroke2 As GDIStroke) As Boolean
        Dim bEqual As Boolean = True

        bEqual = _
            System.Drawing.Color.op_Equality(stroke1.Color, stroke2.Color) AndAlso _
            stroke1.Alignment = stroke2.Alignment AndAlso _
            stroke1.DashCap = stroke2.DashCap AndAlso _
            stroke1.DashStyle = stroke2.DashStyle AndAlso _
            stroke1.Endcap = stroke2.Endcap AndAlso _
            stroke1.LineJoin = stroke2.LineJoin AndAlso _
            stroke1.MiterLimit = stroke2.MiterLimit AndAlso _
            stroke1.Width = stroke2.Width AndAlso _
            stroke1.DashOffset = stroke2.DashOffset

        Return bEqual

    End Function
#End Region
 

#Region "Property Accessors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A pen used to stroke objects that has all properties of the stroke assigned to 
    ''' it.
    ''' </summary>
    ''' <value>A pen that shapes can use to paint their edges with.
    ''' </value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Pen() As Pen
        Get
            If _Pen Is Nothing Then
                UpdateStroke()
            End If

            Return _Pen
        End Get
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets whether to add this stroke to the consolidation set or not.
    ''' </summary>
    ''' <value>A Boolean indicating whether to consolidate this stroke or not</value>
    ''' -----------------------------------------------------------------------------
    <Description("Whether on export to consolidate with identical strokes.  See help for more details.")> _
    Public Property Consolidate() As Boolean
        Get
            Return _Consolidate
        End Get
        Set(ByVal Value As Boolean)
            _Consolidate = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the stroke's dash cap
    ''' </summary>
    ''' <value>A Drawing2D.DashCap</value>
    ''' -----------------------------------------------------------------------------
    <DefaultValue(GetType(Drawing2D.DashCap), "Flat"), Description("Cap on dash points.  Only applies if the stroke has a dash pattern")> _
    Public Overridable Property DashCap() As Drawing2D.DashCap
        Get
            Return _DashCap

        End Get
        Set(ByVal Value As Drawing2D.DashCap)
            _DashCap = Value
            UpdateStroke()
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the stroke's starting line cap
    ''' </summary>
    ''' <value>A Drawing2D.LineCap</value>
    ''' -----------------------------------------------------------------------------
    <DefaultValue(GetType(Drawing2D.LineCap), "Flat"), Description("Cap at the start of a line.  Only applies to open pen paths and lines")> _
    Public Overridable Property Startcap() As Drawing2D.LineCap
        Get
            Return _StartCap
        End Get
        Set(ByVal Value As Drawing2D.LineCap)
            _StartCap = Value
            UpdateStroke()
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets a stroke's end cap
    ''' </summary>
    ''' <value>A Drawing2D.EndCap</value>
    ''' -----------------------------------------------------------------------------  
    <DefaultValue(GetType(Drawing2D.LineCap), "Flat"), Description("Cap at the end of a line.  Only applies to open pen paths and lines")> _
     Public Overridable Property Endcap() As Drawing2D.LineCap
        Get
            Return _EndCap
        End Get
        Set(ByVal Value As Drawing2D.LineCap)
            _EndCap = Value
            UpdateStroke()

        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the stroke's alignment
    ''' </summary>
    ''' <value>A Drawing2D.Alignment</value>
    ''' -----------------------------------------------------------------------------  
    <DefaultValue(GetType(Drawing2D.PenAlignment), "Center"), Description("Alignment of the stroke in relation to the points it passes through")> _
        Public Overridable Property Alignment() As Drawing2D.PenAlignment
        Get
            Return _Alignment
        End Get
        Set(ByVal Value As Drawing2D.PenAlignment)
            _Alignment = Value
            UpdateStroke()
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets a stroke's LineJoin
    ''' </summary>
    ''' <value>A Drawing2D.LineJoin</value>
    ''' -----------------------------------------------------------------------------  
    <DefaultValue(GetType(Drawing2D.LineJoin), "Miter"), Description("The style of the edges where points on the stroke meet")> _
        Public Overridable Property LineJoin() As Drawing2D.LineJoin
        Get
            Return _LineJoin
        End Get
        Set(ByVal Value As Drawing2D.LineJoin)
            _LineJoin = Value
            UpdateStroke()
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the dash offset of the stroke.
    ''' </summary>
    ''' <value>A Single (float) of how much to offset the dash pattern by.</value>
    ''' -----------------------------------------------------------------------------
    <DefaultValue(0.0!), Description("Offset of the dash pattern from start point")> _
    Public Overridable Property DashOffset() As Single
        Get
            Return _DashOffset
        End Get
        Set(ByVal Value As Single)

            _DashOffset = Value
            UpdateStroke()
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the dash style of the stroke.
    ''' </summary>
    ''' <value>A Drawing2D.DashStyle</value>
    ''' -----------------------------------------------------------------------------
    <DefaultValue(GetType(Drawing2D.DashStyle), "Solid"), Description("Dash pattern of the stroke")> _
    Public Overridable Property DashStyle() As Drawing2D.DashStyle
        Get
            Return _DashStyle
        End Get
        Set(ByVal Value As Drawing2D.DashStyle)
            _DashStyle = Value
            UpdateStroke()
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the width of a stroke
    ''' </summary>
    ''' <value>A single (float) representing the width of the stroke.</value>
    ''' -----------------------------------------------------------------------------
    <DefaultValue(1.0!), Description("Width of the stroke")> _
    Public Overridable Property Width() As Single
        Get
            Return _Width
        End Get
        Set(ByVal Value As Single)
            If Value > 0 Then
                _Width = Value
                UpdateStroke()
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the color used to fill a stroke
    ''' </summary>
    ''' <value>A valid Drawing.Color</value>
    ''' -----------------------------------------------------------------------------
    <Description("Color of the stroke")> _
    Public Overridable Property Color() As Color
        Get
            Return _Color
        End Get
        Set(ByVal Value As Color)
            _Color = Value
            UpdateStroke()
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the MiterLimit of a stroke (see MSDN for more details about MiterLimit)
    ''' </summary>
    ''' <value>A Single (float) representing the miter limit.</value>
    ''' -----------------------------------------------------------------------------
    <BrowsableAttribute(False), DefaultValue(10.0!)> _
    Public Overridable Property MiterLimit() As Single
        Get
            Return _MiterLimit
        End Get
        Set(ByVal Value As Single)
            _MiterLimit = Value
            UpdateStroke()
        End Set
    End Property


#End Region

#Region "Implementation"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new pen object as stroke properties are changed.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub UpdateStroke()
        If _Pen Is Nothing Then
            _Pen = New Pen(_Color, _Width)
        End If

        With _Pen
            .Width = _Width
            .Color = _Color
            .DashStyle = _DashStyle
            .Alignment = _Alignment
            .DashOffset = _DashOffset
            .StartCap = _StartCap
            .EndCap = _EndCap
            .DashCap = _DashCap
            .LineJoin = _LineJoin
            .MiterLimit = _MiterLimit
        End With

        RaiseEvent StrokeUpdated(Me, EventArgs.Empty)
    End Sub

#End Region


#Region "Cleanup and Disposal"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the GDIStroke.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the GDIStroke, specifically disposing of the _Pen.
    ''' </summary>
    ''' <param name="disposing">Whether managed resources are being disposed or not.</param>
    ''' -----------------------------------------------------------------------------
    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        ' Check to see if Dispose has already been called.
        If Not _Disposed Then

            If disposing Then
                ' Dispose managed resources
                If Not _Pen Is Nothing Then
                    _Pen.Dispose()
                End If
            End If
            ' Release unmanaged resources.    

        End If

        _Disposed = True
    End Sub

#End Region

End Class
