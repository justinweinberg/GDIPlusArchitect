Imports System.CodeDom
Imports System.Drawing.Drawing2D

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIGradientFill
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A simple two color gradient fill for fillable GDIObjects.  Contains a start color and 
''' end color, a gradient style, and optional gradient angle.  Also contains a point from 
''' which to draw the gradient from depending on whether the GradientFill is tracking its 
''' parent location.
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public Class GDIGradientFill
    Inherits GDIFill

#Region "Type Declarations"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An Enumeration of Gradient Directions.  All but Custom map to the drawing 
    ''' namespace's LinearGradientMode mode.  Custom means an angle will be used instead 
    ''' of a gradient mode.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Enum EnumGradientMode As Integer
        ''' <summary>Maps to LinearGradientMode.Horizontal</summary>
        Horizontal = 0
        ''' <summary>Maps to LinearGradientMode.Horizontal</summary>
         BackwardDiagonal = 3
        ''' <summary>Maps to LinearGradientMode.BackwardDiagonal</summary>
        ForwardDiagonal = 2
        ''' <summary>Maps to LinearGradientMode.ForwardDiagonal</summary>
        Vertical = 1
        ''' <summary>Indicates a custom angle argument will be specified 
        ''' instead of a using a gradient mode.</summary>
        Custom = 101
    End Enum

#End Region

#Region "Local Fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The area over which to paint the gradient.  
    ''' When "track fill" is set this becomes the parent's bounds.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _GradientArea As Rectangle = New Rectangle(0, 0, 100, 100)

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The first color in the two color gradient.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Color1 As Color = Color.Black
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The second color in the two color gradient.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Color2 As Color = Color.White
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The gradient mode (maps to the build in System.Drawing.Drawing2D.LinearGradientMode
    ''' which the exception of EnumGradientMode.Custom which indicates an angle will 
    ''' be specified for the gradient).
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _GradientMode As EnumGradientMode

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The angle to project the gradient at (used only  when _GradientMode is set 
    ''' to custom.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _GradientAngle As Single = 90.0!
#End Region


#Region "Constructors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new gradient fill given the fill colors and a gradient mode
    ''' </summary>
    ''' <param name="parent">The parent object being filled with this fill</param>
    ''' <param name="color1">The first color in the two color gradient fill</param>
    ''' <param name="color2">The second color in the two color gradient fill</param>
    ''' <param name="eMode">The fill mode</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal color1 As Color, ByVal color2 As Color, ByVal eMode As EnumGradientMode)
        MyBase.New()
        Trace.WriteLineIf(Session.Tracing, "GradientFill.New.General")

        _GradientArea = New Rectangle(0, 0, 100, 100)
        _Color1 = color1
        _Color2 = color2
        _GradientMode = eMode

     End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new gradient fill given a second gradient fill to base it on.
    ''' </summary>
    ''' <param name="parent">The parent object to be filled.</param>
    ''' <param name="fill">The fill to base the gradient on.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal fill As GDIGradientFill)
        MyBase.New()
        Trace.WriteLineIf(Session.Tracing, "GradientFill.New.Clone")

        _Color1 = fill.Color1
        _Color2 = fill.Color2
        _GradientMode = fill.GradientMode
        _GradientArea = fill.Bounds
        _GradientAngle = fill.Angle
    End Sub

#End Region

#Region "Code Generation Related methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if two gradients share export equality for purposes of consolidation
    ''' </summary>
    ''' <param name="fill1">The first gradient fill to examine</param>
    ''' <param name="fill2">The second gradient fill to examine</param>
    ''' <returns>A Boolean indicating whether the two gradients are equivalent.</returns>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function ExportEquality(ByVal fill1 As GDIGradientFill, ByVal fill2 As GDIGradientFill) As Boolean

        Return op_Equality(fill1, fill2) AndAlso Rectangle.op_Equality(fill1._GradientArea, fill2._GradientArea)

    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts the gradient fill to an XML source for SVG display.
    ''' </summary>
    ''' <param name="xmlDoc">See base class</param>
    ''' <param name="defs">See base class</param>
    ''' <param name="parent">See base class</param>
    ''' ----------------------------------------------------------------------------- 
    Public Overrides Sub toXML(ByVal xmlDoc As Xml.XmlDocument, ByVal defs As Xml.XmlNode, ByVal parent As Xml.XmlNode)
        '<linearGradient id="Gradient01">
        '    <linearGradient id="MyGradient">
        '<stop offset="5%" stop-color="#F60" />
        ' <stop offset="95%" stop-color="#FF6" />
        '</linearGradient>

        '   </linearGradient>
        ' <stop offset="20%" stop-color="#39F" />

        Dim attr As Xml.XmlAttribute

        Dim nodeGrad As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "linearGradient", String.Empty)

        defs.AppendChild(nodeGrad)


        attr = xmlDoc.CreateAttribute("id")
        attr.Value = parent.Attributes("id").Value & "Brush"
        nodeGrad.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("gradientUnits")
        attr.Value = "userSpaceOnUse"
        nodeGrad.Attributes.Append(attr)




        Dim x1, x2, y1, y2 As Int32


        Select Case Me.GradientMode
            Case EnumGradientMode.BackwardDiagonal
                x1 = Me.Bounds.X + Me.Bounds.Width
                x2 = Me.Bounds.X
                y1 = Me.Bounds.Y
                y2 = Me.Bounds.Y + Me.Bounds.Height

            Case EnumGradientMode.ForwardDiagonal
                x1 = Me.Bounds.X
                x2 = Me.Bounds.X + Me.Bounds.Width
                y1 = Me.Bounds.Y
                y2 = Me.Bounds.Y + Me.Bounds.Height

            Case EnumGradientMode.Horizontal
                x1 = Me.Bounds.X
                x2 = Me.Bounds.X + Me.Bounds.Width
                y1 = Me.Bounds.Y
                y2 = Me.Bounds.Y
            Case EnumGradientMode.Vertical
                x1 = Me.Bounds.X
                x2 = Me.Bounds.X
                y1 = Me.Bounds.Y
                y2 = Me.Bounds.Y + Me.Bounds.Height
            Case EnumGradientMode.Custom

        End Select

        attr = xmlDoc.CreateAttribute("x1")
        attr.Value = x1.ToString
        nodeGrad.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("y1")
        attr.Value = y1.ToString
        nodeGrad.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("x2")
        attr.Value = x2.ToString
        nodeGrad.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("y2")
        attr.Value = y2.ToString
        nodeGrad.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("GDIGradientBounds")
        attr.Value = String.Format("{0}, {1}, {2}, {3}", Me.Bounds.X, Me.Bounds.Y, Me.Bounds.Width, Me.Bounds.Height)
        nodeGrad.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("GDIGradientMode")
        attr.Value = String.Format("{0}", System.Enum.GetName(GetType(GDIGradientFill.EnumGradientMode), Me.GradientMode))
        nodeGrad.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("spreadMethod")
        attr.Value = "repeat"
        nodeGrad.Attributes.Append(attr)

        Dim stopNode As Xml.XmlNode

        'StopNode 1
        stopNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "stop", String.Empty)

        attr = xmlDoc.CreateAttribute("offset")
        attr.Value = "0"
        stopNode.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("stop-color")
        If Me.Color1.IsNamedColor Then
            attr.Value = Me.Color1.Name
        Else
            attr.Value = String.Format("rgb({0},{1}, {2})", Me.Color1.R, Me.Color1.G, Me.Color1.B)
        End If

        stopNode.Attributes.Append(attr)

        If Me.Color1.A < 255 Then
            attr = xmlDoc.CreateAttribute("stop-opacity")
            attr.Value = FormatNumber(Me.Color1.A / 255, 3)
            stopNode.Attributes.Append(attr)
        End If

        stopNode.Attributes.Append(attr)

        nodeGrad.AppendChild(stopNode)


        'StopNode 2
        stopNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "stop", String.Empty)

        attr = xmlDoc.CreateAttribute("offset")
        attr.Value = "1"
        stopNode.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("stop-color")
        If Me.Color2.IsNamedColor Then
            attr.Value = Me.Color2.Name
        Else
            attr.Value = String.Format("rgb({0},{1}, {2})", Me.Color2.R, Me.Color2.G, Me.Color2.B)
        End If

        stopNode.Attributes.Append(attr)

        If Me.Color2.A < 255 Then
            attr = xmlDoc.CreateAttribute("stop-opacity")
            attr.Value = FormatNumber(Me.Color2.A / 255, 3)
            stopNode.Attributes.Append(attr)
        End If



        stopNode.Attributes.Append(attr)

        nodeGrad.AppendChild(stopNode)


        attr = xmlDoc.CreateAttribute("fill")
        attr.Value = String.Format("url(#{0})", parent.Attributes("id").Value & "Brush")
        parent.Attributes.Append(attr)
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a gradient fill to code.
    ''' </summary>
    ''' <param name="obj">See base class</param>
    ''' <param name="declarations">See base class</param>
    ''' <param name="InitGraphics">See base class</param>
    ''' <param name="RenderGDI">See base class</param>
    ''' <param name="DisposeGDI">See base class</param>
    ''' <param name="ExportSettings">See base class</param>
    ''' <param name="Consolidated">See base class</param>
    ''' <returns>See base class</returns>
        '''  -----------------------------------------------------------------------------
    Friend Overloads Overrides Function emit(ByVal obj As GDIObject, _
  ByVal declarations As CodeDom.CodeTypeMemberCollection, _
  ByVal InitGraphics As CodeDom.CodeMemberMethod, _
  ByVal RenderGDI As CodeDom.CodeStatementCollection, _
  ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
  ByVal ExportSettings As ExportSettings, _
  ByVal Consolidated As ExportConsolidate) As String

        Trace.WriteLineIf(Session.Tracing, "GradientFill.emit.Consolidated")

        If Consolidated.hasFillMatch(Me) And Me.Consolidate Then
            Return Consolidated.getFillName(Me)
        Else


            Dim fillDeclaration As New CodeMemberField
            Dim fillInit As New CodeObjectCreateExpression
            Dim fillDispose As New CodeDom.CodeMethodInvokeExpression

            Dim fillRectParam As New CodeDom.CodeObjectCreateExpression

            With fillRectParam
                .CreateType = New CodeTypeReference(GetType(System.Drawing.Rectangle))
                .Parameters.Add(New CodePrimitiveExpression(_GradientArea.X))
                .Parameters.Add(New CodePrimitiveExpression(_GradientArea.Y))
                .Parameters.Add(New CodePrimitiveExpression(_GradientArea.Width))
                .Parameters.Add(New CodePrimitiveExpression(_GradientArea.Height))
            End With

            With fillInit
                .CreateType = New CodeTypeReference(GetType(System.Drawing.Drawing2D.LinearGradientBrush))
                .Parameters.Add(fillRectParam)
                .Parameters.Add(getColorAssignment(Me.Color1))
                .Parameters.Add(getColorAssignment(Me.Color2))

                If Me.GradientMode = EnumGradientMode.Custom Then
                    .Parameters.Add(New CodeDom.CodePrimitiveExpression(_GradientAngle))

                Else
                    Dim gradTypeParam As New CodeDom.CodeFieldReferenceExpression

                    With gradTypeParam
                        .TargetObject = New CodeTypeReferenceExpression(GetType(System.Drawing.Drawing2D.LinearGradientMode))
                        .FieldName = System.Enum.GetName(GetType(System.Drawing.Drawing2D.LinearGradientMode), Me.GradientMode)
                    End With

                    .Parameters.Add(gradTypeParam)
                End If

            End With

            'Create the member variable for the stroke
            With fillDeclaration
                .Name = obj.ExportName & "Brush"
                .Type = New CodeTypeReference(GetType(System.Drawing.Drawing2D.LinearGradientBrush))
                .Attributes = MemberAttributes.Private
                .InitExpression = fillInit
            End With

            'create dispose call
            With fillDispose
                .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, fillDeclaration.Name)
                .Method.MethodName = "Dispose"
            End With

            'add object disposal
            With DisposeGDI
                .Statements.Add(fillDispose)
            End With

            With declarations
                .Add(fillDeclaration)
            End With

            Return fillDeclaration.Name
        End If
    End Function





    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a consolidated gradient fill to code.
    ''' </summary>
    ''' <param name="sSharedName">See base class</param>
    ''' <param name="declarations">See base class</param>
    ''' <param name="InitGraphics">See base class</param>
    ''' <param name="RenderGDI">See base class</param>
    ''' <param name="DisposeGDI">See base class</param>
    ''' <param name="ExportSettings">See base class</param>
    ''' <returns>See base class</returns>
    '''  -----------------------------------------------------------------------------
    Friend Overloads Overrides Function emit(ByVal sSharedName As String, _
  ByVal declarations As CodeDom.CodeTypeMemberCollection, _
  ByVal InitGraphics As CodeDom.CodeMemberMethod, _
  ByVal RenderGDI As CodeDom.CodeStatementCollection, _
  ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
  ByVal ExportSettings As ExportSettings) As String


        Trace.WriteLineIf(Session.Tracing, "GradientFill.emit")

        Dim fillDeclaration As New CodeMemberField
        Dim fillInit As New CodeObjectCreateExpression
        Dim fillDispose As New CodeDom.CodeMethodInvokeExpression

        Dim fillRectParam As New CodeDom.CodeObjectCreateExpression

        With fillRectParam
            .CreateType = New CodeTypeReference(GetType(System.Drawing.Rectangle))
            .Parameters.Add(New CodePrimitiveExpression(_GradientArea.X))
            .Parameters.Add(New CodePrimitiveExpression(_GradientArea.Y))
            .Parameters.Add(New CodePrimitiveExpression(_GradientArea.Width))
            .Parameters.Add(New CodePrimitiveExpression(_GradientArea.Height))
        End With

        With fillInit
            .CreateType = New CodeTypeReference(GetType(System.Drawing.Drawing2D.LinearGradientBrush))
            .Parameters.Add(fillRectParam)
            .Parameters.Add(getColorAssignment(Me.Color1))
            .Parameters.Add(getColorAssignment(Me.Color2))

            If Me.GradientMode = EnumGradientMode.Custom Then
                .Parameters.Add(New CodeDom.CodePrimitiveExpression(_GradientAngle))

            Else
                Dim gradTypeParam As New CodeDom.CodeFieldReferenceExpression

                With gradTypeParam
                    .TargetObject = New CodeTypeReferenceExpression(GetType(System.Drawing.Drawing2D.LinearGradientMode))
                    .FieldName = System.Enum.GetName(GetType(System.Drawing.Drawing2D.LinearGradientMode), Me.GradientMode)
                End With

                .Parameters.Add(gradTypeParam)
            End If

        End With

        'Create the member variable for the stroke
        With fillDeclaration
            .Name = sSharedName
            .Type = New CodeTypeReference(GetType(System.Drawing.Drawing2D.LinearGradientBrush))
            .Attributes = MemberAttributes.Private
            .InitExpression = fillInit
        End With

        'create dispose call
        With fillDispose
            .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, fillDeclaration.Name)
            .Method.MethodName = "Dispose"
        End With

        'add object disposal
        With DisposeGDI
            .Statements.Add(fillDispose)
        End With

        With declarations
            .Add(fillDeclaration)
        End With

        Return fillDeclaration.Name

    End Function


#End Region


#Region "Base Class Implementers"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes to the parent GDIFilledShape
    ''' </summary>
    ''' <param name="obj">The parent GDIObject</param>
    ''' -----------------------------------------------------------------------------
    Public Overrides Sub OnParentUpdated(ByVal obj As GDIFilledShape)
        If obj.TrackFill Then
            Me.Bounds = obj.Bounds
            UpdateFill()
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to requests from the base class to recreate a brush used on the surface.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub UpdateFill()
        Dim gradbrush As LinearGradientBrush

        If Not _Brush Is Nothing Then
            _Brush.Dispose()
        End If

        If _GradientMode = EnumGradientMode.Custom Then
            gradbrush = New LinearGradientBrush(_GradientArea, _Color1, _Color2, _GradientAngle)
        Else
            gradbrush = New LinearGradientBrush(_GradientArea, _Color1, _Color2, CType(_GradientMode, Drawing2D.LinearGradientMode))
        End If

        _Brush = gradbrush
        NotifyFillUpdated()
    End Sub
#End Region

#Region "Property Accessors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the angle of the gradient used in custom fill scenarios.
    ''' </summary>
    ''' <value>A single (float) representing the angle of the gradient fill.</value>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.Description("Angle of gradient (only applies if GradientMode is set to custom")> _
    Public Property Angle() As Single
        Get
            Return _GradientAngle
        End Get
        Set(ByVal Value As Single)
            _GradientAngle = Value
            _GradientMode = EnumGradientMode.Custom
            Trace.WriteLineIf(Session.Tracing, "GradientFill.Angle.Set: " & Value.ToString)
            UpdateFill()
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets a rectangle used to bound the filled gradient area.  
    ''' This will usually equal the parent rectangle unless trackfill has been disabled. 
    ''' </summary>
    ''' <value>The rectangle bounding the gradient fill.</value>
    ''' ----------------------------------------------------------------------------- 
    <ComponentModel.Description("Rectangle filled by the gradient.  Will be overwritten if trackfill is set to true")> _
    Public Property Bounds() As Rectangle
        Get
            Return _GradientArea
        End Get
        Set(ByVal Value As Rectangle)
            Trace.WriteLineIf(Session.Tracing, "GradientFill.Bounds.Set: " & Value.ToString)

            _GradientArea = New Rectangle(Value.Location, Value.Size)
            UpdateFill()
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the gradient mode used to populate the gradient.  
    ''' </summary>
    ''' <value>an EnumGradientMode, which is equiv. to the System.Drawing GradientModes with 
    ''' the exception of a "Custom" setting used when a custom gradient angle is being used </value>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.Description("Gradient mode.  Set to custom to use the angle property")> _
Public Property GradientMode() As EnumGradientMode
        Get
            Return _GradientMode
        End Get
        Set(ByVal Value As EnumGradientMode)
            Trace.WriteLineIf(Session.Tracing, "GradientFill.Mode.Set: " & Value.ToString)

            _GradientMode = Value
            UpdateFill()
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the first color used in a gradient fill.
    ''' </summary>
    ''' <value>A System.Drawing Color</value>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.Description("Left most / top most color in the gradient")> _
Public Property Color1() As Color
        Get
            Return _Color1
        End Get
        Set(ByVal Value As Color)
            Trace.WriteLineIf(Session.Tracing, "GradientFill.Color1.Set: " & Value.ToString)

            _Color1 = Value
            UpdateFill()
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the second color used in a gradient fill.
    ''' </summary>
    ''' <value>A System.Drawing Color</value>
    '''  -----------------------------------------------------------------------------
    <ComponentModel.Description("Right most / bottom most most color in the gradient")> _
Public Property Color2() As Color
        Get
            Return _Color2
        End Get
        Set(ByVal Value As Color)
            Trace.WriteLineIf(Session.Tracing, "GradientFill.Color2.Set: " & Value.ToString)

            _Color2 = Value
            UpdateFill()
        End Set
    End Property

#End Region


#Region "Misc Methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a string representation of this object.
    ''' </summary>
    ''' <returns>A string representation of the object</returns>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function ToString() As String
        Return "Gradient Fill"
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if two gradients fills are equal.
    ''' </summary>
    ''' <param name="fill1">The first gradient fill to examine</param>
    ''' <param name="fill2">The second gradient fill to examine</param>
    ''' <returns>A Boolean indicating if the two fills are equivalent.</returns>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function op_Equality(ByVal fill1 As GDIGradientFill, ByVal fill2 As GDIGradientFill) As Boolean
        Dim bEquality As Boolean = True
        bEquality = bEquality And Color.op_Equality(fill1.Color1, fill2.Color1) And Color.op_Equality(fill1.Color2, fill2.Color2)
        If bEquality = False Then
            Return bEquality
        End If
        bEquality = bEquality And (fill1._GradientMode = fill2._GradientMode)
        bEquality = bEquality And (fill1._GradientAngle = fill2._GradientAngle)

        Return bEquality
    End Function
#End Region
End Class
