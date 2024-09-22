Imports System.CodeDom
Imports System.ComponentModel

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIRect
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Renders a rectangle to the drawing surface.  Contains properties and methods related 
''' to drawing rectangles, as well as drawing rounded rectangles depending on the rounded 
''' property.
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable(), DefaultProperty("Name")> _
Public Class GDIRect
    Inherits GDIFilledShape
   

#Region "Non serialized members"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Next valid integer suffix for a GDIRect object (Rectangle1, Rectangle2, etc)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _NextnameID As Int32 = 0
#End Region

#Region "Local Fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Percent the rectangle's corners are rounded by as a float (single).
    ''' </summary>
    ''' <remarks>If the rectangle is rounded at all 
    ''' (a value other than 0), both drawing and code emission are drastically 
    ''' different.  Instead of using a rectangle, a graphics path is used with 
    ''' rounded corners.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private _CornerRoundess As Single
#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new instance of a GDIRect
    ''' </summary>
    ''' <param name="rect">The rectangle to use to as the bounds of the GDIRect.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal rect As Rectangle)
        MyBase.New()
        Trace.WriteLineIf(Session.Tracing, "Rectangle.New")
        Trace.WriteLineIf(Session.Tracing, "rect:" & rect.ToString)

        Me.Bounds = rect
    End Sub
#End Region

#Region "Code generation methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a GDIRect to code when the rectangle is rounded (where _CornerRoundess is not 0)
    ''' </summary>
    ''' <param name="declarations">Declarations section of the outgoing class.</param>
    ''' <param name="InitGraphics">The GDI+ Architect InitGraphics method.  
    '''                            used to setup drawing</param>
    ''' <param name="RenderGDI">The GDI+ Architect render graphics method. Where actual drawing 
    ''' takes place.</param>
    ''' <param name="DisposeGDI">The GDI+ Architect disposal method to dispose of 
    ''' graphics resources that have .Dispose methods.</param>
    ''' <param name="ExportSettings">The current export settings of the parent document.</param>
    ''' <param name="Consolidated">The set of consolidated objects thus far.  
    ''' Consolidated objects are objects that are identical and can optionally be declared 
    ''' once instead of multiple times.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub emitRoundedRect(ByVal declarations As CodeDom.CodeTypeMemberCollection, _
    ByVal InitGraphics As CodeDom.CodeMemberMethod, _
    ByVal RenderGDI As CodeDom.CodeStatementCollection, _
    ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
    ByVal ExportSettings As ExportSettings, _
    ByVal Consolidated As ExportConsolidate)

        Trace.WriteLineIf(Session.Tracing, "Rectangle.emitRoundedRect")

        Dim pathDeclaration As New CodeMemberField
        Dim pathInit As New CodeObjectCreateExpression
        Dim pathDispose As New CodeDom.CodeMethodInvokeExpression

        Dim invokeStroke As New CodeMethodInvokeExpression
        Dim invokeFill As New CodeMethodInvokeExpression

        Dim pointDeclaration As New CodeArrayCreateExpression(GetType(PointF))
        Dim typeDeclaration As New CodeArrayCreateExpression(GetType(System.Byte))


        'Sets up an array of bytes used to initialize paths
        For Each bt As Byte In _Path.PathTypes
            typeDeclaration.Initializers.Add(New CodeDom.CodePrimitiveExpression(bt))
        Next

        'Sets up an array of points used to initialize paths

        For Each pt As PointF In _Path.PathPoints
            Dim pointInit As CodeObjectCreateExpression

            pointInit = New CodeObjectCreateExpression

            With pointInit
                .CreateType = New CodeTypeReference(GetType(PointF))
                .Parameters.Add(New CodePrimitiveExpression(pt.X))
                .Parameters.Add(New CodePrimitiveExpression(pt.Y))
            End With

            pointDeclaration.Initializers.Add(pointInit)
        Next

        'Create the graphics path intializer
        With pathInit
            .CreateType = New CodeTypeReference(GetType(Drawing2D.GraphicsPath))
            .Parameters.Add(pointDeclaration)
            .Parameters.Add(typeDeclaration)
        End With


        'Setup the path declaration
        With pathDeclaration
            .InitExpression = pathInit
            .Name = MyBase.ExportName()
            .Attributes = MyBase.getScope
            .Type = New CodeTypeReference(GetType(System.Drawing.Drawing2D.GraphicsPath))
        End With

        'Create dispose path method invoke expression
        With pathDispose
            .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, pathDeclaration.Name)
            .Method.MethodName = "Dispose"
        End With

        'Add the path declaration to declarations
        declarations.Add(pathDeclaration)

        'Add the path disposal code to disposeGDI
        DisposeGDI.Statements.Add(pathDispose)


        If Rotation > 0 Then
            MyBase.emitRotationDeclaration(declarations)
            MyBase.emitInvokeBeginRotation(RenderGDI)
        End If

        'emit fill
        If _DrawFill Then
            Dim sFillName As String = Me.Fill.emit(Me, declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings, Consolidated)

            With invokeFill
                .Method.TargetObject = New CodeArgumentReferenceExpression("g")
                .Method.MethodName = "FillPath"
                .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sFillName))
                .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, pathDeclaration.Name))
            End With

            RenderGDI.Add(invokeFill)

        End If

        'emit stroke
        If _DrawStroke Then
            Dim sStrokeName As String = Me.Stroke.emit(Me, declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings, Consolidated)

            With invokeStroke
                .Method.TargetObject = New CodeArgumentReferenceExpression("g")
                .Method.MethodName = "DrawPath"
                .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sStrokeName))
                .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, pathDeclaration.Name))
            End With

            RenderGDI.Add(invokeStroke)
        End If

        If Rotation > 0 Then
            MyBase.emitInvokeEndRotation(RenderGDI)
        End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a GDIRect to code when there is no rounding (_CornerRoundess = 0)
    ''' </summary>
    ''' <param name="declarations">Declarations section of the outgoing class.</param>
    ''' <param name="InitGraphics">The GDI+ Architect InitGraphics method.  
    '''                            used to setup drawing</param>
    ''' <param name="RenderGDI">The GDI+ Architect render graphics method. Where actual drawing 
    ''' takes place.</param>
    ''' <param name="DisposeGDI">The GDI+ Architect disposal method to dispose of 
    ''' graphics resources that have .Dispose methods.</param>
    ''' <param name="ExportSettings">The current export settings of the parent document.</param>
    ''' <param name="Consolidated">The set of consolidated objects thus far.  
    ''' Consolidated objects are objects that are identical and can optionally be declared 
    ''' once instead of multiple times.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub emitRect(ByVal declarations As CodeDom.CodeTypeMemberCollection, _
    ByVal InitGraphics As CodeDom.CodeMemberMethod, _
    ByVal RenderGDI As CodeDom.CodeStatementCollection, _
    ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
    ByVal ExportSettings As ExportSettings, _
    ByVal Consolidated As ExportConsolidate)

        Trace.WriteLineIf(Session.Tracing, "Rectangle.emitRect")

        Dim rectDeclaration As New CodeMemberField
        Dim rectInit As New CodeObjectCreateExpression

        Dim invokeStroke As New CodeMethodInvokeExpression
        Dim invokeFill As New CodeMethodInvokeExpression




        'Create Rectangle Initializer
        With rectInit
            .CreateType = New CodeTypeReference(GetType(System.Drawing.Rectangle))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.X))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.Y))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.Width))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.Height))
        End With

        'Create Rectangle declaration
        With rectDeclaration
            .InitExpression = rectInit
            .Name = MyBase.ExportName()
            .Attributes = MyBase.getScope
            .Type = New CodeTypeReference(GetType(System.Drawing.Rectangle))
        End With

        If Rotation > 0 Then
            MyBase.emitRotationDeclaration(declarations)
            MyBase.emitInvokeBeginRotation(RenderGDI)
        End If

        'Create Fill, if applicable 
        If _DrawFill Then
            Dim sFillname As String = Me.Fill.emit(Me, declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings, Consolidated)

            With invokeFill
                .Method.TargetObject = New CodeArgumentReferenceExpression("g")
                .Method.MethodName = "FillRectangle"
                .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sFillname))
                .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, rectDeclaration.Name))
            End With


            RenderGDI.Add(invokeFill)
        End If

        'Create Stroke, if applicable 
        If _DrawStroke Then
            Dim sStrokeName As String = Me.Stroke.emit(Me, declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings, Consolidated)

            With invokeStroke
                .Method.TargetObject = New CodeArgumentReferenceExpression("g")
                .Method.MethodName = "DrawRectangle"
                .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sStrokeName))
                .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, rectDeclaration.Name))
            End With

            RenderGDI.Add(invokeStroke)
        End If

        If Rotation > 0 Then
            MyBase.emitInvokeEndRotation(RenderGDI)
        End If

        declarations.Add(rectDeclaration)


    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a GDIRect to code.
    ''' </summary>
    ''' <param name="declarations">Declarations section of the outgoing class.</param>
    ''' <param name="InitGraphics">The GDI+ Architect InitGraphics method.  
    '''                            used to setup drawing</param>
    ''' <param name="RenderGDI">The GDI+ Architect render graphics method. Where actual drawing 
    ''' takes place.</param>
    ''' <param name="DisposeGDI">The GDI+ Architect disposal method to dispose of 
    ''' graphics resources that have .Dispose methods.</param>
    ''' <param name="ExportSettings">The current export settings of the parent document.</param>
    ''' <param name="Consolidated">The set of consolidated objects thus far.  
    ''' Consolidated objects are objects that are identical and can optionally be declared 
    ''' once instead of multiple times.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overloads Overrides Sub emit( _
    ByVal declarations As CodeDom.CodeTypeMemberCollection, _
    ByVal InitGraphics As CodeDom.CodeMemberMethod, _
    ByVal RenderGDI As CodeDom.CodeStatementCollection, _
    ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
    ByVal ExportSettings As ExportSettings, _
    ByVal Consolidated As ExportConsolidate)

        Trace.WriteLineIf(Session.Tracing, "Rectangle.emit")
        If _CornerRoundess = 0 Then
            emitRect(declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings, Consolidated)
        Else
            emitRoundedRect(declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings, Consolidated)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts a GDIRect to SVG XML and appends it to the outgoing SVG code.
    ''' </summary>
    ''' <param name="xmlDoc">The SVG document to append the rect to.</param>
    ''' <param name="defs">The definitions section of the SVG document.</param>
    ''' <param name="group">The group to append the XML spec of the rect to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub toXML(ByVal xmlDoc As Xml.XmlDocument, ByVal defs As Xml.XmlNode, ByVal group As Xml.XmlNode)
        Dim attr As Xml.XmlAttribute

        Dim node As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "rect", String.Empty)

        attr = xmlDoc.CreateAttribute("id")
        attr.Value = Me.Name
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("x")
        attr.Value = String.Format("{0}", Me.Bounds.X)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("y")
        attr.Value = String.Format("{0}", Me.Bounds.Y)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("width")
        attr.Value = String.Format("{0}", Me.Bounds.Width)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("height")
        attr.Value = String.Format("{0}", Me.Bounds.Height)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("shape-rendering")
        attr.Value = "crispEdges"
        node.Attributes.Append(attr)

        If Me.Roundness > 0 Then
            Dim offsetX As Single = (_Bounds.Right - _Bounds.Left) * _CornerRoundess / 100
            Dim offsety As Single = (_Bounds.Bottom - _Bounds.Top) * _CornerRoundess / 100


            attr = xmlDoc.CreateAttribute("rx")
            attr.Value = offsetX.ToString
            node.Attributes.Append(attr)

            attr = xmlDoc.CreateAttribute("rx")
            attr.Value = offsety.ToString
            node.Attributes.Append(attr)
        End If
        If Me.Rotation > 0 Then
            attr = xmlDoc.CreateAttribute("transform")
            attr.Value = "rotate(" & Me.Rotation & " " & RotationPoint.X & " " & Me.RotationPoint.Y & ")"
            node.Attributes.Append(attr)
        End If

        If Me.DrawFill Then
            Me.Fill.toXML(xmlDoc, defs, node)
        Else
            attr = xmlDoc.CreateAttribute("fill")
            attr.Value = String.Format("none")
            node.Attributes.Append(attr)
        End If

        If Me.DrawStroke Then
            Me.Stroke.toXML(xmlDoc, node)
        End If

        group.AppendChild(node)
    End Sub
#End Region

#Region "Base Class Implementers"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a name used when displaying the class in the property browser.
    ''' </summary>
    ''' <value>A string containing the word "Rectangle"</value>
    ''' -----------------------------------------------------------------------------   
    Public Overrides ReadOnly Property ClassName() As String
        Get
            Return "Rectangle"
        End Get
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the next valid name for a new GDIRect
    ''' </summary>
    ''' <returns>A string containing the next valid name.</returns>
    ''' 
    ''' -----------------------------------------------------------------------------
    Protected Overrides Function NextName() As String
        _NextnameID += 1
        Return "Rect" & _NextnameID
    End Function

#End Region


#Region "Property Accessors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The amount to round the corner edges of the rectangle by
    ''' </summary>
    ''' <value>A Single (float) indicating the extent of rounding.</value>
    ''' <remarks>Rounding must be between 0 and 100.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    <DefaultValue(0), Description("Specifies corner roundness from 0 - 100.  Adding any rounding causes the rectangle to export as a Drawing2D.GraphicsPath")> _
    Public Property Roundness() As Single
        Get
            Return _CornerRoundess
        End Get
        Set(ByVal Value As Single)

            Trace.WriteLineIf(Session.Tracing, "Rectangle.Roundness.Set")

            If Value <= 100 AndAlso Value >= 0 Then
                _CornerRoundess = Value
            End If
        End Set
    End Property

#End Region

#Region "Drawing related methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Recreates the path used to draw the GDIRect to a surface.
    ''' </summary>
    ''' <remarks>Notice that if rounding is not 0, a more complicated series of events 
    ''' take place here in order to create the rectangle's rounded edges.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub createPath()
        MyBase.resetPath()

        If _CornerRoundess = 0 Then
            _Path.AddRectangle(_Bounds)
        Else

            Dim offsetX As Single = (_Bounds.Right - _Bounds.Left) * _CornerRoundess / 100
            Dim offsety As Single = (_Bounds.Bottom - _Bounds.Top) * _CornerRoundess / 100

            With _Bounds
                _Path.AddArc(.Right - offsetX, .Top, offsetX, offsety, 270, 90)
                _Path.AddArc(.Right - offsetX, .Bottom - offsety, offsetX, offsety, 0, 90)
                _Path.AddArc(.Left, .Bottom - offsety, offsetX, offsety, 90, 90)
                _Path.AddArc(.Left, .Top, offsetX, offsety, 180, 90)
                _Path.CloseFigure()
            End With


        End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders a rectangle to a specific graphics surface with a specific draw mode.
    ''' </summary>
    ''' <param name="g">The graphics context to render the GDIRect to.</param>
    ''' <param name="eDrawMode">The current drawing mode (See EnumDrawMode for more information).</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub Draw(ByVal g As System.Drawing.Graphics, ByVal eDrawMode As EnumDrawMode)
        MyBase.BeginDraw(g)


        If _ResetPath Then
            createPath()
        End If

        If _DrawFill Then
            g.FillPath(Fill.Brush(), _Path)
        End If

        If _DrawStroke Then
            g.DrawPath(Stroke.Pen, _Path)
        End If

        MyBase.EndDraw(g)

    End Sub

#End Region

End Class
