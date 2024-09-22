Imports System.CodeDom

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIHatchFill
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A Hatch fill for a fillable GDIObject.  Corresponds to a Drawing2D.HatchStyle 
''' enumeration set.  Contains a start and end color as well as a hatch setting.
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public Class GDIHatchFill
    Inherits GDIFill


#Region "Local Fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The back color to use in the hatch. 
    ''' </summary>
    ''' <remarks>Hatches are composed of both a fore and a back color.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private _HatchBackColor As Color = Color.White

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The fore color to use in the hatch.
    ''' </summary>
    ''' <remarks>Hatches are composed of both a fore and a back color.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private _HatchForeColor As Color = Color.Black

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The hatch style to use in the hatch.  Corresponds to a Drawing2D.HatchStyle.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _HatchStyle As Drawing2D.HatchStyle

#End Region

#Region "Constructors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a GDIHatchFill given a hatch style, back color and fore color.
    ''' </summary>
    ''' <param name="parent">The parent node to populate.</param>
    ''' <param name="forecolor">The forecolor of the hatch pattern</param>
    ''' <param name="backcolor">The backcolor of the hatch pattern</param>
    ''' <param name="style">The hatch style of the hatch pattern</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal forecolor As Color, ByVal backcolor As Color, ByVal style As Drawing2D.HatchStyle)
        MyBase.New()
        Trace.WriteLineIf(Session.Tracing, "HatchFill.New.General")

        _HatchForeColor = forecolor
        _HatchBackColor = backcolor
        _HatchStyle = style
        UpdateFill()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new GDIHatchFill given another GDIHatchFill to copy.
    ''' </summary>
    ''' <param name="parent">The parent shape this fill is being used with.</param>
    ''' <param name="fill">The GDIHatch fill to copy.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal fill As GDIHatchFill)
        MyBase.New()
        Trace.WriteLineIf(Session.Tracing, "HatchFill.New.Clone")

        _HatchBackColor = fill.HatchBackColor
        _HatchStyle = fill.HatchStyle
        _HatchForeColor = fill.HatchForeColor

    End Sub

#End Region

#Region "Code Generation related methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a consolidated hatch fill to code.
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

        Trace.WriteLineIf(Session.Tracing, "HatchFill.emit")


        Dim fillDeclaration As New CodeMemberField
        Dim fillInit As New CodeObjectCreateExpression
        Dim fillDispose As New CodeDom.CodeMethodInvokeExpression

        Dim fillHatchStyleParam As New CodeDom.CodeFieldReferenceExpression

        With fillHatchStyleParam
            .TargetObject = New CodeTypeReferenceExpression(GetType(Drawing2D.HatchStyle))
            .FieldName = System.Enum.GetName(GetType(Drawing2D.HatchStyle), Me.HatchStyle)
        End With

        'Create the initialization args and assignment  for stroke
        With fillInit
            .CreateType = New CodeTypeReference(GetType(System.Drawing.Drawing2D.HatchBrush))
            .Parameters.Add(fillHatchStyleParam)
            .Parameters.Add(getColorAssignment(Me.HatchForeColor))
            .Parameters.Add(getColorAssignment(Me.HatchBackColor))
        End With

        'Create the member variable for the stroke
        With fillDeclaration
            .Name = sSharedName
            .Type = New CodeTypeReference(GetType(System.Drawing.Drawing2D.HatchBrush))
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

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a hatch fill to code.
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

        Trace.WriteLineIf(Session.Tracing, "HatchFill.emit.Consolidated")

        If Consolidated.hasFillMatch(Me) And Me.Consolidate Then
            Return Consolidated.getFillName(Me)
        Else

            Dim fillDeclaration As New CodeMemberField
            Dim fillInit As New CodeObjectCreateExpression
            Dim fillDispose As New CodeDom.CodeMethodInvokeExpression

            Dim fillHatchStyleParam As New CodeDom.CodeFieldReferenceExpression

            With fillHatchStyleParam
                .TargetObject = New CodeTypeReferenceExpression(GetType(Drawing2D.HatchStyle))
                .FieldName = System.Enum.GetName(GetType(Drawing2D.HatchStyle), Me.HatchStyle)
            End With

            'Create the initialization args and assignment  for stroke
            With fillInit
                .CreateType = New CodeTypeReference(GetType(System.Drawing.Drawing2D.HatchBrush))
                .Parameters.Add(fillHatchStyleParam)
                .Parameters.Add(getColorAssignment(Me.HatchForeColor))
                .Parameters.Add(getColorAssignment(Me.HatchBackColor))
            End With

            'Create the member variable for the stroke
            With fillDeclaration
                .Name = obj.ExportName & "Brush"
                .Type = New CodeTypeReference(GetType(System.Drawing.Drawing2D.HatchBrush))
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
    ''' Converts the hatch fill to an XML source for SVG display.
    ''' </summary>
    ''' <param name="xmlDoc">See base class</param>
    ''' <param name="defs">See base class</param>
    ''' <param name="parent">See base class</param>
    ''' -----------------------------------------------------------------------------
    Public Overrides Sub toXML(ByVal xmlDoc As Xml.XmlDocument, ByVal defs As Xml.XmlNode, ByVal parent As Xml.XmlNode)

        Dim attr As Xml.XmlAttribute


        Dim nodeHatch As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "pattern", String.Empty)

        attr = xmlDoc.CreateAttribute("id")
        attr.Value = parent.Attributes("id").Value & "Brush"
        nodeHatch.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("x")
        attr.Value = "0"
        nodeHatch.Attributes.Append(attr)


        attr = xmlDoc.CreateAttribute("y")
        attr.Value = "0"
        nodeHatch.Attributes.Append(attr)


        attr = xmlDoc.CreateAttribute("width")
        attr.Value = "8"
        nodeHatch.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("height")
        attr.Value = "8"
        nodeHatch.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("patternUnits")
        attr.Value = "userSpaceOnUse"
        nodeHatch.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("shape-rendering")
        attr.Value = "crispEdges"
        nodeHatch.Attributes.Append(attr)


        buildStyle(xmlDoc, defs, parent, nodeHatch)

        attr = xmlDoc.CreateAttribute("fill")
        attr.Value = String.Format("url(#{0})", parent.Attributes("id").Value & "Brush")
        parent.Attributes.Append(attr)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Used in the SVG code generation to create a polygon pattern
    ''' </summary>
    ''' <param name="xmldoc">Parent SVG document</param>
    ''' <param name="parentnode">Node that this pattern should append to </param>
    ''' <param name="pts">The point set used in creating the pattern</param>
    ''' -----------------------------------------------------------------------------
    Private Sub AppendPatternPolygon(ByVal xmldoc As Xml.XmlDocument, ByVal parentnode As Xml.XmlNode, ByVal pts As String)
        ' <polygon style="stroke:#000000; stroke-width:2; fill:none;"
        ' points="230,200 230,50 330,50 230,200" />
        Dim attr As Xml.XmlAttribute

        Dim node As Xml.XmlNode = xmldoc.CreateNode(Xml.XmlNodeType.Element, "polygon", String.Empty)

        attr = xmldoc.CreateAttribute("points")
        attr.Value = String.Format("{0}", pts)
        node.Attributes.Append(attr)
        appendPatternFillColor(xmldoc, node)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Appends a fill color to an SVG pattern.
    ''' </summary>
    ''' <param name="xmlDoc">The SVG document to append the pattern to</param>
    ''' <param name="node">The node to append to.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub appendPatternFillColor(ByVal xmlDoc As Xml.XmlDocument, ByVal node As Xml.XmlNode)
        Dim arg As String
        Dim attr As Xml.XmlAttribute

        attr = xmlDoc.CreateAttribute("fill")

        If Me.HatchForeColor.IsNamedColor Then
            attr.Value = Me.HatchForeColor.Name
        Else
            attr.Value = String.Format("rgb({0},{1},{2})", Me.HatchForeColor.R, Me.HatchForeColor.G, Me.HatchForeColor.B)
        End If

        node.Attributes.Append(attr)

        If Me.HatchForeColor.A < 255 Then
            attr = xmlDoc.CreateAttribute("opacity")
            arg = FormatNumber(Me.HatchForeColor.A / 255, 3)
            attr.Value = arg
            node.Attributes.Append(attr)
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Appends a line pattern to the SVG document
    ''' </summary>
    ''' <param name="xmlDoc">The document to append to</param>
    ''' <param name="node">The node to append to.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub appendPatternLineColor(ByVal xmlDoc As Xml.XmlDocument, ByVal node As Xml.XmlNode)
        Dim arg As String
        Dim attr As Xml.XmlAttribute

        attr = xmlDoc.CreateAttribute("stroke")

        If Me.HatchForeColor.IsNamedColor Then
            attr.Value = Me.HatchForeColor.Name
        Else
            attr.Value = String.Format("rgb({0},{1},{2})", Me.HatchForeColor.R, Me.HatchForeColor.G, Me.HatchForeColor.B)
        End If

        node.Attributes.Append(attr)

        If Me.HatchForeColor.A < 255 Then
            attr = xmlDoc.CreateAttribute("stroke-opacity")
            arg = FormatNumber(Me.HatchForeColor.A / 255, 3)
            attr.Value = arg
            node.Attributes.Append(attr)
        End If



    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Appends an ellipse pattern to an SVG document.
    ''' </summary>
    ''' <param name="xmlDoc">The document to append to.</param>
    ''' <param name="parentNode">The node to append to</param>
    ''' <param name="cx">The x position of the pattern.</param>
    ''' <param name="cy">The y position of the pattern</param>
    ''' <param name="rx">The radius x of the pattern</param>
    ''' <param name="ry">The radius y of the pattern</param>
    ''' -----------------------------------------------------------------------------
    Private Sub AppendPatternEllipse(ByVal xmlDoc As Xml.XmlDocument, ByVal parentNode As Xml.XmlNode, ByVal cx As Int32, ByVal cy As Int32, ByVal rx As Int32, ByVal ry As Int32)
        Dim attr As Xml.XmlAttribute

        Dim node As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "ellipse", String.Empty)

        attr = xmlDoc.CreateAttribute("cx")
        attr.Value = String.Format("{0}", cx)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("cy")
        attr.Value = String.Format("{0}", cy)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("rx")
        attr.Value = String.Format("{0}", rx)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("ry")
        attr.Value = String.Format("{0}", ry)
        node.Attributes.Append(attr)


        parentNode.AppendChild(node)

        appendPatternFillColor(xmlDoc, node)
        appendPatternLineColor(xmlDoc, node)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Appends a line pattern to an SVG document that has dash settings.
    ''' </summary>
    ''' <param name="xmlDoc">The SVG document to append to</param>
    ''' <param name="parentNode">The node to append to</param>
    ''' <param name="x1">The first point's x position in the pattern</param>
    ''' <param name="y1">The first point's y position in the pattern</param>
    ''' <param name="x2">The second point's x position in the pattern</param>
    ''' <param name="y2">The second point's y position in the pattern</param>
    ''' 
    ''' <param name="dashSet">Dash settings for the line pattern</param>
    ''' -----------------------------------------------------------------------------
    Private Sub AppendPatternLine(ByVal xmlDoc As Xml.XmlDocument, ByVal parentNode As Xml.XmlNode, ByVal x1 As Int32, ByVal y1 As Int32, ByVal x2 As Int32, ByVal y2 As Int32, ByVal dashSet As String)
        Dim attr As Xml.XmlAttribute

        Dim node As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "line", String.Empty)
        '   attr = xmlDoc.CreateAttribute("id")
        '  node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("x1")
        attr.Value = String.Format("{0}", x1)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("y1")
        attr.Value = String.Format("{0}", y1)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("x2")
        attr.Value = String.Format("{0}", x2)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("y2")
        attr.Value = String.Format("{0}", y2)
        node.Attributes.Append(attr)


        parentNode.AppendChild(node)
        appendPatternLineColor(xmlDoc, node)

        attr = xmlDoc.CreateAttribute("stroke-dasharray")
        attr.Value = dashSet
        node.Attributes.Append(attr)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Appends a line pattern to an SVG document
    ''' </summary>
    ''' <param name="xmlDoc">The parent SVG document</param>
    ''' <param name="parentNode">The parent node</param>
    ''' <param name="x1">The first point's x position in the pattern</param>
    ''' <param name="y1">The first point's y position in the pattern</param>
    ''' <param name="x2">The second point's x position in the pattern</param>
    ''' <param name="y2">The second point's y position in the pattern</param>
    ''' -----------------------------------------------------------------------------
    Private Sub AppendPatternLine(ByVal xmlDoc As Xml.XmlDocument, ByVal parentNode As Xml.XmlNode, ByVal x1 As Int32, ByVal y1 As Int32, ByVal x2 As Int32, ByVal y2 As Int32)
        Dim attr As Xml.XmlAttribute

        Dim node As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "line", String.Empty)
        '   attr = xmlDoc.CreateAttribute("id")
        '  node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("x1")
        attr.Value = String.Format("{0}", x1)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("y1")
        attr.Value = String.Format("{0}", y1)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("x2")
        attr.Value = String.Format("{0}", x2)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("y2")
        attr.Value = String.Format("{0}", y2)
        node.Attributes.Append(attr)



        parentNode.AppendChild(node)
        appendPatternLineColor(xmlDoc, node)

        attr = xmlDoc.CreateAttribute("fill")
        attr.Value = "none"
        node.Attributes.Append(attr)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Appends a back ground rectangle to a pattern.  Used when generating SVG code.
    ''' </summary>
    ''' <param name="xmlDoc">The parent SVG document</param>
    ''' <param name="parentNode">The parent node</param>
    ''' <param name="x">The x position of the top left corner of the rectangle</param>
    ''' <param name="y">The y position of the top left corner of the rectangle</param>
    ''' <param name="width">The width of the rectangle</param>
    ''' <param name="height">The height of the rectangle</param>
    ''' -----------------------------------------------------------------------------
    Private Sub AppendBackGroundRect(ByVal xmlDoc As Xml.XmlDocument, ByVal parentNode As Xml.XmlNode, ByVal x As Int32, ByVal y As Int32, ByVal width As Int32, ByVal height As Int32)
        Dim attr As Xml.XmlAttribute

        Dim node As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "rect", String.Empty)
        '   attr = xmlDoc.CreateAttribute("id")
        '  node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("x")
        attr.Value = String.Format("{0}", x)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("y")
        attr.Value = String.Format("{0}", y)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("width")
        attr.Value = String.Format("{0}", width)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("height")
        attr.Value = String.Format("{0}", height)
        node.Attributes.Append(attr)

        parentNode.AppendChild(node)


        Dim arg As String

        attr = xmlDoc.CreateAttribute("fill")

        If Me.HatchBackColor.IsNamedColor Then
            attr.Value = Me.HatchBackColor.Name
        Else
            attr.Value = String.Format("rgb({0},{1},{2})", Me.HatchBackColor.R, Me.HatchBackColor.G, Me.HatchBackColor.B)
        End If

        node.Attributes.Append(attr)

        If Me.HatchBackColor.A < 255 Then
            attr = xmlDoc.CreateAttribute("opacity")
            arg = FormatNumber(Me.HatchBackColor.A / 255, 3)
            attr.Value = arg
            node.Attributes.Append(attr)
        End If


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Appends a rectangle pattern to a pattern (foreground).  Used in SVG document 
    ''' generation.
    ''' </summary>
    ''' <param name="xmlDoc">The parent SVG document</param>
    ''' <param name="parentNode">The parent node</param>
    ''' <param name="x">The x position of the top left corner of the rectangle</param>
    ''' <param name="y">The y position of the top left corner of the rectangle</param>
    ''' <param name="width">The width of the rectangle</param>
    ''' <param name="height">The height of the rectangle</param>
    ''' -----------------------------------------------------------------------------
    Private Sub AppendPatternRect(ByVal xmlDoc As Xml.XmlDocument, ByVal parentNode As Xml.XmlNode, ByVal x As Int32, ByVal y As Int32, ByVal width As Int32, ByVal height As Int32)
        Dim attr As Xml.XmlAttribute

        Dim node As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "rect", String.Empty)
        '   attr = xmlDoc.CreateAttribute("id")
        '  node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("x")
        attr.Value = String.Format("{0}", x)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("y")
        attr.Value = String.Format("{0}", y)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("width")
        attr.Value = String.Format("{0}", width)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("height")
        attr.Value = String.Format("{0}", height)
        node.Attributes.Append(attr)

        parentNode.AppendChild(node)
        appendPatternFillColor(xmlDoc, node)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Builds the hatch style pattern as an SVG XML equivalent (as close as possible)
    ''' </summary>
    ''' <param name="xmlDoc">Parent SVG document</param>
    ''' <param name="defs">The definitions node (definitions section) of the SVG document.</param>
    ''' <param name="parent">The parent node to append the hatch to.</param>
    ''' <param name="nodeHatch">The hatch node where the pattern is being used.</param>
    ''' <returns>An XMLNode that contains everything necessary to render the hatch style in 
    ''' SVG.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function buildStyle(ByVal xmlDoc As Xml.XmlDocument, ByVal defs As Xml.XmlNode, ByVal parent As Xml.XmlNode, ByVal nodeHatch As Xml.XmlNode) As Xml.XmlNode
        Dim attr As Xml.XmlAttribute

        defs.AppendChild(nodeHatch)



        AppendBackGroundRect(xmlDoc, nodeHatch, 0, 0, 8, 8)

        Select Case Me.HatchStyle
            Case HatchStyle.Cross

                AppendPatternLine(xmlDoc, nodeHatch, 4, 0, 4, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 4, 8, 4)

            Case HatchStyle.BackwardDiagonal
                AppendPatternLine(xmlDoc, nodeHatch, 8, 0, 0, 8)

            Case HatchStyle.LightDownwardDiagonal
                AppendPatternLine(xmlDoc, nodeHatch, 4, 0, 8, 4)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 4, 4, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 8, 8)

            Case HatchStyle.DarkDownwardDiagonal
                AppendPatternLine(xmlDoc, nodeHatch, 4, 0, 8, 4)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 4, 4, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 8, 8)

            Case HatchStyle.LightHorizontal
                AppendPatternLine(xmlDoc, nodeHatch, 0, 2, 8, 2)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 6, 8, 6)

            Case HatchStyle.DarkHorizontal
                AppendPatternLine(xmlDoc, nodeHatch, 0, 2, 8, 2)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 6, 8, 6)


            Case HatchStyle.LightUpwardDiagonal
                AppendPatternLine(xmlDoc, nodeHatch, 0, 4, 4, 0)
                AppendPatternLine(xmlDoc, nodeHatch, 4, 8, 8, 4)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 8, 8, 0)

            Case HatchStyle.DarkUpwardDiagonal
                AppendPatternLine(xmlDoc, nodeHatch, 0, 4, 4, 0)
                AppendPatternLine(xmlDoc, nodeHatch, 4, 8, 8, 4)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 8, 8, 0)


            Case HatchStyle.LightVertical
                AppendPatternLine(xmlDoc, nodeHatch, 2, 0, 2, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 6, 0, 6, 8)

            Case HatchStyle.DarkVertical
                AppendPatternLine(xmlDoc, nodeHatch, 2, 0, 2, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 6, 0, 6, 8)


            Case HatchStyle.DashedDownwardDiagonal
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 4, 4)
                AppendPatternLine(xmlDoc, nodeHatch, 4, 0, 8, 4)

            Case HatchStyle.DashedHorizontal
                AppendPatternLine(xmlDoc, nodeHatch, 0, 2, 4, 2)
                AppendPatternLine(xmlDoc, nodeHatch, 4, 6, 8, 6)

            Case HatchStyle.DashedUpwardDiagonal
                AppendPatternLine(xmlDoc, nodeHatch, 4, 0, 0, 4)
                AppendPatternLine(xmlDoc, nodeHatch, 8, 0, 4, 4)

            Case HatchStyle.DashedVertical
                AppendPatternLine(xmlDoc, nodeHatch, 2, 0, 2, 4)
                AppendPatternLine(xmlDoc, nodeHatch, 6, 4, 6, 8)

            Case HatchStyle.DiagonalBrick
                AppendPatternLine(xmlDoc, nodeHatch, 0, 8, 8, 0)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 4, 4)

            Case HatchStyle.DiagonalCross
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 8, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 8, 0, 0, 8)

            Case HatchStyle.Divot
                AppendPatternLine(xmlDoc, nodeHatch, 2, 2, 4, 4)
                AppendPatternLine(xmlDoc, nodeHatch, 4, 4, 2, 6)


            Case HatchStyle.DottedDiamond
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 8, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 8, 8, 0)

            Case HatchStyle.DottedGrid
                AppendPatternLine(xmlDoc, nodeHatch, 4, 0, 4, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 4, 8, 4)


            Case HatchStyle.ForwardDiagonal
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 8, 8)

            Case HatchStyle.Horizontal
                AppendPatternLine(xmlDoc, nodeHatch, 0, 4, 8, 4)

            Case HatchStyle.HorizontalBrick
                AppendPatternLine(xmlDoc, nodeHatch, 0, 3, 8, 3)
                AppendPatternLine(xmlDoc, nodeHatch, 3, 0, 3, 3)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 3, 0, 7)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 7, 7, 7)


            Case HatchStyle.LargeCheckerBoard
                AppendPatternRect(xmlDoc, nodeHatch, 0, 0, 3, 3)
                AppendPatternRect(xmlDoc, nodeHatch, 4, 4, 4, 4)

            Case HatchStyle.LargeConfetti
                AppendPatternRect(xmlDoc, nodeHatch, 0, 0, 1, 1)
                AppendPatternRect(xmlDoc, nodeHatch, 2, 3, 1, 1)
                AppendPatternRect(xmlDoc, nodeHatch, 5, 2, 1, 1)
                AppendPatternRect(xmlDoc, nodeHatch, 6, 6, 1, 1)


            Case HatchStyle.NarrowHorizontal
                AppendPatternLine(xmlDoc, nodeHatch, 0, 1, 8, 1)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 3, 8, 3)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 5, 8, 5)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 7, 8, 7)


            Case HatchStyle.NarrowVertical
                AppendPatternLine(xmlDoc, nodeHatch, 1, 0, 1, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 3, 0, 3, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 5, 0, 5, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 7, 0, 7, 8)


            Case HatchStyle.OutlinedDiamond
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 8, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 8, 0, 0, 8)


            Case HatchStyle.Plaid
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 8, 0)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 3, 8, 3)
                AppendPatternRect(xmlDoc, nodeHatch, 0, 4, 3, 3)

            Case HatchStyle.Shingle
                AppendPatternLine(xmlDoc, nodeHatch, 0, 2, 2, 0)
                AppendPatternLine(xmlDoc, nodeHatch, 2, 0, 7, 5)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 3, 3, 7)


            Case HatchStyle.SmallCheckerBoard
                AppendPatternRect(xmlDoc, nodeHatch, 0, 0, 1, 1)
                AppendPatternRect(xmlDoc, nodeHatch, 4, 4, 1, 1)
                AppendPatternRect(xmlDoc, nodeHatch, 4, 0, 1, 1)
                AppendPatternRect(xmlDoc, nodeHatch, 0, 4, 1, 1)


            Case HatchStyle.SmallConfetti
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 2, 2)
                AppendPatternLine(xmlDoc, nodeHatch, 7, 3, 5, 5)
                AppendPatternLine(xmlDoc, nodeHatch, 2, 6, 4, 4)


            Case HatchStyle.SmallGrid
                AppendPatternLine(xmlDoc, nodeHatch, 0, 2, 8, 2)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 6, 8, 6)
                AppendPatternLine(xmlDoc, nodeHatch, 2, 0, 2, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 6, 0, 6, 8)

            Case HatchStyle.SolidDiamond
                AppendPatternPolygon(xmlDoc, nodeHatch, "3 0 6 3 3 6 0 3")

            Case HatchStyle.Sphere
                AppendPatternEllipse(xmlDoc, nodeHatch, 3, 3, 2, 2)

            Case HatchStyle.Trellis
                AppendPatternLine(xmlDoc, nodeHatch, 0, 1, 8, 1)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 3, 8, 3)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 5, 8, 5)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 7, 8, 7)

            Case HatchStyle.Vertical
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 0, 8)


            Case HatchStyle.Wave
                AppendPatternLine(xmlDoc, nodeHatch, 0, 4, 3, 2)
                AppendPatternLine(xmlDoc, nodeHatch, 3, 2, 8, 4)


            Case HatchStyle.Weave
                AppendPatternLine(xmlDoc, nodeHatch, 0, 4, 4, 0)
                AppendPatternLine(xmlDoc, nodeHatch, 8, 4, 4, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 0, 4)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 4, 4, 8)


            Case HatchStyle.WideDownwardDiagonal
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 8, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 1, 8, 9)
                AppendPatternLine(xmlDoc, nodeHatch, 7, 0, 8, 1)


            Case HatchStyle.WideUpwardDiagonal
                AppendPatternLine(xmlDoc, nodeHatch, 8, 0, 0, 8)
                AppendPatternLine(xmlDoc, nodeHatch, 8, 1, 0, 9)
                AppendPatternLine(xmlDoc, nodeHatch, 0, 1, -1, 0)


            Case HatchStyle.ZigZag
                AppendPatternLine(xmlDoc, nodeHatch, 0, 4, 4, 0)
                AppendPatternLine(xmlDoc, nodeHatch, 4, 0, 8, 4)


            Case HatchStyle.Percent05
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 1, 0)
                AppendPatternLine(xmlDoc, nodeHatch, 4, 4, 5, 4)


            Case HatchStyle.Percent10
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 1, 0)
                AppendPatternLine(xmlDoc, nodeHatch, 4, 2, 5, 2)
                AppendPatternLine(xmlDoc, nodeHatch, 2, 4, 3, 4)
                AppendPatternLine(xmlDoc, nodeHatch, 6, 6, 7, 6)

            Case HatchStyle.Percent20
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 1, 0)
                AppendPatternLine(xmlDoc, nodeHatch, 4, 2, 5, 2)
                AppendPatternLine(xmlDoc, nodeHatch, 2, 4, 3, 4)
                AppendPatternLine(xmlDoc, nodeHatch, 5, 6, 7, 6)


            Case HatchStyle.Percent25
                AppendPatternLine(xmlDoc, nodeHatch, 0, 0, 3, 0)
                AppendPatternLine(xmlDoc, nodeHatch, 4, 2, 6, 2)
                AppendPatternLine(xmlDoc, nodeHatch, 2, 4, 5, 4)
                AppendPatternLine(xmlDoc, nodeHatch, 5, 6, 7, 6)


            Case HatchStyle.Percent30
                AppendPatternRect(xmlDoc, nodeHatch, 0, 0, 3, 1)
                AppendPatternLine(xmlDoc, nodeHatch, 4, 2, 6, 2)
                AppendPatternRect(xmlDoc, nodeHatch, 2, 4, 3, 1)
                AppendPatternLine(xmlDoc, nodeHatch, 5, 6, 7, 6)

            Case HatchStyle.Percent40
                AppendPatternRect(xmlDoc, nodeHatch, 0, 0, 3, 1)
                AppendPatternRect(xmlDoc, nodeHatch, 4, 2, 3, 1)
                AppendPatternRect(xmlDoc, nodeHatch, 2, 4, 3, 1)
                AppendPatternRect(xmlDoc, nodeHatch, 5, 6, 3, 1)

            Case HatchStyle.Percent50
                AppendPatternRect(xmlDoc, nodeHatch, 0, 0, 3, 3)
                AppendPatternRect(xmlDoc, nodeHatch, 4, 4, 4, 4)

            Case HatchStyle.Percent60
                AppendPatternRect(xmlDoc, nodeHatch, 0, 0, 4, 3)
                AppendPatternRect(xmlDoc, nodeHatch, 4, 4, 4, 4)


            Case HatchStyle.Percent70
                AppendPatternRect(xmlDoc, nodeHatch, 0, 0, 4, 5)
                AppendPatternRect(xmlDoc, nodeHatch, 4, 4, 4, 4)

            Case HatchStyle.Percent75
                AppendPatternRect(xmlDoc, nodeHatch, 0, 0, 7, 3)
                AppendPatternRect(xmlDoc, nodeHatch, 0, 2, 3, 7)

            Case HatchStyle.Percent80
                AppendPatternRect(xmlDoc, nodeHatch, 0, 0, 7, 4)
                AppendPatternRect(xmlDoc, nodeHatch, 0, 2, 4, 7)

            Case HatchStyle.Percent90
                AppendPatternRect(xmlDoc, nodeHatch, 0, 0, 7, 5)
                AppendPatternRect(xmlDoc, nodeHatch, 0, 2, 5, 7)

        End Select


        Return nodeHatch

    End Function



#End Region

#Region "Misc Methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a string representation of this object.
    ''' </summary>
    ''' <returns>A string representation of the object</returns>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function ToString() As String
        Return "Hatch Fill"
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if two GDIHatchFill objects are equivalent.
    ''' </summary>
    ''' <param name="fill1">The first GDIHatchFill object to compare</param>
    ''' <param name="fill2">The second GDIHatchFill object to compare</param>
    ''' <returns>A Boolean indicating whether the two fills are identical.</returns>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function op_Equality(ByVal fill1 As GDIHatchFill, ByVal fill2 As GDIHatchFill) As Boolean
        Dim bEquality As Boolean = True
        bEquality = bEquality And Color.op_Equality(fill1._HatchBackColor, fill2._HatchBackColor) And Color.op_Equality(fill1._HatchForeColor, fill2._HatchForeColor)

        If bEquality = False Then
            Return bEquality
        End If

        bEquality = bEquality And fill1._HatchStyle = fill2._HatchStyle

        Return bEquality
    End Function
#End Region

#Region "Base Class Implementers"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Allows the hatch fill to respond to changes in its parent state.  In this case,
    ''' nothing happens since the hatch fill has no dependencies on its parent like the 
    ''' texture or gradient fill.
    ''' </summary>
    ''' <param name="obj">The parent GDIFilledShape</param>
    ''' -----------------------------------------------------------------------------
    Public Overrides Sub OnParentUpdated(ByVal obj As GDIFilledShape)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a brush capable of drawing this fill.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub UpdateFill()
        If Not _Brush Is Nothing Then
            _Brush.Dispose()
        End If
        _Brush = New Drawing2D.HatchBrush(_HatchStyle, _HatchForeColor, _HatchBackColor)
        NotifyFillUpdated()
    End Sub
#End Region

#Region "Property Accessors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the forecolor of the hatch style 
    ''' </summary>
    ''' <value>A System.Drawing color.</value>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.Description("Forecolor of hatch fill")> _
  Public Property HatchForeColor() As Color
        Get
            Return _HatchForeColor
        End Get
        Set(ByVal Value As Color)
            Trace.WriteLineIf(Session.Tracing, "HatchFill.HatchForeColor.Set: " & Value.ToString)

            _HatchForeColor = Value
            UpdateFill()
        End Set
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the backcolor of the hatch style.
    ''' </summary>
    ''' <value>A System.Drawing color</value>
    '''  -----------------------------------------------------------------------------
    <ComponentModel.Description("backcolor of hatch fill")> _
Public Property HatchBackColor() As Color
        Get
            Return _HatchBackColor
        End Get
        Set(ByVal Value As Color)
            Trace.WriteLineIf(Session.Tracing, "HatchFill.HatchBackColor.Set: " & Value.ToString)

            _HatchBackColor = Value
            UpdateFill()
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the Hatch style of the hatch fill.
    ''' </summary>
    ''' <value>A Drawing2D.hatchstyle</value>
    '''  -----------------------------------------------------------------------------
    <ComponentModel.Description("hatch style")> _
    Public Property HatchStyle() As Drawing2D.HatchStyle
        Get
            Return _HatchStyle
        End Get
        Set(ByVal Value As Drawing2D.HatchStyle)
            Trace.WriteLineIf(Session.Tracing, "HatchFill.HatchStyle.Set: " & Value.ToString)

            _HatchStyle = Value
            UpdateFill()
        End Set
    End Property
#End Region
End Class
