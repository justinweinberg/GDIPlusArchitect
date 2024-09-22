Imports System.CodeDom
Imports System.ComponentModel


''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIText
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Contains everything necessary to render text to a GDI+ Architect surface as well 
''' as provides emit functionality for Text.
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public Class GDIText
    Inherits GDITextBaseClass

#Region "non serialized members"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Holds the next valid integer suffix for text objects (Text1, Text2, etc)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _NextnameID As Int32 = 0
#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Creates a new instance of a GDIText object
    ''' </summary>
    ''' <param name="Text">Text to place inside the object</param>
    ''' <param name="Font">Font to use with the text</param>
    ''' <param name="rectBounds">Initial bounds of the GDIText object</param>
    ''' <param name="wrap">Whether the text should initially wrap or not to its bounds</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal Text As String, ByVal Font As Font, ByVal rectBounds As Rectangle, ByVal wrap As Boolean)
        MyBase.New(Text, Font, rectBounds, wrap)
    End Sub
#End Region

#Region "Base Class Implementers"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the next valid name for a new GDIText object
    ''' </summary>
    ''' <returns>A string containing the next valid name.</returns>
    ''' 
    ''' -----------------------------------------------------------------------------
    Protected Overrides Function NextName() As String
        _NextnameID += 1

        Return "Text" & _NextnameID
    End Function
#End Region

#Region "Code generation related methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a variable width GDIText object
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
    Private Sub emitVariableWidth(ByVal declarations As CodeDom.CodeTypeMemberCollection, _
    ByVal InitGraphics As CodeDom.CodeMemberMethod, _
    ByVal RenderGDI As CodeDom.CodeStatementCollection, _
    ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
    ByVal ExportSettings As ExportSettings, _
    ByVal Consolidated As ExportConsolidate)

        Trace.WriteLineIf(Session.Tracing, "Text.EmitVariableWidth")

        Dim sFontName As String = String.Empty
        Dim sFormatName As String = String.Empty

        Dim XDeclaration As New CodeMemberField
        Dim YDeclaration As New CodeMemberField

        Dim textDeclaration As New CodeMemberField

        Dim invokeFill As New CodeMethodInvokeExpression

        With textDeclaration
            .Name = MyBase.ExportName()
            .Attributes = MyBase.getScope
            .Type = New CodeTypeReference(GetType(System.String))
            .InitExpression = New CodePrimitiveExpression(_Text)
        End With

        With XDeclaration
            .Name = MyBase.ExportName() & "PointX"
            .Type = New CodeTypeReference(GetType(System.Single))
            .Attributes = MemberAttributes.Private
            .InitExpression = New CodePrimitiveExpression(_Bounds.X)
        End With

        With YDeclaration
            .Name = MyBase.ExportName() & "PointY"
            .Type = New CodeTypeReference(GetType(System.Single))
            .Attributes = MemberAttributes.Private
            .InitExpression = New CodePrimitiveExpression(_Bounds.Y)
        End With


        declarations.Add(textDeclaration)
        declarations.Add(XDeclaration)
        declarations.Add(YDeclaration)



        If Consolidated.hasFontMatch(Me.Font) AndAlso Me._ConsolidateFont Then
            sFontName = Consolidated.getFontName(Me.Font)
        Else
            'Create the font object
            Dim fontDeclaration As New CodeMemberField

            With fontDeclaration
                .Name = MyBase.ExportName() & "Font"
                .Type = New CodeTypeReference(GetType(System.Drawing.Font))
                .Attributes = MemberAttributes.Private
                .InitExpression = getFontInitializer(ExportSettings)
            End With

            declarations.Add(fontDeclaration)
            sFontName = fontDeclaration.Name

            With DisposeGDI
                .Statements.Add(getFontDisposal(fontDeclaration.Name))
            End With
        End If



        Dim sFillName As String = Me.Fill.emit(Me, declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings, Consolidated)

        With invokeFill
            .Method.TargetObject = New CodeArgumentReferenceExpression("g")
            .Method.MethodName = "DrawString"
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, textDeclaration.Name))
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sFontName))
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sFillName))
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, XDeclaration.Name))
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, YDeclaration.Name))
        End With


        If Rotation > 0 Then
            MyBase.emitRotationDeclaration(declarations)
            MyBase.emitInvokeBeginRotation(RenderGDI)
        End If

        RenderGDI.Add(invokeFill)

        If Rotation > 0 Then
            MyBase.emitInvokeEndRotation(RenderGDI)
        End If

    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a fixed width GDIText object (one with an explicit width bounds).
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
    Private Sub emitFixedWidth(ByVal declarations As CodeDom.CodeTypeMemberCollection, _
    ByVal InitGraphics As CodeDom.CodeMemberMethod, _
    ByVal RenderGDI As CodeDom.CodeStatementCollection, _
    ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
    ByVal ExportSettings As ExportSettings, _
    ByVal Consolidated As ExportConsolidate)

        Trace.WriteLineIf(Session.Tracing, "Text.emitFixedWidth")

        Dim sFontName As String = String.Empty
        Dim sFormatName As String = String.Empty

        Dim rectDeclaration As New CodeMemberField
        Dim textDeclaration As New CodeMemberField

        Dim rectInit As New CodeObjectCreateExpression


        Dim invokeFill As New CodeMethodInvokeExpression

        With rectInit
            .CreateType = New CodeTypeReference(GetType(RectangleF))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.X))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.Y))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.Width))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.Height))
        End With


        With textDeclaration
            .Name = MyBase.ExportName()
            .Type = New CodeTypeReference(GetType(System.String))
            .Attributes = MyBase.getScope
            .InitExpression = New CodePrimitiveExpression(_Text)
        End With

        With rectDeclaration
            .Name = MyBase.ExportName() & "Rect"
            .Type = New CodeTypeReference(GetType(System.Drawing.RectangleF))
            .Attributes = MemberAttributes.Private
            .InitExpression = rectInit
        End With




        declarations.Add(textDeclaration)
        declarations.Add(rectDeclaration)


        If Consolidated.hasStringFormatMatch(Me.StringFormat) AndAlso Me._ConsolidateFormat Then
            sFormatName = Consolidated.getStringFormatName(Me.StringFormat)
        Else

            Dim formatDeclaration As New CodeMemberField
            Dim formatInit As New CodeObjectCreateExpression(GetType(System.Drawing.StringFormat))
            formatInit.Parameters.Add(New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(System.Drawing.StringFormat)), "GenericTypographic"))


            With formatDeclaration
                .Name = MyBase.ExportName() & "Format"
                .Type = New CodeTypeReference(GetType(System.Drawing.StringFormat))
                .Attributes = MemberAttributes.Private
                .InitExpression = formatInit
            End With
            declarations.Add(formatDeclaration)
            Dim formatDispose As New CodeMethodInvokeExpression


            'create dispose call
            With formatDispose
                .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, formatDeclaration.Name)
                .Method.MethodName = "Dispose"
            End With

            With DisposeGDI
                .Statements.Add(formatDispose)
            End With

            sFormatName = formatDeclaration.Name


            With InitGraphics

                Dim AlignAssignment As New CodeDom.CodeAssignStatement

                AlignAssignment.Left = New CodeFieldReferenceExpression(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sFormatName), "Alignment")
                AlignAssignment.Right = New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(System.Drawing.StringAlignment)), _StringFormat.Alignment.ToString)

                .Statements.Add(AlignAssignment)
            End With
        End If

        If Consolidated.hasFontMatch(Me.Font) AndAlso Me._ConsolidateFont Then
            sFontName = Consolidated.getFontName(Me.Font)
        Else
            'Create the font object
            Dim fontDeclaration As New CodeMemberField

            With fontDeclaration
                .Name = MyBase.ExportName() & "Font"
                .Type = New CodeTypeReference(GetType(System.Drawing.Font))
                .Attributes = MemberAttributes.Private
                .InitExpression = getFontInitializer(ExportSettings)
            End With

            declarations.Add(fontDeclaration)
            sFontName = fontDeclaration.Name

            With DisposeGDI
                .Statements.Add(getFontDisposal(fontDeclaration.Name))
            End With
        End If


        Dim sFillName As String = Me.Fill.emit(Me, declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings, Consolidated)


        With invokeFill
            .Method.TargetObject = New CodeArgumentReferenceExpression("g")
            .Method.MethodName = "DrawString"

            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, textDeclaration.Name))
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sFontName))
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sFillName))

            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, rectDeclaration.Name))
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sFormatName))
        End With

        If Rotation > 0 Then
            MyBase.emitRotationDeclaration(declarations)
            MyBase.emitInvokeBeginRotation(RenderGDI)
        End If

        RenderGDI.Add(invokeFill)


        If Rotation > 0 Then
            MyBase.emitInvokeEndRotation(RenderGDI)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a GDIText object to code.
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
    Friend Overrides Sub emit( _
    ByVal declarations As CodeDom.CodeTypeMemberCollection, _
    ByVal InitGraphics As CodeDom.CodeMemberMethod, _
    ByVal RenderGDI As CodeDom.CodeStatementCollection, _
    ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
    ByVal ExportSettings As ExportSettings, _
    ByVal Consolidated As ExportConsolidate)

        Trace.WriteLineIf(Session.Tracing, "Text.emit")


        If _DefinedWidth Then
            emitFixedWidth(declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings, Consolidated)
        Else
            emitVariableWidth(declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings, Consolidated)
        End If


    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts  a GDIText bit to SVG XML and appends it to the outgoing SVG code.
    ''' </summary>
    ''' <param name="xmlDoc">The SVG document to append the text to.</param>
    ''' <param name="defs">The definitions section of the SVG document.</param>
    ''' <param name="group">The group to append the XML spec of the text to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub toXML(ByVal xmlDoc As System.Xml.XmlDocument, ByVal defs As System.Xml.XmlNode, ByVal group As System.Xml.XmlNode)
        '<text x="0" y="13" fill="red" text-anchor="start">Text</text>

        MyBase.UpdateCharRanges()
        Dim attr As Xml.XmlAttribute

        Dim node As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "text", String.Empty)

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

        If Me.Rotation > 0 Then
            attr = xmlDoc.CreateAttribute("transform")
            attr.Value = "rotate(" & Me.Rotation & " " & Me.RotationPoint.X & " " & RotationPoint.Y & ")"
            node.Attributes.Append(attr)
        End If


        attr = xmlDoc.CreateAttribute("baseline-shift")
        attr.Value = "-86%"
        node.Attributes.Append(attr)



        attr = xmlDoc.CreateAttribute("font-family")
        attr.Value = Me.Font.Name
        node.Attributes.Append(attr)

        'em, ex, px, pt, pc, cm, mm,
        attr = xmlDoc.CreateAttribute("font-size")
        Dim unit As String

        Select Case Me.Font.Unit
            Case GraphicsUnit.Display
                unit = "em"
            Case GraphicsUnit.Document
                unit = "em"
            Case GraphicsUnit.Inch
                unit = "in"
            Case GraphicsUnit.Millimeter
                unit = "mm"
            Case GraphicsUnit.Pixel
                unit = "px"
            Case GraphicsUnit.Point
                unit = "pt"
            Case GraphicsUnit.World
                unit = "em"
        End Select

        attr.Value = String.Format("{0}{1}", Me.Font.Size, unit)
        node.Attributes.Append(attr)

        If Not Me.Alignment = StringAlignment.Near Then
            Dim sAnchor As String
            attr = xmlDoc.CreateAttribute("text-anchor")
            Select Case Me.Alignment

                Case StringAlignment.Far
                    sAnchor = "middle"
                Case StringAlignment.Near
                    sAnchor = "end"
            End Select

            attr.Value = sAnchor
            node.Attributes.Append(attr)

        End If

        If Me.Font.Bold Then

            attr = xmlDoc.CreateAttribute("font-weight")
            attr.Value = "bold"
            node.Attributes.Append(attr)

        End If

        If Me.Font.Italic Then
            attr = xmlDoc.CreateAttribute("font-style")
            attr.Value = "italic"
            node.Attributes.Append(attr)
        End If


        If Me.Font.Underline And Me.Font.Strikeout Then
            attr = xmlDoc.CreateAttribute("text-decoration")
            attr.Value = "underline line-through"
            node.Attributes.Append(attr)
        ElseIf Me.Font.Underline Then
            attr = xmlDoc.CreateAttribute("text-decoration")
            attr.Value = "underline"
            node.Attributes.Append(attr)
        ElseIf Me.Font.Strikeout Then
            attr = xmlDoc.CreateAttribute("text-decoration")
            attr.Value = "line-through"
            node.Attributes.Append(attr)
        End If

        If Me.Wrap Then
            Dim tspanNode As Xml.XmlNode


            Dim lastx As Int32
            Dim lastY As Int32
            Dim sLineText As String

            lastx = Me.Bounds.X
            lastY = Me.Bounds.Y
            tspanNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "tspan", String.Empty)

            attr = xmlDoc.CreateAttribute("x")
            attr.Value = String.Format("{0}", Me.Bounds.X.ToString)
            tspanNode.Attributes.Append(attr)

            attr = xmlDoc.CreateAttribute("y")
            attr.Value = String.Format("{0}", Me.Bounds.Y.ToString)
            tspanNode.Attributes.Append(attr)

            node.AppendChild(tspanNode)

            For i As Int32 = 0 To _RectCharRanges.Count - 1
                Dim rectf As Rectangle = Rectangle.Round(DirectCast(_RectCharRanges(i), RectangleF))


                If rectf.Y > lastY Then
                    'end current tspan
                    tspanNode.InnerText = sLineText
                    sLineText = Me.Text.Chars(i)
                    lastY = rectf.Y

                    tspanNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "tspan", String.Empty)

                    attr = xmlDoc.CreateAttribute("x")
                    attr.Value = Me.Bounds.X.ToString
                    tspanNode.Attributes.Append(attr)

                    attr = xmlDoc.CreateAttribute("y")
                    attr.Value = lastY.ToString
                    tspanNode.Attributes.Append(attr)

                    node.AppendChild(tspanNode)
                Else
                    sLineText &= Me.Text.Chars(i)

                End If

            Next


            If Not sLineText = String.Empty Then
                tspanNode.InnerText = sLineText
            End If
        Else
            node.InnerText = Me.Text
        End If


        Me.Fill.toXML(xmlDoc, defs, node)

        group.AppendChild(node)
        Dim g As Graphics
        Dim x As StringFormat

    End Sub

#End Region

#Region "Drawing related methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the text in tool mode.  Used by the GDIPlus Architect project to draw the 
    ''' text tool without having to either expose Draw for the entire GDIObject hierarchy
    ''' or alternatively duplicate work in the GDIPlus Architect project
    ''' </summary>
    ''' <param name="g">System.Drawing.Graphics context to draw against.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub RenderTool(ByVal g As Graphics)
        Me.Draw(g, EnumDrawMode.eNormal)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws the GDIText to a surface for a specific drawmode.
    ''' </summary>
    ''' <param name="g">The graphics surface to draw the GDIText object to.</param>
    ''' <param name="eDrawMode">The mode (see EnumDrawMode) that indicates how to render the GDIText object.</param>
    ''' <remarks>GDIText is one of the few objects that actually makes use of the EnumDrawMode.  Whereas most
    ''' objects draw the same regardless of the EnumDrawMode, a couple things are different with this 
    ''' instance.  
    ''' 
    ''' For printing, the font differ slightly than screen fonts depending on if they are measured in non screen 
    ''' units (units besides Pixel and World).
    ''' 
    ''' For drawing to the surface, not only is the text rendered, but depending on user settings, borders are 
    ''' drawn around the text as helpers for determining the position of the text.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub Draw(ByVal g As System.Drawing.Graphics, ByVal eDrawMode As EnumDrawMode)


        MyBase.BeginDraw(g)

        Select Case eDrawMode

            Case EnumDrawMode.ePrinting
                Dim printerFont As Font

                If _Font.Unit = GraphicsUnit.Pixel OrElse _Font.Unit = GraphicsUnit.World Then
                    printerFont = New Font(_Font.Name, _Font.Size, _Font.Style, _Font.Unit, _Font.GdiCharSet)
                Else
                    printerFont = New Font(_Font.Name, _Font.Size * (Session.Settings.DPIY / 100), _Font.Style, _Font.Unit, _Font.GdiCharSet)
                End If


                If _DefinedWidth Then
                    g.DrawString(_Text, printerFont, Fill.Brush, New RectangleF(_Bounds.X, _Bounds.Y, _Bounds.Width, _Bounds.Height), _StringFormat)
                Else
                    g.DrawString(_Text, printerFont, Fill.Brush, _Bounds.X, _Bounds.Y, _StringFormat)
                End If

                printerFont.Dispose()

            Case EnumDrawMode.eGraphicExport
                If _DefinedWidth Then
                    g.DrawString(_Text, _Font, Fill.Brush, New RectangleF(_Bounds.X, _Bounds.Y, _Bounds.Width, _Bounds.Height), _StringFormat)
                Else
                    g.DrawString(_Text, _Font, Fill.Brush, _Bounds.X, _Bounds.Y, _StringFormat)
                End If


            Case Else
                If Session.Settings.DrawTextFieldBorders Then

                    Dim pDotted As Pen


                    pDotted = New Pen(Color.Gray, 1)

                    pDotted.DashStyle = Drawing2D.DashStyle.Dash

                    g.DrawRectangle(pDotted, _Bounds)

                    'reset the graphics object to its original value
                    pDotted.Dispose()
                End If

                If _DefinedWidth Then
                    g.DrawString(_Text, _Font, Fill.Brush, New RectangleF(_Bounds.X, _Bounds.Y, _Bounds.Width, _Bounds.Height), _StringFormat)
                Else
                    g.DrawString(_Text, _Font, Fill.Brush, _Bounds.X, _Bounds.Y, _StringFormat)
                End If


        End Select

        MyBase.EndDraw(g)



    End Sub

#End Region


End Class
