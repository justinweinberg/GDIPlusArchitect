Imports System.CodeDom

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIEllipse
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Represents a circle or ellipse on the drawing surface.
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public Class GDIEllipse
    Inherits GDIFilledShape

#Region "Non serialized members"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Next valid ellipse integer suffix (Ellipse1, Ellipse2, etc)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _NextnameID As Int32 = 0


#End Region

#Region "Constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new instance of a GDIEllipse.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()
        Trace.WriteLineIf(Session.Tracing, "Creating ellipse")
    End Sub
#End Region

#Region "Code generation related methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts a GDIEllipse to SVG XML and appends it to the outgoing SVG code.
    ''' </summary>
    ''' <param name="xmlDoc">The SVG document to append the ellipse code to.</param>
    ''' <param name="defs">The definitions section of the SVG document.</param>
    ''' <param name="group">The group to append the XML spec of the ellipse to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub toXML(ByVal xmlDoc As Xml.XmlDocument, ByVal defs As Xml.XmlNode, ByVal group As Xml.XmlNode)
        '<ellipse cx="150" cy="150" rx="100" ry="50" style="fill:red; opacity:0.2; stroke:black; stroke-width:1pt;"/>

        Dim node As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "ellipse", String.Empty)

        Dim attr As Xml.XmlAttribute
        attr = xmlDoc.CreateAttribute("id")
        attr.Value = Me.Name
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("cx")
        attr.Value = String.Format("{0}", Me.Bounds.X + Me.Bounds.Width / 2)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("cy")
        attr.Value = String.Format("{0}", Me.Bounds.Y + Me.Bounds.Height / 2)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("rx")
        attr.Value = String.Format("{0}", Me.Bounds.Width / 2)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("ry")
        attr.Value = String.Format("{0}", Me.Bounds.Height / 2)
        node.Attributes.Append(attr)

        If Me.Rotation > 0 Then
            attr = xmlDoc.CreateAttribute("transform")
            attr.Value = "rotate(" & Me.Rotation & " " & Me.RotationPoint.X & " " & Me.RotationPoint.Y & ")"
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

        ' Dim sXML As String = String.Format("<rect x='{0}px' y='{1}px' rx='{2}px' ry='{3}px' ", Me.Bounds.X, Me.Bounds.Y, Me.Bounds.Width / 2, Me.Bounds.Height / 2)

        group.AppendChild(node)

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a GDIEllipse to code.
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

        'Declare the rectangle
        Dim declareEllipse As New CodeMemberField
        Dim ellipseInit As New CodeObjectCreateExpression

        Dim invokeStroke As New CodeMethodInvokeExpression
        Dim invokeFill As New CodeMethodInvokeExpression

        With ellipseInit
            .CreateType = New CodeTypeReference(GetType(Rectangle))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.X))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.Y))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.Width))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.Height))
        End With

        With declareEllipse
            .InitExpression = ellipseInit
            .Name = MyBase.ExportName()
            .Attributes = MyBase.getScope
            .Type = New CodeTypeReference(GetType(System.Drawing.Rectangle))
        End With

        If Rotation > 0 Then
            MyBase.emitRotationDeclaration(declarations)
            MyBase.emitInvokeBeginRotation(RenderGDI)
        End If

        If _DrawFill Then
            Dim sFillName As String = Me.Fill.emit(Me, declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings, Consolidated)

            With invokeFill
                .Method.TargetObject = New CodeArgumentReferenceExpression("g")
                .Method.MethodName = "FillEllipse"
                .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sFillName))
                .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, declareEllipse.Name))
            End With


            RenderGDI.Add(invokeFill)

        End If

        If _DrawStroke Then
            Dim sStrokeName As String = Me.Stroke.emit(Me, declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings, Consolidated)

            With invokeStroke
                .Method.TargetObject = New CodeArgumentReferenceExpression("g")
                .Method.MethodName = "DrawEllipse"
                .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sStrokeName))
                .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, declareEllipse.Name))
            End With

            RenderGDI.Add(invokeStroke)
        End If

        If Rotation > 0 Then
            MyBase.emitInvokeEndRotation(RenderGDI)
        End If

        declarations.Add(declareEllipse)

    End Sub

#End Region


#Region "Base class Implementers"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Recreates the graphics path used to render the ellipse to a drawing surface.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub createPath()
        MyBase.resetPath()
        _Path.AddEllipse(Bounds)
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a name used when displaying the class in the property browser.
    ''' </summary>
    ''' <value>A string containing the words "Ellipse"</value>
    ''' -----------------------------------------------------------------------------   
    Public Overrides ReadOnly Property ClassName() As String
        Get
            Return "Ellipse"
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the next valid name for a new GDIEllipse
    ''' </summary>
    ''' <returns>A string containing the next valid name.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Function NextName() As String
        _NextnameID += 1

        Return "Ellipse" & _NextnameID
    End Function
#End Region



#Region "Drawing related methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws the GDIEllipse to a graphics surface given a DrawingMode.
    ''' </summary>
    ''' <param name="g">Graphics context to render the ellipse to.</param>
    ''' <param name="eDrawMode">The current drawing mode (See EnumDrawMode for more information)</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub Draw(ByVal g As System.Drawing.Graphics, ByVal eDrawMode As EnumDrawMode)
        MyBase.BeginDraw(g)

        If _ResetPath Then
            createPath()
        End If

        If _DrawFill Then
            g.FillPath(Fill.Brush, _Path)
        End If

        If _DrawStroke Then
            g.DrawPath(Stroke.Pen, _Path)
        End If

        MyBase.EndDraw(g)

    End Sub

#End Region



End Class
