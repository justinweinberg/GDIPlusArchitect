Imports System.CodeDom
Imports System.ComponentModel



''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDILine
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Represents a Line.  Note that the line is rendered to the surface by its base 
''' class, GDIOpenPath
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public Class GDILine
    Inherits GDIOpenPath
    


#Region "Non serialized members"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Next valid integer suffix for naming lines (Line1, Line2)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _NextnameID As Int32 = 0
#End Region

#Region "Constructors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs an instance of a GDILine from a start and end point.
    ''' </summary>
    ''' <param name="pt1">The first point in the line.</param>
    ''' <param name="pt2">The second point in the line.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal pt1 As Point, ByVal pt2 As Point)
        MyBase.New(pt1)

        Trace.WriteLineIf(Session.Tracing, "Line.New")
        Trace.Indent()
        Trace.WriteLineIf(Session.Tracing, "PT1: " & pt1.ToString)
        Trace.WriteLineIf(Session.Tracing, "PT2: " & pt1.ToString)
        Trace.Unindent()

        _Points = New PointF() {Point.op_Implicit(pt1), Point.op_Implicit(pt2)}
        _PointTypes = New Byte() {0, 1}

    End Sub

#End Region


#Region "Code Emitting Methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a GDILine into code. 
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
        Dim declarePointStart As New CodeMemberField
        Dim declarePointEnd As New CodeMemberField

        Dim PointStartInit As New CodeObjectCreateExpression
        Dim PointEndInit As New CodeObjectCreateExpression

        Dim invokeStroke As New CodeMethodInvokeExpression
        Dim invokeFill As New CodeMethodInvokeExpression

        With PointStartInit
            .CreateType = New CodeTypeReference(GetType(PointF))
            .Parameters.Add(New CodePrimitiveExpression(Me._Points(0).X))
            .Parameters.Add(New CodePrimitiveExpression(Me._Points(0).Y))
        End With

        With PointEndInit
            .CreateType = New CodeTypeReference(GetType(PointF))
            .Parameters.Add(New CodePrimitiveExpression(Me._Points(1).X))
            .Parameters.Add(New CodePrimitiveExpression(Me._Points(1).Y))
        End With

        With declarePointStart
            .InitExpression = PointStartInit
            .Name = MyBase.ExportName() & "PointStart"
            .Attributes = MyBase.getScope
            .Type = New CodeTypeReference(GetType(System.Drawing.PointF))
        End With


        With declarePointEnd
            .InitExpression = PointEndInit
            .Name = MyBase.ExportName() & "PointEnd"
            .Attributes = MyBase.getScope
            .Type = New CodeTypeReference(GetType(System.Drawing.PointF))
        End With


        Dim sStrokeName As String = Me.Stroke.emit(Me, declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings, Consolidated)

        With invokeStroke
            .Method.TargetObject = New CodeArgumentReferenceExpression("g")
            .Method.MethodName = "DrawLine"
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sStrokeName))
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, declarePointStart.Name))
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, declarePointEnd.Name))
        End With

        RenderGDI.Add(invokeStroke)

        declarations.Add(declarePointStart)
        declarations.Add(declarePointEnd)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts a GDILine to SVG XML and appends it to the outgoing SVG code.
    ''' </summary>
    ''' <param name="xmlDoc">The SVG document to append the line to.</param>
    ''' <param name="defs">The definitions section of the SVG document.</param>
    ''' <param name="group">The group to append the XML spec of the line to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub toXML(ByVal xmlDoc As Xml.XmlDocument, ByVal defs As Xml.XmlNode, ByVal group As Xml.XmlNode)

        Dim attr As Xml.XmlAttribute

        Dim node As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "line", String.Empty)
        attr = xmlDoc.CreateAttribute("id")
        attr.Value = Me.Name
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("x1")
        attr.Value = String.Format("{0}", Me.Points(0).X)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("y1")
        attr.Value = String.Format("{0}", Me.Points(0).Y)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("x2")
        attr.Value = String.Format("{0}", Me.Points(1).X)
        node.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("y2")
        attr.Value = String.Format("{0}", Me.Points(1).Y)
        node.Attributes.Append(attr)



        Me.Stroke.toXML(xmlDoc, node)

        group.AppendChild(node)
        ' Return String.Format("<line x1='{0}px' y1='{1}px' x2='{2}px' y2='{3}px' {4}/>", _
        '  Me.Points(0).X, Me.Points(0).Y, Me.Points(1).X, Me.Points(1).Y, sStyle)
    End Sub

#End Region


#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a name used when displaying the class in the property browser.
    ''' </summary>
    ''' <value>A string containing the word "Line"</value>
    ''' -----------------------------------------------------------------------------    
    Public Overrides ReadOnly Property ClassName() As String
        Get
            Return "Line"
        End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the first point in the line.
    ''' </summary>
    ''' <value>The first point in the set.</value>
    ''' -----------------------------------------------------------------------------
    <Description("Start Point of the line")> _
    Public Property StartPoint() As Point
        Get
            Return Point.Round(_Points(0))
        End Get
        Set(ByVal Value As Point)
            Try
                Trace.WriteLineIf(Session.Tracing, "Line.StartPoint.Set: " & Value.ToString)

                _Points = New PointF() {Point.op_Implicit(Value), _Points(1)}
                _PointTypes = New Byte() {0, 1}
                _ResetPath = True
            Catch ex As Exception

            End Try
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the second point in the line.
    ''' </summary>
    ''' <value>The second point in the set.</value>
    ''' -----------------------------------------------------------------------------
    <Description("End Point of the line")> _
        Public Property EndPoint() As Point
        Get
            Return Point.Round(_Points(1))
        End Get
        Set(ByVal Value As Point)
            Try
                Trace.WriteLineIf(Session.Tracing, "Line.EndPoint.Set: " & Value.ToString)

                _Points = New PointF() {_Points(0), Point.op_Implicit(Value)}
                _PointTypes = New Byte() {0, 1}
                _ResetPath = True
            Catch ex As Exception

            End Try
        End Set
    End Property

#End Region

#Region "Base class Implementers"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the next valid name for a new GDILine.
    ''' </summary>
    ''' <returns>A string containing the next valid name.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Function NextName() As String
        _NextnameID += 1
        Return "Line" & _NextnameID
    End Function


#End Region

End Class
 