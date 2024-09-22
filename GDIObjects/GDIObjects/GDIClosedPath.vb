Imports System.CodeDom

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIClosedPath
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A closed path (A path whos starting point connects to its end point).  
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public Class GDIClosedPath
    Inherits GDIFilledShape

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a closed path given an open path which has been marked 
    ''' "to be closed".  This means that the user moved the mouse cursor within range of the 
    ''' start point while drawing an open path with the pen tool and clicked indicating they 
    ''' wish to closed the path.
    ''' </summary>
    ''' <param name="openpath">The open path to close on itself.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal openpath As GDIOpenPath)
        MyBase.New()
        Trace.WriteLineIf(Session.Tracing, "Creating closed path from openpath")

        ReDim _Points(openpath.Points.Length - 1)
        ReDim _PointTypes(openpath.Types.Length - 1)

        'set points and types from open path
        For i As Int32 = 0 To _Points.Length - 1
            _Points(i) = openpath.Points(i)
        Next

        For i As Int32 = 0 To _PointTypes.Length - 1
            _PointTypes(i) = openpath.Types(i)
        Next

        resetPath()

        'get the current in use fill
        Me.Fill = Session.Fill

        'flag for a path reset
        _ResetPath = True


    End Sub

#End Region

#Region "Local Fields"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An array of points that make up the path 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _Points As PointF()
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An array of types of the points in the path.  These types correspond to 
    ''' Drawing2D.PathPointType types.  For more information on the types of points 
    ''' that can be used in a path, see this enumeration.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _PointTypes As Byte()

#End Region

#Region "Non serialized fields"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Next valid closed path integer suffix (Path1, Path2, etc)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _NextnameID As Int32 = 0

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Used to supplement dragging of closed paths.  Maintains an array of points indicating 
    ''' by how much each point in the path has been offset by the drag operation.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
      Protected _OffSets As PointF()

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Records the last point set by hitting one of the handles in the path. This is used 
    ''' to note which point will be dragged if the user continues a drag operation from 
    ''' the down point.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
   Private _HitIndex As Int32 = -1
#End Region


#Region "Code Generation related methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a GDIClosedPath to code.
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
    Friend Overrides Sub emit(ByVal declarations As CodeDom.CodeTypeMemberCollection, _
       ByVal InitGraphics As CodeDom.CodeMemberMethod, _
       ByVal RenderGDI As CodeDom.CodeStatementCollection, _
       ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
       ByVal ExportSettings As ExportSettings, _
       ByVal Consolidated As ExportConsolidate)

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

    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts a GDIClosedPath to SVG XML and appends it to the outgoing SVG code.
    ''' </summary>
    ''' <param name="xmlDoc">The SVG document to append the closed path to.</param>
    ''' <param name="defs">The definitions section of the SVG document.</param>
    ''' <param name="group">The group to append the XML spec of the closed path to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub toXML(ByVal xmlDoc As System.Xml.XmlDocument, ByVal defs As System.Xml.XmlNode, ByVal group As System.Xml.XmlNode)

        Dim attr As Xml.XmlAttribute
        Dim node As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "path", String.Empty)

        attr = xmlDoc.CreateAttribute("id")
        attr.Value = Me.Name
        node.Attributes.Append(attr)

        Dim sPathData As String = String.Empty

        sPathData = "M" & Me.Points(0).X.ToString & " " & Me.Points(0).Y.ToString & " "

        For i As Int32 = 1 To Me.Points.Length - 1


            If _PointTypes(i) = CByte(Drawing2D.PathPointType.Bezier) Then
                Dim bez1 As PointF = _Points(i)
                Dim bez2 As PointF = _Points(i + 1)
                Dim bezClose As PointF = _Points(i + 2)
                sPathData &= "C" & bez1.X.ToString & "," & bez1.Y.ToString & " " & _
                bez2.X.ToString & "," & bez2.Y.ToString & " " & _
                bezClose.X.ToString & "," & bezClose.Y.ToString & " "
                i += 2
            Else
                sPathData &= "L" & Me.Points(i).X.ToString & " " & Me.Points(i).Y.ToString & " "
            End If


        Next

        sPathData &= " z"
        attr = xmlDoc.CreateAttribute("d")
        attr.Value = sPathData
        node.Attributes.Append(attr)




        If Me.DrawFill Then
            Me.Fill.toXML(xmlDoc, defs, node)
        Else
            attr = xmlDoc.CreateAttribute("fill")
            attr.Value = "none"
            node.Attributes.Append(attr)
        End If

        If Me.DrawStroke Then
            Me.Stroke.toXML(xmlDoc, node)
        End If
        group.AppendChild(node)


    End Sub

#End Region

#Region "Base class Implementers"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the next valid name for a new GDIClosedPath
    ''' </summary>
    ''' <returns>A string containing the next valid name.</returns>
    ''' 
    ''' -----------------------------------------------------------------------------
    Protected Overrides Function NextName() As String
        _NextnameID += 1

        Return "ClosedPath" & _NextnameID
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a name used when displaying the class in the property browser.
    ''' </summary>
    ''' <value>A string containing the words "Closed Path"</value>
    ''' -----------------------------------------------------------------------------   
    Public Overrides ReadOnly Property ClassName() As String
        Get
            Return "Closed Path"
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Resets the path, which means recreating the visually displayed graphics path.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub resetPath()
        If Not _Path Is Nothing Then
            _Path.Dispose()
        End If

        _Path = New Drawing2D.GraphicsPath(_Points, _PointTypes)

        If Me.TrackFill Then
            Fill.OnParentUpdated(Me)
        End If
        _ResetPath = False

    End Sub
#End Region




#Region "Drawing Related methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the closed path to a surface
    ''' </summary>
    ''' <param name="g">Graphics context to draw against</param>
    ''' <param name="eDrawMode">The current draw mode</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub Draw(ByVal g As System.Drawing.Graphics, ByVal eDrawMode As EnumDrawMode)

        If _ResetPath Then
            resetPath()
        End If

        If _DrawFill Then
            g.FillPath(Fill.Brush, _Path)
        End If

        If _DrawStroke Then
            g.DrawPath(Stroke.Pen, _Path)
        End If

    End Sub

#End Region

#Region "Handles and hit testing"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a series of HitHandle objects relevant to this object
    ''' </summary>
    ''' <param name="fscale">The current zoom factor of the surface</param>
    ''' <returns>An array of HitHandle object</returns>
    ''' -----------------------------------------------------------------------------
    Friend Overloads Overrides Function getHandles(ByVal fscale As Single) As HitHandle()
        Dim hw As Single = 6 / fscale
        Dim hwhalf As Single = hw / 2

        Dim hHandles(_Points.Length - 1) As HitHandle

        For i As Int32 = 0 To hHandles.Length - 1
            Dim rect As RectangleF = New RectangleF(_Points(i).X - hwhalf, _Points(i).Y - hwhalf, hw, hw)

            Select Case _PointTypes(i)
                Case 3
                    hHandles(i) = New HitHandle(rect, HitHandle.EnumHandletypes.eCurvePoint)
                Case Else
                    hHandles(i) = New HitHandle(rect, HitHandle.EnumHandletypes.eNormal)
            End Select

        Next

        Return hHandles

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the last recorded hithandle
    ''' </summary>
    ''' <value>Returns the last hit handle</value>
    ''' -----------------------------------------------------------------------------
    <System.ComponentModel.Browsable(False)> _
 Public ReadOnly Property LastHitHandle() As PointF
        Get
            If _HitIndex > -1 Then
                Return _Points(_HitIndex)
            Else
                Return New PointF(0, 0)
            End If
        End Get

    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a hit test over a series of handles given a point.  Assigns the _HitIndex
    ''' field to the handle hit, if any, or -1 if no handle was hit.
    ''' </summary>
    ''' <param name="ptHit">The point to test handles against</param>
    ''' <param name="fScale">The current zoom factor of the surface</param>
    ''' <returns>The type of handle hit.  For Closed paths this is either a curve or normal 
    ''' point.</returns>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function HitTestHandles(ByVal ptHit As PointF, ByVal fScale As Single) As EnumDragHandles
        Dim h() As HitHandle = getHandles(fScale)

        For i As Int32 = 0 To h.Length - 1
            If h(i).Contains(ptHit) Then
                _HitIndex = i
                Return EnumDragHandles.ePointHandle
            End If
        Next

        _HitIndex = -1

        Return EnumDragHandles.eNone


    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a hit test for the GDIClosedPath against a rectangle. 
    ''' </summary>
    ''' <param name="rect">A rectangle to compare intersection with the GDIOpenPath 
    ''' against.</param>
    ''' <returns>A Boolean indicating if the rectangle and the path intersect.</returns>
    ''' -----------------------------------------------------------------------------
    Public Overloads Overrides Function HitTest(ByVal rect As Rectangle) As Boolean
        Dim g As Graphics = Session.GraphicsManager.getTempGraphics()
        Dim bIntersect As Boolean


        Dim r As New Region(_Path)

        r.Intersect(rect)


        bIntersect = Not r.IsEmpty(g)
        r.Dispose()
        g.Dispose()


        Return bIntersect

    End Function




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs hit testing on a closed path given a point and a surface scale.
    ''' </summary>
    ''' <param name="pt">A point in object coordinate space to hit test against</param>
    ''' <param name="fScale">The current zoom factor of the surface</param>
    ''' <returns>A Boolean indicating if the point lies on the path.</returns>
    ''' -----------------------------------------------------------------------------
    Public Overloads Overrides Function HitTest(ByVal pt As System.Drawing.PointF, ByVal fScale As Single) As Boolean
        Dim hw As Single = 6 / fScale
        Dim bIntersectPath As Boolean = False

        Dim myPen As New Pen(Me.Stroke.Color, Me.Stroke.Width * hw)

        If _Path.IsVisible(pt) OrElse HitTestHandles(pt, fScale) = EnumDragHandles.ePointHandle Then
            Return True
        End If

        myPen.Dispose()


    End Function
#End Region

#Region "Property Accessors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Shadows the rotation property.  This is because rotation of ClosedPaths has not 
    ''' been implemented.
    ''' </summary>
    ''' <value>A single (not used)</value>
    ''' -----------------------------------------------------------------------------
    <System.ComponentModel.Browsable(False)> _
         Public Shadows Property Rotation() As Single
        Get
            Return 0
        End Get
        Set(ByVal Value As Single)

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the bounds of the closed path as a rectangle.  
    ''' </summary>
    ''' <value>A rectangle containing the path</value>
    ''' <remarks>A get operation here is simple and returns the path bounds. However,
    ''' a set operation is relatively complex.  This is because of the user's expectation of 
    ''' what should happen to the individual points in the path when bounds change.  
    ''' Points should scale like blowing up or deflating a balloon.
    ''' </remarks>
    ''' ----------------------------------------------------------------------------- 
    <System.ComponentModel.Browsable(False)> _
   Public Overrides Property Bounds() As Rectangle
        Get
            Return Rectangle.Round(_Path.GetBounds())

        End Get
        Set(ByVal Value As Rectangle)

            Dim dW As Single
            Dim dH As Single

            Dim curbounds As Rectangle = Rectangle.Round(_Path.GetBounds())

            If _ResetPath Then
                resetPath()
            End If

            If curbounds.Width > 0 Then
                dW = CSng(Value.Width / curbounds.Width)
            Else
                dW = 0
            End If

            If curbounds.Height > 0 Then
                dH = CSng(Value.Height / curbounds.Height)
            Else
                dH = 0
            End If

            Dim scaledWidth As Single = curbounds.Width * dW
            Dim scaledHeight As Single = curbounds.Height * dH

            Dim originalCenterX As Int32 = curbounds.X + (curbounds.Width \ 2)
            Dim originalCenterY As Int32 = curbounds.Y + (curbounds.Height \ 2)

            Dim scaledRect As New RectangleF(originalCenterX - (scaledWidth / 2), originalCenterY - scaledHeight / 2, scaledWidth, scaledHeight)


            'The distance between each original.x and object.x should be scaled by fprcW and then added to scaledRect.X
            For i As Int32 = 0 To _Points.Length - 1
                Dim distOriginalX As Single = _Points(i).X - curbounds.X
                Dim distOriginalY As Single = _Points(i).Y - curbounds.Y
                _Points(i).X = scaledRect.X + (distOriginalX * dW)
                _Points(i).Y = scaledRect.Y + (distOriginalY * dH)
                _Points(i).X += (Value.X - scaledRect.X)
                _Points(i).Y += (Value.Y - scaledRect.Y)
            Next

            _ResetPath = True
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the points that make up the closed path
    ''' </summary>
    ''' <value>The points that make up the path</value>
    ''' -----------------------------------------------------------------------------   
    <System.ComponentModel.Browsable(False)> _
      Public ReadOnly Property Points() As PointF()
        Get
            Return _Points
        End Get
    End Property

#End Region

#Region "Dragging related methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Overrides the default drag behavior to provide custom drag functionality.  
    ''' The point past in as the parameter is a point in object  coordinate space 
    ''' from which the drag operation is initially invoked.
    ''' </summary>
    ''' <param name="ptObject">The point from which the drag operation is beginning</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub startDrag(ByVal ptObject As Point)
        ReDim _OffSets(_Points.Length)

        For i As Int32 = 0 To _Points.Length - 1

            Dim newpoint As New PointF(_Points(i).X - ptObject.X, _Points(i).Y - ptObject.Y)
            _OffSets(i) = newpoint
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates a drag operation given the latest drag point recorded in object coordinate 
    ''' space.
    ''' </summary>
    ''' <param name="ptObject">The last recorded drag point from the user interface.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub updateDrag(ByVal ptObject As Point)

        For i As Int32 = 0 To _Points.Length - 1
            _Points(i).X = _OffSets(i).X + ptObject.X
            _Points(i).Y = _OffSets(i).Y + ptObject.Y
        Next

        _ResetPath = True

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles an explicit point set operation on a point in the closed path.  This 
    ''' occurs when the user is dragging a point in the path as opposed to the entire 
    ''' path object which is a drag operation
    ''' </summary>
    ''' <param name="ptObject">A point in object coordinate space that contains the new 
    ''' point to assign to the last hit point recorded in the _HitIndex field</param>
    ''' -----------------------------------------------------------------------------
    Public Overrides Sub handlePointSet(ByVal ptObject As Point)
        _Points(_HitIndex) = Point.op_Implicit(ptObject)

        _ResetPath = True
    End Sub
#End Region



#Region "Path Creation"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Regenerates the path drawn to surfaces and used in hit testing.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub createPath()
        If Not _Path Is Nothing Then
            _Path.Dispose()
        End If

        _Path = New Drawing2D.GraphicsPath(_Points, _PointTypes)
    End Sub

#End Region


End Class
