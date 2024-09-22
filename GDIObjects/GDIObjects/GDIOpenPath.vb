Imports System.CodeDom
Imports System.ComponentModel


''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIOpenPath
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Represents an Open Path.  This class is juxtaposed to the GDIFilledShape class.
''' Whereas the GDIFilledShape represents a closed fillable shape, the OpenPath represents
''' objects where a fill is inappropriate such as with lines and open pen paths that 
''' do not close.
''' </summary>
''' <remarks>The key to the path based functionality in GDI+ Architect revolves around 
''' having an array of points that make up the path and a complementary array of point 
''' types which correspond to Drawing2D.PathPointType instances.
''' </remarks>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public Class GDIOpenPath
    Inherits GDIShape

#Region "Non Serialized Members"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Array of points used to aid in drag operations.  Contains the offsets from the 
    ''' original location of the path
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Protected _Offsets As PointF()


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Used to note the index position in the point array of a hit point. 
    ''' 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _HitIndex As Int32 = -1

#End Region


#Region "Local Fields"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The next valid integer name suffix (Openpath1, OpenPath2, ...)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _NextnameID As Int32 = 0

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

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a GDIOpenPath.  While the path is valid at this point 
    ''' from a GDI+ standpoint, it isn't interesting until segments have been added, 
    ''' since a path of a single point will not render
    ''' </summary>
    ''' <param name="ptOrigin">The first point in the path</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal ptOrigin As Point)
        MyBase.New()
        Trace.WriteLineIf(Session.Tracing, "OpenPath.New")
        Trace.WriteLineIf(Session.Tracing, "ptOrigin:" & ptOrigin.ToString)
        Trace.WriteLineIf(Session.Tracing, "OpenPath.New")


        ReDim _PointTypes(0)
        ReDim _Points(0)

        _Points(0) = Point.op_Implicit(ptOrigin)
        _PointTypes(0) = CByte(Drawing2D.PathPointType.Start)
    End Sub

#End Region

#Region "Code emit related methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a GDIOpenPath to code.
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
        'Open paths are not filled
        '  Dim invokeFill As New CodeMethodInvokeExpression

        Dim pointDeclaration As New CodeArrayCreateExpression(GetType(PointF))
        Dim typeDeclaration As New CodeArrayCreateExpression(GetType(System.Byte))

        Trace.WriteLineIf(Session.Tracing, "OpenPath.emit")

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

        'emit stroke
        Dim sStrokeName As String = Me.Stroke.emit(Me, declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings, Consolidated)

        With invokeStroke
            .Method.TargetObject = New CodeArgumentReferenceExpression("g")
            .Method.MethodName = "DrawPath"
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sStrokeName))
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, pathDeclaration.Name))
        End With

        RenderGDI.Add(invokeStroke)

    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts a GDIOpenPath to SVG XML outgoing SVG code.
    ''' </summary>
    ''' <param name="xmlDoc">The SVG document to append the open path to.</param>
    ''' <param name="defs">The definitions section of the SVG document.</param>
    ''' <param name="group">The group to append the XML spec of the open path to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub toXML(ByVal xmlDoc As System.Xml.XmlDocument, ByVal defs As System.Xml.XmlNode, ByVal group As System.Xml.XmlNode)

        Dim attr As Xml.XmlAttribute
        Dim node As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "path", String.Empty)


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

        attr = xmlDoc.CreateAttribute("d")
        attr.Value = sPathData
        node.Attributes.Append(attr)




        attr = xmlDoc.CreateAttribute("fill")
        attr.Value = "none"
        node.Attributes.Append(attr)

        Me.Stroke.toXML(xmlDoc, node)

        group.AppendChild(node)


    End Sub

#End Region

#Region "Base class Implementers"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the next valid name for a new GDIOpenPath
    ''' </summary>
    ''' <returns>A string containing the next valid name.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Function NextName() As String
        _NextnameID += 1
        Return "OpenPath" & _NextnameID
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a name used when displaying the class in the property browser.
    ''' </summary>
    ''' <value>A string containing the word "Open Path"</value>
    ''' -----------------------------------------------------------------------------   
    Public Overrides ReadOnly Property ClassName() As String
        Get
            Return "Open Path"
        End Get
    End Property

#End Region

#Region "Misc Members"


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

        _ResetPath = False

    End Sub

#End Region

#Region "Drawing related members"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders a GDIOpenPath for a specific draw mode.
    ''' </summary>
    ''' <param name="g">The graphics context to draw to.</param>
    ''' <param name="eDrawMode">The draw mode.  See the EnumDrawMode enumeration for more details.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub Draw(ByVal g As System.Drawing.Graphics, ByVal eDrawMode As EnumDrawMode)

        If _ResetPath Then
            resetPath()
        End If

        g.DrawPath(Stroke.Pen, _Path)

    End Sub
#End Region


#Region "Handles and hit testing"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles an explicit coordinate set on a point at the currently selected 
    ''' index.
    ''' </summary>
    ''' <param name="pt">The new location to place the point at.</param>
    ''' -----------------------------------------------------------------------------
    Public Overrides Sub handlePointSet(ByVal pt As Point)
        _Points(_HitIndex) = Point.op_Implicit(pt)

        _ResetPath = True
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the set of hit handles for the GDIOpenPath 
    ''' as an array of HitHandle objects
    ''' </summary>
    ''' <param name="fscale">The zoom factor of the surface the object resides on.</param>
    ''' <returns>An array of hit handles to use in hit tests.</returns>
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
    ''' Gets the index of the last hit handle.
    ''' </summary>
    ''' <value>An Int32 index location of the hit handle.</value>
    ''' -----------------------------------------------------------------------------
    <System.ComponentModel.Browsable(False)> _
 Public ReadOnly Property HitHandleIndex() As Int32
        Get
            If _HitIndex > -1 Then
                Return _HitIndex
            Else
                Return -1
            End If
        End Get

    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the point in the array of points that make up the path that matches the 
    ''' last hit handle.  This is the point at _HitIndex.  Since PointF is a structure,
    ''' to indicate that there is no point, an empty point is returned if there is no 
    ''' current hit index.
    ''' </summary>
    ''' <value>The point in the path that was last recorded as a hit or an empty point 
    ''' if no handles have been hit.</value>
    ''' -----------------------------------------------------------------------------
    <System.ComponentModel.Browsable(False)> _
    Public ReadOnly Property LastHitHandle() As PointF
        Get
            If _HitIndex > -1 Then
                Return _Points(_HitIndex)
            Else
                Return PointF.Empty
            End If
        End Get

    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a hit test over the handles in this object.  Returns the type of point 
    ''' hit if there is a collision, or the empty value (EnumDragHandles.eNone) if there 
    ''' is no collision.
    ''' </summary>
    ''' <param name="ptHit">The point to hit test against.</param>
    ''' <param name="fScale">The zoom factor of the surface this path is rendered on.</param>
    ''' <returns>The type of point hit (see EnumDragHandles).  This is a bit of a stretch, 
    ''' but it is here for consistency with other GDIObject hit tests.</returns>
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
    ''' Performs a hit test for the GDIOpenpath against a rectangle. 
    ''' </summary>
    ''' <param name="rect">A rectangle to compare intersection with the GDIOpenPath 
    ''' against.</param>
    ''' <returns>A Boolean indicating if the rectangle and the path intersect.</returns>
    ''' <remarks>Notice that the path is widened here.  This is because the inherited line 
    ''' for some reason will not perform a hit test correctly without increasing the thickness
    ''' of the path.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Overloads Overrides Function HitTest(ByVal rect As Rectangle) As Boolean


        Dim tempPath As New System.Drawing.Drawing2D.GraphicsPath(_Path.PathPoints, _Path.PathTypes)
        Dim tempPen As New System.Drawing.Pen(Color.Black, 0.1)

        tempPath.Widen(tempPen)

        Dim g As Graphics = Session.GraphicsManager.getTempGraphics()

        Dim r As New Region(rect)

        r.Intersect(tempPath)

        Dim bIntersect As Boolean = Not r.IsEmpty(g)

        tempPath.Dispose()
        tempPen.Dispose()
         r.Dispose()
        g.Dispose()



        Return bIntersect


    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if the GDIOpenPath intersects with a specific point.
    ''' </summary>
    ''' <param name="pt">The point to check intersection for.</param>
    ''' <param name="fScale">The zoom factor of the surface the GDIOpenPath lies on.</param>
    ''' <returns>A Boolean indicating if the given point at the given scale intersected 
    ''' the GDIOpenPath.</returns>
    ''' -----------------------------------------------------------------------------
    Public Overloads Overrides Function HitTest(ByVal pt As System.Drawing.PointF, ByVal fScale As Single) As Boolean
        Dim bHit As Boolean = False

        'This pen is used to draw the stroke on a path and hit test against  
        Dim p As New Pen(Color.Black, (Me.Stroke.Width * 6) / fScale)
 
        Dim bIntersectPath As Boolean = _Path.IsOutlineVisible(pt, p)


        If bIntersectPath OrElse HitTestHandles(pt, fScale) = EnumDragHandles.ePointHandle Then
            bHit = True
        Else
            bHit = False
        End If

        p.Dispose()

        Return bHit
        

    End Function

#End Region

#Region "Property Accessors"

 

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or Sets rotation - however - GDIOpenPaths are not allowed to be rotated at this 
    ''' time.  To enforce this the property is made non browsable and always returns 0 
    ''' for rotation.
    ''' </summary>
    ''' <value>A rotation value, which is ignored.</value>
    ''' -----------------------------------------------------------------------------
    ''' 
    <Browsable(False)> _
          Public Shadows Property Rotation() As Single
        Get
            Return 0
        End Get
        Set(ByVal Value As Single)

        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a Boolean value indicating whether to draw the stroke.  In GDIOpenPaths, 
    ''' the stroke is always drawn, so this property always returns true.
    ''' </summary>
    ''' <value>A Boolean indicating whether to draw the stroke or not.</value>
    ''' -----------------------------------------------------------------------------
    <Browsable(False)> _
    Public Overrides Property DrawStroke() As Boolean
        Get
            Return True
        End Get
        Set(ByVal Value As Boolean)
            _DrawStroke = True
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the array of pointF structures that make up the GDIOpenPath.
    ''' </summary>
    ''' <value>An array of pointf values.</value>
    ''' -----------------------------------------------------------------------------
    <System.ComponentModel.Browsable(False)> _
       Public ReadOnly Property Points() As PointF()
        Get
            Return _Points
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the types of points that make up the OpenPath
    ''' </summary>
    ''' <value>An array of Byte values corresponding to Drawing2D.PathPointType
    ''' values.</value>
    ''' -----------------------------------------------------------------------------
    <System.ComponentModel.Browsable(False)> _
     Friend ReadOnly Property Types() As Byte()
        Get
            Return _PointTypes
        End Get
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets a rectangle that tightly bounds the points in the GDIOpenPath.
    ''' </summary>
    ''' <value>A rectangle that bounds the points in the GDIOpenPath</value>
    ''' <remarks>There is a lot of activity taking place in the Set portion of this statement.
    ''' This is because by setting this rectangle, each point in the path must be scaled 
    ''' appropriately from its position in the current bounds to the new bounds.  
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    <System.ComponentModel.Browsable(False)> _
     Public Overrides Property Bounds() As Rectangle
        Get
            Dim rect As RectangleF = _Path.GetBounds()

            'Horizontal and vertical lines will not act as intended unless this is done.
            If rect.Height = 0 Then rect.Height = 1
            If rect.Width = 0 Then rect.Width = 1


            Return Rectangle.Round(rect)
        End Get
        Set(ByVal Value As Rectangle)
            Trace.WriteLineIf(Session.Tracing, "OpenPath.Bounds.Set: " & Value.ToString)

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


#End Region


#Region "Drag related members"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Starts a drag operation on a GDIOpenPath.
    ''' </summary>
    ''' <param name="ptOrigin">The origin of the drag.
    ''' (where the mouse pointer went down)</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub startDrag(ByVal ptOrigin As Point)
        ReDim _Offsets(_Points.Length)

        For i As Int32 = 0 To _Points.Length - 1
            Dim newpoint As New PointF(_Points(i).X - ptOrigin.X, _Points(i).Y - ptOrigin.Y)
            _Offsets(i) = newpoint
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates a drag operation on a GDIOpenpath
    ''' </summary>
    ''' <param name="dragPoint">The point at which the mouse currently resides.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub updateDrag(ByVal dragPoint As Point)

        For i As Int32 = 0 To _Points.Length - 1
            _Points(i).X = _Offsets(i).X + dragPoint.X
            _Points(i).Y = _Offsets(i).Y + dragPoint.Y
        Next

        _ResetPath = True

    End Sub

#End Region


#Region "Path Contruction members"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Recreates the graphics path object used to draw the GDIOpenPath to a surface.  
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub createPath()
        If Not _Path Is Nothing Then
            _Path.Dispose()
        End If

        _Path = New Drawing2D.GraphicsPath(_Points, _PointTypes)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the curvature being drawn to form a full Bezier curve. 
    ''' </summary>
    ''' <param name="ptCurvature">The end location the curvature point should reside at.</param>
    ''' <remarks>
    ''' If a curvature is being drawn, this makes sure all points needed to define the curve
    ''' are created.  In GDI+, a curve is composed of a series of point definitions 
    ''' marked as Bezier (see the code below). 
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub UpdateCurveSegment(ByVal ptCurvature As PointF)
        Trace.WriteLineIf(Session.Tracing, "Openpath.UpdateCurveSegment" & ptCurvature.ToString)

        If _Points.Length > 1 Then
            Dim ptlast As PointF = _Points(_Points.Length - 1)
            Dim pt2ndlast As PointF = _Points(_Points.Length - 2)


            _Points(_Points.Length - 1) = ptlast
            _PointTypes(_Points.Length - 1) = CByte(Drawing2D.PathPointType.Bezier)

            _Points(_Points.Length - 2) = ptCurvature
            _PointTypes(_Points.Length - 2) = CByte(Drawing2D.PathPointType.Bezier)


            _Points(_Points.Length - 3) = pt2ndlast
            _PointTypes(_PointTypes.Length - 3) = CByte(Drawing2D.PathPointType.Bezier)

            _ResetPath = True

        End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a closed path version of the GDIOpenPath.  This is used for when the user 
    ''' has drawn an open path and then links to the start point in the path indicating a desire 
    ''' to create a closed path instead of an open path.
    ''' </summary>
    ''' <returns>A GDIClosedPath object.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function GetClosedPath() As GDIClosedPath

        _PointTypes(_PointTypes.Length - 1) = CByte(Drawing2D.PathPointType.CloseSubpath) Or CByte(_PointTypes(_PointTypes.Length - 1))

        Return New GDIClosedPath(Me)
    End Function




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Starts the process of creating a curved segment.  This happens in response to the 
    ''' user executing a "click and hold" which pulls the curve segment away from the initial 
    ''' path and indicates a new curve segment should be created.
    ''' </summary>
    ''' <param name="ptcurvature">The point at which to begin creating the curvature.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub BeginCurveSegment(ByVal ptcurvature As PointF)
        Trace.WriteLineIf(Session.Tracing, "Openpath.BeginCurveSegment" & ptcurvature.ToString)

        If _Points.Length > 1 Then

            Dim ptlast As PointF = _Points(_Points.Length - 1)
            Dim pt2ndlast As PointF = _Points(_Points.Length - 2)


            ReDim Preserve _Points(_Points.Length + 1)
            ReDim Preserve _PointTypes(_PointTypes.Length + 1)

            _Points(_Points.Length - 1) = ptlast
            _PointTypes(_Points.Length - 1) = CByte(Drawing2D.PathPointType.Bezier3)

            _Points(_Points.Length - 2) = ptcurvature
            _PointTypes(_Points.Length - 2) = CByte(Drawing2D.PathPointType.Bezier3)


            _Points(_Points.Length - 3) = pt2ndlast
            _PointTypes(_PointTypes.Length - 3) = CByte(Drawing2D.PathPointType.Bezier3)

            _ResetPath = True

        End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the open path in tool mode.  Used by the GDIPlus Architect project to draw the 
    ''' tool without having to either expose Draw for the entire GDIObject hierarchy
    ''' or alternatively duplicate work in the GDIPlus Architect project
    ''' </summary>
    ''' <param name="g">System.Drawing.Graphics context to draw against.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub RenderTool(ByVal g As Graphics)
        Me.Draw(g, EnumDrawMode.eNormal)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a straight line segment to the current path.
    ''' </summary>
    ''' <param name="ptSnapped">The point at which to create the segment.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub addLineSegment(ByVal ptSnapped As Point) ', ByVal type As Drawing2D.GraphicsPath)
        Trace.WriteLineIf(Session.Tracing, "Openpath.addLineSegment " & ptSnapped.ToString)

        ReDim Preserve _Points(_Points.Length)
        ReDim Preserve _PointTypes(_PointTypes.Length)

        _Points(_Points.Length - 1) = Point.op_Implicit(ptSnapped)
        _PointTypes(_Points.Length - 1) = CByte(Drawing2D.PathPointType.Line)

        _ResetPath = True
    End Sub


#End Region


End Class
