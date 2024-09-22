Imports System.CodeDom
Imports System.ComponentModel

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIObject
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' GDIObject is a virtual class from which all objects on the design surface 
''' inherit.  As much shared functionality as possible has been placed inside this base
''' class.  Note that GDIObjects include images and textual objects as well as shapes.
''' There are further levels of abstraction down the inheritence heirarchy.
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable(), DefaultProperty("Name")> _
Public MustInherit Class GDIObject
    Implements IDisposable


#Region "Non Serialized Fields"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether this object has been disposed or not
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _Disposed As Boolean = False

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Offset during a drag operation.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
   Protected _DragOffset As New Point



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Used wrap rotation inside a graphics container
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _GraphicsContainer As Drawing2D.GraphicsContainer

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Used store the original matrix prior to a rotation operation.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _OriginalMatrix As Drawing2D.Matrix


#End Region

#Region "Local fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A rectangle defining the bounds of this particular GDIObject.  This rectangle 
    ''' is used extensively and rebuilt frequently.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _Bounds As Rectangle


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' For objects that are rotatable, the amount of rotation.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _Rotation As Single = 0

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of this object when emitted to code.  This corresponds to the name of the object 
    ''' in the property grid
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Name As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The scope of this object when it is emitted to code
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _Scope As EnumScope = EnumScope.Private

#End Region


#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new instance of a GDIObject.  This constructor is intended for 
    ''' deserialization
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Sub New()
        Me.Name = NextName()
        _Scope = Session.Settings.MemberScope
    End Sub

#End Region

#Region "Must Implement Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Virtual function inheritors must implement.  Responsible for rendering the object 
    ''' to a drawing surface.
    ''' </summary>
    ''' <param name="g">A System.Drawing graphics context to draw itself against.</param>
    ''' <param name="eDrawMode">The type of surface the object is drawing itself to.  Depending 
    ''' on the surface, the object will render itself differently, if needed.  For example, 
    ''' when drawing to a print preview surface, borders on text are suppressed, etc.
    ''' </param>
    ''' -----------------------------------------------------------------------------
    Friend MustOverride Sub Draw(ByVal g As Graphics, ByVal eDrawMode As EnumDrawMode)

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Virtual function inheritors must implement. Returns the next usable name for a new 
    ''' object placed on the drawing surface (Rectangle1, Rectangle2, ...)
    ''' </summary>
    ''' <returns>The next valid name for an object of this type</returns>
    ''' -----------------------------------------------------------------------------
    Protected MustOverride Function NextName() As String


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Virtual function inheritors must implement to render SVG to the surface.
    ''' </summary>
    ''' <param name="xmlDoc">The document the object should append its SVG to.</param>
    ''' <param name="defs">Definitions section node of the SVG document</param>
    ''' <param name="group">The group to which the object should append itself</param>
    ''' -----------------------------------------------------------------------------
    Friend MustOverride Sub toXML(ByVal xmlDoc As Xml.XmlDocument, ByVal defs As Xml.XmlNode, ByVal group As Xml.XmlNode)


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Virtual property inheritors must implement. Returns the name of the class used 
    ''' when displaying the GDIObject  in the property browser.  For example, GDIRect 
    ''' returns the word "Rectangle" instead of "GDIRect"
    ''' </summary>
    ''' <value>A friendly string to display in the property browser</value>
    ''' -----------------------------------------------------------------------------   
    <Browsable(False)> _
       Public MustOverride ReadOnly Property ClassName() As String


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Virtual definition that all GDIObjects must implement in order to emit code 
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
    Friend MustOverride Overloads Sub emit( _
    ByVal declarations As CodeDom.CodeTypeMemberCollection, _
    ByVal InitGraphics As CodeDom.CodeMemberMethod, _
    ByVal RenderGDI As CodeDom.CodeStatementCollection, _
    ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
    ByVal ExportSettings As ExportSettings, _
    ByVal Consolidated As ExportConsolidate)


#End Region

#Region "Common emit related functionality "


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the name that should be used for export purposes
    ''' </summary>
    ''' <returns>The name in string format</returns>
    ''' -----------------------------------------------------------------------------
    Friend Function ExportName() As String
        Return Session.DocumentManager.MemberPrefix & Me.Name
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a scope attribute for a specific object (Private, protected, etc)
    ''' </summary>
    ''' <returns>A MemberAttribute with the correct scope.</returns>
    ''' <remarks>Note that ExportOptions fully handles this.  This is because the 
    ''' final scope settings are determined by settings at the application level.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Protected Function getScope() As CodeDom.MemberAttributes


        Trace.WriteLineIf(Session.Tracing, "ExportOptions.TranslateScope" & _Scope.ToString)

        Select Case _Scope
            Case EnumScope.FriendFamilyInternal
                Return CodeDom.MemberAttributes.Assembly

            Case EnumScope.Private
                Return CodeDom.MemberAttributes.Private

            Case EnumScope.Protected
                Return CodeDom.MemberAttributes.Family

            Case EnumScope.ProtectedFriendFamilyInternal
                Return CodeDom.MemberAttributes.FamilyOrAssembly
            Case EnumScope.Public
                Return CodeDom.MemberAttributes.Public
        End Select

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits the rotate related fields as needed for rotated objects 
    ''' </summary>
    ''' <param name="declarations">The area in the class to append class
    '''  level declarations to</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub emitRotationDeclaration(ByVal declarations As CodeDom.CodeTypeMemberCollection)
        'Create Rectangle Initializer
        Dim rotatePointInit As New CodeObjectCreateExpression
        Dim rotateValueInit As New CodeObjectCreateExpression

        Dim rotatePointDeclaration As New CodeMemberField
        Dim rotateValueDeclaration As New CodeMemberField

        With rotatePointInit
            .CreateType = New CodeTypeReference(GetType(System.Drawing.PointF))
            .Parameters.Add(New CodePrimitiveExpression(Me.RotationPoint.X))
            .Parameters.Add(New CodePrimitiveExpression(Me.RotationPoint.Y))
        End With

        'Create Rectangle declaration
        With rotatePointDeclaration
            .InitExpression = rotatePointInit
            .Name = ExportName() & "RotationPoint"
            .Attributes = MemberAttributes.Private
            .Type = New CodeTypeReference(GetType(PointF))
        End With



        'Create Rectangle declaration
        With rotateValueDeclaration
            .InitExpression = New CodePrimitiveExpression(Me.Rotation)
            .Name = ExportName() & "Rotation"
            .Attributes = MemberAttributes.Private
            .Type = New CodeTypeReference(GetType(System.Single))
        End With

        declarations.Add(rotatePointDeclaration)
        declarations.Add(rotateValueDeclaration)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a begin rotate statement for the specified object
    ''' </summary>
    ''' <param name="RenderGDI">The RenderGDI method to append the begin rotate statement to</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub emitInvokeBeginRotation(ByVal RenderGDI As CodeDom.CodeStatementCollection)
        Dim invokeBegin As New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeThisReferenceExpression, "beginRotation"))
        With invokeBegin
            .Parameters.Add(New CodeArgumentReferenceExpression("g"))
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, ExportName() & "Rotation"))
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, ExportName() & "RotationPoint"))
        End With

        RenderGDI.Add(invokeBegin)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits an end rotate statement for the object.
    ''' </summary>
    ''' <param name="RenderGDI">The RenderGDI method to append the end rotate statement to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub emitInvokeEndRotation(ByVal RenderGDI As CodeDom.CodeStatementCollection)

        Dim invokeEnd As New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeThisReferenceExpression, "endRotation"))
        With invokeEnd
            .Parameters.Add(New CodeArgumentReferenceExpression("g"))
        End With

        RenderGDI.Add(invokeEnd)
    End Sub



#End Region

#Region "Misc Methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a scale transform on an object.  Inheritors can override this method 
    ''' to perform custom scaling as needed.  Note the original bounds argument is not the 
    ''' original bounds of the object, but the original bounds of the set of selected objects that 
    ''' this object belongs to.  There may be one or more other selected objects, and users expect 
    ''' that they "scale together" when a scale is performed and maintain position based upon the 
    ''' rectangle that bound the set of objects.
    ''' </summary>
    ''' <param name="rectOriginalBounds">The original bounds of the selected object prior to 
    ''' scaling.</param>
    ''' <param name="fprcW">Percentage width to scale as single (float)</param>
    ''' <param name="fprcH">Percentage height to scale as single (float)</param>
    ''' -----------------------------------------------------------------------------
    Friend Overridable Sub ScaleTransform(ByVal rectOriginalBounds As Rectangle, ByVal fprcW As Single, ByVal fprcH As Single)
        Trace.WriteLineIf(Session.Tracing, "Performing scale transformation")
        Trace.Indent()
        Trace.WriteLineIf(Session.Tracing, "Original scale" & rectOriginalBounds.ToString)
        Trace.WriteLineIf(Session.Tracing, "fprcW" & fprcW)
        Trace.WriteLineIf(Session.Tracing, "fprcW" & fprcH)
        Trace.Unindent()


        Dim scaledWidth As Int32 = CInt(rectOriginalBounds.Width * fprcW)
        Dim scaledHeight As Int32 = CInt(rectOriginalBounds.Height * fprcH)

        Dim originalCenterX As Int32 = rectOriginalBounds.X + (rectOriginalBounds.Width \ 2)
        Dim originalCenterY As Int32 = rectOriginalBounds.Y + (rectOriginalBounds.Height \ 2)

        Dim scaledRect As New Rectangle(originalCenterX - (scaledWidth \ 2), originalCenterY - scaledHeight \ 2, scaledWidth, scaledHeight)


        'The distance between each original.x and object.x should be scaled by fprcW and then added to scaledRect.X

        Dim distOriginalX As Int32 = Me.Bounds.X - rectOriginalBounds.X
        Dim distOriginalY As Int32 = Me.Bounds.Y - rectOriginalBounds.Y


        Me.Bounds = Rectangle.Round(New RectangleF(scaledRect.X + (distOriginalX * fprcW), scaledRect.Y + (distOriginalY * fprcH), Me.Bounds.Width * fprcW, Me.Bounds.Height * fprcH))

        '  Me.Bounds = New Rectangle(Me.Bounds.X, Me.Bounds.Y, CInt(Me.Bounds.Width * (wt / 100)), CInt(Me.Bounds.Height * (ht / 100)))
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a string representation of the object
    ''' </summary>
    ''' <returns>A string containing the name of the object.</returns>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function ToString() As String
        Return Me.Name
    End Function


#End Region

#Region "Serialization and Deserialization"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Attempts to deserialize an object on a page.  
    ''' </summary>
    ''' <param name="doc">The parent document this object belongs to</param>
    ''' <param name="pg">The page this object exists on.</param>
    ''' <returns>A Boolean indicating if the object was deserialized.  For the generic case, this 
    ''' is always true, but inheritors may have other restrictions.
    ''' For example, an object filled with a texturefill will return false if its associated 
    ''' image cannot be located.</returns>
    ''' -----------------------------------------------------------------------------
    Friend Overridable Function deserialize(ByVal doc As GDIDocument, ByVal pg As GDIPage) As Boolean
        '  Me.mParentCollection = pg.GDIObjects
        Trace.WriteLineIf(Session.Tracing, "Deserialing object")

        If doc.NameExists(Me, _Name) Then
            Trace.WriteLineIf(Session.Tracing, "Duplicate name found - " & _Name)
            Name = NextName()
        End If



        Return True
    End Function

#End Region

#Region "Drawing related code"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws an object in its selected state.  Selected is defined as being in the 
    ''' selected set of objects.  Selected objects may need to draw hit handles or highlights.
    ''' Inheritors can override this method to draw handle custom selected look and feel 
    ''' scenarios.
    ''' </summary>
    ''' <param name="g">Graphics to draw to</param>
    ''' <param name="fscale">The current zoom factor of the surface.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overridable Sub DrawSelectedObject(ByVal g As Graphics, ByVal fscale As Single)
        Dim bHandle As New SolidBrush(Session.Settings.DragHandleColor)
        Dim bCurve As New SolidBrush(Session.Settings.CurveColor)

        Dim mtx As System.Drawing.Drawing2D.Matrix
        Dim gcon As Drawing2D.GraphicsContainer

        Dim h As HitHandle() = getHandles(fscale)

        If _Rotation > 0 Then
            gcon = g.BeginContainer()
            mtx = g.Transform()
            mtx.RotateAt(_Rotation, RotationPoint, Drawing2D.MatrixOrder.Append)
            g.Transform = mtx
            mtx.Dispose()
        End If

        For Each hhandle As HitHandle In h
            If hhandle.HandleType = HitHandle.EnumHandletypes.eCurvePoint Then
                g.FillRectangle(bCurve, hhandle.HandleRect)
            Else
                g.FillRectangle(bHandle, hhandle.HandleRect)
            End If

        Next

        If _Rotation > 0 Then
            g.EndContainer(gcon)
        End If

        bHandle.Dispose()
        bCurve.Dispose()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws highlights upon an object.  GDI+ Architect offers an object to highlight objects 
    ''' as they are hit upon the surface in response to the mouse moving over them.  This is a 
    ''' different look than selected, which is when an object has actually been chosen to participate
    ''' in the selected set.
    ''' </summary>
    ''' <param name="g">A graphics context to draw the highlight to.</param>
    ''' <param name="fscale">The surfaces current zoom factor.</param>
    ''' <param name="highlightColor">Color set by the user to highlight objects with.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overridable Sub HighlightObject(ByVal g As Graphics, ByVal fscale As Single, ByVal highlightColor As Color)
        Dim bDragRect As New SolidBrush(highlightColor)
        Dim mtx As System.Drawing.Drawing2D.Matrix
        Dim gCon As Drawing2D.GraphicsContainer

        If _Rotation > 0 Then
            gCon = g.BeginContainer()
            mtx = g.Transform()
            mtx.RotateAt(_Rotation, RotationPoint)
            g.Transform = mtx
            mtx.Dispose()
        End If


        Dim h As HitHandle() = getHandles(fscale)

        For Each hhandle As HitHandle In h
            g.FillRectangle(bDragRect, hhandle.HandleRect)
        Next

        If _Rotation > 0 Then
            g.EndContainer(gCon)
        End If


        bDragRect.Dispose()

    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Starts a draw operation for the object.  For the generic case, this checks 
    ''' if a rotation container is needed and start this rotation prior to drawing.  
    ''' </summary>
    ''' <param name="g">A System.Drawing.Graphics context to draw against</param>
    ''' -----------------------------------------------------------------------------
    Protected Overridable Sub BeginDraw(ByVal g As Graphics)


        If _Rotation > 0 Then

            _GraphicsContainer = g.BeginContainer()
            _OriginalMatrix = g.Transform()
            _OriginalMatrix.RotateAt(_Rotation, RotationPoint)
            g.Transform = _OriginalMatrix
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends a draw operation for the current GDIObject. For the generic case, this 
    ''' ends any rotation containers.
    ''' </summary>
    ''' <param name="g">Graphics context to draw against</param>
    ''' -----------------------------------------------------------------------------
    Protected Overridable Sub EndDraw(ByVal g As Graphics)
        If _Rotation > 0 Then
            g.EndContainer(_GraphicsContainer)

            _OriginalMatrix.Dispose()
            _OriginalMatrix = Nothing
            _GraphicsContainer = Nothing
        End If

    End Sub

#End Region

#Region "Drag Related Code"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Called when an end drag operation is performed.  Inheritors can override this 
    ''' method when they are interested in responding to an EndDrag statement.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Overridable Sub endDrag()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Begins a drag statement for a GDI object.  Inheritors can override 
    ''' this method to perform custom drag implementations
    ''' </summary>
    ''' <param name="ptObject">The origin point where the drag is starting.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overridable Sub startDrag(ByVal ptObject As Point)
        _DragOffset = New Point(Bounds.X - ptObject.X, Bounds.Y - ptObject.Y)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the drag of the object given a current drag point.  Inheritors can override 
    ''' this method to perform custom placement based on a drag operation.
    ''' </summary>
    ''' <param name="dragPoint">The last recorded point for a specific object.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overridable Sub updateDrag(ByVal dragPoint As Point)

        dragPoint.Offset(_DragOffset.X, _DragOffset.Y)


        Me.Bounds = New Rectangle(dragPoint.X, dragPoint.Y, _
                                        Bounds.Width, Bounds.Height)

    End Sub



#End Region


#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the current rotation of a GDIObject.
    ''' </summary>
    ''' <value>A single (float) indicating the amount of rotation from the origin.</value>
    ''' -----------------------------------------------------------------------------
    Public Overridable Property Rotation() As Single
        Get
            Return _Rotation
        End Get
        Set(ByVal Value As Single)
            If Value >= 0 AndAlso Value < 360 Then
                _Rotation = Value
            End If
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the name of the GDIObject.  The name is used when code is emitted to 
    ''' give a unique declaration, as well as shows in the property grid inside GDI+ Architect.
    ''' </summary>
    ''' <value>The name of the object</value>
    ''' <remarks>Notice that the Set portion of the property does quite a bit of work.
    ''' All of this extra code is to guarantee that the name is unique across other objects within 
    ''' the GDIDocument.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    <Category(""), System.ComponentModel.ParenthesizePropertyName(True), Description("Name of the graphic item"), MergablePropertyAttribute(False)> _
    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal Value As String)

            Dim tempName As String = Value.Replace(" ", "")

            If Not _Name = tempName Then
                If tempName.Trim = String.Empty Then
                    Throw New FormatException("Invalid field name")
                Else
                    If Session.DocumentManager.NameConflictExists(tempName) Then
                        Name = NextName()
                    Else
                        _Name = tempName
                    End If
                End If


            End If

        End Set
    End Property




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the scope of the declaration of the object when code is emitted for it.
    ''' </summary>
    ''' <value>An EnumScope value.  See EnumScope for more information.</value>
    '''  -----------------------------------------------------------------------------
    <DefaultValue(GetType(EnumScope), "Private"), Browsable(True), Description("Scope of the primary attribute generated for this object.  For text, the attribute is the associated string. For paths, the path.  For lines, the points, and for rectangles and ellipses, the bounding rectangle.")> _
    Public Property Scope() As EnumScope
        Get
            Return _Scope
        End Get
        Set(ByVal Value As EnumScope)
            _Scope = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     ''' Returns an array of points representing the rectangular bounds of the object after rotation.
    ''' </summary>
    ''' <value>An array of points bounding the object after rotation</value>
    ''' -----------------------------------------------------------------------------
    <Browsable(False)> _
    Public Overridable ReadOnly Property RotatedBoundPoints() As Point()
        Get
            Dim mtx As New Drawing2D.Matrix

            Dim pts As Point() = BoundPoints()
            mtx.RotateAt(_Rotation, RotationPoint)

            mtx.TransformPoints(pts)
            mtx.Dispose()

            Return pts

        End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the current bounds of the object (a rectangle that surrounds the 
    ''' object).  Inheritors can optionally override this method to perform custom 
    ''' operations.
    ''' </summary>
    ''' <value>A rectangle with the bounds of the object.</value>
    ''' -----------------------------------------------------------------------------
    '''  
    <Category(""), Description("A rectangle that bounds the graphic")> _
       Public Overridable Property Bounds() As System.Drawing.Rectangle
        <DebuggerStepThrough()> Get
            Return _Bounds
        End Get
        <DebuggerStepThrough()> Set(ByVal Value As System.Drawing.Rectangle)
            _Bounds = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the point about which the object should be rotated.  This is typically 
    ''' the center of the object, but inheritors can override this if there is a more 
    ''' appropriate point to rotate the object about.
    ''' </summary>
    ''' <value>A point about which to rotate the object.</value>
    ''' -----------------------------------------------------------------------------
    <Browsable(False)> _
   Public Overridable ReadOnly Property RotationPoint() As PointF
        Get
            Return New PointF(_Bounds.X + _Bounds.Width \ 2, _Bounds.Y + _Bounds.Height \ 2)
        End Get
    End Property




#End Region

#Region "Hit Test Related Code"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Allows inheritors to response to explicit point set operations (where a point 
    ''' is selected and dragged to another location).
    ''' </summary>
    ''' <param name="pt">The new location to set the point to.</param>
    ''' -----------------------------------------------------------------------------
    Public Overridable Sub handlePointSet(ByVal pt As Point)

    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the points in the bounds rectangle as an array of points.
    ''' </summary>
    ''' <returns>An array of points containing the four corner points of the bounds of the object.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function BoundPoints() As Point()
        Dim ptArr() As Point = {New Point(_Bounds.X, _Bounds.Y), New Point(_Bounds.Right, _Bounds.Y), New Point(_Bounds.Right, _Bounds.Bottom), New Point(_Bounds.Left, _Bounds.Bottom)}
        Return ptArr
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a sit of hit handles.  Hit handles are usually pinned on the bounding box 
    ''' around the object, but inheritors can override this method for more complex hit handle 
    ''' operations.
    ''' </summary>
    ''' <param name="fscale">The current surface scale.</param>
    ''' <returns>An array of hithandles that hit testing on handles can be performed over.</returns>
    ''' -----------------------------------------------------------------------------
    Friend Overridable Function getHandles(ByVal fscale As Single) As HitHandle()
        Dim hw As Single = 6 / fscale
        Dim hwhalf As Single = hw / 2

        Dim hHandles As HitHandle() = New HitHandle() { _
        New HitHandle(New RectangleF(_Bounds.X - hwhalf, _Bounds.Y + (_Bounds.Height \ 2) - hwhalf, hw, hw), HitHandle.EnumHandletypes.eNormal), _
        New HitHandle(New RectangleF(_Bounds.X - hwhalf, _Bounds.Y - hwhalf, hw, hw), HitHandle.EnumHandletypes.eNormal), _
        New HitHandle(New RectangleF(_Bounds.X + (_Bounds.Width \ 2) - hwhalf, _Bounds.Y - hwhalf, hw, hw), HitHandle.EnumHandletypes.eNormal), _
        New HitHandle(New RectangleF(_Bounds.X + _Bounds.Width - hwhalf, _Bounds.Y - hwhalf, hw, hw), HitHandle.EnumHandletypes.eNormal), _
        New HitHandle(New RectangleF(_Bounds.X + _Bounds.Width - hwhalf, _Bounds.Y + (_Bounds.Height \ 2) - hwhalf, hw, hw), HitHandle.EnumHandletypes.eNormal), _
        New HitHandle(New RectangleF(_Bounds.X + _Bounds.Width - hwhalf, _Bounds.Y + _Bounds.Height - hwhalf, hw, hw), HitHandle.EnumHandletypes.eNormal), _
        New HitHandle(New RectangleF(_Bounds.X + (_Bounds.Width \ 2) - hwhalf, _Bounds.Y + _Bounds.Height - hwhalf, hw, hw), HitHandle.EnumHandletypes.eNormal), _
        New HitHandle(New RectangleF(_Bounds.X - hwhalf, _Bounds.Y + _Bounds.Height - hwhalf, hw, hw), HitHandle.EnumHandletypes.eNormal) _
        }

        Return hHandles


    End Function


    Private Function testSingleHandle(ByVal mtx As Drawing2D.Matrix, ByVal rect As RectangleF, ByVal ptHit As PointF) As Boolean
        Dim bHit As Boolean = False

        Dim gPath As New Drawing2D.GraphicsPath

        gPath.AddRectangle(rect)
        gPath.Transform(mtx)


        If gPath.IsVisible(ptHit) Then
            bHit = True
        End If

        gPath.Dispose()

        Return bHit
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a hit test on the handles of rotated objects
    ''' </summary>
    ''' <param name="ptHit">The point to hit test against</param>
    ''' <param name="fScale">The current zoom factor of the surface</param>
    ''' <returns>The drag handle hit, or EnumDragHandles.None if no handle was hit</returns>
    ''' -----------------------------------------------------------------------------
    Private Function HitTestRotatedHandles(ByVal ptHit As PointF, ByVal fScale As Single) As EnumDragHandles
        Dim hHandles As HitHandle() = getHandles(fScale)

        Dim mtx As New System.Drawing.Drawing2D.Matrix

        'Transform  by the rotation point of the object
        mtx.RotateAt(_Rotation, Me.RotationPoint)


        For i As Int32 = 0 To hHandles.Length - 1
            Dim bHit As Boolean = testSingleHandle(mtx, hHandles(i).HandleRect, ptHit)

            If bHit Then
                mtx.Dispose()

                Dim x As EnumDragHandles
                x = CType(i, EnumDragHandles)
                Return x
            End If

        Next

        mtx.Dispose()

        Return EnumDragHandles.eNone

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a hit test over an objects handles.  The handles are the points the user clicks to resize 
    ''' objects along the surface.  For example, a rectangle GDIObject would simply have the hit handles along its 
    ''' edges, whereas a complex curvature would have a collection of points along the curve.
    ''' </summary>
    ''' <param name="ptHit">A point to test if it intersects with the hit rectangle</param>
    ''' <param name="fScale">The current zoom factor on the interface.</param>
    ''' <returns>An enumeration value indicating which type of handle was hit.  See EnumDragHandles for more information. </returns>
    ''' -----------------------------------------------------------------------------
    Public Overridable Function HitTestHandles(ByVal ptHit As PointF, ByVal fScale As Single) As EnumDragHandles

        'If there is rotation, use the more expensive rotated based hit testing code.
        'Otherwise use the simple code.
        If _Rotation > 0 Then
            Return HitTestRotatedHandles(ptHit, fScale)
        Else


            Dim hHandles() As HitHandle = getHandles(fScale)

            For i As Int32 = 0 To hHandles.Length - 1
                If hHandles(i).Contains(ptHit) Then
                    Return CType(i, EnumDragHandles)
                End If

            Next

            Return EnumDragHandles.eNone
        End If

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a hit test between a GDIObject and a specific rectangle.
    ''' </summary>
    ''' <param name="rectObject">The rectangle to check for an overlap with the object for.  This 
    ''' rectangle is express in object coordinates.</param>
    ''' <returns>A Boolean indicating whether the rectangle and the hit test object intersected</returns>
    ''' -----------------------------------------------------------------------------
    Public Overridable Function HitTest(ByVal rectObject As Rectangle) As Boolean
        If _Rotation > 0 Then

            Dim mtx As New System.Drawing.Drawing2D.Matrix
            Dim g As Graphics = Session.GraphicsManager.getTempGraphics
            Dim r As Region
            Dim pth As New Drawing2D.GraphicsPath(Drawing2D.FillMode.Alternate)
 
            mtx.RotateAt(_Rotation, RotationPoint)

            With pth
                .AddRectangle(Me.Bounds)
                .Transform(mtx)
            End With

            r = New Region(pth)
            r.Intersect(rectObject)

            Dim bIntersect As Boolean = Not r.IsEmpty(g)

            mtx.Dispose()
            pth.Dispose()
            r.Dispose()
            g.Dispose()

            Return bIntersect

        Else
            Return rectObject.IntersectsWith(Me.Bounds)

        End If


    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a hit test for the object given a current zoom factor and the hit test point.
    ''' </summary>
    ''' <param name="ptHit">The hit test point</param>
    ''' <param name="fScale">The current zoom scale</param>
    ''' <returns>A Boolean indicating if the object was hit or not.</returns>
    ''' -----------------------------------------------------------------------------
    Public Overridable Function HitTest(ByVal ptHit As PointF, ByVal fScale As Single) As Boolean
        Dim hw As Single = 6 / fScale

        Dim inflatedRect As RectangleF = New RectangleF(_Bounds.X, _Bounds.Y, _Bounds.Width, _Bounds.Height)
        inflatedRect.Inflate(hw, hw)

        If _Rotation > 0 Then
            Dim mtx As New System.Drawing.Drawing2D.Matrix
            mtx.RotateAt(_Rotation, RotationPoint)
            Dim pth As New Drawing2D.GraphicsPath(Drawing2D.FillMode.Alternate)

            pth.AddRectangle(inflatedRect)
            pth.Transform(mtx)

            Dim r As New Region(pth)

            Dim bVisible As Boolean = r.IsVisible(ptHit)

            pth.Dispose()
            r.Dispose()
            mtx.Dispose()

            Return bVisible
        Else
            Return inflatedRect.Contains(ptHit)

        End If

    End Function

#End Region

#Region "Cleanup code"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of a GDIObject
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the object.  At the GDIObject level this doesn't do anything, but 
    ''' inheritors can override this method to perform custom disposal.
    ''' </summary>
    ''' <param name="disposing">Whether unmanaged resources are being disposed or not.</param>
    ''' -----------------------------------------------------------------------------
    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        _Disposed = True
    End Sub


#End Region

End Class
