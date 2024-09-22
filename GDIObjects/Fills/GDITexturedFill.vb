Imports System.CodeDom

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDITexturedFill
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A textured fill for fillable GDIObjects.  Texture fills paint objects using 
''' a graphic pattern with a specific winding mode and start point for the pattern.
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public Class GDITexturedFill
    Inherits GDIFill




#Region "Local Fields"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A string containing the file name portion of the path to the image source to use 
    ''' as a fill.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _ImageSource As String = String.Empty



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The wrap mode used to fill this texture
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _TextureWrapMode As Drawing2D.WrapMode = Drawing2D.WrapMode.Tile
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The type of runtime source - See EnumLinkType for more information.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _RunTimeSource As EnumLinkType = EnumLinkType.Embedded

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A path to the image source used to render the texture.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Path As String


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' X component of a transformation, if any, applied to this texture fill.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _XTransform As Single = 0.0!
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Y component of a transformation, if any, applied to this texture fill.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _YTransform As Single = 0.0!

#End Region


#Region "non Serialized Fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A non serialized bitmap used to render the texture inside GDI+ Architect
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _Image As Bitmap

#End Region

#Region "Constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a textured fill given a parent shape and another textured fill to base 
    ''' the instance off of.
    ''' </summary>
    ''' <param name="parent">The parent shape this fill is responsible for filling.</param>
    ''' <param name="fill">The texture fill to copy properties from.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal fill As GDITexturedFill)

        Trace.WriteLineIf(Session.Tracing, "TextureFill.New.Clone")

        With Me
            _Image = New Bitmap(100, 100)
            ImageSource = fill.ImageSource
            _TextureWrapMode = fill.WrapMode
            RuntimeSource = fill.RuntimeSource

        End With

        UpdateFill()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a textured fill given a parent fill and a source image to use as the 
    ''' texture.
    ''' </summary>
    ''' <param name="parent">The parent shape this fill is responsible for filling.</param>
    ''' <param name="sImageSource">A path to a source image to use to fill with.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal sImageSource As String)
        MyBase.New()
        Trace.WriteLineIf(Session.Tracing, "TextureFill.New")

        With Me
            ImageSource = sImageSource
            RuntimeSource = Session.Settings.TextureLinkType
        End With

        UpdateFill()
    End Sub
#End Region

#Region "Code Generation Related Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if Two texture fills are equivalent for the purposes of export.
    ''' </summary>
    ''' <param name="fill1">The first texture fill to compare.</param>
    ''' <param name="fill2">The second texture fill to compare.</param>
    ''' <returns>A Boolean indicating if the two textures are equiv.</returns>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function ExportEquality(ByVal fill1 As GDITexturedFill, ByVal fill2 As GDITexturedFill) As Boolean

        Return op_Equality(fill1, fill2) AndAlso fill1._XTransform = fill2._XTransform AndAlso fill1._YTransform = fill2._YTransform

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a consolidated textured fill to code.
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

        Dim fillDeclaration As New CodeMemberField
        Dim BitmapDeclaration As New CodeMemberField
        Dim fillDispose As New CodeDom.CodeMethodInvokeExpression
        Dim BitmapDispose As New CodeDom.CodeMethodInvokeExpression

        Dim fillInit As New CodeObjectCreateExpression
        Dim bitmapInit As New CodeMethodInvokeExpression

        Dim assignBitmap As New CodeAssignStatement
        Dim assignBrush As New CodeAssignStatement

        Dim drawModeParam As New CodeDom.CodeFieldReferenceExpression

        Trace.WriteLineIf(Session.Tracing, "TextureFill.emit")

        With BitmapDeclaration
            .Name = sSharedName & "BrushTexture"
            .Type = New CodeTypeReference(GetType(System.Drawing.Bitmap))
            .Attributes = MemberAttributes.Private
            .InitExpression = Nothing
        End With

        'Create the member variable for the stroke
        With fillDeclaration
            .Name = sSharedName & "Brush"
            .Type = New CodeTypeReference(GetType(System.Drawing.TextureBrush))
            .Attributes = MemberAttributes.Private
            .InitExpression = Nothing
        End With


        Select Case Me.RuntimeSource
            Case EnumLinkType.AbsolutePath
                bitmapInit = New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeThisReferenceExpression, "ImageFromAbsolutePath"))
            Case EnumLinkType.Embedded
                bitmapInit = New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeThisReferenceExpression, "ImageFromEmbedded"))
            Case EnumLinkType.RelativePath
                bitmapInit = New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeThisReferenceExpression, "ImageFromRelativePath"))
        End Select

        bitmapInit.Parameters.Add(New CodePrimitiveExpression(Me._Path))

        assignBitmap.Left = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, BitmapDeclaration.Name)
        assignBitmap.Right = bitmapInit

        assignBrush.Left = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, fillDeclaration.Name)
        assignBrush.Right = fillInit


        Dim translateBrush As New CodeDom.CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, fillDeclaration.Name), "TranslateTransform"))
        translateBrush.Parameters.Add(New CodePrimitiveExpression(_XTransform))
        translateBrush.Parameters.Add(New CodePrimitiveExpression(_YTransform))

        With InitGraphics
            .Statements.Add(assignBitmap)
            .Statements.Add(assignBrush)
        End With

        With drawModeParam
            .TargetObject = New CodeTypeReferenceExpression(GetType(System.Drawing.Drawing2D.WrapMode))
            .FieldName = System.Enum.GetName(GetType(System.Drawing.Drawing2D.WrapMode), Me.WrapMode)
        End With

        With fillInit
            .CreateType = New CodeTypeReference(GetType(System.Drawing.TextureBrush))
            .Parameters.Add(New CodeDom.CodeFieldReferenceExpression(New CodeThisReferenceExpression, BitmapDeclaration.Name))
            .Parameters.Add(drawModeParam)
        End With

        With fillDispose
            .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, fillDeclaration.Name)
            .Method.MethodName = "Dispose"
        End With


        With BitmapDispose
            .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, BitmapDeclaration.Name)
            .Method.MethodName = "Dispose"
        End With

        With DisposeGDI
            .Statements.Add(BitmapDispose)
            .Statements.Add(fillDispose)
        End With

        With declarations
            .Add(BitmapDeclaration)
            .Add(fillDeclaration)
        End With

        InitGraphics.Statements.Add(translateBrush)

        Return fillDeclaration.Name

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a textured fill to code.
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


        Trace.WriteLineIf(Session.Tracing, "TextureFill.emit.Consolidated")


        If Consolidated.hasFillMatch(Me) And Me.Consolidate Then
            Return Consolidated.getFillName(Me)
        Else


            Dim fillDeclaration As New CodeMemberField
            Dim BitmapDeclaration As New CodeMemberField
            Dim fillDispose As New CodeDom.CodeMethodInvokeExpression
            Dim BitmapDispose As New CodeDom.CodeMethodInvokeExpression

            Dim fillInit As New CodeObjectCreateExpression
            Dim bitmapInit As New CodeMethodInvokeExpression

            Dim assignBitmap As New CodeAssignStatement
            Dim assignBrush As New CodeAssignStatement

            Dim drawModeParam As New CodeDom.CodeFieldReferenceExpression


            With BitmapDeclaration
                .Name = obj.ExportName & "BrushTexture"
                .Type = New CodeTypeReference(GetType(System.Drawing.Bitmap))
                .Attributes = MemberAttributes.Private
                .InitExpression = Nothing
            End With

            'Create the member variable for the stroke
            With fillDeclaration
                .Name = obj.ExportName & "Brush"
                .Type = New CodeTypeReference(GetType(System.Drawing.TextureBrush))
                .Attributes = MemberAttributes.Private
                .InitExpression = Nothing
            End With


            Select Case Me.RuntimeSource
                Case EnumLinkType.AbsolutePath
                    bitmapInit = New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeThisReferenceExpression, "ImageFromAbsolutePath"))
                Case EnumLinkType.Embedded
                    bitmapInit = New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeThisReferenceExpression, "ImageFromEmbedded"))
                Case EnumLinkType.RelativePath
                    bitmapInit = New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeThisReferenceExpression, "ImageFromRelativePath"))
            End Select

            bitmapInit.Parameters.Add(New CodePrimitiveExpression(Me._Path))

            assignBitmap.Left = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, BitmapDeclaration.Name)
            assignBitmap.Right = bitmapInit

            assignBrush.Left = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, fillDeclaration.Name)
            assignBrush.Right = fillInit


            Dim translateBrush As New CodeDom.CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, fillDeclaration.Name), "TranslateTransform"))
            translateBrush.Parameters.Add(New CodePrimitiveExpression(_XTransform))
            translateBrush.Parameters.Add(New CodePrimitiveExpression(_YTransform))

            With InitGraphics
                .Statements.Add(assignBitmap)
                .Statements.Add(assignBrush)
            End With

            With drawModeParam
                .TargetObject = New CodeTypeReferenceExpression(GetType(System.Drawing.Drawing2D.WrapMode))
                .FieldName = System.Enum.GetName(GetType(System.Drawing.Drawing2D.WrapMode), Me.WrapMode)
            End With

            With fillInit
                .CreateType = New CodeTypeReference(GetType(System.Drawing.TextureBrush))
                .Parameters.Add(New CodeDom.CodeFieldReferenceExpression(New CodeThisReferenceExpression, BitmapDeclaration.Name))
                .Parameters.Add(drawModeParam)
            End With

            With fillDispose
                .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, fillDeclaration.Name)
                .Method.MethodName = "Dispose"
            End With


            With BitmapDispose
                .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, BitmapDeclaration.Name)
                .Method.MethodName = "Dispose"
            End With

            With DisposeGDI
                .Statements.Add(BitmapDispose)
                .Statements.Add(fillDispose)
            End With

            With declarations
                .Add(BitmapDeclaration)
                .Add(fillDeclaration)
            End With

            InitGraphics.Statements.Add(translateBrush)

            Return fillDeclaration.Name
        End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts the textured fill to an XML source for SVG display.
    ''' </summary>
    ''' <param name="xmlDoc">See base class</param>
    ''' <param name="defs">See base class</param>
    ''' <param name="parent">See base class</param>
    ''' -----------------------------------------------------------------------------
    Public Overrides Sub toXML(ByVal xmlDoc As Xml.XmlDocument, ByVal defs As Xml.XmlNode, ByVal parent As Xml.XmlNode)

        '<filter id="MyFilter" filterUnits="userSpaceOnUse" x="0" y="0" width="12" height="12">

        '<feImage width="40" height="30" xlink:href="0.jpg" result="image" />
        '<feTile in="image" result="tile" />
        '</filter>

        '<pattern id="Rect1Brush" patternUnits="userSpaceOnUse" viewport="50 50">
        '<rect id="Rect1" x="0" y="0" width="50" height="50" filter="url(#MyFilter)"/>
        '</pattern>

        Dim img As Image = Me.GetImage
        Dim attr As Xml.XmlAttribute

        Dim nodeFilter As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "filter", String.Empty)
        Dim nodePattern As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "pattern", String.Empty)
        Dim nodePatternRect As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "rect", String.Empty)


        attr = xmlDoc.CreateAttribute("x")
        attr.Value = "0"
        nodePatternRect.Attributes.Append(attr)


        attr = xmlDoc.CreateAttribute("y")
        attr.Value = "0"
        nodePatternRect.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("width")
        attr.Value = img.Width.ToString
        nodePatternRect.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("height")
        attr.Value = img.Height.ToString
        nodePatternRect.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("filter")
        attr.Value = String.Format("url(#{0})", parent.Attributes("id").Value & "BrushFilter")
        nodePatternRect.Attributes.Append(attr)

        nodePattern.AppendChild(nodePatternRect)



        attr = xmlDoc.CreateAttribute("id")
        attr.Value = parent.Attributes("id").Value & "Brush"
        nodePattern.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("id")
        attr.Value = parent.Attributes("id").Value & "BrushFilter"
        nodeFilter.Attributes.Append(attr)




        attr = xmlDoc.CreateAttribute("filterUnits")
        attr.Value = "userSpaceOnUse"
        nodeFilter.Attributes.Append(attr)

        'Image filter
        Dim nodeImage As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "feImage", String.Empty)

        attr = xmlDoc.CreateAttribute("width")
        attr.Value = img.Width.ToString
        nodeImage.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("height")
        attr.Value = img.Height.ToString
        nodeImage.Attributes.Append(attr)


        Dim ele As Xml.XmlElement = CType(nodeImage, Xml.XmlElement)

        ele.SetAttribute("href", "http://www.w3.org/1999/xlink", Me.ImageSource)


        attr = xmlDoc.CreateAttribute("result")
        attr.Value = "image"
        nodeImage.Attributes.Append(attr)

        'Tile filter
        Dim nodeTile As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "feTile", String.Empty)
        attr = xmlDoc.CreateAttribute("in")
        attr.Value = "image"
        nodeTile.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("result")
        attr.Value = "tile"
        nodeTile.Attributes.Append(attr)

        nodeFilter.AppendChild(nodeImage)
        nodeFilter.AppendChild(nodeTile)


        attr = xmlDoc.CreateAttribute("fill")
        attr.Value = String.Format("url(#{0})", parent.Attributes("id").Value & "Brush")
        parent.Attributes.Append(attr)

        defs.AppendChild(nodeFilter)
        defs.AppendChild(nodePattern)

    End Sub


#End Region


#Region "Misc Methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a string representation of this object.
    ''' </summary>
    ''' <returns>A string representation of the object</returns>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function ToString() As String
        Return "Textured Fill"
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Initializes the path to a texture resource based upon a chosen EnumLinkType.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub setupPath()
        If _ImageSource.Length > 0 Then

            Trace.WriteLineIf(Session.Tracing, "TextureFill.setupPath")
            Dim fInfo As New System.IO.FileInfo(_ImageSource)

            Select Case Me.RuntimeSource
                Case EnumLinkType.AbsolutePath
                    Me.Path = Session.Settings.TextureAbsolutePath & "\" & fInfo.Name

                Case EnumLinkType.RelativePath
                    If Session.Settings.TextureRelativePath = String.Empty Then
                        Me.Path = fInfo.Name
                    Else
                        Me.Path = Session.Settings.TextureRelativePath & "\" & fInfo.Name
                    End If
                Case EnumLinkType.Embedded
                    If Session.DocumentManager.HaveCurrentDocument Then
                        Me.Path = Session.DocumentManager.RootNameSpace & "." & fInfo.Name
                    End If
            End Select
        End If
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if two texture fills are equal (have equal properties)
    ''' </summary>
    ''' <param name="fill1">The first texture fill to compare</param>
    ''' <param name="fill2">The second texture fill to compare</param>
    ''' <returns>A Boolean indicating if the two textures are equal</returns>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function op_Equality(ByVal fill1 As GDITexturedFill, ByVal fill2 As GDITexturedFill) As Boolean
        Dim bEquality As Boolean = True

        bEquality = bEquality And fill1._ImageSource = fill2._ImageSource
        bEquality = bEquality And fill1._TextureWrapMode = fill2._TextureWrapMode
        bEquality = bEquality And fill1._RunTimeSource = fill2._RunTimeSource
        bEquality = bEquality And fill1._Path = fill2._Path

        Return bEquality
    End Function




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns an image used to fill the texture.
    ''' </summary>
    ''' <returns>An image used to fill the texture.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function GetImage() As System.Drawing.Image


        Try
            If _Image Is Nothing Then
                _Image = Session.GraphicsManager.ImageFromAbsolutePath(_ImageSource)
            End If

            Return _Image

        Catch ex As System.Exception
            Trace.WriteLineIf(Session.Tracing, "TextureFill.GetImage.Exception" & ex.Message)

            Throw New System.FormatException("Invalid Property Value")
        End Try

    End Function
#End Region

#Region "Property Accessors"






    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the Path to the image used in this texture fill.  Note that this property
    ''' must be used in conjunction with the EnumLinkType value to determine how it should 
    ''' be interpreted.
    ''' </summary>
    ''' <value>A path to the resource.</value>
    ''' -----------------------------------------------------------------------------
    Public Property Path() As String
        Get
            Return _Path
        End Get
        Set(ByVal Value As String)
            _Path = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the X Offset value of a fill transformation.
    ''' </summary>
    ''' <value>A Single (float) of the amount to offset the X position by.</value>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.Description("X Offset of the texture.  Turn off TrackFill to manually set this.")> _
    Public Property XOffSet() As Single
        Get
            Return _XTransform
        End Get
        Set(ByVal Value As Single)
            Trace.WriteLineIf(Session.Tracing, "TextureFill.XOffSet.Set" & Value.ToString)

            _XTransform = Value
            UpdateFill()
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the Y Offset value of a fill transformation.
    ''' </summary>
    ''' <value>A Single (float) of the amount to offset the Y position by.</value>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.Description("Y Offset of the texture.  Turn off TrackFill to manually set this.")> _
 Public Property YOffSet() As Single
        Get
            Return _YTransform
        End Get
        Set(ByVal Value As Single)
            Trace.WriteLineIf(Session.Tracing, "TextureFill.YOffSet.Set" & Value.ToString)

            _YTransform = Value
            UpdateFill()
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the Runtimesource type of this TextureFill.  See EnumLinkType for more 
    ''' information.
    ''' </summary>
    ''' <value>An EnumLinkType enumeration value.</value>
    ''' <remarks>Note when this is set it cascades some property changes using the private 
    ''' setupPath method.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Property RuntimeSource() As EnumLinkType
        Get
            Return _RunTimeSource
        End Get
        Set(ByVal Value As EnumLinkType)
            Trace.WriteLineIf(Session.Tracing, "TextureFill.Runtimesource.Set: " & Value.ToString)

            _RunTimeSource = Value

            setupPath()

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the Drawing2D wrapmode of this texture fill.
    ''' </summary>
    ''' <value>a Drawing2D wrapmode</value>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.Description("Texture Wrapping Mode")> _
    Public Property WrapMode() As Drawing2D.WrapMode
        Get
            Return _TextureWrapMode
        End Get
        Set(ByVal Value As Drawing2D.WrapMode)
            Trace.WriteLineIf(Session.Tracing, "TextureFill.WrapMode.Set: " & Value.ToString)

            _TextureWrapMode = Value
            UpdateFill()
        End Set
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the ImageSource (the location of an image file used to render the texture)
    ''' </summary>
    ''' <value>A string value containing a path to the image resource.</value>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.Description("Image Source location"), ComponentModel.EditorAttribute(GetType(GDIImage.ImageFileChooser), GetType(Drawing.Design.UITypeEditor))> _
    Public Property ImageSource() As String
        Get
            Return _ImageSource
        End Get
        Set(ByVal Value As String)
            Trace.WriteLineIf(Session.Tracing, "TextureFill.ImageSource.Set: " & Value.ToString)

            'Verify it's a bitmap
            Dim tempBMP As Bitmap
            If Value <> _ImageSource Then
                Try
                    Dim fInfo As New System.IO.FileInfo(Value)
                    If fInfo.Exists() Then
                        tempBMP = New Bitmap(Value)
                        tempBMP.Dispose()
                        tempBMP = Nothing

                        _ImageSource = Value
                        setupPath()

                        If Not _Image Is Nothing Then
                            _Image.Dispose()
                            _Image = Nothing
                        End If

                    Else
                        Throw New System.FormatException("Invalid Property Value")
                    End If

                Catch ex As Exception
                    Trace.WriteLineIf(Session.Tracing, "TextureFill.ImageSource.Exception: " & ex.Message.ToString)

                    MsgBox("Invalid Property Value")
                Finally
                    If Not tempBMP Is Nothing Then
                        tempBMP.Dispose()
                        tempBMP = Nothing
                    End If
                End Try
            End If


            UpdateFill()

        End Set
    End Property
#End Region


#Region "Base class Implementers"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles an update of the fill's parent (the shape being filled by the fill)
    ''' </summary>
    ''' <param name="obj">The parent object to fill with.</param>
    ''' -----------------------------------------------------------------------------
    Public Overrides Sub OnParentUpdated(ByVal obj As GDIFilledShape)

        If obj.TrackFill Then
            _XTransform = obj.Bounds.X
            _YTransform = obj.Bounds.Y
            UpdateFill()
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the _Brush member used to actually paint the fill to textures.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub UpdateFill()
        If Not _Brush Is Nothing Then
            _Brush.Dispose()
        End If

        Dim textureBrush As TextureBrush
        Dim img As Image = Me.GetImage

        If Not img Is Nothing Then
            textureBrush = New TextureBrush(img, _TextureWrapMode)
            textureBrush.TranslateTransform(_XTransform, _YTransform)

            _Brush = textureBrush
        Else
            _Brush = New SolidBrush(Color.Beige)
        End If

        NotifyFillUpdated()

    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deserializes a Texture fill
    ''' </summary>
    ''' <param name="parent">the parent shape this fill is responsible for filling.</param>
    ''' <returns>A Boolean indicating if deserialization was successful</returns>
    ''' <remarks>Unlike the other GDIFill deserialize methods, this one has functionality.
    ''' It is responsible for checking if the texture source used to fill the image is still
    ''' valid.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Function deserialize() As Boolean
        Trace.WriteLineIf(Session.Tracing, "TextureFill.Deserialize: " & _ImageSource)
        Dim finfo As New System.IO.FileInfo(_ImageSource)

        If finfo.Exists() Then
            Return True
        Else
            _ImageSource = Session.GraphicsManager.AttemptLoadImage("Could not find texture: " & _ImageSource & ". Please select a different bitmap source.")
            If _ImageSource = String.Empty Then
                Return False
            Else
                Return True
            End If
        End If
    End Function

#End Region


#Region "Disposal and Cleanup"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the textured fill, in turn disposing of the image field.
    ''' </summary>
    ''' <param name="disposing">Whether unmanaged resources are being disposed or not.</param>
    ''' -----------------------------------------------------------------------------
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not _Image Is Nothing Then
                _Image.Dispose()
            Else
                _Image = Nothing
            End If
        End If

        MyBase.Dispose(disposing)
    End Sub

#End Region
End Class
