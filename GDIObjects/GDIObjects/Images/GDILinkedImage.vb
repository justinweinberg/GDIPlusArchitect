Imports System.Drawing
Imports System.CodeDom

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDILinkedImage
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Handles images where the path to an image file is specified in text.
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public Class GDILinkedImage
    Inherits GDIImage


#Region "Local Fields"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  The source of where the image is located when using GDI+ Architect.  
    ''' This is different than _Path which is the path to the resource at runtime.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _ImageSource As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the next valid integer suffix for a GDILinkedImage (Image1, Image2, etc)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _NextnameID As Int32 = 0

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Used only in code export.  This is the path that the exported 
    ''' image should be found at, depending on its runtime type.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Path As String = String.Empty

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the runtime source type for generated code. 
    ''' See EnumLinkType for more information.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _RunTimeSource As EnumLinkType


#End Region
 

#Region "Constructors"
 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructor to use for images that are being placed within a specific rectangle on a surface 
    ''' rather than with its "natural" dimensions.
    ''' </summary>
    ''' <param name="xDPI">the x DPI of the image</param>
    ''' <param name="yDPI">The y DPI of the image</param>
    ''' <param name="rect">A rectangle that defines the bounds of the image.</param>
    ''' <param name="imgsrc">A full path to the image resource.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal xDPI As Int32, ByVal yDPI As Int32, ByVal rect As Rectangle, ByVal imgsrc As String)
        MyBase.New()
        Trace.WriteLineIf(Session.Tracing, "LinkedImage.New.rect")

        _XDPI = xDPI
        _YDPI = yDPI

        If Not imgsrc.Length = 0 Then
            _ImageSource = imgsrc
        End If

        Me.Bounds = rect
        _Sized = True

        Me.RuntimeSource = Session.Settings.ImageLinkType

        createImage()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructor for when the image should assume its natural size on a surface.
    ''' </summary>
    ''' <param name="xDPI">The x DPI of the image</param>
    ''' <param name="yDPI">The y DPI of the image.</param>
    ''' <param name="posX">Top Left corner X position.</param>
    ''' <param name="posY">Top Left corner Y position</param>
    ''' <param name="imgSrc">A full path to the image resource.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal xDPI As Int32, ByVal yDPI As Int32, ByVal posX As Integer, ByVal posY As Integer, ByVal imgSrc As String)
        MyBase.New()
        Trace.WriteLineIf(Session.Tracing, "LinkedImage.New.X.Y")

        _XDPI = xDPI
        _YDPI = yDPI


        _Bounds.X = posX
        _Bounds.Y = posY

        If Not imgSrc.Length = 0 Then
            _ImageSource = imgSrc
        End If

        RuntimeSource = Session.Settings.ImageLinkType
        createImage()
    End Sub

#End Region

#Region "Base class Implementers"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the local _Image parameter
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub disposeImage()
        If Not _image Is Nothing Then
            _image.Dispose()
            _image = Nothing
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Assigned the local _Image field with the image specified in the _ImageSource field.
    ''' </summary>
    ''' <returns>A bitmap that can be drawn against a GDI+ graphics object.</returns>
    ''' -----------------------------------------------------------------------------
    Private Sub createImage()
        Try
            disposeImage()
            _Image = Session.GraphicsManager.ImageFromAbsolutePath(_ImageSource)

        Catch ex As System.Exception
            Throw New System.FormatException("Invalid Property Value")
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deserializes an image resource.  Returns false if the image cannot be deserialized.
    ''' </summary>
    ''' <param name="doc">The parent GDIDocument the image belongs to</param>
    ''' <param name="pg">The page the image was saved on.</param>
    ''' <returns>A Boolean indicating if deserialization was successful or not.</returns>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Function deserialize(ByVal doc As GDIDocument, ByVal pg As GDIPage) As Boolean
        'A linked file may have been destroyed, moved, etc.
        Trace.WriteLineIf(Session.Tracing, "LinkedImage.Deserialize")


        'Verify the file still exists at the specified location
        Dim finfo As New System.IO.FileInfo(_ImageSource)



        If finfo.Exists() Then
            'if it does, create it.
            createImage()
            Return True And MyBase.deserialize(doc, pg)
        Else
            Trace.WriteLineIf(Session.Tracing, "LinkedImage.Deserialize.Failed")

            _ImageSource = Session.GraphicsManager.AttemptLoadImage("Could not find the bitmap: " & _ImageSource & ". Please select a different bitmap source.")
            If _ImageSource = String.Empty Then
                Return False
            Else
                createImage()
                Return True And MyBase.deserialize(doc, pg)
            End If
        End If

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the next valid name for a new GDILinkedImage
    ''' </summary>
    ''' <returns>A string containing the next valid name.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Function NextName() As String
        _NextnameID += 1
        Return "Image" & _NextnameID
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a name used when displaying the class in the property browser.
    ''' </summary>
    ''' <value>A string containing the word "Image"</value>
    ''' -----------------------------------------------------------------------------   
    Public Overrides ReadOnly Property ClassName() As String
        Get
            Return "Image"
        End Get
    End Property
#End Region

#Region "Property Accessors"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the source where the image file to be used is located.
    ''' </summary>
    ''' <value>A string representing a path to the resource.</value>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.EditorAttribute(GetType(ImageFileChooser), GetType(Drawing.Design.UITypeEditor))> _
    Public Property ImageSource() As String
        Get
            Return _ImageSource
        End Get
        Set(ByVal Value As String)
            Trace.WriteLineIf(Session.Tracing, "LinkedImage.ImageSource.Set")

            'Verify it's a bitmap

            Dim tempBMP As Bitmap

            If Not Value.ToLower = _ImageSource.ToLower Then
                Try

                    Dim fInfo As New System.IO.FileInfo(Value)

                    If fInfo.Exists() Then

                        tempBMP = New Bitmap(Value)
                        tempBMP.Dispose()
                        tempBMP = Nothing

                        _ImageSource = Value

                        setupPath()
                        createImage()

                    Else
                        Throw New System.FormatException("Invalid Property Value")
                    End If

                Catch ex As Exception
                    Trace.WriteLineIf(Session.Tracing, "LinkedImage.ImageSource.Set.Exception: " & ex.Message)

                    MsgBox("Invalid Property Value")
                Finally
                    If Not tempBMP Is Nothing Then
                        tempBMP.Dispose()
                        tempBMP = Nothing
                    End If
                End Try
            End If

        End Set
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a path to the resource to use in generated code to get at the resource
    ''' </summary>
    ''' <value>A string representing the path to use to get at the resource</value>
    ''' <remarks>This meaning of the value is dependent upon the RunTimeSource property.
    ''' - An absolute RunTimeSource means this will contain a full path to the resource.
    ''' - A relative RunTimeSource will contain a relative path from the application path.
    ''' - An embedded RunTimeSource means this will contain the path used to retrieve the embedded 
    ''' resource.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.Description("If Source is set to relative path, this should be the relative below the executing assembly directory.  If you have chosen absolute path, this should be a complete path to the image.")> _
    Public Property Path() As String
        Get
            Return _Path
        End Get
        Set(ByVal Value As String)
            Trace.WriteLineIf(Session.Tracing, "LinkedImage.Path.Set" & Value)

            _Path = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set how the generated code will find the linked image bitmap.  The options are 
    ''' an EnumLinkType enumeration containing the choices of embedded, absolutepath, or 
    ''' relative path.
    ''' </summary>
    ''' <value>The runtime source specified for the object</value>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.Description("How the generated code will find the bitmap. Choose Embed if you will add the bitmap to the project and set its build action to Embed.    Choose AbsolutePath if you will load it from a fixed path (e.g. c:\..) Choose RelativePath if you will include a subdirectory to the image from the executingassembly path.")> _
    Public Property RuntimeSource() As EnumLinkType
        Get
            Return _RunTimeSource
        End Get
        Set(ByVal Value As EnumLinkType)
            Trace.WriteLineIf(Session.Tracing, "LinkedImage.RuntimeSource.Set" & Value)

            _RunTimeSource = Value
            setupPath()
        End Set
    End Property

#End Region

#Region "Code Generation related methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts a GDILinkedImage to SVG XML and appends it to the outgoing SVG code.
    ''' </summary>
    ''' <param name="xmlDoc">The SVG document to append the image to.</param>
    ''' <param name="defs">The definitions section of the SVG document.</param>
    ''' <param name="group">The group to append the XML spec of the image to.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub toXML(ByVal xmlDoc As Xml.XmlDocument, ByVal defs As Xml.XmlNode, ByVal group As Xml.XmlNode)
        '<image x="0" y="20" width="200" height="180" xlink:href="cat.png" /

        Dim attr As Xml.XmlAttribute
        Dim imageNode As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "image", String.Empty)

        attr = xmlDoc.CreateAttribute("x")
        attr.Value = Me.Bounds.X.ToString
        imageNode.Attributes.Append(attr)


        attr = xmlDoc.CreateAttribute("y")
        attr.Value = Me.Bounds.Y.ToString
        imageNode.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("width")
        attr.Value = Me.Bounds.Width.ToString
        imageNode.Attributes.Append(attr)

        attr = xmlDoc.CreateAttribute("height")
        attr.Value = Me.Bounds.Height.ToString
        imageNode.Attributes.Append(attr)

        If Me.Rotation > 0 Then
            attr = xmlDoc.CreateAttribute("transform")
            attr.Value = "rotate(" & Me.Rotation & " " & Me.RotationPoint.X & " " & RotationPoint.Y & ")"
            imageNode.Attributes.Append(attr)
        End If

        Dim ele As Xml.XmlElement = CType(imageNode, Xml.XmlElement)

        ele.SetAttribute("href", "http://www.w3.org/1999/xlink", Me.ImageSource)



        group.AppendChild(imageNode)

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a GDILinkedImage to code.
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

        Trace.WriteLineIf(Session.Tracing, "LinkedImage.emit")

        'Declare the rectangle
        Dim rectDeclaration As New CodeMemberField
        Dim bitmapDeclaration As New CodeMemberField

        Dim rectInit As New CodeObjectCreateExpression
        Dim bitmapInit As CodeMethodInvokeExpression

        Dim bitmapAssign As New CodeAssignStatement

        Dim invokeDrawImage As New CodeMethodInvokeExpression

        Dim bitmapDispose As New CodeMethodInvokeExpression

        With bitmapDeclaration
            .Name = MyBase.ExportName()
            .Type = New CodeTypeReference(GetType(System.Drawing.Bitmap))
            .Attributes = MyBase.getScope
            .InitExpression = Nothing
        End With

        With rectInit
            .CreateType = New CodeTypeReference(GetType(System.Drawing.RectangleF))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.X))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.Y))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.Width))
            .Parameters.Add(New CodePrimitiveExpression(Me.Bounds.Height))
        End With

        With rectDeclaration
            .Name = MyBase.ExportName() & "Rectangle"
            .Type = New CodeTypeReference(GetType(System.Drawing.RectangleF))
            .Attributes = MemberAttributes.Private
            .InitExpression = rectInit
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

        bitmapAssign.Left = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, bitmapDeclaration.Name)
        bitmapAssign.Right = bitmapInit

        With invokeDrawImage
            .Method.TargetObject = New CodeArgumentReferenceExpression("g")
            .Method.MethodName = "DrawImage"
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, bitmapDeclaration.Name))
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, rectDeclaration.Name))
        End With



        'create dispose call
        With bitmapDispose
            .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, bitmapDeclaration.Name)
            .Method.MethodName = "Dispose"
        End With

        declarations.Add(rectDeclaration)
        declarations.Add(bitmapDeclaration)



        InitGraphics.Statements.Add(bitmapAssign)


        If Rotation > 0 Then
            MyBase.emitRotationDeclaration(declarations)
            MyBase.emitInvokeBeginRotation(RenderGDI)
        End If
        RenderGDI.Add(invokeDrawImage)
        If Rotation > 0 Then
            MyBase.emitInvokeEndRotation(RenderGDI)
        End If
        DisposeGDI.Statements.Add(bitmapDispose)

    End Sub

#End Region


#Region "Misc Methods"

 


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles situations where the path of the runtime code must be adjusted due to 
    ''' changes in the source image or changes in the Runtime source type.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub setupPath()
        If _ImageSource.Length > 0 Then
            Dim fInfo As New System.IO.FileInfo(_ImageSource)

            Select Case Me.RuntimeSource
                Case EnumLinkType.AbsolutePath
                    Me.Path = Session.Settings.ImageAbsolutePath & "\" & fInfo.Name
                Case EnumLinkType.RelativePath
                    If Session.Settings.ImageRelativePath = String.Empty Then
                        Me.Path = fInfo.Name
                    Else
                        Me.Path = Session.Settings.ImageRelativePath & "\" & fInfo.Name
                    End If
                Case EnumLinkType.Embedded
                    Me.Path = Session.DocumentManager.RootNameSpace & "." & fInfo.Name
            End Select
        End If
    End Sub
#End Region


End Class
