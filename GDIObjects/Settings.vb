
''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : Settings
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Holds application wide settings relevant to the GDIObjects project
''' </summary>
''' -----------------------------------------------------------------------------
Public Class Settings

#Region "Local Fields"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to draw sample text in fields or not 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _DrawSampleText As Boolean = False
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The Y DPI of the system.  Retrieved from a System.Drawing.Graphics object.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _DPIY As Single = 96
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The X DPI of the system.  Retrieved from a System.Drawing.Graphics object.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _DPIX As Single = 96

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Current Absolute path specified in the parent project's absolute texture path option
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _TextureAbsolutePath As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Current image run time type specified in the parent project's options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _ImageLinkType As EnumLinkType

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Current absolute image path specified in the parent project's options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _ImageAbsolutePath As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Current relative image page specified in the parent project's options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _ImageRelativePath As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether fill consolidation is currently enabled or not in the parent project's options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _ConsolidateFills As Boolean
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether stroke consolidation is currently enabled or not in the parent project's options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _ConsolidateStrokes As Boolean

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether font consolidation is currently enabled or not in the parent project's options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Consolidatefonts As Boolean


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Whether string format consolidation is currently enabled or
    '''  not in the parent project's options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _ConsolidateStringFormats As Boolean

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Color of drag handles defined in the parent project's options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _DragHandleColor As Color
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Curved handle color defined in the parent project's options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _CurveColor As Color
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Guide color defined in the parent project's options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _GuideColor As Color

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Default local field scope defined in the parent project's options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _MemberScope As EnumScope
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Default texture runtime source defined in the parent project's options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _TextureLinkType As EnumLinkType
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Default texture relative path specified in the parent project's options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _TextureRelativePath As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Default number of undo steps specified in the parent project's options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _UndoSteps As Int32 = 20
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to draw text field borders or not.  Specified in the parent project's options.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _DrawTextFieldBorders As Boolean

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Default field scope for new GDIFields.  Defined in the parent project's options.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _FieldScope As EnumScope


#End Region

#Region "Public Methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Assigns the X and Y DPI  values
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub SetupDPI()
        Dim g As Graphics = Session.GraphicsManager.getTempGraphics()
        _DPIY = g.DpiY
        _DPIX = g.DpiX

        g.Dispose()

    End Sub

#End Region


#Region "Property Accessors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The number of undo steps to record.  The default is 20.
    ''' </summary>
    ''' <value>The number of undo steps to record in documents.
    ''' </value>
    '''<remarks>
    '''Due to the way documents consume history, this property doesn't take effect for 
    '''document until they have been reopened.
    '''</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property UndoSteps() As Int32
        Get
            Return _UndoSteps
        End Get
        Set(ByVal Value As Int32)
            _UndoSteps = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to draw text item borders on GDIText and GDIField objects.
    ''' </summary>
    ''' <value>A Boolean indicating whether to draw borders on fields or not.
    ''' </value>
    ''' -----------------------------------------------------------------------------
    Public Property DrawTextFieldBorders() As Boolean
        Get
            Return _DrawTextFieldBorders
        End Get
        Set(ByVal Value As Boolean)
            _DrawTextFieldBorders = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Default scope for newly created GDIField objects.
    ''' </summary>
    ''' <value>Default field scope for new GDIFields.
    ''' </value>
    ''' -----------------------------------------------------------------------------
    Public Property FieldScope() As EnumScope
        Get
            Return _FieldScope
        End Get
        Set(ByVal Value As EnumScope)
            _FieldScope = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' X DPI from the graphics context
    ''' </summary>
    ''' <value>A single containing the X DPI</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property DPIX() As Single
        Get
            Return _DPIX
        End Get

    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Y DPI from the graphics context.
    ''' </summary>
    ''' <value>A single containing the Y DPI</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property DPIY() As Single
        Get
            Return _DPIY
        End Get

    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Default scope for all GDIObjects aside from GDIFields.
    ''' </summary>
    ''' <value>The default memberscope for new items.
    ''' </value>
    ''' -----------------------------------------------------------------------------
    Public Property MemberScope() As EnumScope
        Get
            Return _MemberScope
        End Get
        Set(ByVal Value As EnumScope)
            _MemberScope = Value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The runtime texture path for new textures (set by user options)
    ''' </summary>
    ''' <value>The default runtime source for textures.
    ''' </value>
    ''' -----------------------------------------------------------------------------
    Public Property TextureLinkType() As EnumLinkType
        Get
            Return _TextureLinkType
        End Get
        Set(ByVal Value As EnumLinkType)
            _TextureLinkType = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The relative texture path for new textures (set by user options)
    ''' </summary>
    ''' <value>The default relative path for textures.
    ''' </value>
    ''' -----------------------------------------------------------------------------
    Public Property TextureRelativePath() As String
        Get
            Return _TextureRelativePath
        End Get
        Set(ByVal Value As String)
            _TextureRelativePath = Value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The absolute texture path for new textures (set by user options)
    ''' </summary>
    ''' <value>The default absolute path for textures.
    ''' </value>
    ''' -----------------------------------------------------------------------------
    Public Property TextureAbsolutePath() As String
        Get
            Return _TextureAbsolutePath
        End Get
        Set(ByVal Value As String)
            _TextureAbsolutePath = Value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The default runtime type for placed images.
    ''' </summary>
    ''' <value>An EnumLinkType indicating the default image source type.
    ''' </value>
    ''' -----------------------------------------------------------------------------
    Public Property ImageLinkType() As EnumLinkType
        Get
            Return _ImageLinkType
        End Get
        Set(ByVal Value As EnumLinkType)
            _ImageLinkType = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The default absolute path for placed images.
    ''' </summary>
    ''' <value>The default absolute path to retrieve images from.
    ''' </value>
    ''' -----------------------------------------------------------------------------
    Public Property ImageAbsolutePath() As String
        Get
            Return _ImageAbsolutePath
        End Get
        Set(ByVal Value As String)
            _ImageAbsolutePath = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The default relative path for placed images.
    ''' </summary>
    ''' <value>The default relative path to use for retrieving images in code.
    ''' </value>
    ''' -----------------------------------------------------------------------------
    Public Property ImageRelativePath() As String
        Get
            Return _ImageRelativePath
        End Get
        Set(ByVal Value As String)
            _ImageRelativePath = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether by default to consolidate GDIFills.
    ''' </summary>
    ''' <value>A Boolean indicating whether new GDIFills should be consolidated or not
    ''' </value>
    ''' -----------------------------------------------------------------------------
    Public Property ConsolidateFills() As Boolean
        Get
            Return _ConsolidateFills
        End Get
        Set(ByVal Value As Boolean)
            _ConsolidateFills = Value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether by default to consolidate strokes.
    ''' </summary>
    ''' <value>A Boolean indicating whether new GDIStrokes should be consolidated or not
    ''' </value>
    ''' -----------------------------------------------------------------------------
    Public Property ConsolidateStrokes() As Boolean
        Get
            Return _ConsolidateStrokes
        End Get
        Set(ByVal Value As Boolean)
            _ConsolidateStrokes = Value
        End Set
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether by default to consolidate fonts.
    ''' </summary>
    ''' <value>A Boolean indicating whether new fonts should be consolidated or not
    ''' </value>
    ''' -----------------------------------------------------------------------------
    Public Property ConsolidateFonts() As Boolean
        Get
            Return _Consolidatefonts
        End Get
        Set(ByVal Value As Boolean)
            _Consolidatefonts = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether by default to consolidate string format statements.
    ''' </summary>
    ''' <value>A Boolean indicating whether new string formats should be consolidated or not
    ''' </value>
    ''' -----------------------------------------------------------------------------
    Public Property ConsolidateStringFormats() As Boolean
        Get
            Return _ConsolidateStringFormats
        End Get
        Set(ByVal Value As Boolean)
            _ConsolidateStringFormats = Value
        End Set
    End Property




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The color to draw drag handles in
    ''' </summary>
    ''' <value>The color to render drag handles in</value>
    ''' -----------------------------------------------------------------------------
    Public Property DragHandleColor() As Color
        Get
            Return _DragHandleColor
        End Get
        Set(ByVal Value As Color)
            _DragHandleColor = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The color to draw curvature points in
    ''' </summary>
    ''' <value>The color to render curve points in</value>
    ''' -----------------------------------------------------------------------------
    Public Property CurveColor() As Color
        Get
            Return _CurveColor
        End Get
        Set(ByVal Value As Color)
            _CurveColor = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The color to draw guides in.
    ''' </summary>
    ''' <value>The color to render guides in.</value>
    ''' -----------------------------------------------------------------------------
    Public Property GuideColor() As Color
        Get
            Return _GuideColor
        End Get
        Set(ByVal Value As Color)
            _GuideColor = Value
        End Set
    End Property







    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to draw sample text into GDIFields.  Sample text does not emit as code 
    ''' but serves to help understand if the text assigned at runtime will fit within the
    ''' GDIField 
    ''' </summary>
    ''' <value>Boolean indicating if sample text should be drawn.</value>
    ''' -----------------------------------------------------------------------------
    Public Property DrawSampleText() As Boolean
        Get
            Return _DrawSampleText
        End Get
        Set(ByVal Value As Boolean)
            _DrawSampleText = Value
        End Set
    End Property

#End Region

End Class
