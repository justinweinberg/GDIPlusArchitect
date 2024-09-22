Imports System.IO
Imports System.IO.IsolatedStorage
Imports System.Runtime.Serialization
Imports System.Xml.Serialization
Imports GDIObjects



''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : Options
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a mechanism for persisting and getting at application wide settings.
''' Responsible for managing all of the settings that users can specify.
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable(), Xml.Serialization.XmlRootAttribute()> _
Public Class OptionsManager
    Implements IDisposable

#Region "Event Declarations"
    '''<summary>Used to notify interested parties that the options managed 
    ''' by the optionmanager have changed.</summary>
    Public Event OptionsChanged(ByVal s As Object, ByVal e As EventArgs)

#End Region


#Region "Options dialog Invocation"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invoked dgOptions, and broadcasts an OptionsChanged event if the options 
    ''' may have changed.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub InvokeOptionsdialog()
        Dim opt As New dgOptions
        opt.ShowDialog()
        If opt.OptionsChanged Then
            RaiseEvent OptionsChanged(Me, EventArgs.Empty)
        End If
    End Sub
#End Region

  

#Region "Non Serialized members"

    '''<summary>Whether resources have been disposed or not.</summary>
    <NonSerialized()> _
     Private _Disposed As Boolean = False


    '''<summary>Trace listener.</summary>
    <NonSerialized()> _
    Private Shared _TraceListener As TextWriterTraceListener


#End Region

#Region "Local Fields"


    '''<summary>Whether tracing is enabled or not</summary>
    Private _Tracing As Boolean = False

    '''<summary>Whether font consolidation is enabled or not </summary>
    Protected _ConsolidateFonts As Boolean
    '''<summary>Whether string format consolidation is enabled or not </summary>
    Protected _ConsolidateStringFormats As Boolean
    '''<summary>Whether stroke consolidation is enabled or not </summary>
    Protected _ConsolidateStrokes As Boolean
    '''<summary>Whether fill consolidation is enabled or not </summary>
    Protected _ConsolidateFills As Boolean


    '''<summary>The current visual menu style </summary>
    Protected _MenuStyle As Crownwood.Magic.common.VisualStyle


    '''<summary>Whether icons are enabled on menu items or not</summary>
    Protected _ShowMenuIcons As Boolean = True
    '''<summary>Whether tool tips are enabled or not</summary>
    Protected _ShowTooltips As Boolean = True

    '''<summary>The current scope for newly created field objects</summary>
    Protected _GDIFieldScope As EnumScope
    '''<summary>The currnet scope of the RenderGDI method </summary>
    Protected _RenderGDIScope As EnumScope
    '''<summary>The current scope of new members (aside from fields)</summary>
    Protected _DefaultMemberScope As EnumScope


    '''<summary>Recent file list item #1</summary>
    Protected _RecentFile1 As String = String.Empty
    '''<summary>Recent file list item #2</summary>
    Protected _RecentFile2 As String = String.Empty
    '''<summary>Recent file list item #3</summary>
    Protected _RecentFile3 As String = String.Empty
    '''<summary>Recent file list item #4</summary>
    Protected _RecentFile4 As String = String.Empty

    '''<summary>The current smoothing mode for new documents</summary>
    Protected _SmoothingMode As Drawing2D.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

    '''<summary>The current text hint for new documents</summary>
    Protected _TextRenderHint As System.Drawing.Text.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAlias


    '''<summary>The member prefix to append to code generated members</summary>
    Protected _MemberPrefix As String = String.Empty
    '''<summary>The member prefix to append to fields when code is generated </summary>
    Protected _FieldPrefix As String = String.Empty

    '''<summary>The default code generation language (used for quick code and new documents)</summary>
    Protected _CodeLanguage As EnumCodeTypes

    '''<summary>The default name space for embedded resources.  This is used in code  
    'generation to create embedded resource references.</summary>
    Protected _CodeNameSpace As String = "MynameSpace"

    '''<summary>Default path for relative referenced textures</summary>
    Protected _TextureRelativePath As String
    '''<summary>Default path for absolute referenced textures </summary>
    Protected _TextureAbsolutePath As String
    '''<summary>Default run time source type for textures</summary>
    Protected _TextureLinkType As EnumLinkType

    '''<summary>Default relative path for images placed on surfacess</summary>
    Protected _ImageRelativePath As String
    '''<summary>Default absolute path for images placed on surfaces</summary>
    Protected _ImageAbsolutePath As String
    '''<summary>Default run time source type for images placed on surfaces</summary>
    Protected _ImageLinkType As EnumLinkType

    '''<summary>Whether to use the custom color picker or the windows color picker.</summary>
    Protected _UseCustomColorPicker As Boolean = True

    '''<summary>Whether snap to margins enabled or not </summary>
    Protected _SnapToMargins As Boolean = False
    '''<summary>Whether snap to guides is enabled or not </summary>
    Protected _SnapToGuides As Boolean = False
    '''<summary>Whether snapping to the major grid is enabled or not </summary>
    Protected _SnapToMajorGrid As Boolean = False
    '''<summary>Whether snapping to the minor grid is enabled or not </summary>
    Protected _SnapToMinorGrid As Boolean = False

    '''<summary>The number of undo steps to record</summary>
    Protected _UndoSteps As Int32 = 30
    '''<summary>Whether to draw sample text in GDField objects</summary>
    Protected _ShowSampleText As Boolean
    '''<summary>Whether to draw borders around text objets on the surface </summary>
    Protected _ShowTextBorders As Boolean
    '''<summary>Whether to disable automatic code generation for selected items</summary>
    Protected _DisableCodePanel As Boolean
    '''<summary>Whether to draw handles on mousing over objects</summary>
    Protected _ShowMouseOverHandles As Boolean

    '''<summary>Whether to set fills when an item is selected 
    ''' and the application wide fill changes.  Currently not used.</summary>
    Protected _ApplyFillToSelections As Boolean

    '''<summary>Whether to update the stroke of items when an item is selected 
    ''' and the application wide stroke changes.  Currently not used.</summary>
    Protected _ApplyStrokeToSelections As Boolean

    '''<summary>Color to outline objects in</summary>
    Protected _ColorOutline As Color
    '''<summary>Color to paint the minor grid with </summary>
    Protected _ColorMinorGrid As Color
    '''<summary>Color to paint the major grid with </summary>
    Protected _ColorMajorGrid As Color
    '''<summary>Color to paint the margins with </summary>
    Protected _ColorMargin As Color
    '''<summary>Color to paint the non document (nonprint) area with </summary>
    Protected _ColorNonPrintArea As Color
    '''<summary>Color to paint handles on mouse over</summary>
    Protected _ColorMouseOver As Color
    '''<summary>Color to paint handles on selected items </summary>
    Protected _ColorSelected As Color
    '''<summary>Color to paint curvature points with</summary>
    Protected _ColorCurve As Color
    '''<summary>Color to paint guides with </summary>
    Protected _ColorGuide As Color

    '''<summary>Whether to show the major grid </summary>
    Protected _ShowGrid As Boolean
    '''<summary>Whether to show the minor grid</summary>
    Protected _ShowMinorGrid As Boolean
    '''<summary>Whether to show margins</summary>
    Protected _ShowMargins As Boolean
    '''<summary>Whether to show guides </summary>
    Protected _ShowGuides As Boolean

    '''<summary>Default sample text to insert in new GDIFields</summary>
    Protected _SampleText As String


    '''<summary>Size of the major grid units as a Single (float)</summary>
    Protected _MajorGridSize As Single
    '''<summary>Size of the minor grid units as a single (float)</summary>
    Protected _MinorGridSize As Single



    '''<summary>The current font </summary>
    Protected _Font As Font = New Font(FontFamily.GenericSerif, 10, FontStyle.Regular)


    '''<summary>The amount to nudge items on a "large nudge" (not holding ctrl)</summary>
    Protected _LargeNudge As Int32
    '''<summary>The amount to nudge items on a "small nudge" (holding ctrl)</summary>
    Protected _SmallNudge As Int32

    '''<summary>Snap to elasticity for grids in pixels </summary>
    Protected _GridElasticity As Int32
    '''<summary>Snap to elasticity for guides in pixels</summary>
    Protected _GuideElasticity As Int32
#End Region


#Region "Defaults"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Reverts options to their default values and broadcasts OptionsChanged.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub Revert()
        setDefaults()
        RaiseEvent OptionsChanged(Me, EventArgs.Empty)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the options to the state they initially shipped with.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub setDefaults()
        Me.Font = New Font("Arial", 10, FontStyle.Regular)
        Me.Tracing = False

        _ConsolidateFills = True
        _ConsolidateFonts = True
        _ConsolidateStrokes = True
        _ConsolidateStringFormats = True
        _ShowMenuIcons = True
        _ShowTooltips = True
        _MenuStyle = Crownwood.Magic.Common.VisualStyle.IDE
        _DefaultMemberScope = EnumScope.Protected
        _GDIFieldScope = EnumScope.FriendFamilyInternal
        _RenderGDIScope = EnumScope.Protected

        _MemberPrefix = String.Empty
        _FieldPrefix = String.Empty
        _CodeLanguage = EnumCodeTypes.eVB
        _TextureRelativePath = "images"
        _TextureAbsolutePath = "c:\images"
        _TextureLinkType = EnumLinkType.Embedded

        _ImageRelativePath = "images"
        _ImageAbsolutePath = "c:\images"
        _ImageLinkType = EnumLinkType.Embedded

        _ApplyFillToSelections = True
        _ApplyStrokeToSelections = True

        _TextRenderHint = Drawing.Text.TextRenderingHint.AntiAlias
        _SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

        _UseCustomColorPicker = True
        _UndoSteps = 30
        _SmallNudge = 1
        _LargeNudge = 10
        _GridElasticity = 3
        _GuideElasticity = 10
        _SnapToMargins = True
        _ShowGuides = True
        _ShowSampleText = True
        _ShowTextBorders = True
        _ShowMouseOverHandles = True

        _ColorCurve = Color.Pink
        _ColorSelected = Color.Aqua
        _ColorGuide = Color.LightGreen
        _ColorOutline = Color.Aqua
        _ColorMinorGrid = Color.LightGray
        _ColorMajorGrid = Color.Lavender
        _ColorMouseOver = Color.Red
        _ColorMargin = Color.Green
        _ColorNonPrintArea = Color.Gray

        _SnapToMajorGrid = False
        _SnapToMinorGrid = False
        _SnapToMargins = True
        _SnapToGuides = True

        _ShowGrid = True
        _ShowMinorGrid = False
        _ShowMargins = True

        _SampleText = "Sample Text"

        _MajorGridSize = 50
        _MinorGridSize = 10



        _DisableCodePanel = False

    End Sub

#End Region


#Region "Trace Related Functions"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Starts application wide tracing.  Creates a trace listener and logs to the 
    ''' interface log file.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub BeginTracing()
        If Not _Tracing Then

            _TraceListener = New TextWriterTraceListener(RuntimePath & "\GDIInterface.log")

            Trace.AutoFlush = True
            Trace.Listeners.Add(_TraceListener)
            Trace.WriteLine("")
            Trace.WriteLine("")

            Trace.WriteLine("Beginning Tracing at " & Date.Now.ToString)
            Session.Tracing = True
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' ends application wide tracing.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub EndTracing()
        If _Tracing Then
            If Not _TraceListener Is Nothing Then
                _TraceListener.Close()
                _TraceListener.Dispose()
            End If

            Trace.Listeners.Clear()
            Session.Tracing = False
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether tracing is enabled or not.  On Set,if true, begins tracing.  on false,
    ''' ends tracing.
    ''' </summary>
    ''' <value>A boolean indicating if tracing is enabled or not.</value>
    ''' -----------------------------------------------------------------------------
    Public Property Tracing() As Boolean
        Get
            Return _Tracing
        End Get
        Set(ByVal Value As Boolean)

            If Value Then
                BeginTracing()
            Else
                EndTracing()
            End If

            _Tracing = Value

        End Set
    End Property




#End Region

#Region "Misc Methods"





    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Assures that at least one texture exists and returns this texture.
    ''' </summary>
    ''' <returns>A string containing a valid texture path.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function GetFirstTexture() As String

        For Each sImage As String In System.IO.Directory.GetFiles(RuntimePath & "\textures")
            Dim fInfo As New System.IO.FileInfo(sImage)
            If fInfo.Extension = ".gif" OrElse fInfo.Extension = ".jpg" OrElse fInfo.Extension = ".jpeg" OrElse fInfo.Extension = ".png" OrElse fInfo.Extension = ".bmp" Then
                Return fInfo.FullName
            End If

        Next

    End Function







    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the GDIObject's Session.Settings properties based on the current options
    ''' defined in the interface project.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub setSessionProps()

        With Session.Settings
            .DrawSampleText = Me.ShowSampleText
            .DrawTextFieldBorders = Me.ShowTextBorders

            .UndoSteps = Me.UndoSteps
            .CurveColor = Me.ColorCurve
            .DragHandleColor = Me.ColorSelected
            .GuideColor = Me.ColorGuide

            .TextureAbsolutePath = Me.TextureAbsolutePath
            .TextureRelativePath = Me.TextureRelativePath
            .TextureLinkType = Me.TextureLinkType

            .ImageAbsolutePath = Me.ImageAbsolutePath
            .ImageRelativePath = Me.ImageRelativePath
            .ImageLinkType = Me.ImageLinkType


            .MemberScope = Me.MemberScope
            .FieldScope = Me.FieldScope

            .SetupDPI()
        End With


    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a recent file to the recent file list.  
    ''' </summary>
    ''' <param name="sPath">The path to append to the list.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub addRecentFile(ByVal sPath As String)
        'check if its in the list already
        If sPath = _RecentFile1 OrElse sPath = _RecentFile2 OrElse sPath = _RecentFile3 OrElse sPath = _RecentFile4 Then
            Return
        End If
        If _RecentFile1.Length = 0 Then
            _RecentFile1 = sPath
            Return
        End If
        If _RecentFile2.Length = 0 Then
            _RecentFile2 = sPath
            Return
        End If
        If _RecentFile3.Length = 0 Then
            _RecentFile3 = sPath
            Return
        End If
        If _RecentFile4.Length = 0 Then
            _RecentFile4 = sPath
            Return
        End If

        _RecentFile4 = _RecentFile3
        _RecentFile3 = _RecentFile2
        _RecentFile2 = _RecentFile1
        _RecentFile1 = sPath



    End Sub


#End Region

#Region "Serialize and Deserialize methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Attempts to load options from isolated storage.  If this fails, returns an instance 
    ''' of options with default settings.
    ''' </summary>
    ''' <returns></returns>
    ''' -----------------------------------------------------------------------------
    Public Shared Function LoadOptions() As OptionsManager

        Dim opt As OptionsManager
        Try
            opt = GetFromIsolatedStorage()
        Catch
            opt = New OptionsManager
            opt.setDefaults()
        End Try

        opt.setSessionProps()

        Return opt

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts a color to an XML representation
    ''' </summary>
    ''' <param name="color">The color to convert to an XML representation</param>
    ''' <returns>A color in XML format</returns>
    ''' -----------------------------------------------------------------------------
    Private Function SerializeColor(ByVal color As Color) As String

        If (color.IsNamedColor) Then
            Return String.Format("{0}:{1}", ColorFormat.NamedColor, color.Name)
        Else
            Return String.Format("{0}:{1}:{2}:{3}:{4}", ColorFormat.ARGBColor, color.A, color.R, color.G, color.B)
        End If
    End Function




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts from an XML representation of a color to a System.Drawing.Color
    ''' </summary>
    ''' <param name="sColor">The string of XML representing a color to convert to 
    ''' a color.</param>
    ''' <returns>A color matching the XML.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function DeserializeColor(ByVal sColor As String) As Color

        Dim a, r, g, b As Byte

        Dim pieces() As String = sColor.Split(New Char() {":"c})

        Dim colorType As ColorFormat = CType(System.Enum.Parse(GetType(ColorFormat), pieces(0), True), ColorFormat)

        Select Case colorType
            Case ColorFormat.NamedColor
                Return Color.FromName(pieces(1))



            Case ColorFormat.ARGBColor
                a = Byte.Parse(pieces(1))
                r = Byte.Parse(pieces(2))
                g = Byte.Parse(pieces(3))
                b = Byte.Parse(pieces(4))

                Return Color.FromArgb(a, r, g, b)
        End Select

        Return Color.Empty
    End Function




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Saves the options to isolated storage as XML
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub SaveOptions()

        Dim stream As IsolatedStorageFileStream
        Dim serializer As XmlSerializer = New XmlSerializer(GetType(OptionsManager))
         Dim writer As System.Xml.XmlTextWriter

        Try
            stream = New IsolatedStorageFileStream(Constants.CONST_OPTIONS_FILE_NAME, System.IO.FileMode.Create)
            writer = New System.Xml.XmlTextWriter(stream, System.Text.Encoding.UTF8)
            writer.Formatting = System.Xml.Formatting.Indented
            serializer.Serialize(writer, Me)
         Catch

        Finally
            If Not writer Is Nothing Then
                writer.Close()
            End If
        End Try

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retrieves previously saved options from isolated storage
    ''' </summary>
    ''' <returns>A fully populated OptionsManager class</returns>
    ''' -----------------------------------------------------------------------------
    Private Shared Function GetFromIsolatedStorage() As OptionsManager

        Dim serializer As XmlSerializer = New XmlSerializer(GetType(OptionsManager))
        Dim stream As IsolatedStorageFileStream

        Try
            stream = New IsolatedStorageFileStream( _
            Constants.CONST_OPTIONS_FILE_NAME, _
            System.IO.FileMode.OpenOrCreate)

            Dim opt As OptionsManager = _
            CType(serializer.Deserialize(stream), OptionsManager)

            Return opt

        Catch ex As Exception
            Throw New Exception("Could not retrieve Options From isolated Storage")
        Finally

            If Not stream Is Nothing Then
                stream.Close()
            End If



        End Try

    End Function

#End Region

#Region "Property Accessors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the application runtime path.
    ''' </summary>
    ''' <value><A string containing the runtime path/value>
    ''' -----------------------------------------------------------------------------
    <XmlIgnore()> _
        Public ReadOnly Property RuntimePath() As String
        Get
            Return System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
        End Get
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the texture library path.
    ''' </summary>
    ''' <value><A string containing a path to the texture library/value>
    ''' -----------------------------------------------------------------------------
    <XmlIgnore()> _
    Public ReadOnly Property TexturePath() As String
        Get
            Return Me.RuntimePath & "\textures"
        End Get
    End Property

#Region "Color Related Property Accessors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of the curve point handles
    ''' </summary>
    ''' <value>A color to render point handles with</value>
    ''' -----------------------------------------------------------------------------
    <XmlIgnore()> _
    Public Property ColorCurve() As Color
        Get
            Return _ColorCurve
        End Get
        Set(ByVal Value As Color)
            _ColorCurve = Value
            Session.Settings.CurveColor = _ColorCurve
        End Set
    End Property

  
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of the curve point handles in XML.
    ''' </summary>
    ''' <value>A color to render point handles with</value>
    ''' <remarks>This variant exists because Color's cannot be serialized to XML 
    ''' without a bit of help.</remarks> 
    ''' -----------------------------------------------------------------------------
    <XmlElement("ColorCurve")> _
     Public Property XMLColorCurve() As String
        Get
            Return Me.SerializeColor(ColorCurve)
        End Get
        Set(ByVal Value As String)
            Me.ColorCurve = Me.DeserializeColor(Value)
        End Set
    End Property


 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of guids in XML
    ''' </summary>
    ''' <value>A color to render guides with</value>
    ''' -----------------------------------------------------------------------------
    <XmlIgnore()> _
    Public Property ColorGuide() As Color
        Get
            Return _ColorGuide
        End Get
        Set(ByVal Value As Color)
            _ColorGuide = Value

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of guids in XML
    ''' </summary>
    ''' <value>A color to render guides with</value>
    ''' <remarks>This variant exists because Color's cannot be serialized to XML 
    ''' without a bit of help.</remarks>
    ''' -----------------------------------------------------------------------------
    <XmlElement("ColorGuide")> _
   Public Property XMLColorGuide() As String
        Get
            Return Me.SerializeColor(ColorGuide)
        End Get
        Set(ByVal Value As String)
            Me.ColorGuide = Me.DeserializeColor(Value)
        End Set
    End Property

 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of selected objects, such as drag handles.
    ''' </summary>
    ''' <value>A color to render selected objects with</value>
    ''' -----------------------------------------------------------------------------
    <XmlIgnore()> _
    Public Property ColorSelected() As Color
        Get
            Return _ColorSelected
        End Get
        Set(ByVal Value As Color)
            _ColorSelected = Value
            Session.Settings.DragHandleColor = _ColorSelected
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of selected objects, such as drag handles.
    ''' </summary>
    ''' <value>A color to render selected objects with</value>
    ''' <remarks>This variant exists because Color's cannot be serialized to XML 
    ''' without a bit of help.</remarks>
    ''' -----------------------------------------------------------------------------
    <XmlElement("ColorSelected")> _
     Public Property XMLColorSelected() As String
        Get
            Return Me.SerializeColor(ColorSelected)
        End Get
        Set(ByVal Value As String)
            Me.ColorSelected = Me.DeserializeColor(Value)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of outlines.  Outlines are used when the pen is drawing 
    ''' or manipulate objects and a "virtual" shape is drawn.
    ''' </summary>
    ''' <value>A color to render outlines in.</value>
    ''' -----------------------------------------------------------------------------
     <XmlIgnore()> _
    Public Property ColorOutline() As Color
        Get
            Return _ColorOutline
        End Get
        Set(ByVal Value As Color)
            _ColorOutline = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of outlines.  Outlines are used when the pen is drawing 
    ''' or manipulate objects and a "virtual" shape is drawn.
    ''' </summary>
    ''' <value>A color to render outlines in.</value>
    ''' <remarks>This variant exists because Color's cannot be serialized to XML 
    ''' without a bit of help.</remarks>
    ''' -----------------------------------------------------------------------------
    <XmlElement("ColorOutline")> _
  Public Property XMLColorOutline() As String
        Get
            Return Me.SerializeColor(ColorOutline)
        End Get
        Set(ByVal Value As String)
            Me.ColorOutline = Me.DeserializeColor(Value)
        End Set
    End Property

  
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of the no print area.  This is the area outside of the 
    ''' document surface.
    ''' </summary>
    ''' <value>A color to render the non printed area with</value>
    ''' -----------------------------------------------------------------------------
    <XmlIgnore()> _
    Public Property ColorNonPrintArea() As Color
        Get
            Return _ColorNonPrintArea
        End Get
        Set(ByVal Value As Color)
            _ColorNonPrintArea = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of the no print area.  This is the area outside of the 
    ''' document surface.
    ''' </summary>
    ''' <value>A color to render the non printed area with</value>
    ''' <remarks>This variant exists because Color's cannot be serialized to XML 
    ''' without a bit of help.</remarks>
    ''' -----------------------------------------------------------------------------
    <XmlElement("ColorNonPrintArea")> _
   Public Property XMLColorNonPrintArea() As String
        Get
            Return Me.SerializeColor(ColorNonPrintArea)
        End Get
        Set(ByVal Value As String)
            Me.ColorNonPrintArea = Me.DeserializeColor(Value)
        End Set
    End Property

  
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of the minor grid.
    ''' </summary>
    ''' <value>A color to render the minor grid with</value>
    ''' -----------------------------------------------------------------------------
    <XmlIgnore()> _
    Public Property ColorMinorGrid() As Color
        Get
            Return _ColorMinorGrid
        End Get
        Set(ByVal Value As Color)
            _ColorMinorGrid = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of the minor grid.
    ''' </summary>
    ''' <value>A color to render the minor grid with</value>
    ''' <remarks>This variant exists because Color's cannot be serialized to XML 
    ''' without a bit of help.</remarks>
    ''' -----------------------------------------------------------------------------
    <XmlElement("ColorMinorGrid")> _
      Public Property XMLColorMinorGrid() As String
        Get
            Return Me.SerializeColor(ColorMinorGrid)
        End Get
        Set(ByVal Value As String)
            Me.ColorMinorGrid = Me.DeserializeColor(Value)
        End Set
    End Property

   
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of the major grid.
    ''' </summary>
    ''' <value>A color to render the major grid with</value>
    ''' -----------------------------------------------------------------------------
    <XmlIgnore()> _
    Public Property ColorMajorGrid() As Color
        Get
            Return _ColorMajorGrid
        End Get
        Set(ByVal Value As Color)
            _ColorMajorGrid = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of the major grid.
    ''' </summary>
    ''' <value>A color to render the major grid with</value>
    ''' <remarks>This variant exists because Color's cannot be serialized to XML 
    ''' without a bit of help.</remarks>
    ''' -----------------------------------------------------------------------------
    <XmlElement("ColorMajorGrid")> _
      Public Property XMLColorMajorGrid() As String
        Get
            Return Me.SerializeColor(ColorMajorGrid)
        End Get
        Set(ByVal Value As String)
            Me.ColorMajorGrid = Me.DeserializeColor(Value)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color to use when the mouse moves over objects
    ''' </summary>
    ''' <value>A color to render hovered objects with.</value>
    ''' -----------------------------------------------------------------------------
    <XmlIgnore()> _
    Public Property ColorMouseOver() As Color
        Get
            Return _ColorMouseOver
        End Get
        Set(ByVal Value As Color)
            _ColorMouseOver = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color to use when the mouse moves over objects
    ''' </summary>
    ''' <value>A color to render hovered objects with.</value>
    ''' <remarks>This variant exists because Color's cannot be serialized to XML 
    ''' without a bit of help.</remarks>
    ''' -----------------------------------------------------------------------------
    <XmlElement("ColorMouseOver")> _
        Public Property XMLColorMouseOver() As String
        Get
            Return Me.SerializeColor(ColorMouseOver)
        End Get
        Set(ByVal Value As String)
            Me.ColorMouseOver = Me.DeserializeColor(Value)
        End Set
    End Property


    
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of the margins on PrintDocument style GDIDocuments.
    ''' </summary>
    ''' <value>A color to render margins with.</value>
    ''' -----------------------------------------------------------------------------
    <XmlIgnore()> _
    Public Property ColorMargin() As Color
        Get
            Return _ColorMargin
        End Get
        Set(ByVal Value As Color)
            _ColorMargin = Value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the color of the margins on PrintDocument style GDIDocuments.
    ''' </summary>
    ''' <value>A color to render margins with.</value>
    ''' <remarks>This variant exists because Color's cannot be serialized to XML 
    ''' without a bit of help. </remarks>
    ''' -----------------------------------------------------------------------------
    <XmlElement("ColorMargin")> _
        Public Property XMLColorMargin() As String
        Get
            Return Me.SerializeColor(ColorMargin)
        End Get
        Set(ByVal Value As String)
            ColorMargin = Me.DeserializeColor(Value)
        End Set
    End Property
#End Region




#Region "File list"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the 1st listed recent file
    ''' </summary>
    ''' <value>A string containing the 1st recent file.</value>
    ''' -----------------------------------------------------------------------------
    Public Property RecentFile1() As String
        Get
            Return _RecentFile1
        End Get
        Set(ByVal Value As String)
            _RecentFile1 = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the 2nd listed recent file
    ''' </summary>
    ''' <value>A string containing the 2nd recent file.</value>
    ''' -----------------------------------------------------------------------------
    Public Property RecentFile2() As String
        Get
            Return _RecentFile2
        End Get
        Set(ByVal Value As String)
            _RecentFile2 = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the 3rd listed recent file
    ''' </summary>
    ''' <value>A string containing the 3rd recent file.</value>
    ''' -----------------------------------------------------------------------------
    Public Property RecentFile3() As String
        Get
            Return _RecentFile3
        End Get
        Set(ByVal Value As String)
            _RecentFile3 = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the 4th listed recent file
    ''' </summary>
    ''' <value>A string containing the 4th recent file.</value>
    ''' -----------------------------------------------------------------------------
    Public Property RecentFile4() As String
        Get
            Return _RecentFile4
        End Get
        Set(ByVal Value As String)
            _RecentFile4 = Value
        End Set
    End Property

#End Region


#Region "Image and Texture Settings"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the default runtime type for newly created GDIImage objects.
    ''' </summary>
    ''' <value>An EnumLinkType enumeration</value>
    ''' <remarks>Updates the GDIObjects associated property on set</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property ImageLinkType() As EnumLinkType
        Get
            Return _ImageLinkType
        End Get
        Set(ByVal Value As EnumLinkType)
            _ImageLinkType = Value
            Session.Settings.ImageLinkType = _ImageLinkType
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the default runtime type for newly created texture fills.
    ''' </summary>
    ''' <value>An EnumLinkType enumeration</value>
    ''' <remarks>Updates the GDIObjects associated property on set</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property TextureLinkType() As EnumLinkType
        Get
            Return _TextureLinkType
        End Get
        Set(ByVal Value As EnumLinkType)
            _TextureLinkType = Value
            Session.Settings.TextureLinkType = _TextureLinkType
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the default absolute path for images that have their 
    ''' path specified as absolute.
    ''' </summary>
    ''' <value>A string containing an absolute path</value>
    ''' <remarks>Updates the GDIObjects associated property on set</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property ImageAbsolutePath() As String
        Get
            Return _ImageAbsolutePath
        End Get
        Set(ByVal Value As String)
            _ImageAbsolutePath = Value
            Session.Settings.ImageAbsolutePath = _ImageAbsolutePath
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the default image relative path for images 
    ''' that have their path specified as relative.
    ''' </summary>
    ''' <value>A string containing a relative path</value>
    ''' <remarks>Updates the GDIObjects associated property on set</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property ImageRelativePath() As String
        Get
            Return _ImageRelativePath
        End Get
        Set(ByVal Value As String)
            _ImageRelativePath = Value
            Session.Settings.ImageRelativePath = _ImageRelativePath
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the default relative texture fill path for texture filled object 
    ''' that have their fill specified as a relative path.
    ''' </summary>
    ''' <value>A string containing a relative path</value>
    ''' <remarks>Updates the GDIObjects associated property on set</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property TextureRelativePath() As String
        Get
            Return _TextureRelativePath
        End Get
        Set(ByVal Value As String)
            _TextureRelativePath = Value
            Session.Settings.TextureRelativePath = _TextureRelativePath
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the default absolute texture fill path for texture filled object 
    ''' that have their fill specified as an absolute path.
    ''' </summary>
    ''' <value>A string containing an absolute path</value>
    ''' <remarks>Updates the GDIObjects associated property on set</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property TextureAbsolutePath() As String
        Get
            Return _TextureAbsolutePath
        End Get
        Set(ByVal Value As String)
            _TextureAbsolutePath = Value
            Session.Settings.TextureAbsolutePath = _TextureAbsolutePath
        End Set
    End Property

#End Region




#Region "Consolidation Options"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set whether to consolidate strokes for new stroked objects.
    ''' </summary>
    ''' <value>True to consolidate new strokes, false otherwise.</value>
    ''' <remarks>Updates the GDIObjects associated property on set</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property ConsolidateStrokes() As Boolean
        Get
            Return _ConsolidateStrokes
        End Get
        Set(ByVal Value As Boolean)
            _ConsolidateStrokes = Value
            Session.Settings.ConsolidateStrokes = _ConsolidateStrokes
        End Set
    End Property




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set whether to consolidate fills for newly created fillable shapes.
    ''' </summary>
    ''' <value>True to consolidate new fills, false otherwise.</value>
    ''' <remarks>Updates the GDIObjects associated property on set</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property ConsolidateFills() As Boolean
        Get
            Return _ConsolidateFills
        End Get
        Set(ByVal Value As Boolean)
            _ConsolidateFills = Value
            Session.Settings.ConsolidateFills = _ConsolidateFills
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set whether to consolidate string formats for new GDIText and GDIField 
    ''' objects.
    ''' </summary>
    ''' <value>True to consolidate formats, false otherwise.</value>
    ''' <remarks>Updates the GDIObjects associated property on set</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property ConsolidateStringFormats() As Boolean
        Get
            Return _ConsolidateStringFormats
        End Get
        Set(ByVal Value As Boolean)
            _ConsolidateStringFormats = Value
            Session.Settings.ConsolidateStringFormats = _ConsolidateStringFormats
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set whether to consolidate fonts for new GDIText and GDIField 
    ''' objects.
    ''' </summary>
    ''' <value>True to consolidate fonts, false otherwise.</value>
    ''' <remarks>Updates the GDIObjects associated property on set</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property ConsolidateFonts() As Boolean
        Get
            Return _ConsolidateFonts
        End Get
        Set(ByVal Value As Boolean)
            _ConsolidateFonts = Value
            Session.Settings.ConsolidateFonts = _ConsolidateFonts
        End Set
    End Property


#End Region

#Region "Interface Style"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set whether to draw handles when objects are moused over (highlighted)
    ''' </summary>
    ''' <value>Boolean indicating whether to draw handles or not.</value>
    ''' -----------------------------------------------------------------------------
    Public Property ShowMouseOverHandles() As Boolean
        Get
            Return _ShowMouseOverHandles
        End Get
        Set(ByVal Value As Boolean)
            _ShowMouseOverHandles = Value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set whether to draw sample text in fields 
    ''' </summary>
    ''' <value>Boolean indicating whether to draw sample text or not.</value>
    ''' <remarks>Update the GDIObject Project's related property on set.</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property ShowSampleText() As Boolean
        Get
            Return _ShowSampleText
        End Get
        Set(ByVal Value As Boolean)
            _ShowSampleText = Value
            Session.Settings.DrawSampleText = Value
        End Set
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set whether to draw borders around fields. 
    ''' </summary>
    ''' <value>Boolean indicating whether to draw borders or not.</value>
    ''' <remarks>Update the GDIObject Project's related property on set.</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property ShowTextBorders() As Boolean
        Get
            Return _ShowTextBorders
        End Get
        Set(ByVal Value As Boolean)
            _ShowTextBorders = Value
            Session.Settings.DrawTextFieldBorders = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets Magic's interface style. 
    ''' </summary>
    ''' <value>A Crownwood.Magic.Common.VisualStyle</value>
    ''' -----------------------------------------------------------------------------
    Public Property MenuStyle() As Crownwood.Magic.Common.VisualStyle
        Get
            Return _MenuStyle
        End Get
        Set(ByVal Value As Crownwood.Magic.Common.VisualStyle)
            _MenuStyle = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set whether to use the custom color picker or the default windows 
    ''' color picker.
    ''' </summary>
    ''' <value>True to use the custom picker, false otherwise</value>
    ''' -----------------------------------------------------------------------------
    Public Property UseCustomColorPicker() As Boolean
        Get
            Return _UseCustomColorPicker
        End Get
        Set(ByVal Value As Boolean)
            _UseCustomColorPicker = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set a value indicating whether to render tool tips or not
    ''' </summary>
    ''' <value>A boolean indicating whether to render tips or not.</value>
    ''' -----------------------------------------------------------------------------
    Public Property ShowToolTips() As Boolean
        Get
            Return _ShowTooltips
        End Get
        Set(ByVal Value As Boolean)
            _ShowTooltips = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set a value indicating whether to draw menu icons or not.
    ''' </summary>
    ''' <value>A boolean indicating whether to draw icons or not.</value>
    ''' -----------------------------------------------------------------------------
    Public Property ShowMenuIcons() As Boolean
        Get
            Return _ShowMenuIcons
        End Get
        Set(ByVal Value As Boolean)
            _ShowMenuIcons = Value
        End Set
    End Property


#End Region








#Region "Snapping, Grids, Guides, and Margins"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the elasticity of the grid.  This is how close objects 
    ''' have to be to grid lines before they snap.
    ''' </summary>
    ''' <value>An int32 containing the elasticity.</value>
    ''' -----------------------------------------------------------------------------
    Public Property GridElasticity() As Int32
        Get
            Return _GridElasticity
        End Get
        Set(ByVal Value As Int32)
            _GridElasticity = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the the major grid size.  Determines the size of the major grid. 
    ''' </summary>
    ''' <value>Single (float) containing the grid size.</value>
    ''' -----------------------------------------------------------------------------
    Public Property MajorGridSize() As Single
        Get
            Return _MajorGridSize
        End Get
        Set(ByVal Value As Single)
            _MajorGridSize = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the the minor grid size.  Determines the size of the minor grid. 
    ''' </summary>
    ''' <value>Single (float) containing the grid size.</value>
    ''' -----------------------------------------------------------------------------
    Public Property MinorGridSize() As Single
        Get
            Return _MinorGridSize
        End Get
        Set(ByVal Value As Single)
            _MinorGridSize = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set a value indicating if objects should snap to guides or not.
    ''' </summary>
    ''' <value>True to snap to guides, false otherwise</value>
    ''' -----------------------------------------------------------------------------
    Public Property SnapToGuides() As Boolean
        Get
            Return _SnapToGuides
        End Get
        Set(ByVal Value As Boolean)
            _SnapToGuides = Value
        End Set
    End Property




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set a value indicating if objects should snap to the minor grid or not.
    ''' </summary>
    ''' <value>True to snap to the minor grid, false otherwise</value>
    ''' -----------------------------------------------------------------------------
    Public Property SnapToMinorGrid() As Boolean
        Get
            Return _SnapToMinorGrid
        End Get
        Set(ByVal Value As Boolean)
            _SnapToMinorGrid = Value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set a value indicating if objects should snap to the major grid or not.
    ''' </summary>
    ''' <value>True to snap to the major grid, false otherwise</value>
    ''' -----------------------------------------------------------------------------
    Public Property SnapToMajorGrid() As Boolean
        Get
            Return _SnapToMajorGrid
        End Get
        Set(ByVal Value As Boolean)
            _SnapToMajorGrid = Value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set a value indicating if objects should snap to margins or not.
    ''' </summary>
    ''' <value>True to snap to margins, false otherwise</value>
    ''' -----------------------------------------------------------------------------
    Public Property SnapToMargins() As Boolean
        Get
            Return _SnapToMargins
        End Get
        Set(ByVal Value As Boolean)
            _SnapToMargins = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the snap to elasticity of guides.  This is how close to a guide 
    ''' objects must get before the snap to behavior occurs.
    ''' </summary>
    ''' <value>An int32 containing the snap to elasticity</value>
    ''' -----------------------------------------------------------------------------
    Public Property GuideElasticity() As Int32
        Get
            Return _GuideElasticity
        End Get
        Set(ByVal Value As Int32)
            _GuideElasticity = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the show major grid option.  Determined 
    ''' major grid visibility.
    ''' </summary>
    ''' <value>Boolean whether to show the major grid or not.</value>
    ''' -----------------------------------------------------------------------------
    Public Property ShowGrid() As Boolean
        Get
            Return _ShowGrid
        End Get
        Set(ByVal Value As Boolean)
            _ShowGrid = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the show minor grid option.  Determined 
    ''' the minor grid is visible or not.
    ''' </summary>
    ''' <value>Boolean whether to show the minor grid or not.</value>
    ''' -----------------------------------------------------------------------------
    Public Property ShowMinorGrid() As Boolean
        Get
            Return _ShowMinorGrid
        End Get
        Set(ByVal Value As Boolean)
            _ShowMinorGrid = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the show guides option.  Determined 
    ''' guide visibility in PrintDocuments
    ''' </summary>
    ''' <value>Boolean whether to show guides or not.</value>
    ''' -----------------------------------------------------------------------------
    Public Property ShowGuides() As Boolean
        Get
            Return _ShowGuides
        End Get
        Set(ByVal Value As Boolean)
            _ShowGuides = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the show margins option.  Determined 
    ''' margin visibility in PrintDocuments
    ''' </summary>
    ''' <value>Boolean whether to show margins or not.</value>
    ''' -----------------------------------------------------------------------------
    Public Property ShowMargins() As Boolean
        Get
            Return _ShowMargins
        End Get
        Set(ByVal Value As Boolean)
            _ShowMargins = Value
        End Set
    End Property



#End Region


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the unit increment for a large nudge  (holding CTRL down)
    ''' </summary>
    ''' <value>An int32 containing the increment for a large nudge.</value>
    ''' -----------------------------------------------------------------------------
    Public Property LargeNudge() As Int32
        Get
            Return _LargeNudge
        End Get
        Set(ByVal Value As Int32)
            _LargeNudge = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the unit increment for a small nudge (not holding CTRL down)
    ''' </summary>
    ''' <value>An int32 containing the increment for a small nudge.</value>
    ''' -----------------------------------------------------------------------------
    Public Property SmallNudge() As Int32
        Get
            Return _SmallNudge
        End Get
        Set(ByVal Value As Int32)
            _SmallNudge = Value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the default sample text to display in GDIFields.
    ''' </summary>
    ''' <value>A string containing  the sample text to display.</value>
    ''' ----------------------------------------------------------------------------- 
    Public Property SampleText() As String
        Get
            Return _SampleText
        End Get
        Set(ByVal Value As String)
            _SampleText = Value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the the default number of undo steps to record.
    ''' </summary>
    ''' <value>An int32 containing the number of undo steps to record to history.</value>
    ''' <remarks>Update the GDIObject Project's related property on set.  Additionally, 
    ''' because of the way GDIDocument's declare history, this property doesn't take effect 
    ''' until documents are closed and reopened.</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property UndoSteps() As Int32
        Get
            Return _UndoSteps
        End Get
        Set(ByVal Value As Int32)
            _UndoSteps = Value
            Session.Settings.UndoSteps = _UndoSteps
        End Set
    End Property

#Region "Code Generation Related"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the default prefix for members (field level variables) exported to code.
    ''' </summary>
    ''' <value>A string containing a prefix to append to the front 
    ''' of exported member variables.</value>
    ''' -----------------------------------------------------------------------------
    Public Property MemberPrefix() As String
        Get
            Return _MemberPrefix
        End Get
        Set(ByVal Value As String)
            _MemberPrefix = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the default prefix for GDIFields exported to code.
    ''' </summary>
    ''' <value>A string containing a prefix to append to the front of 
    ''' exported GDIFields.</value>
    ''' -----------------------------------------------------------------------------
    Public Property FieldPrefix() As String
        Get
            Return _FieldPrefix
        End Get
        Set(ByVal Value As String)
            _FieldPrefix = Value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the default code language (C# or VB.NET)
    ''' </summary>
    ''' <value>A code language from the EnumCodeTypes enumeration.</value>
    ''' -----------------------------------------------------------------------------
    Public Property CodeLanguage() As EnumCodeTypes
        Get
            Return _CodeLanguage
        End Get
        Set(ByVal Value As EnumCodeTypes)
            _CodeLanguage = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the the default member scope for exported objects.  Member scope 
    ''' is really "field scope".
    ''' </summary>
    ''' <value>A scope value specified in the EnumScope enumeration.</value>
    ''' <remarks>Update the GDIObject Project's related property on set.</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property MemberScope() As EnumScope
        Get
            Return _DefaultMemberScope
        End Get
        Set(ByVal Value As EnumScope)
            _DefaultMemberScope = Value
            Session.Settings.MemberScope = _DefaultMemberScope

        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the the default field scope for new GDIFields.
    ''' </summary>
    ''' <value>A scope value specified in the EnumScope enumeration.</value>
    ''' <remarks>Update the GDIObject Project's related property on set.</remarks>
    ''' -----------------------------------------------------------------------------
    Public Property FieldScope() As EnumScope
        Get
            Return _GDIFieldScope
        End Get
        Set(ByVal Value As EnumScope)
            _GDIFieldScope = Value
            Session.Settings.FieldScope = _GDIFieldScope

        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the the default smoothing mode for new GDIDocuments.  Smoothing mode 
    ''' has a big effect on performance characteristics of a rendering documents.
    ''' </summary>
    ''' <value>A valid Drawing2D.SmoothingMode.</value>
    ''' -----------------------------------------------------------------------------
    Public Property SmoothingMode() As Drawing2D.SmoothingMode
        Get
            Return _SmoothingMode

        End Get
        Set(ByVal Value As Drawing2D.SmoothingMode)
            _SmoothingMode = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the the default text render hint for new GDIDocuments.  
    ''' Controls how fonts are drawn.
    ''' </summary>
    ''' <value>A valid Drawing.Text.TextRenderingHint.</value>
    ''' -----------------------------------------------------------------------------
    Public Property TextRenderingHint() As Drawing.Text.TextRenderingHint
        Get
            Return _TextRenderHint
        End Get
        Set(ByVal Value As Drawing.Text.TextRenderingHint)
            _TextRenderHint = Value
        End Set
    End Property

#End Region



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set 
    ''' </summary>
    ''' <value></value>
    ''' -----------------------------------------------------------------------------

    Public Property DisableCodePanel() As Boolean
        Get
            Return _DisableCodePanel
        End Get
        Set(ByVal Value As Boolean)
            _DisableCodePanel = Value
        End Set
    End Property




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the application wide font.  Notice the XmlIgnore attributes of the Font property.
    ''' To make this property serializable to XML would require a custom serialization
    ''' formatter similar to how colors are serialized in this class.
    ''' </summary>
    ''' <value>A valid font.</value>
    ''' <remarks>When this property is set, it asks the ActiveDocument's selected set 
    ''' to update fonts for selected objects</remarks>
    ''' -----------------------------------------------------------------------------
    <XmlIgnore()> _
    Public Property Font() As Font
        Get
            Return _Font
        End Get
        Set(ByVal Value As Font)

            If Not _Font Is Nothing Then
                _Font.Dispose()
            End If

            _Font = DirectCast(Value.Clone, Font)

            If Not MDIMain.ActiveDocument Is Nothing Then
                MDIMain.ActiveDocument.Selected.setFont(_Font)
            End If


        End Set
    End Property




#End Region



#Region "Disposal and Cleanup"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the OptionsManager
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the OptionsManager, specifically handling the _font.
    ''' </summary>
    ''' <param name="disposing">Whether unmanaged resources are being disposed or not.</param>
    ''' -----------------------------------------------------------------------------
    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        ' Check to see if Dispose has already been called.
        If Not _Disposed Then

            If disposing Then
                If Not _Font Is Nothing Then
                    _Font.Dispose()
                End If
            End If
            ' Release unmanaged resources.    

        End If

        _Disposed = True
    End Sub



#End Region

End Class



