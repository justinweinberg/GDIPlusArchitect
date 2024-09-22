''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : ExportSettings
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' ExportSettings serves as a container for information relevant to exporting 
''' documents to code but not directly related to the content and objects contained 
''' in the document.  There is one ExportSetting object per document, and it is 
''' persisted with the document.
''' </summary>
''' <remarks></remarks>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public Class ExportSettings

#Region "Local Fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Root namespace to use when exporting.  This is relevant for C# as well as when 
    ''' resources are retrieved via embedded.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _NameSpace As String
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The type of document to export (Graphics class or print document)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _DocumentType As EnumDocumentTypes
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the class to create in the code file (e.g. class foo)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _ClassName As String = String.Empty
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The language to export to (at this time this is only C# or VB.NET)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _ExportLanguage As EnumCodeTypes
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Optional field prefix qualifier for exported field code.
    ''' </summary>
    ''' <remarks>
    ''' Note there is some confusion on naming here.  In this context a "field"
    ''' refers to a property wrapper for a GDIField.  See GDIField for more 
    ''' information 
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private _GDIFieldPrefix As String
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Optional member prefix qualifier for exported member code.
    ''' </summary>
    ''' <remarks>The MemberPrefix applies only to private variables.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private _MemberPrefix As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Optional override of consolidating fills.  If true, all 
    ''' fill consolidation will be ignored.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _OverrideConsolidateFill As Boolean = False

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Optional override of consolidating stroke .  If true, all 
    ''' stroke consolidation will be ignored.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _OverrideConsolidateStroke As Boolean = False

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Optional override of consolidating fonts. If true, all 
    ''' font consolidation will be ignored.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _OverrideConsolidateFont As Boolean = False

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Optional override of consolidating string formats. If true, all 
    ''' string format consolidation will be ignored.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _OverrideConsolidateStringFormats As Boolean = False


#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of an ExportSettings class.
    ''' </summary>
    ''' <param name="sRootspace">The root namespace to use in code generation</param>
    ''' <param name="doctype">The type of document that will be exported (graphics class or printdocument)</param>
    ''' <param name="sClassname">The name of the generated class.</param>
    ''' <param name="elanguage">The language to generated code in (C# or VB.NET).</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal sRootspace As String, ByVal doctype As EnumDocumentTypes, ByVal sClassname As String, ByVal elanguage As EnumCodeTypes)
        _NameSpace = sRootspace
        _DocumentType = doctype
        _ClassName = sClassname
        _ExportLanguage = elanguage
    End Sub

#End Region

#Region "Property Accessors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets whether to override consolidating fills.  If true, all 
    ''' fill consolidation will be ignored.
    ''' </summary>
    ''' <value>Boolean indicating whether to ignore fill consolidation settings</value>
    ''' -----------------------------------------------------------------------------
    Public Property OverrideConsolidateFill() As Boolean
        Get
            Return _OverrideConsolidateFill
        End Get
        Set(ByVal Value As Boolean)
            _OverrideConsolidateFill = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets whether to override consolidating strokes.  If true, all 
    ''' stroke consolidation will be ignored.
    ''' </summary>
    ''' <value>Boolean indicating whether to ignore stroke consolidation settings</value>
    ''' -----------------------------------------------------------------------------     
    Public Property OverrideConsolidateStroke() As Boolean
        Get
            Return _OverrideConsolidateStroke
        End Get
        Set(ByVal Value As Boolean)
            _OverrideConsolidateStroke = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets whether to override consolidating fonts.  If true, all 
    ''' font consolidation will be ignored.
    ''' </summary>
    ''' <value>Boolean indicating whether to ignore font consolidation settings</value>
    ''' -----------------------------------------------------------------------------    
    Public Property OverrideConsolidateFont() As Boolean
        Get
            Return _OverrideConsolidateFont
        End Get
        Set(ByVal Value As Boolean)
            _OverrideConsolidateFont = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets whether to override consolidating string formats.  If true, all 
    ''' string format consolidation will be ignored.
    ''' </summary>
    ''' <value>Boolean indicating whether to ignore string format consolidation settings</value>
    ''' -----------------------------------------------------------------------------    
    Public Property OverrideConsolidateStringFormats() As Boolean
        Get
            Return _OverrideConsolidateStringFormats
        End Get
        Set(ByVal Value As Boolean)
            _OverrideConsolidateStringFormats = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the prefix to preface local class declarations with.
    ''' </summary>
    ''' <value>The prefix to use.</value>
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
    ''' Gets or sets an optional prefix for outputted GDIField property wrappers.
    ''' </summary>
    ''' <value>The prefix to use.</value>
    ''' -----------------------------------------------------------------------------
    Public Property FieldPrefix() As String
        Get
            Return _GDIFieldPrefix
        End Get
        Set(ByVal Value As String)
            _GDIFieldPrefix = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the export language to use when generating code.
    ''' </summary>
    ''' <value>An EnumCodeTypes enumeration (C# or VB.NET)</value>
    ''' -----------------------------------------------------------------------------
    Public Property Language() As EnumCodeTypes
        Get
            Return _ExportLanguage
        End Get
        Set(ByVal Value As EnumCodeTypes)
            _ExportLanguage = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the root namespace.  This is used when embedded resources are to 
    ''' be retrieved.
    ''' </summary>
    ''' <value>The root namespace.</value>
    ''' -----------------------------------------------------------------------------
    Public Property RootnameSpace() As String
        Get
            Return _NameSpace
        End Get
        Set(ByVal Value As String)
            _NameSpace = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the target export document type (Graphics class or PrintDocument)
    ''' </summary>
    ''' <value>An EnumDocumentTypes specifying the export doc type.</value>
    ''' -----------------------------------------------------------------------------
    Public Property DocumentType() As EnumDocumentTypes
        Get
            Return _DocumentType
        End Get
        Set(ByVal Value As EnumDocumentTypes)
            _DocumentType = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the name of the class that code generation will produce.
    ''' </summary>
    ''' <value>A string containing the desired name of the class</value>
    ''' -----------------------------------------------------------------------------
    Public Property ClassName() As String
        Get
            Return _ClassName
        End Get
        Set(ByVal Value As String)
            _ClassName = Value
        End Set
    End Property

#End Region
End Class
