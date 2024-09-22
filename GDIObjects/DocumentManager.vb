''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : DocumentManager
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Responsible for providing access to properties of the current document being worked with 
''' in the GDIArchitect project.  Limits the access to the current document so that 
''' only a small set of actions can be performed directly from the GDIObject project.
''' </summary>
''' -----------------------------------------------------------------------------
Public Class DocumentManager


#Region "Local Fields"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A reference to the currently selected document in the UI portion of the project.  
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _CurrentDocument As GDIDocument

#End Region


#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the member prefix (local field prefix) of the current document's export settings
    ''' </summary>
    ''' <value>A string containing the member prefix</value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property MemberPrefix() As String
        Get
            Return _CurrentDocument.ExportSettings.MemberPrefix
        End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the FieldPrefix in the current document's ExportSettings
    ''' </summary>
    ''' <value>A string containing the field prefix</value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property FieldPrefix() As String
        Get
            Return _CurrentDocument.ExportSettings.FieldPrefix
        End Get
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the RootNamespace defined in the current document's export settings
    ''' </summary>
    ''' <value>A string containing the root namespace</value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property RootNameSpace() As String
        Get
            Return _CurrentDocument.ExportSettings.RootnameSpace
        End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the application wide current document.
    ''' </summary>
    ''' <value>The GDIDocument to make the system wide current document.</value>
    ''' <remarks>The WriteOnly style is used 
    ''' to make the point clear that the GDIObject project should not access the 
    ''' current document directly.  Instead, it should only consume those methods and properties
    ''' available through this class.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property CurrentDocument() As GDIDocument
        Set(ByVal Value As GDIDocument)
            _CurrentDocument = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a Boolean indicating if a current document exists or not.
    ''' </summary>
    ''' <value>A Boolean indicating if there is a current document.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property HaveCurrentDocument() As Boolean
        Get
            Return Not _CurrentDocument Is Nothing
        End Get
    End Property
#End Region

#Region "Public Methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if a naming conflict exists for a suggested name.
    ''' </summary>
    ''' <param name="proposedName">The proposed name to check</param>
    ''' <returns>A Boolean indicating if a naming conflict exists</returns>
    ''' <remarks>This exists for when the user explicitly sets a name in the property grid 
    ''' and objects need to check if a name conflict exists.</remarks>
    ''' -----------------------------------------------------------------------------
    Public Shared Function NameConflictExists(ByVal proposedName As String) As Boolean
        Return Not _CurrentDocument Is Nothing AndAlso _CurrentDocument.NameExists(proposedName)
    End Function
#End Region

End Class
