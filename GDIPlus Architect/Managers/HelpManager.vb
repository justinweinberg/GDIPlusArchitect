Imports GDIObjects


''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : HelpManager
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a wrapper for help throughout the GDI+ Architect project.  Maintains 
''' a local help provider and exposes methods to invoke help for specific topics.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class HelpManager


#Region "Local fields"

    '''<summary>An F1 help provider object.  Used to invoke help</summary>
    Private _hp As System.Windows.Forms.HelpProvider

    '''<summary>Path to the GDI+ Architect help file.</summary>
    Private _HelpPath As String = String.Empty

#End Region

#Region "Constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs the HelpManager.
    ''' </summary>
    ''' <remarks>In the constructor, the help path is assigned the F1 Help Provider is 
    ''' created.  Finally, the help manager assigns help to the panels in the application 
    ''' through the panel manager's setHelpProvier method.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub New()

        _HelpPath = App.Options.RuntimePath & "\GDIPlus Architect.chm"

        _hp = New System.Windows.Forms.HelpProvider
        _hp.HelpNamespace = _HelpPath

        App.PanelManager.SetHelpProvider(Me._hp)
    End Sub
#End Region

#Region "Help Invocation"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invoked the generic content section of the help provider.  This is for when 
    ''' there isn't a specific topic to invoke.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Overloads Sub InvokeHelpContents()
        If Not App.MDIMain Is Nothing Then
            Help.ShowHelp(App.MDIMain, _HelpPath)
        End If
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invokes the contents from the help provider for a specific topic
    ''' </summary>
    ''' <param name="sTopic">The topic to investigate help for</param>
    ''' -----------------------------------------------------------------------------
    Public Overloads Sub InvokeHelpContents(ByVal sTopic As String)
        If Not App.MDIMain Is Nothing Then
            Help.ShowHelp(App.MDIMain, _HelpPath, sTopic)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invokes the general index of help.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub InvokeHelpIndex()
        If Not App.MDIMain Is Nothing Then
            Help.ShowHelp(App.MDIMain, _HelpPath, HelpNavigator.Index)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invokes search from the Help Provider.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub InvokeHelpSearch()
        If Not App.MDIMain Is Nothing Then
            Help.ShowHelp(App.MDIMain, _HelpPath, HelpNavigator.Find, "")
        End If
    End Sub




#End Region

End Class
