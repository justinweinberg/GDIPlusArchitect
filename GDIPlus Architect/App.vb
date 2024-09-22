''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : App
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Holds references and brokers application wide objects.  Every object in the 
''' Globals class is a single instanced object and is only created here.  There is a 
''' simliar object in the GDIObjects Project which holds single instanced objects relevant 
''' to the GDIObject scope as well as the User Interface scope.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class App


#Region "Application Entry Point"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Main entry point for the application.
    ''' </summary>
    ''' <param name="CmdArgs"></param>
    ''' <returns></returns>
    ''' -----------------------------------------------------------------------------
    <STAThread()> Public Shared Function Main(ByVal CmdArgs() As String) As Integer
        Dim splsh As New Splash

        splsh.Show()
        splsh.Refresh()

        _MDIMain = New MDIMain
        _GraphicsManager = New GraphicsManager

        _Options = Options.LoadOptions
        _AlignManager = New AlignManager
        _ToolTipManager = New ToolTipManager

        _IconManager = New IconManager
        _Panelmanager = New PanelManager
        _HelpManager = New HelpManager
        _ToolManager = New ToolManager




        _MDIMain.startApplication(CmdArgs)

        If Not splsh Is Nothing Then
            splsh.Close()
        End If

        AddHandler _Options.OptionsChanged, AddressOf Options_Changed

        Application.Run(App.MDIMain)

    End Function


#End Region

#Region "Instance Members"

    '''<summary>The application wide main form in which all of the panels, the GDIMenu, and each child  GDIDoc objcet exists in.</summary>
    Private Shared _MDIMain As MDIMain

    '''<summary>Helper class for graphics operations</summary>
    Private Shared _GraphicsManager As GraphicsManager


    '''<summary>The application wide panel manager.  Has methods, event and properties to control all of the various panels in GDI+ Architect</summary>
    Private Shared _Panelmanager As PanelManager

    '''<summary>handles application alignment settings </summary>
    Private Shared _AlignManager As AlignManager

    '''<summary>The application wide option settings.</summary>
    Private Shared WithEvents _Options As OptionsManager

    '''<summary>Holds the application wide tooltip manager.  Responsible for assigning tool tips to panels as well as populating tool tips for dialogs</summary>
    Private Shared _ToolTipManager As ToolTipManager

    '''<summary>The application wide help manager.  Responsible for providing help for the entire application</summary>
    Private Shared _HelpManager As HelpManager

    '''<summary>Application wide Icon manager.  Holds all icon information needed by each individual object in the application</summary>
    Private Shared _IconManager As IconManager

    '''<summary>The Tool manager.  Responsibe for managing the global tool state as well as encapsulating the functionality of each individual tool used on surfaces.</summary>
    Private Shared _ToolManager As ToolManager


#End Region

#Region "Property Accessors"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Hold the application wide main MDI form in which all of the panels, 
    ''' the GDIMenu, and all child forms exists in.  There is only one instance of this 
    ''' object in the entire application.
    ''' </summary>
    ''' <value>Returns a reference to the MDIMain form</value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property MDIMain() As MDIMain
        Get
            Return _MDIMain
        End Get
    End Property


    Public Shared ReadOnly Property AlignManager() As AlignManager
        Get
            Return _AlignManager
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Hold the application wide GraphicsManager.  The GraphicsManager is responsible for 
    ''' supplementing graphic operations that don't have a defined place in the 
    ''' application.  It's basically a graphics utility class.
    ''' </summary>
    ''' <value>Returns a reference to the GraphicsManager</value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property GraphicsManager() As GraphicsManager
        Get
            Return _GraphicsManager
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Hold the application wide PanelManager.  The PanelManager is responsible for 
    ''' creating the panel instances, brokering events and providing methods to interact
    '''  with panels.  No panel is exposed in scope beyond the panel manager.
    ''' </summary>
    ''' <value>Returns a reference to the MDIMain form</value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property PanelManager() As PanelManager
        Get
            Return _Panelmanager
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a reference to application wide Options.  The options class holds all 
    ''' of the user defined options in the application, notifies the GDIObjects related 
    ''' classes of changes to options, and is responsible for serializing itself to 
    ''' Isolated Storage.
    ''' </summary>
    ''' <value></value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property Options() As OptionsManager
        Get
            Return _Options
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the application wide tool tip manager.  All tool tip interaction in the 
    ''' application uses this object to assign its tips.   The actual tips themselves are 
    ''' contained within an external XML file.  Every tool tip in the application depends on 
    ''' this object to function.
    ''' </summary>
    ''' <value>Returns a reference to the tool tip manager</value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property ToolTipManager() As ToolTipManager
        Get
            Return _ToolTipManager
        End Get
    End Property




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the application wide help manager.  The help manager is responsible for 
    ''' brokering all help interactions in the application.  
    ''' </summary>
    ''' <value></value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property HelpManager() As HelpManager
        Get
            Return _HelpManager
        End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the application wide IconManager.  The icon manager is responsible for 
    ''' managing all of the icons used in the application.  This manager holds an imagelist
    ''' which each item that needs an icon (or bitmap) references.  It also is responsible
    ''' for providing methods to assign icons and bitmaps to objects that request them.
    ''' </summary>
    ''' <value></value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property IconManager() As IconManager
        Get
            Return _IconManager
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' This manager is the most complex. It is responsible for holding the current tool 
    ''' state and wrapping all tool interactions.  For example, when the user clicks the 
    ''' rectangle drawing tool, the toolmanager is notified of a new state.  When the user 
    ''' then clicks on the surface, the manager is responsible for handling the drawing of 
    ''' the tool to the surface and responding to an end tool requests from the surface.
    ''' </summary>
    ''' <value>Returns a reference to the tool manager</value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property ToolManager() As ToolManager
        Get
            Return _ToolManager
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether tracing us current enabled or not.
    ''' </summary>
    ''' <value></value>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property TraceOn() As Boolean
        Get
            If _Options Is Nothing Then
                Return False
            End If
            Return _Options.Tracing
        End Get
    End Property
#End Region

    Private Shared Sub Options_Changed(ByVal s As Object, ByVal e As EventArgs)
        _Panelmanager.OnOptionschanged()
        _MDIMain.onOptionsChanged()
    End Sub



End Class
