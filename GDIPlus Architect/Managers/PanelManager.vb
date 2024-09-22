Imports Crownwood.Magic.Docking
Imports System.IO
Imports System.IO.IsolatedStorage
Imports System.Runtime.Serialization
Imports System.Xml.Serialization
Imports GDIObjects


''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : PanelManager
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a wrapper to interact wit all of the various user interface panels 
''' that are used throughout the application.
''' </summary>
''' <remarks>Instead of having each interface panel respond to events, the panel 
''' manager is informed of callbacks and lets the interested panels respond to 
''' these events.
''' </remarks>
''' -----------------------------------------------------------------------------
Friend Class PanelManager


 
#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the PanelManager.
    ''' </summary>
    ''' <remarks>Creates the various windows and assigns tool tips.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        _HistoryWin = New HistoryWin
        _PropertyWin = New PropertyWin
        _ToolBoxWin = New ToolboxWin
        _ArrangeWin = New ArrangeWin
         _PageWin = New PageManagerWin
        _StrokeWin = New StrokeWin
        _FillWin = New FillWin
        _TextWin = New TextWin(App.Options.Font)
        _CodeWin = New CodeWin

        assignToolTips()





    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets up the CrownWood docking manager used to manage the panels in the MDIMain form.
    ''' Sets docking defaults from isolated storage if they are found, or sets the 
    ''' layout to default if not found.
    ''' </summary>
    ''' <param name="outerForm">The form panels will be docked to.  MDIMain.</param>
    ''' <param name="outerControls">Controls that exist inside MDIMain but should be 
    ''' considered outer for the purposes of docking panels.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub InitializePanels(ByVal outerForm As Form, ByVal outerControls() As Control)

        ' Create the object that manages the docking state
        _DockingManager = New DockingManager(outerForm, Crownwood.Magic.Common.VisualStyle.IDE)

        DisablePanels()

        For i As Int32 = 0 To outerControls.Length - 1
            _DockingManager.OuterControl = outerControls(i)
        Next

        Try
            SetupDocking()

        Catch ex As Exception
            ResetPanels()
        End Try

        setupicons()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a series of content managers to hold the dockable panels and assigns 
    ''' them to the dockingmanager.contents panel.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub SetupDocking()

        With _DockingManager.Contents
            _ctFill = .Add(_FillWin, "Fill")
            _ctStroke = .Add(_StrokeWin, "Stroke")
            _ctText = .Add(_TextWin, "Text")
            _ctAlign = .Add(_ArrangeWin, "Arrange")
            _ctCode = .Add(_CodeWin, "Code")
            _ctToolbar = .Add(_ToolBoxWin, "Toolbox")
            _ctProperties = .Add(_PropertyWin, "Properties")
            _ctHistory = .Add(_HistoryWin, "History")
            _ctPages = .Add(_PageWin, "Pages")

        End With

        LoadPanelLayout()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Assigns tooltips to each member of the panel using the ToolTipManager.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub assignToolTips()

        App.ToolTipManager.AssignPanelTip("ArrangeWin", _ArrangeWin.Controls)
        App.ToolTipManager.AssignPanelTip("ToolboxWin", _ToolBoxWin.Controls)
        App.ToolTipManager.AssignPanelTip("FillWin", _FillWin.Controls)
        App.ToolTipManager.AssignPanelTip("PageWin", _PageWin.Controls)
        App.ToolTipManager.AssignPanelTip("StrokeWin", _StrokeWin.Controls)
        App.ToolTipManager.AssignPanelTip("TextWin", _TextWin.Controls)
        App.ToolTipManager.AssignPanelTip("ArrangeWin", _ArrangeWin.Controls)

    End Sub


#End Region

#Region "Local Fields"

    '''<summary>DotNetMagic Docking Manager</summary>
    Private _DockingManager As DockingManager

    '''<summary>DotNetMagic UI panel holder.  Holds the fill pane.</summary>
    Private _ctFill As Content
    '''<summary>DotNetMagic UI panel holder.  Holds the stroke pane.</summary>
    Private _ctStroke As Content
    '''<summary>DotNetMagic UI panel holder.  Holds the text pane.</summary>
    Private _ctText As Content
    '''<summary>DotNetMagic UI panel holder.  Holds the alignment pane.</summary>
    Private _ctAlign As Content
    '''<summary>DotNetMagic UI panel holder.  Holds the code pane.</summary>
    Private _ctCode As Content
    '''<summary>DotNetMagic UI panel holder.  Holds the toolbar pane.</summary>
    Private _ctToolbar As Content
    '''<summary>DotNetMagic UI panel holder.  Holds the property pane.</summary>
    Private _ctProperties As Content
    '''<summary>DotNetMagic UI panel holder.  Holds the history pane.</summary>
    Private _ctHistory As Content
    '''<summary>DotNetMagic UI panel holder.  Holds the pages pane.</summary>
    Private _ctPages As Content

    '''<summary>the fill window pane</summary>
    Private _FillWin As FillWin
    '''<summary>The stroke window pane</summary>
    Private _StrokeWin As StrokeWin
    '''<summary>The history window pane</summary>
    Private _HistoryWin As HistoryWin
    '''<summary>The property window pane</summary>
    Private _PropertyWin As PropertyWin
    '''<summary>The alignment window pane</summary>
    Private _ArrangeWin As ArrangeWin

    '''<summary>The toolbox window pane</summary>
    Private _ToolBoxWin As ToolboxWin
    '''<summary>The text window pane</summary>
    Private _TextWin As TextWin
    '''<summary>The page window pane</summary>
    Private _PageWin As PageManagerWin

    '''<summary>The code window pane.</summary>
    Private _CodeWin As CodeWin

#End Region

#Region "Panel Event Responders"


    ''' <summary>Called to notify panels of the pagechanged event.</summary>
    ''' <remarks>This is called when a different page in the ActiveDocument 
    ''' becomes the current page.</remarks>
    Public Sub OnPageChanged()
        _PageWin.refreshPanel()
    End Sub


    ''' <summary>Called to notify panels of the selection changed event</summary>
    ''' <remarks>This is called from MDIMain when the set of selected objects 
    ''' in the active document changes.</remarks>
    Public Sub OnSelectionChanged()
        _PropertyWin.refreshPanel()
        _ArrangeWin.refreshPanel()
        _CodeWin.refreshPanel()
    End Sub

    ''' <summary>Called to notify panels of the history changed event</summary>
    ''' <remarks>This is called from MDIMain whenever the history of the current document 
    ''' changes.  It lets each relevant panel respond to changes in document history.
    ''' </remarks>
    Public Sub OnHistoryChanged()
        _HistoryWin.refreshPanel()
        _PageWin.refreshPanel()
        _PropertyWin.refreshPanel()
        _ArrangeWin.refreshPanel()
        _CodeWin.refreshPanel()
    End Sub



    ''' <summary>Called to notify panels of the document changed event</summary>
    ''' <remarks>This is called from MDIMain every time a new document becomes the 
    ''' active document, whether through focus changes or opening / closing files.
    ''' If there is no active document (the last document was closed), the method 
    ''' still gets called to let UI panels disable themselves</remarks>
    Public Sub OnDocumentChanged()

        'When there is no active document, disable all UI panels
        If MDIMain.ActiveDocument Is Nothing Then
            DisablePanels()
        Else
            EnablePanels()
        End If


        'Refresh panels that depend on the current document
        _HistoryWin.refreshPanel()
        _PageWin.refreshPanel()
        _PropertyWin.refreshPanel()
        _CodeWin.refreshPanel()


    End Sub
    ''' <summary>Called to notify panels of the align to changed event</summary>
    Public Sub OnAlignToChanged()
        _ArrangeWin.refreshPanel()
    End Sub



    ''' <summary>Called to notify panels of the options changed event</summary>
    Public Sub OnOptionschanged()
        assignToolTips()
    End Sub

    ''' <summary>Called to notify panels of the tools changed event</summary>
    Public Sub OnToolsChanged()
        _ToolBoxWin.refreshPanel()
    End Sub

#End Region



#Region "Panel State Serialization"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Attempts to load the last known docking layout.  Throws an anticipated 
    ''' exception if this fails.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub LoadPanelLayout()
        Dim stream As IsolatedStorageFileStream

        Try
            'Attempt to retrieve the last known docking state from isolated storage.
            stream = New IsolatedStorageFileStream(Constants.CONST_FILE_DOCKING, System.IO.FileMode.Open)
            _DockingManager.LoadConfigFromStream(stream)
        Catch ex As Exception
            Throw New Exception("Could not load panels from isolated storage")
        Finally
            If Not stream Is Nothing Then
                stream.Close()
            End If
        End Try



    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Attempts to save the panel layout to isolate storage.
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Justin]	1/10/2005	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub SaveLayout()
        Dim stream As IsolatedStorageFileStream

        Try
            'Attempt to save the docking state to isolated storage
            stream = New IsolatedStorageFileStream(Constants.CONST_FILE_DOCKING, System.IO.FileMode.Create)
            _DockingManager.SaveConfigToStream(stream, System.Text.Encoding.UTF8)

        Catch ex As Exception
            Throw New Exception("Could not save panels to isolated storate")

        Finally
            If Not stream Is Nothing Then
                stream.Close()
            End If

        End Try

    End Sub


#End Region

#Region "Public Methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Makes the panel set visible
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub ShowPanels()
        _DockingManager.ShowAllContents()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Makes the panel set invisible
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub HidePanels()
        _DockingManager.HideAllContents()
    End Sub


#End Region



#Region "Panel Visibility Methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Displays the history panel
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub showHistoryPanel()
        If Not _ctHistory.Visible Then
            _DockingManager.ShowContent(_ctHistory)
        End If
        _ctHistory.BringToFront()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Displays the code panel
    ''' </summary>
    ''' ----------------------------------------------------------------------------- 
    Public Sub showCodePanel()
        If _ctCode.Visible = False Then
            _DockingManager.ShowContent(_ctCode)
        End If
        _ctAlign.BringToFront()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Displays the arange panel
    ''' </summary>
    ''' ----------------------------------------------------------------------------- 
    Public Sub showArrangePanel()
        If _ctAlign.Visible = False Then
            _DockingManager.ShowContent(_ctAlign)
        End If
        _ctAlign.BringToFront()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Displays the tool box
    ''' </summary>
    ''' -----------------------------------------------------------------------------    
    Public Sub showToolboxPanel()
        If _ctToolbar.Visible = False Then
            _DockingManager.ShowContent(_ctToolbar)
        End If

        _ctToolbar.BringToFront()
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Displays the text panel
    ''' </summary>
    ''' -----------------------------------------------------------------------------    
    Public Sub showTextPanel()
        If _ctText.Visible = False Then
            _DockingManager.ShowContent(_ctText)
        End If
        _ctText.BringToFront()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Displays the fill panel
    ''' </summary>
    ''' -----------------------------------------------------------------------------    
    Public Sub showFillPanel()
        If _ctFill.Visible = False Then
            _DockingManager.ShowContent(_ctFill)
        End If
        _ctFill.BringToFront()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Displays the stroke panel
    ''' </summary>
    ''' -----------------------------------------------------------------------------    
    Public Sub showStrokePanel()
        If _ctStroke.Visible = False Then
            _DockingManager.ShowContent(_ctStroke)
        End If
        _ctStroke.BringToFront()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Displays the page panel
    ''' </summary>
    ''' -----------------------------------------------------------------------------    
    Public Sub showPagePanel()
        If _ctPages.Visible = False Then
            _DockingManager.ShowContent(_ctPages)
        End If

        _ctPages.BringToFront()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Displays the property grid panel
    ''' </summary>
    ''' -----------------------------------------------------------------------------    
    Public Sub showPropertyPanel()
        If _ctProperties.Visible = False Then
            _DockingManager.ShowContent(_ctProperties)
        End If

        _ctProperties.BringToFront()

    End Sub

#End Region


#Region "Reset Panels"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Reverts panels to their original layout state (as shipped with the application).
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub ResetPanels()
        Trace.WriteLineIf(App.TraceOn, "MDIMain.ResetPanels")


        With _DockingManager
            'Clear current panels
            .Contents.Clear()

            'Recreate the content pieces
            With _DockingManager.Contents
                _ctFill = .Add(_FillWin, "Fill")
                _ctStroke = .Add(_StrokeWin, "Stroke")
                _ctText = .Add(_TextWin, "Text")
                _ctAlign = .Add(_ArrangeWin, "Arrange")
                _ctCode = .Add(_CodeWin, "Code")
                _ctToolbar = .Add(_ToolBoxWin, "Toolbox")
                _ctProperties = .Add(_PropertyWin, "Properties")
                _ctHistory = .Add(_HistoryWin, "History")
                _ctPages = .Add(_PageWin, "Pages")
            End With

            'Set the outer bands on the docking manager
            Dim wcLeft As WindowContent = _DockingManager.AddContentWithState(_ctToolbar, Crownwood.Magic.Docking.State.DockLeft)
            Dim wcRight As WindowContent = _DockingManager.AddContentWithState(_ctFill, Crownwood.Magic.Docking.State.DockRight)

            'Add the content windows to the docking manager
            .AddContentToWindowContent(_ctStroke, wcRight)
            .AddContentToWindowContent(_ctText, wcRight)
            .AddContentToWindowContent(_ctPages, wcRight)
            .AddContentToWindowContent(_ctHistory, wcRight)
            .AddContentToWindowContent(_ctAlign, wcRight)
            .AddContentToZone(_ctProperties, wcRight.ParentZone, 1)

        End With

        'Apply icons to the panels (not the controls on the panels, but the icons for 
        'the panels themselves
        setupicons()

        'Assign tooltips to panels
        assignToolTips()
    End Sub



#End Region



#Region "Enable/Disable Panels"

    ''' -----------------------------------------------------------------------------
    ''' <summary>Enables all panels </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub EnablePanels()
        _ToolBoxWin.Enabled = True
        _PropertyWin.Enabled = True
        _HistoryWin.Enabled = True
        _FillWin.Enabled = True
        _StrokeWin.Enabled = True
        _TextWin.Enabled = True
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>Disables all panels </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub DisablePanels()
        _ToolBoxWin.Enabled = False
        _PropertyWin.Enabled = False
        _HistoryWin.Enabled = False
        _PageWin.Enabled = False
        _FillWin.Enabled = False
        _StrokeWin.Enabled = False
        _TextWin.Enabled = False
        _PropertyWin.Enabled = False
    End Sub



#End Region



#Region "Disposal and Cleanup"

    ''' -----------------------------------------------------------------------------
    ''' <summary>Releases all panel related resource </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub DisposePanels()

        If Not _FillWin Is Nothing Then
            _FillWin.Dispose()
            _FillWin = Nothing
        End If

        If Not _HistoryWin Is Nothing Then
            _HistoryWin.Dispose()
            _HistoryWin = Nothing
        End If



        If Not _PageWin Is Nothing Then
            _PageWin.Dispose()
            _PageWin = Nothing
        End If

        If Not _PropertyWin Is Nothing Then
            _PropertyWin.Dispose()
            _PropertyWin = Nothing
        End If

        If Not _StrokeWin Is Nothing Then
            _StrokeWin.Dispose()
            _StrokeWin = Nothing
        End If
        If Not _TextWin Is Nothing Then
            _TextWin.Dispose()
            _TextWin = Nothing
        End If
        If Not _ArrangeWin Is Nothing Then
            _ArrangeWin.Dispose()
            _ArrangeWin = Nothing
        End If

        If Not _CodeWin Is Nothing Then
            _CodeWin.Dispose()
            _CodeWin = Nothing
        End If
        If Not _ToolBoxWin Is Nothing Then
            _ToolBoxWin.Dispose()
            _ToolBoxWin = Nothing
        End If
    End Sub


#End Region

#Region "Icons and Help Setup"

    ''' -----------------------------------------------------------------------------
    ''' <summary>Sets help for panels</summary>
    ''' -----------------------------------------------------------------------------
    Public Sub SetHelpProvider(ByVal hp As HelpProvider)



        'Set the appropriate help page in help for each panel

        With hp
            .SetHelpNavigator(_PropertyWin, HelpNavigator.Topic)
            .SetHelpKeyword(_PropertyWin, "panelproperties.htm")
            .SetShowHelp(_PropertyWin, True)

            .SetHelpNavigator(_PageWin, HelpNavigator.Topic)
            .SetHelpKeyword(_PageWin, "panelpage.htm")
            .SetShowHelp(_PageWin, True)

            .SetHelpNavigator(_ArrangeWin, HelpNavigator.Topic)
            .SetHelpKeyword(_ArrangeWin, "ArrangeWin.htm")
            .SetShowHelp(_ArrangeWin, True)

            .SetHelpNavigator(_ToolBoxWin, HelpNavigator.Topic)
            .SetHelpKeyword(_ToolBoxWin, "paneltoolbox.htm")
            .SetShowHelp(_ToolBoxWin, True)


            .SetHelpNavigator(_HistoryWin, HelpNavigator.Topic)
            .SetHelpKeyword(_HistoryWin, "panelhistory.htm")
            .SetShowHelp(_HistoryWin, True)


            .SetHelpNavigator(_StrokeWin, HelpNavigator.Topic)
            .SetHelpKeyword(_StrokeWin, "panelstroke.htm")
            .SetShowHelp(_StrokeWin, True)

            .SetHelpNavigator(_FillWin, HelpNavigator.Topic)
            .SetHelpKeyword(_FillWin, "panelfill.htm")
            .SetShowHelp(_FillWin, True)

            .SetHelpNavigator(_TextWin, HelpNavigator.Topic)
            .SetHelpKeyword(_TextWin, "paneltext.htm")
            .SetShowHelp(_TextWin, True)

        End With

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>Sets icons for panels.</summary>
    ''' -----------------------------------------------------------------------------
    Private Sub setupicons()
        _ctAlign.ImageList = App.IconManager.IconImageList
        _ctToolbar.ImageList = App.IconManager.IconImageList
        _ctFill.ImageList = App.IconManager.IconImageList
        _ctStroke.ImageList = App.IconManager.IconImageList
        _ctText.ImageList = App.IconManager.IconImageList
        _ctToolbar.ImageList = App.IconManager.IconImageList
        _ctProperties.ImageList = App.IconManager.IconImageList
        _ctHistory.ImageList = App.IconManager.IconImageList
        _ctPages.ImageList = App.IconManager.IconImageList
        _ctAlign.ImageList = App.IconManager.IconImageList

        _ctAlign.ImageIndex = IconManager.EnumIcons.toolbox
        _ctToolbar.ImageIndex = IconManager.EnumIcons.toolbox
        _ctFill.ImageIndex = IconManager.EnumIcons.bucket
        _ctStroke.ImageIndex = IconManager.EnumIcons.pen
        _ctText.ImageIndex = IconManager.EnumIcons.alpha
        _ctAlign.ImageIndex = IconManager.EnumIcons.AlignMain
        _ctProperties.ImageIndex = IconManager.EnumIcons.propwin
        _ctHistory.ImageIndex = IconManager.EnumIcons.historywin
        _ctPages.ImageIndex = IconManager.EnumIcons.pagewin
    End Sub

#End Region
End Class
