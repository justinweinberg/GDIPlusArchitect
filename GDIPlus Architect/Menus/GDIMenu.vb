Imports Crownwood.Magic.Menus
Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : GDIMenu
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Implements MDIMain's menu.  Many thanks to Crownwood for their originally free magic controls.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class GDIMenu
    Inherits System.ComponentModel.Component

#Region "Local Fields"

    '''<summary>Image list of icons used in the menu</summary>
    Private _menuIcons As ImageList = App.IconManager.IconImageList

    '''<summary>The top level reference for the entire menu structure</summary>
    Private _TopMenu As New Crownwood.Magic.Menus.MenuControl


    '''<summary>Seperator menu control</summary>
    Private _Sep As MenuCommand = New MenuCommand("-")

#End Region


#Region "Menu Control Declarations"



    '-- Top level menus -- 
    Private mnuFile As New MenuCommand("&File")
    Private mnuEdit As New MenuCommand("&Edit")
    Private mnuView As New MenuCommand("&View")
    Private mnuFormat As New MenuCommand("&Format")
    Private mnuWindow As New MenuCommand("&Window")
    Private mnuHelp As New MenuCommand("&Help")

    Private mnuDocumentProperties As New MenuCommand("Document Properties", _menuIcons, IconManager.EnumIcons.properties, AddressOf OnDocumentProperties)
    Private mnuTransform As New MenuCommand("Transform", Shortcut.CtrlT, AddressOf OnTransform)

    Private mnuNew As MenuCommand = New MenuCommand("&New", Shortcut.CtrlN)

    Private mnuNewPrintDoc As MenuCommand = New MenuCommand("Print Document", _menuIcons, IconManager.EnumIcons.newprintdoc, AddressOf OnNewPrintDoc)
    Private mnuNewClass As MenuCommand = New MenuCommand("Graphics Class", _menuIcons, IconManager.EnumIcons.newclass, AddressOf OnNewGraphicsClass)

    Private WithEvents mnuOpen As New MenuCommand("&Open...", _menuIcons, IconManager.EnumIcons.open, Shortcut.CtrlO, AddressOf OnOpen)

    Private mnuRecent As New MenuCommand("Recent Files")
    Private mnuClose As New MenuCommand("&Close", AddressOf OnClose)
    Private mnuSave As New MenuCommand("&Save", _menuIcons, IconManager.EnumIcons.save, Shortcut.CtrlS, AddressOf OnSave)
    Private mnuSaveAs As New MenuCommand("Save &As", AddressOf OnSaveAs)
    Private mnuSaveAll As New MenuCommand("Save Al&l", _menuIcons, IconManager.EnumIcons.saveall, Shortcut.CtrlShiftS, AddressOf OnSaveAll)
    Private mnuAddPage As New MenuCommand("Add page", AddressOf OnNewPage)
    Private mnuImports As New MenuCommand("&Import...", Shortcut.CtrlI, AddressOf OnPlace)
    Private mnuPrinters As New MenuCommand("P&rinters...", _menuIcons, IconManager.EnumIcons.print, AddressOf OnPrinterSettings)
    Private mnuPreview As New MenuCommand("Pre&view", _menuIcons, IconManager.EnumIcons.printpreview, AddressOf OnPreview)
    Private mnuRevert As New MenuCommand("Revert", AddressOf OnRevert)
    Private mnuExport As New MenuCommand("Export Code", AddressOf OnExport)
    Private mnuExportImage As New MenuCommand("Export Image", AddressOf OnExportImage)
    Private mnuExit As New MenuCommand("&Exit", Shortcut.CtrlQ, AddressOf OnExit)

    '-- Edit menu --
    Private mnuUndo As New MenuCommand("&Undo", _menuIcons, IconManager.EnumIcons.undo, Shortcut.CtrlZ, AddressOf OnUndo)
    Private mnuRedo As New MenuCommand("&Redo", _menuIcons, IconManager.EnumIcons.redo, Shortcut.CtrlY, AddressOf OnRedo)
    Private mnuCut As New MenuCommand("Cu&t", _menuIcons, IconManager.EnumIcons.cut, Shortcut.CtrlX, AddressOf OnCut)
    Private mnuCopy As New MenuCommand("&Copy", _menuIcons, IconManager.EnumIcons.copy, Shortcut.CtrlC, AddressOf OnCopy)
    Private mnuPaste As New MenuCommand("&Paste", _menuIcons, IconManager.EnumIcons.paste, Shortcut.CtrlV, AddressOf OnPaste)
    Private mnuPasteAttributes As New MenuCommand("Paste Att&ributes", Shortcut.CtrlShiftV, AddressOf OnPasteAttributes)

    Private mnuDelete As New MenuCommand("&Delete", _menuIcons, IconManager.EnumIcons.delete, Shortcut.Del, AddressOf OnDelete)
    Private mnuSelectAll As New MenuCommand("Select &All", Shortcut.CtrlA, AddressOf OnSelectAll)
    Private mnuDeselect As New MenuCommand("Deselect", Shortcut.CtrlShiftA, AddressOf OnDeselectAll)

    '-- View menu --
    Private mnuZoomIn As New MenuCommand("&Zoom In", AddressOf OnZoomIn)
    Private mnuZoomOut As New MenuCommand("Zoom &Out", AddressOf OnZoomOut)
    Private mnuZoom As New MenuCommand("&Magnification", _menuIcons, IconManager.EnumIcons.zoom)
    Private mnuGrids As New MenuCommand("&Guides and Grids")

    Private mnuDrawSampleText As New MenuCommand("&Draw Sample &Text", AddressOf onDrawSampleText)
    Private mnuToggleBorders As New MenuCommand("Draw Field and Text Borders", AddressOf onToggleTextBorders)


    '-- Zoom sub menu --
    Private mnuZoom25 As New MenuCommand("25%", AddressOf OnMag25)
    Private mnuZoom50 As New MenuCommand("50%", Shortcut.Ctrl5, AddressOf OnMag50)
    Private mnuZoom75 As New MenuCommand("75%", AddressOf OnMag75)
    Private mnuZoom100 As New MenuCommand("100%", Shortcut.Ctrl1, AddressOf OnMag100)
    Private mnuZoom150 As New MenuCommand("150%", AddressOf OnMag150)
    Private mnuZoom200 As New MenuCommand("200%", Shortcut.Ctrl2, AddressOf OnMag200)
    Private mnuZoom500 As New MenuCommand("500%", AddressOf OnMag500)


    '-- Grid sub menu --
    Private mnuShowGrids As New MenuCommand("Show Major Grids", AddressOf OnShowGrid)
    Private mnuShowMinorGrids As New MenuCommand("Show Minor Grid", AddressOf OnShowMinorGrid)
    Private mnuShowMargins As New MenuCommand("Show Margins", AddressOf OnShowMargins)
    Private mnuShowGuides As New MenuCommand("Show Guides", AddressOf OnShowGuides)

    '-- Format menu --
    Private mnuFont As New MenuCommand("&Font...", _menuIcons, IconManager.EnumIcons.alpha, AddressOf OnFont)
    Private mnuAlign As New MenuCommand("&Align")
    Private mnuArrangeMode As New MenuCommand("Arrange Mode")
    Private mnuSendToBack As New MenuCommand("&Send to Back", _menuIcons, IconManager.EnumIcons.SendToBack, AddressOf OnSendToBack)
    Private mnuBringToFront As New MenuCommand("&Bring to Front", _menuIcons, IconManager.EnumIcons.BringToFront, AddressOf OnBringToFront)

    Private mnuSendBackward As New MenuCommand("Send Backward", _menuIcons, IconManager.EnumIcons.SendBackward, AddressOf OnSendBackward)
    Private mnuBringForward As New MenuCommand("Bring Forward", _menuIcons, IconManager.EnumIcons.BringForward, AddressOf OnBringForward)

    Private mnuOptions As New MenuCommand("&Options...", AddressOf OnOptions)

    '-- Fill and stroke sub menu --
    Private mnuStroke As New MenuCommand("&Stroke...", _menuIcons, IconManager.EnumIcons.pen, AddressOf OnStroke)
    Private mnuFill As New MenuCommand("F&ill...", _menuIcons, IconManager.EnumIcons.bucket, AddressOf OnFill)


    '-- Alignment and Distribution --
    Private mnuAlignLeft As New MenuCommand("Lefts", _menuIcons, IconManager.EnumIcons.AlignLeft, Shortcut.Alt1, AddressOf OnAlignLeft)
    Private mnuCenterVert As New MenuCommand("Middles", _menuIcons, IconManager.EnumIcons.AlignMiddle, Shortcut.Alt2, AddressOf OnCenterVert)
    Private mnuAlignRight As New MenuCommand("Rights", _menuIcons, IconManager.EnumIcons.AlignRight, Shortcut.Alt3, AddressOf OnAlignRight)

    Private mnuAlignTop As New MenuCommand("Tops", _menuIcons, IconManager.EnumIcons.AlignTop, Shortcut.Alt4, AddressOf OnAlignTop)
    Private mnuCenterHoriz As New MenuCommand("Centers", _menuIcons, IconManager.EnumIcons.AlignCenter, Shortcut.Alt5, AddressOf OnCenterHoriz)
    Private mnuAlignBottom As New MenuCommand("Bottoms", _menuIcons, IconManager.EnumIcons.AlignBottom, Shortcut.Alt6, AddressOf OnAlignBottom)

    Private mnuMakeSameSize As New MenuCommand("Make Same Size")
    Private mnuSameSizeWidth As New MenuCommand("Widths", _menuIcons, IconManager.EnumIcons.SameSizeWidth, AddressOf OnMakeSameSizeWidth)
    Private mnuSameSizeHeight As New MenuCommand("Heights", _menuIcons, IconManager.EnumIcons.SameSizeHeight, AddressOf OnMakeSameSizeHeight)
    Private mnuSameSizeBoth As New MenuCommand("Both", _menuIcons, IconManager.EnumIcons.SameSizeBoth, AddressOf OnMakeSameSizeBoth)

    Private mnuArrangeModeCanvas As New MenuCommand("To Canvas", AddressOf OnArrangeModeToCanvas)
    Private mnuArrangeModeMargins As New MenuCommand("To Margins", AddressOf OnArrangeModeToMargins)

    Private mnuDistributeWidths As New MenuCommand("Distribute Widths", _menuIcons, IconManager.EnumIcons.HorizSpaceEqual, Shortcut.Alt7, AddressOf OnDistWidths)
    Private mnuDistributeHeights As New MenuCommand("Distribute Heights", _menuIcons, IconManager.EnumIcons.VertSpaceEqual, Shortcut.Alt9, AddressOf OnDistHeights)


    '--Windows menu --
    Private mnuToolWin As New MenuCommand("Toolbox", _menuIcons, IconManager.EnumIcons.toolbox, Shortcut.CtrlF2, AddressOf OnToolboxWin)
    Private mnuPropWin As New MenuCommand("Properties", _menuIcons, IconManager.EnumIcons.propwin, Shortcut.CtrlF3, AddressOf OnPropWin)
    Private mnuHistoryWin As New MenuCommand("History", _menuIcons, IconManager.EnumIcons.historywin, Shortcut.CtrlF4, AddressOf OnHistoryWin)
    Private mnuPageWin As New MenuCommand("Pages", _menuIcons, IconManager.EnumIcons.pagewin, Shortcut.CtrlF5, AddressOf OnPageWin)
    Private mnuFillWin As New MenuCommand("Fill", _menuIcons, IconManager.EnumIcons.bucket, Shortcut.CtrlF7, AddressOf OnFillWin)
    Private mnuStrokeWin As New MenuCommand("Stroke", _menuIcons, IconManager.EnumIcons.pen, Shortcut.CtrlF8, AddressOf OnStrokeWin)
    Private mnuTextWin As New MenuCommand("Font", _menuIcons, IconManager.EnumIcons.alpha, Shortcut.CtrlF9, AddressOf OnFontWin)
    Private mnuArrangeWin As New MenuCommand("Arrange", _menuIcons, IconManager.EnumIcons.AlignMain, Shortcut.CtrlShiftF2, AddressOf OnArrangeWin)
    Private mnuCodeWin As New MenuCommand("Code", Shortcut.CtrlShift3, AddressOf onCodeWin)
    Private mnuCascade As New MenuCommand("Cascade", AddressOf OnCascade)
    Private mnuTileH As New MenuCommand("Tile Horizontal", AddressOf OnTileHoriz)
    Private mnuTileV As New MenuCommand("Tile Vertical", AddressOf OnTileVert)
    Private mnuMDIList As New Crownwood.Magic.Menus.MenuCommand("MDI List")
    Private mnuResetPanels As New MenuCommand("Reset Panel Layout", AddressOf OnResetPanelLayout)

    '-- Help sub menu --
    Private mnuAbout As New MenuCommand("About GDI+ Architect", _menuIcons, IconManager.EnumIcons.app, AddressOf onAbout)
    Private mnuHelpContents As New MenuCommand("Contents", _menuIcons, IconManager.EnumIcons.helpcontents, Shortcut.F1, AddressOf onHelpContents)
    Private mnuHelpIndex As New MenuCommand("Index", _menuIcons, IconManager.EnumIcons.helpindex, AddressOf onHelpIndex)
    Private mnuHelpSearch As New MenuCommand("Search", _menuIcons, IconManager.EnumIcons.helpsearch, AddressOf onHelpSearch)

    '-- MRGSoft Specific --
    Private mnuCheckForUpdates As New MenuCommand("Check for Updates", AddressOf onCheckForUpdates)
    Private mnuFeedBack As New MenuCommand("Product Feedback", AddressOf onFeedBack)
    Private mnuSupport As New MenuCommand("Technical Support", AddressOf onTechnicalSupport)

#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a GDIMenu class.  Creates the menu heirarchy and 
    ''' sets checked menus based on application wide options.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        Trace.WriteLineIf(App.TraceOn, "GDIMenu.New")


        createFileMenu()
        createViewMenu()
        createEditmenu()
        createFormatMenu()
        createWindowMenu()
        createHelpMenu()

        _TopMenu.MenuCommands.AddRange(New MenuCommand() _
            {mnuFile, mnuEdit, mnuView, mnuFormat, mnuWindow, mnuHelp})
        _TopMenu.MdiContainer = App.MDIMain

        _TopMenu.Dock = DockStyle.Top

        refreshRecentFileList()

        With App.Options
            mnuShowGrids.Checked = .ShowGrid
            mnuShowMinorGrids.Checked = .ShowMinorGrid
            mnuShowMargins.Checked = .ShowMargins
            mnuShowGuides.Checked = .ShowGuides
            mnuDrawSampleText.Checked = .ShowSampleText
        End With


    End Sub

#End Region


    Private Sub refreshRecentFileList()
        mnuRecent.MenuCommands.Clear()



        Dim recentFile1 As String = App.Options.RecentFile1
        Dim recentFile2 As String = App.Options.RecentFile2
        Dim recentFile3 As String = App.Options.RecentFile3
        Dim recentFile4 As String = App.Options.RecentFile4

        If Not recentFile1 = String.Empty Then
            mnuRecent.MenuCommands.Add( _
            New MenuCommand(App.Options.RecentFile1, AddressOf OnRecent))
        End If

        If Not recentFile2 = String.Empty Then
            mnuRecent.MenuCommands.Add( _
            New MenuCommand(App.Options.RecentFile2, AddressOf OnRecent))
        End If

        If Not recentFile3 = String.Empty Then
            mnuRecent.MenuCommands.Add( _
            New MenuCommand(App.Options.RecentFile3, AddressOf OnRecent))
        End If


        If Not recentFile4 = String.Empty Then
            mnuRecent.MenuCommands.Add( _
            New MenuCommand(App.Options.RecentFile4, AddressOf OnRecent))
        End If

        If mnuRecent.MenuCommands.Count = 0 Then
            mnuRecent.Enabled = False
        Else
            mnuRecent.Enabled = True
        End If
    End Sub



#Region "Public Interfaces"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Rebuilds the list of recently saved files under the File menu.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub onDocumentChanged()
        refreshRecentFileList()
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the style of menu items baed on the currently selected options
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub updateStyle()
        _TopMenu.Style = App.Options.MenuStyle
        If App.Options.ShowMenuIcons Then
            recurseAssignImages(_TopMenu.MenuCommands)
        Else
            recurseUnassignImages(_TopMenu.MenuCommands)
        End If
    End Sub
#End Region


#Region "Property Accessors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a reference to the top level menucontrol
    ''' </summary>
    ''' <value></value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property TopLevelMenu() As Crownwood.Magic.Menus.MenuControl
        Get
            Return _TopMenu
        End Get

    End Property


#End Region


#Region "Menu Events"


    'Sets the arrangmode to align to canvas
    Private Sub OnArrangeModeToCanvas(ByVal s As Object, ByVal e As EventArgs)
        mnuArrangeModeCanvas.Checked = Not mnuArrangeModeCanvas.Checked
        If mnuArrangeModeCanvas.Checked Then
            App.AlignManager.AlignMode = EnumAlignMode.eCanvas
            mnuArrangeModeMargins.Checked = False
        Else
            If mnuArrangeModeMargins.Checked = False Then
                App.AlignManager.AlignMode = EnumAlignMode.eNormal
            End If
        End If

        App.PanelManager.OnAlignToChanged()

    End Sub

    'Sets the arrange mode to align to margins
    Private Sub OnArrangeModeToMargins(ByVal s As Object, ByVal e As EventArgs)
        mnuArrangeModeMargins.Checked = Not mnuArrangeModeMargins.Checked
        If mnuArrangeModeMargins.Checked Then
            App.AlignManager.AlignMode = EnumAlignMode.eMargins
            mnuArrangeModeCanvas.Checked = False
        Else
            If mnuArrangeModeCanvas.Checked = False Then
                App.AlignManager.AlignMode = EnumAlignMode.eNormal
            End If
        End If

        App.PanelManager.OnAlignToChanged()

    End Sub


    'Adds a page to PrintDocuments
    Private Sub OnNewPage(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.addNewPage()
        End If
    End Sub

    'Invokes save all
    Private Sub OnSaveAll(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            For Each doc As GDIDOC In App.MDIMain.MdiChildren
                doc.saveDocument()
            Next
        End If
    End Sub

    'Invokes save as
    Private Sub OnSaveAs(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            MDIMain.SaveActiveDocument(MDIMain.ActiveDocument.FileName)
        End If
    End Sub



    'Launches the printer settings dialog
    Private Sub OnPrinterSettings(ByVal s As Object, ByVal e As EventArgs)

        App.MDIMain.InvokePrinterSettings()

    End Sub

    'Handles an open document request
    Private Sub OnOpen(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            App.MDIMain.InvokeOpen()
        End If
    End Sub

    'Handles a request for a new graphics class style GDIDocument
    Private Sub OnNewGraphicsClass(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            App.MDIMain.InvokeNewGraphicClass()
        End If
    End Sub

    'Invokes document properties
    Private Sub OnDocumentProperties(ByVal s As Object, ByVal e As EventArgs)
        App.MDIMain.InvokeDocumentSettings()
    End Sub


    'Invokes the tranform dialog
    Private Sub OnTransform(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            Dim popTransform As New dgTransform
            popTransform.ShowDialog()
        End If
    End Sub

    'Handles a request for a new print document style GDIDocument
    Private Sub OnNewPrintDoc(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            App.MDIMain.InvokeNewPrintDocument()
        End If
    End Sub

    'Closes the current GDIDoc window
    Private Sub OnClose(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            MDIMain.CloseActiveWindow()
        End If
    End Sub


    'Begins the process of exiting the application.
    Private Sub OnExit(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            App.MDIMain.Close()
        End If
    End Sub


    'Invokes an image place
    Private Sub OnPlace(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            invokePlace()
        End If
    End Sub


    Private Sub placeImage(ByVal imgSrc As String)
        Trace.WriteLineIf(App.TraceOn, "GDIDoc.PlaceImage")

        If GDIImage.IsBitmap(imgSrc) Then

            Dim imgObj As New GDILinkedImage(96, 96, 0, 0, imgSrc)
            MDIMain.ActiveDocument.AddObjectToPage(imgObj, "Image Inserted")
         Else
            MsgBox("Invalid file format")
        End If

    End Sub

    Public Sub invokePlace()
        Trace.WriteLineIf(App.TraceOn, "MDIMain.InvokePlace")

        Dim dgImgOpen As New OpenFileDialog
        dgImgOpen.Filter = GDIImage.CONST_IMG_FILTER

        Dim iresp As DialogResult = dgImgOpen.ShowDialog()
        Try
            If iresp = DialogResult.OK Then
                placeImage(dgImgOpen.FileName)
            End If
        Catch ex As Exception
            MsgBox("Invalid file type")
        Finally
            dgImgOpen.Dispose()
        End Try
    End Sub

    'Sets zoom to 25%
    Private Sub OnMag25(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom = 0.25
    End Sub

    'Sets zoom to 50%
    Private Sub OnMag50(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom = 0.5
    End Sub

    'Sets zoom to 75%
    Private Sub OnMag75(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom = 0.75
    End Sub

    'Sets zoom to 100%
    Private Sub OnMag100(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom = 1
    End Sub

    'Sets zoom to 150%
    Private Sub OnMag150(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom = 1.5
    End Sub

    'Sets zoom to 200%
    Private Sub OnMag200(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom = 2
    End Sub

    'Sets zoom to 500%
    Private Sub OnMag500(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom = 5
    End Sub

    'Deselects all objects on the surface and sets the current tool to the pointer
    Private Sub OnDeselectAll(ByVal s As Object, ByVal e As EventArgs)

        If App.ToolManager.ToolInUse = False Then
            App.ToolManager.GlobalMode = EnumTools.ePointer
            MDIMain.ActiveDocument.Selected.DeselectAll()
            MDIMain.RefreshActiveDocument()
        End If
    End Sub

    'Hides or shows the primary grid
    Private Sub OnShowGrid(ByVal s As Object, ByVal e As EventArgs)
        App.Options.ShowGrid = Not App.Options.ShowGrid
    End Sub

    'Hides or shows document guides
    Private Sub OnShowGuides(ByVal s As Object, ByVal e As EventArgs)
        App.Options.ShowGuides = Not App.Options.ShowGuides
    End Sub

    'Hides or shows the margins on PrintDocument style GDIDocuments
    Private Sub OnShowMargins(ByVal s As Object, ByVal e As EventArgs)
        App.Options.ShowMargins = Not App.Options.ShowMargins
    End Sub


    'Hides or shows the minor grid
    Private Sub OnShowMinorGrid(ByVal s As Object, ByVal e As EventArgs)
        App.Options.ShowMinorGrid = Not App.Options.ShowMinorGrid
    End Sub

    'Selects all items for the current document on the current page
    Private Sub OnSelectAll(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            App.ToolManager.GlobalMode = EnumTools.ePointer
            MDIMain.ActiveDocument.SelectAll()
            MDIMain.RefreshActiveDocument()
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to a Cut command.  If editing text, sends the cut command to the surface.
    ''' Otherwise lets the document's selected set handle the cut operation.  
    ''' </summary>
    ''' <param name="s"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub OnCut(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.InTextEdit Then
            App.ToolManager.NotifyTextCut()
        ElseIf App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.Cut()
        End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to a Paste command.  If editing text, sends the paste command to the surface.
    ''' Otherwise lets the document's selected set handle the paste operation.  
    ''' </summary>
    ''' <param name="s"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub OnPaste(ByVal s As Object, ByVal e As EventArgs)
        Dim d As IDataObject = Clipboard.GetDataObject()

        If d.GetDataPresent(DataFormats.Bitmap) Then

        ElseIf d.GetDataPresent(GDIObjCol.Format.Name) Then
            If App.ToolManager.ToolInUse = False Then
                MDIMain.ActiveDocument.Paste(d)
            End If

        ElseIf d.GetDataPresent(DataFormats.Text) Then
            If App.ToolManager.InTextEdit Then
                App.ToolManager.notifyTextPaste(CStr(d.GetData(DataFormats.Text, True)))
            End If
        End If

    End Sub

    'Invokes a distribute widths
    Private Sub OnDistWidths(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.DistributeWidths()
    End Sub

    'Invokes a distribute heights
    Private Sub OnDistHeights(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.DistributeHeights()
    End Sub

    'Invokes the fill picker dialog
    Private Sub OnFill(ByVal s As Object, ByVal e As EventArgs)
        dgFillPicker.GO()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Launches the font dialog to let the user select a font.  The same functionality
    ''' can be obtained using the TextWin.
    ''' </summary>
    ''' <param name="s"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub OnFont(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            Dim dgFont As New FontDialog
            dgFont.Font = DirectCast(App.Options.Font.Clone, Font)
            Dim iresp As DialogResult = dgFont.ShowDialog()


            If iresp = DialogResult.OK Then
                App.Options.Font = dgFont.Font
            End If

            dgFont.Dispose()
        End If


    End Sub

    'Launches the stroke picker
    Private Sub OnStroke(ByVal s As Object, ByVal e As EventArgs)
        dgStrokePicker.GO()

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a step of undo.  If in text mode, notifies the surface of the undo,
    ''' otherwise notifies the document.
    ''' </summary>
    ''' <param name="s"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub OnUndo(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.InTextEdit Then
            App.ToolManager.notifyTextUndo()
        Else
            If App.ToolManager.ToolInUse = False Then
                MDIMain.ActiveDocument.undo()
            End If
        End If

    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a step of redo.  If in text mode, notifies the surface of the redo,
    ''' otherwise notifies the document.
    ''' </summary>
    ''' <param name="s"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub OnRedo(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.InTextEdit Then
            App.ToolManager.notifyTextRedo()
        Else
            If App.ToolManager.ToolInUse = False Then

                MDIMain.ActiveDocument.redo()
            End If
        End If

    End Sub

    Private Sub onCheckForUpdates(ByVal s As Object, ByVal e As EventArgs)
        'Removed links to MRGSoft.com
    End Sub

    Private Sub onFeedBack(ByVal s As Object, ByVal e As EventArgs)
        'Removed links to MRGSoft.com
    End Sub

    Private Sub onTechnicalSupport(ByVal s As Object, ByVal e As EventArgs)
        'Removed links to MRGSoft.com
    End Sub


    'Shows the about box
    Private Sub onAbout(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            Dim fmAbout As New frmAbout
            fmAbout.ShowDialog()
        End If
    End Sub


    Private Sub onRegister(ByVal s As Object, ByVal e As EventArgs)
        'Removed links to MRGSoft.com
    End Sub

    'Shows help contents
    Private Sub onHelpContents(ByVal s As Object, ByVal e As EventArgs)
        App.HelpManager.InvokeHelpContents()
    End Sub

    'Reverts panels to their shipped defaults
    Private Sub OnResetPanelLayout(ByVal s As Object, ByVal e As EventArgs)
        Dim iresp As MsgBoxResult = MsgBox("Are you sure you wish to reset your panel layout?", MsgBoxStyle.OKCancel)
        If iresp = MsgBoxResult.OK Then
            Dim outerControls() As Control = {Me.TopLevelMenu, App.MDIMain.statusBar}

            App.PanelManager.ResetPanels()
        End If
    End Sub

    'Invokes the help index
    Private Sub onHelpIndex(ByVal s As Object, ByVal e As EventArgs)
        App.HelpManager.InvokeHelpIndex()
    End Sub

    'Invokes help search
    Private Sub onHelpSearch(ByVal s As Object, ByVal e As EventArgs)
        App.HelpManager.InvokeHelpSearch()
    End Sub


    'Tiles selected objects vertically
    Private Sub OnTileVert(ByVal s As Object, ByVal e As EventArgs)
        App.MDIMain.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    'Tile selected objects horizontally
    Private Sub OnTileHoriz(ByVal s As Object, ByVal e As EventArgs)
        App.MDIMain.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    'Cascades windows
    Private Sub OnCascade(ByVal s As Object, ByVal e As EventArgs)
        App.MDIMain.LayoutMdi(MdiLayout.Cascade)
    End Sub

    'Doubles the current zoom
    Private Sub OnZoomIn(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom *= 2
    End Sub


    'Shows or hides borders on text objects
    Private Sub onToggleTextBorders(ByVal s As Object, ByVal e As EventArgs)

        App.Options.ShowTextBorders = Not App.Options.ShowTextBorders

        mnuToggleBorders.Checked = App.Options.ShowTextBorders

    End Sub

    'Toggles drawing sample text in GDIField objects
    Private Sub onDrawSampleText(ByVal s As Object, ByVal e As EventArgs)
        App.Options.ShowSampleText = Not App.Options.ShowSampleText
    End Sub


    'Zooms out by a factor of 2
    Private Sub OnZoomOut(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom /= 2
    End Sub

    'Shows the prop window if hidden
    Private Sub OnPropWin(ByVal s As Object, ByVal e As EventArgs)
        App.PanelManager.showPropertyPanel()
    End Sub


    'Invokes code export on the current document
    Private Sub OnExport(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then

            App.MDIMain.invokeExport()
        End If
    End Sub

    'Invokes image export on the current document
    Private Sub OnExportImage(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            dgExportImage.GO()
        End If
    End Sub


    'Reverts a document to its previously saved state
    Private Sub OnRevert(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            App.MDIMain.InvokeRevert()
        End If
    End Sub

    'Invokes print preview on PrintDocuments
    Private Sub OnPreview(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Preview()

        End If
    End Sub

    'Shows the fill win if hidden
    Private Sub OnFillWin(ByVal s As Object, ByVal e As EventArgs)
        App.PanelManager.showFillPanel()
    End Sub

    'Shows the stroke win if hidden
    Private Sub OnStrokeWin(ByVal s As Object, ByVal e As EventArgs)
        App.PanelManager.showStrokePanel()
    End Sub

    'Shows the font win if hidden
    Private Sub OnFontWin(ByVal s As Object, ByVal e As EventArgs)
        App.PanelManager.showTextPanel()
    End Sub


    'Shows the page win if hidden
    Private Sub OnPageWin(ByVal s As Object, ByVal e As EventArgs)
        App.PanelManager.showPagePanel()
    End Sub




    'Shows the history win if hidden
    Private Sub OnHistoryWin(ByVal s As Object, ByVal e As EventArgs)
        App.PanelManager.showHistoryPanel()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a delete command.  If in text edit mode, sends the delete to the 
    ''' surface.  Otherwise sends it to the selected object set.
    ''' </summary>
    ''' <param name="s"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub OnDelete(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.InTextEdit Then
            App.ToolManager.notifyTextDelete()
        Else
            If App.ToolManager.ToolInUse = False Then
                MDIMain.ActiveDocument.Selected.DeleteAll()
            End If
        End If

    End Sub

    'Sizes selected objects the same for both height and width
    Private Sub OnMakeSameSizeBoth(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.SameSizeBoth()
    End Sub



    'Aligns selected objects top
    Private Sub OnAlignTop(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.AlignTop()
    End Sub

    'Aligns selected objects right
    Private Sub OnAlignRight(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.AlignRight()
    End Sub

    'Aligns selected objects bottom
    Private Sub OnAlignBottom(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.AlignBottom()
    End Sub

    'Handles a save.  Passes the save request to MDIMain
    Private Sub OnSave(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            App.MDIMain.invokesave()
        End If
    End Sub

    'Centers objects horizontally
    Private Sub OnCenterHoriz(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.AlignHorizCenter()
    End Sub

    'Centers objects vertically
    Private Sub OnCenterVert(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.AlignVertCenter()
    End Sub

    'Shows the code panel, if hidden
    Private Sub onCodeWin(ByVal s As Object, ByVal e As EventArgs)
        App.PanelManager.showCodePanel()
    End Sub

    'Shows the arrange panel, if hidden
    Private Sub OnArrangeWin(ByVal s As Object, ByVal e As EventArgs)
        App.PanelManager.showArrangePanel()
    End Sub

    'Shows the toolbox, if hidden
    Private Sub OnToolboxWin(ByVal s As Object, ByVal e As EventArgs)
        App.PanelManager.showToolboxPanel()
    End Sub
    'Sends selected object backward in Z-Order
    Private Sub OnSendBackward(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.SendBackward()
    End Sub

    'Sends selected object forward in Z-Order
    Private Sub OnBringForward(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.BringForward()

    End Sub
    'Sends selected object to the Z-Order back
    Private Sub OnSendToBack(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.SendToBack()

    End Sub


    'Brings selected objects to the Z-Order front
    Private Sub OnBringToFront(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.BringToFront()
    End Sub

    'Invokes the GDI+ Architect options dialog
    Private Sub OnOptions(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.ToolInUse = False Then
            App.Options.InvokeOptionsdialog()
        End If
    End Sub


    'Aligns selected objects left
    Private Sub OnAlignLeft(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.AlignLeft()
    End Sub

    'Makes selected object have the same width
    Private Sub OnMakeSameSizeWidth(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.SameSizeWidth()

    End Sub

    'Makes selected objects have the same height
    Private Sub OnMakeSameSizeHeight(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.SameSizeHeight()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to a copy command.  If editing text, sends the copy command 
    ''' to the surface. Otherwise lets the document's selected set handle the 
    ''' copy operation.  
    ''' </summary>
    ''' <param name="s"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub OnCopy(ByVal s As Object, ByVal e As EventArgs)
        If App.ToolManager.InTextEdit Then
            App.ToolManager.notifyTextCopy()
        Else
            MDIMain.ActiveDocument.Selected.Copy()
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to a paste attributes menu command.  Asks the current GDIDocument 
    ''' to handle the pasteAttributes request.
    ''' </summary>
    ''' <param name="s"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub OnPasteAttributes(ByVal s As Object, ByVal e As EventArgs)
        Dim d As IDataObject = Clipboard.GetDataObject()

        If d.GetDataPresent(GDIObjCol.Format.Name) Then
            If App.ToolManager.ToolInUse = False Then
                MDIMain.ActiveDocument.Selected.PasteAttributes(d)
            End If
        End If
    End Sub

    'Handles a recent file menu click.  Asks MDIMain to open the file 
    Public Sub OnRecent(ByVal s As Object, ByVal e As EventArgs)
        Dim mnucommand As MenuCommand = CType(s, MenuCommand)

        App.MDIMain.InvokeOpenrecent(mnucommand.Text)
    End Sub

#End Region

#Region "Menu Refresh"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in application wide Options from the GDISession object
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub OnOptionsChanged()
        With App.Options
            mnuShowGrids.Checked = .ShowGrid
            mnuShowMinorGrids.Checked = .ShowMinorGrid
            mnuShowMargins.Checked = .ShowMargins
            mnuShowGuides.Checked = .ShowGuides
            mnuDrawSampleText.Checked = .ShowSampleText
        End With

        updateStyle()

    End Sub


    Public Sub Refresh()
        refreshmenuView()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Refreshes align to mode
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub refreshAlignTo()
        mnuArrangeModeCanvas.Checked = App.AlignManager.AlignMode = EnumAlignMode.eCanvas
        mnuArrangeModeMargins.Checked = App.AlignManager.AlignMode = EnumAlignMode.eMargins

    End Sub





    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Refreshes the entire menu, disabling and enabling menu commands and marking checks 
    ''' appropriate  for the current state of the incoming document and 
    ''' application wide option settings.
    ''' </summary>
    ''' <param name="doc">The current GDIDocument</param>
    ''' <remarks>Menu commands may be disabled or enabled for a number of reasons.
    ''' 1) Is there a document at all 
    ''' 2) Are any items on the document selected
    ''' 3) Are two or more items on the document selected
    ''' 4) Is the type of the current GDIDocument (graphics class or PrintDocument 
    ''' style) appropriate for a  menu command
    ''' </remarks>
        ''' -----------------------------------------------------------------------------
    Private Sub refreshmenuView()
        Trace.WriteLineIf(App.TraceOn, "GDIMenu.Refresh")


        Dim bHaveInstance As Boolean = Not MDIMain.ActiveDocument Is Nothing
        Dim bHaveSelection As Boolean
        Dim bTwoOrMoreSelected As Boolean
        Dim bIsPrintDocument As Boolean



        If bHaveInstance Then
            bHaveSelection = MDIMain.ActiveDocument.Selected.Count > 0
            bTwoOrMoreSelected = MDIMain.ActiveDocument.Selected.Count > 1
            bIsPrintDocument = MDIMain.ActiveDocument.ExportSettings.DocumentType = EnumDocumentTypes.ePrintDocument

        Else
            bHaveSelection = False
            bTwoOrMoreSelected = False
            bIsPrintDocument = False
        End If

        mnuNew.Enabled = True
        mnuOpen.Enabled = True

        mnuToggleBorders.Checked = App.Options.ShowTextBorders

        mnuPrinters.Enabled = True
        mnuExit.Enabled = True

        '-- Edit menu --

        mnuRevert.Enabled = bHaveInstance AndAlso MDIMain.ActiveDocument.Saved

        mnuPreview.Enabled = bIsPrintDocument



        mnuDocumentProperties.Enabled = bHaveInstance

        mnuExport.Enabled = bHaveInstance
        mnuExportImage.Enabled = bHaveInstance
        mnuClose.Enabled = bHaveInstance
        mnuSave.Enabled = bHaveInstance

 
        mnuSaveAs.Enabled = bHaveInstance

        mnuSaveAll.Enabled = bHaveInstance
        mnuImports.Enabled = bHaveInstance
        mnuAddPage.Enabled = bHaveInstance

        mnuZoomIn.Enabled = bHaveInstance
        mnuZoomOut.Enabled = bHaveInstance
        mnuZoom.Enabled = bHaveInstance
        mnuTransform.Enabled = bHaveSelection
        mnuPasteAttributes.Enabled = bHaveSelection

        mnuUndo.Enabled = (bHaveInstance AndAlso MDIMain.ActiveDocument.HasUndo) _
        OrElse App.ToolManager.InTextEdit

        mnuRedo.Enabled = (bHaveInstance AndAlso MDIMain.ActiveDocument.HasRedo) _
        OrElse App.ToolManager.InTextEdit

        mnuCut.Enabled = bHaveSelection Or App.ToolManager.InTextEdit

        mnuCopy.Enabled = bHaveSelection Or App.ToolManager.InTextEdit

        mnuPaste.Enabled = bHaveInstance Or App.ToolManager.InTextEdit


        mnuDelete.Enabled = bHaveSelection Or App.ToolManager.InTextEdit

        mnuSelectAll.Enabled = bHaveInstance
        mnuDeselect.Enabled = bHaveSelection
        mnuSendToBack.Enabled = bHaveSelection
        mnuBringToFront.Enabled = bHaveSelection
        mnuSendBackward.Enabled = bHaveSelection
        mnuBringForward.Enabled = bHaveSelection
        mnuAlign.Enabled = bTwoOrMoreSelected OrElse _
        (bHaveSelection AndAlso _
        (Not App.AlignManager.AlignMode = EnumAlignMode.eNormal))
        mnuMakeSameSize.Enabled = bTwoOrMoreSelected OrElse _
        (bHaveSelection AndAlso _
        (Not App.AlignManager.AlignMode = EnumAlignMode.eNormal))

        mnuArrangeModeCanvas.Checked = App.AlignManager.AlignMode = EnumAlignMode.eCanvas
        mnuArrangeModeMargins.Checked = App.AlignManager.AlignMode = EnumAlignMode.eMargins
    End Sub

#End Region

#Region "Menu Creation"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Assigns all of the menu's imagelists
    ''' causing the menu items to display images.
    ''' </summary>
    ''' <param name="mnucol"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub recurseAssignImages(ByVal mnucol As Crownwood.Magic.Collections.MenuCommandCollection)
        For Each mnu As MenuCommand In mnucol
            If mnu.IsParent Then
                recurseAssignImages(mnu.MenuCommands)
            End If
            mnu.ImageList = App.IconManager.IconImageList
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Removes all references to the image list from all menu items, in effet removing 
    ''' icons from the displayed menus.
    ''' </summary>
    ''' <param name="mnucol"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub recurseUnassignImages(ByVal mnucol As Crownwood.Magic.Collections.MenuCommandCollection)
        For Each mnu As MenuCommand In mnucol
            If mnu.IsParent Then
                recurseUnassignImages(mnu.MenuCommands)
            End If

            mnu.ImageList = Nothing
        Next


    End Sub


    'Builds the file menu
    Private Sub createFileMenu()
        mnuNew.MenuCommands.AddRange(New MenuCommand() {mnuNewPrintDoc, mnuNewClass})

        mnuFile.MenuCommands.AddRange(New MenuCommand() _
        {mnuNew, mnuOpen, mnuClose, _Sep, _
        mnuSave, mnuSaveAs, mnuSaveAll, _Sep, _
        mnuAddPage, _Sep, _
        mnuImports, _Sep, _
    mnuPrinters, mnuPreview, _Sep, mnuRevert, _Sep, mnuExport, _
    mnuExportImage, _Sep, mnuDocumentProperties, _Sep, mnuRecent, _Sep, _
        mnuExit})
    End Sub

    'Builds the edit menu
    Private Sub createEditmenu()
        mnuEdit.MenuCommands.AddRange(New MenuCommand() _
        {mnuUndo, mnuRedo, _Sep, _
     mnuCut, mnuCopy, mnuPaste, mnuPasteAttributes, _Sep, mnuDelete, _Sep, _
     mnuSelectAll, mnuDeselect})
    End Sub

    'Builds the view menu
    Private Sub createViewMenu()
        mnuZoom.MenuCommands.AddRange(New MenuCommand() _
        {mnuZoom25, mnuZoom50, mnuZoom75, mnuZoom100, mnuZoom150, mnuZoom200, mnuZoom500})

        mnuGrids.MenuCommands.AddRange(New MenuCommand() _
        {mnuShowGrids, mnuShowMinorGrids, mnuShowMargins, mnuShowGuides})

        mnuView.MenuCommands.AddRange(New MenuCommand() _
        {mnuZoomIn, mnuZoomOut, mnuZoom, mnuGrids, mnuDrawSampleText, mnuToggleBorders})
    End Sub

    'Builds the format menu
    Private Sub createFormatMenu()

        mnuAlign.MenuCommands.AddRange(New MenuCommand() _
        {mnuAlignLeft, mnuCenterVert, mnuAlignRight, _Sep, _
        mnuAlignTop, mnuCenterHoriz, mnuAlignBottom, _Sep, _
        mnuDistributeWidths, mnuDistributeHeights})

        mnuArrangeMode.MenuCommands.AddRange(New MenuCommand() _
        {mnuArrangeModeCanvas, mnuArrangeModeMargins})

        mnuMakeSameSize.MenuCommands.AddRange(New MenuCommand() _
        {mnuSameSizeWidth, mnuSameSizeHeight, mnuSameSizeBoth})

        mnuFormat.MenuCommands.AddRange(New MenuCommand() _
        {mnuFont, _Sep, mnuStroke, mnuFill, mnuTransform, mnuAlign, mnuMakeSameSize, _Sep, _
        mnuArrangeMode, _Sep, _
        mnuSendToBack, mnuBringToFront, _Sep, _
        mnuBringForward, mnuSendBackward, _Sep, _
        mnuOptions})
    End Sub

    'Builds the window menu
    Private Sub createWindowMenu()
        mnuWindow.MenuCommands.AddRange(New MenuCommand() _
        {mnuToolWin, mnuPropWin, mnuHistoryWin, mnuPageWin, _Sep, _
        mnuFillWin, mnuStrokeWin, mnuTextWin, mnuArrangeWin, mnuCodeWin, _Sep, _
        mnuResetPanels, _Sep, _
         mnuCascade, mnuTileH, mnuTileV})
    End Sub

    Private Sub createHelpMenu()
        mnuHelp.MenuCommands.AddRange(New MenuCommand() _
        {mnuHelpContents, mnuHelpIndex, mnuHelpSearch, _Sep, _
        mnuFeedBack, mnuCheckForUpdates, mnuSupport, _Sep, _
        mnuAbout})
    End Sub

#End Region




#Region "Disposal and Cleanup"
    Public Overloads Sub Dispose()
        Me.Dispose(True)
    End Sub



    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then

            _menuIcons = Nothing
            _TopMenu.Dispose()
            mnuFile.Dispose()
            mnuEdit.Dispose()
            mnuView.Dispose()
            mnuFormat.Dispose()
            mnuRecent.Dispose()
            mnuWindow.Dispose()
            mnuHelp.Dispose()
            mnuDocumentProperties.Dispose()
            mnuTransform.Dispose()

            mnuNew.Dispose()

            mnuNewPrintDoc.Dispose()
            mnuNewClass.Dispose()

            mnuOpen.Dispose()
            mnuClose.Dispose()
            mnuSave.Dispose()
            mnuSave.Dispose()
            mnuSaveAll.Dispose()
            mnuAddPage.Dispose()
            mnuImports.Dispose()
            mnuPrinters.Dispose()
            mnuPreview.Dispose()
            mnuRevert.Dispose()
            mnuExport.Dispose()
            mnuExportImage.Dispose()
            mnuExit.Dispose()


            mnuUndo.Dispose()
            mnuRedo.Dispose()
            mnuCut.Dispose()
            mnuCopy.Dispose()
            mnuPaste.Dispose()
            mnuPasteAttributes.Dispose()

            mnuDelete.Dispose()
            mnuSelectAll.Dispose()
            mnuDeselect.Dispose()


            mnuZoomIn.Dispose()
            mnuZoomOut.Dispose()
            mnuZoom.Dispose()
            mnuGrids.Dispose()

            mnuDrawSampleText.Dispose()
            mnuZoom25.Dispose()
            mnuZoom50.Dispose()
            mnuZoom75.Dispose()
            mnuZoom100.Dispose()
            mnuZoom150.Dispose()
            mnuZoom200.Dispose()
            mnuZoom500.Dispose()

            mnuShowGrids.Dispose()
            mnuShowMinorGrids.Dispose()
            mnuShowMargins.Dispose()
            mnuShowGuides.Dispose()


            mnuFont.Dispose()
            mnuAlign.Dispose()
            mnuArrangeMode.Dispose()
            mnuArrangeModeCanvas.Dispose()
            mnuArrangeModeMargins.Dispose()
            mnuSendToBack.Dispose()
            mnuBringToFront.Dispose()
            mnuOptions.Dispose()


            mnuStroke.Dispose()
            mnuFill.Dispose()


            mnuAlignLeft.Dispose()
            mnuAlignTop.Dispose()
            mnuAlignRight.Dispose()
            mnuSameSizeWidth.Dispose()
            mnuSameSizeHeight.Dispose()
            mnuSameSizeBoth.Dispose()
            mnuAlignBottom.Dispose()
            mnuDistributeWidths.Dispose()
            mnuDistributeHeights.Dispose()

            mnuToolWin.Dispose()
            mnuArrangeWin.Dispose()
            mnuCodeWin.Dispose()
            mnuPropWin.Dispose()
            mnuHistoryWin.Dispose()
            mnuPageWin.Dispose()

            mnuCut.Dispose()

            mnuTileH.Dispose()
            mnuTileV.Dispose()

            mnuAbout.Dispose()

            mnuHelpContents.Dispose()
            mnuHelpSearch.Dispose()
            mnuHelpIndex.Dispose()
        End If


        MyBase.Dispose(disposing)
    End Sub

#End Region


End Class
