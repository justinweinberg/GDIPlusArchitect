Imports Crownwood.Magic.Menus
Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : SurfaceContextMenu
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Implements the context menu that shows in response to a right click on the 
''' drawing surface.
''' </summary>
''' <remarks>Rather than document each one liner, an overview of what this class 
''' does is more valuable.
''' </remarks>
''' -----------------------------------------------------------------------------
Friend Class SurfaceContextMenu

#Region "Local Fields"
    '''<summary>Whether this class has been disposed or not.</summary>
    Private mbDisposed As Boolean = False


#End Region

#Region "Menu Commands"
    Private sep As MenuCommand = New MenuCommand("-")


    Protected _TopMenu As New Crownwood.Magic.Menus.PopupMenu

    Private sepGuideDelete As MenuCommand = New MenuCommand("-")
    Private sepGuideSelect As MenuCommand = New MenuCommand("-")

    Private mnuGuides As New MenuCommand("Insert Guide")


    Private mnuEdit As New MenuCommand("Edit")
    Private mnuAlign As New MenuCommand("Align")
    Private mnuZoom As New MenuCommand("Magnification")

    '-- guides --
    Private mnuNewHGuide As New MenuCommand("Insert Horiz Guide", AddressOf OnInsertHGuide)
    Private mnuNewVGuide As New MenuCommand("Insert Vert Guide", AddressOf OnInsertVGuide)


    '-- Edit menu --
    Private mnuCut As New MenuCommand("Cu&t", Shortcut.CtrlX, AddressOf OnCut)
    Private mnuCopy As New MenuCommand("&Copy", Shortcut.CtrlC, AddressOf OnCopy)
    Private mnuPaste As New MenuCommand("&Paste", Shortcut.CtrlV, AddressOf OnPaste)
    Private mnuDelete As New MenuCommand("&Delete", Shortcut.Del, AddressOf OnDelete)
    Private mnuSelectAll As New MenuCommand("Select &All", Shortcut.CtrlA, AddressOf OnSelectAll)
    Private mnuDeselect As New MenuCommand("Deselect", Shortcut.CtrlShiftA, AddressOf OnDeselectAll)

    '--View menu
    Private mnuZoomIn As New MenuCommand("Zoom &In", AddressOf OnZoomIn)
    Private mnuZoomOut As New MenuCommand("Zoom &Out", AddressOf OnZoomOut)

    Private mnuSense As New MenuCommand("Focused Entry")
    Private mnuGuideSelect As New MenuCommand("Select Guide", AddressOf OnSelectGuide)
    Private mnuGuideDelete As New MenuCommand("Delete Guide", AddressOf onDeleteGuide)
    Private mnuRevertSize As New MenuCommand("Revert To Original Size", AddressOf Onrevert)

    '-- Zoom sub menu
    Private mnuZoom25 As New MenuCommand("25%", AddressOf OnMag25)
    Private mnuZoom50 As New MenuCommand("50%", AddressOf OnMag50)
    Private mnuZoom75 As New MenuCommand("75%", AddressOf OnMag75)
    Private mnuZoom100 As New MenuCommand("100%", AddressOf OnMag100)
    Private mnuZoom150 As New MenuCommand("150%", AddressOf OnMag150)
    Private mnuZoom200 As New MenuCommand("200%", AddressOf OnMag200)
    Private mnuZoom500 As New MenuCommand("500%", AddressOf OnMag500)


    'Alignment
    Private mnuAlignLeft As New MenuCommand("Lefts", AddressOf OnAlignLeft)
    Private mnuCenterVert As New MenuCommand("Center Vertical", AddressOf OnCenterVert)
    Private mnuAlignRight As New MenuCommand("Rights", AddressOf OnAlignRight)

    Private mnuAlignTop As New MenuCommand("Tops", AddressOf OnAlignTop)
    Private mnuCenterHoriz As New MenuCommand("Center Horizontal", AddressOf OnCenterHoriz)

    Private mnuAlignBottom As New MenuCommand("Bottoms", AddressOf OnAlignBottom)
    Private mnuDistributeWidths As New MenuCommand("Distribute Widths", AddressOf OnDistWidths)
    Private mnuDistributeHeights As New MenuCommand("Distribute Heights", AddressOf OnDistHeights)


    Private mnuSendToBack As New MenuCommand("&Send to Back", AddressOf OnSendToBack)
    Private mnuBringToFront As New MenuCommand("&Bring to Front", AddressOf OnBringToFront)

    Private mnuSendBackward As New MenuCommand("Send Backward", AddressOf OnSendBackward)
    Private mnuBringForward As New MenuCommand("Bring Forward", AddressOf OnBringForward)


    'Quick Code
    Private mnuQuickCode As New MenuCommand("Quick Code", AddressOf OnQuickViewCode)
    Private mnuQuickSVG As New MenuCommand("Quick SVG", AddressOf OnQuickSVG)
    Private mnuSVGBrowser As New MenuCommand("View SVG in Browser", AddressOf OnSVGBrowser)

#End Region


#Region "Constructor"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the SurfaceContextMenu object.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        Trace.WriteLineIf(App.TraceOn, "ContextMenu.New")
        createMenus()
    End Sub

#End Region

#Region "Public Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Shows the context menu at the point specified bt pt.
    ''' </summary>
    ''' <param name="pt">The point the mouse was invoked at to show the context menu.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub show(ByVal pt As Point)

        refreshmenuView()

        If IntPtr.op_Equality(_TopMenu.Handle, System.IntPtr.Zero) Then
            _TopMenu.TrackPopup(pt)
        Else
            _TopMenu.DestroyHandle()
            _TopMenu.TrackPopup(pt)
        End If

    End Sub

#End Region

#Region "Menu creation"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates the menu items used in the surface context menu.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub createMenus()
        buildMenuHierarchy()
        recreateSenseMenu()
        _TopMenu.MenuCommands.AddRange(New MenuCommand() {mnuGuideSelect, sepGuideSelect, mnuGuideDelete, sepGuideDelete, mnuGuides, mnuEdit, mnuZoom, mnuAlign, sep, mnuBringToFront, mnuSendToBack, sep, mnuBringForward, mnuSendBackward, sep, mnuQuickCode, sep, mnuQuickSVG, mnuSVGBrowser})
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates the menu hierarchy items that fall under the main menu types.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub buildMenuHierarchy()
        mnuEdit.MenuCommands.AddRange(New MenuCommand() _
        {mnuCut, mnuCopy, mnuPaste, mnuDelete, sep, _
        mnuSelectAll, mnuDeselect})

        mnuGuides.MenuCommands.AddRange(New MenuCommand() {mnuNewHGuide, mnuNewVGuide})
        mnuZoom.MenuCommands.AddRange(New MenuCommand() _
            {mnuZoomIn, mnuZoomOut, sep, mnuZoom25, mnuZoom50, mnuZoom75, mnuZoom100, mnuZoom150, mnuZoom200, mnuZoom500})

        mnuAlign.MenuCommands.AddRange(New MenuCommand() _
{mnuAlignLeft, mnuCenterVert, mnuAlignRight, sep, mnuAlignTop, mnuCenterHoriz, mnuAlignBottom, sep, mnuDistributeWidths, mnuDistributeHeights})


    End Sub

#End Region

#Region "Menu Refresh"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Recreates the context sensitive menu based upon the current selection.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub recreateSenseMenu()

        If Not MDIMain.ActiveDocument Is Nothing Then

            If _TopMenu.MenuCommands.Contains(mnuSense) Then
                _TopMenu.MenuCommands.Remove(mnuSense)
            End If


            If Not mnuSense Is Nothing Then
                mnuSense.Dispose()
                mnuSense = Nothing
            End If

            If TypeOf MDIMain.ActiveDocument.Selected.LastSelected Is GDIImage Then
                mnuSense = New MenuCommand("Image Options")
                mnuSense.MenuCommands.AddRange(New MenuCommand() {mnuRevertSize})
                _TopMenu.MenuCommands.AddRange(New MenuCommand() {mnuSense})
                Return
            End If

            If Not App.ToolManager.SelectedGuide Is Nothing Then
                mnuGuideSelect.Visible = True
                sepGuideSelect.Visible = True
                mnuGuideDelete.Visible = True
                sepGuideDelete.Visible = True
            Else
                mnuGuideSelect.Visible = False
                mnuGuideDelete.Visible = False
                sepGuideDelete.Visible = False
                sepGuideSelect.Visible = False
            End If
        End If

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Enables and disabled menu items based upon the current state of the application.
    ''' </summary>
    ''' <remarks>The requirements for enabled vary based upon the type of command.  
    ''' There are three different situations that can occur.
    ''' 1) A menu item may require a document.
    ''' 2) A menu item may require a selected item
    ''' 3) A menu item may require more than one selected item.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub refreshmenuView()
        Dim bHaveIntance As Boolean = Not MDIMain.ActiveDocument Is Nothing
        Dim bHaveSelection As Boolean = MDIMain.ActiveDocument.Selected.Count > 0
        Dim bTwoOrMoreSelected As Boolean = MDIMain.ActiveDocument.Selected.Count > 1

        mnuZoomIn.Enabled = bHaveIntance
        mnuZoomOut.Enabled = bHaveIntance
        mnuZoom.Enabled = bHaveIntance

        mnuCut.Enabled = bHaveSelection
        mnuCopy.Enabled = bHaveSelection
        mnuPaste.Enabled = bHaveIntance
        mnuDelete.Enabled = bHaveSelection
        mnuSelectAll.Enabled = bHaveIntance
        mnuDeselect.Enabled = bHaveSelection
        mnuSendToBack.Enabled = bHaveSelection
        mnuBringToFront.Enabled = bHaveSelection

        mnuSendBackward.Enabled = bHaveSelection
        mnuBringForward.Enabled = bHaveSelection


        mnuQuickCode.Enabled = bHaveSelection
        mnuQuickSVG.Enabled = bHaveSelection
        mnuSVGBrowser.Enabled = bHaveSelection

        mnuAlign.Enabled = bTwoOrMoreSelected OrElse _
        (bHaveSelection AndAlso _
        (Not App.AlignManager.AlignMode = EnumAlignMode.eNormal))

        recreateSenseMenu()

    End Sub


#End Region
#Region "Menu Event Handlers"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a paste command from the clip board
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub OnPaste(ByVal s As Object, ByVal e As EventArgs)
        Dim d As IDataObject = Clipboard.GetDataObject()

        If d.GetDataPresent(DataFormats.Bitmap) Then

        ElseIf d.GetDataPresent(GDIObjCol.Format.Name) Then

            MDIMain.ActiveDocument.Paste(d)

        ElseIf d.GetDataPresent(DataFormats.Text) Then

        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the quick code command.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub OnQuickViewCode(ByVal s As Object, ByVal e As EventArgs)

        Dim sQuickCode As String = MDIMain.ActiveDocument.QuickCode
        Dim dgQuickCode As New dgQuickCode(sQuickCode)
        dgQuickCode.ShowDialog()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the Quick SVG command.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub OnQuickSVG(ByVal s As Object, ByVal e As EventArgs)

        Dim sQuickXML As String = MDIMain.ActiveDocument.QuickSVG
        Dim dgQuickCode As New dgQuickCode(sQuickXML)
        dgQuickCode.ShowDialog()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handle the SVG "browse" command
    ''' </summary>
    ''' <param name="s"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub OnSVGBrowser(ByVal s As Object, ByVal e As EventArgs)

        Dim sQuickXML As String = MDIMain.ActiveDocument.QuickSVG

        Dim sTemp As String = IO.Path.GetTempFileName() & ".svg"

        Dim fs As New System.IO.FileStream(sTemp, IO.FileMode.Create)

        If fs.CanWrite Then
            Dim writer As New System.IO.StreamWriter(fs)
            writer.Write(sQuickXML)

            writer.Close()
            fs.Close()

            Dim procInfo As New ProcessStartInfo("iexplore.exe", sTemp)
            Dim iexplore As Process = Process.Start(procInfo)

        End If


    End Sub

    Private Sub OnInsertHGuide(ByVal s As Object, ByVal e As EventArgs)
        App.MDIMain.InsertHGuide()
    End Sub

    Private Sub OnInsertVGuide(ByVal s As Object, ByVal e As EventArgs)
        App.MDIMain.InsertVGuide()
    End Sub

    Private Sub OnCut(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveDocument.Selected.Cut()
    End Sub

    Private Sub OnCopy(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveDocument.Selected.Copy()
    End Sub

    Private Sub OnDistWidths(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.DistributeWidths()
    End Sub

    Private Sub OnDistHeights(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.DistributeHeights()
    End Sub

    Private Sub OnUndo(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveDocument.undo()
    End Sub

    Private Sub OnRedo(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveDocument.redo()
    End Sub

    Private Sub OnZoomIn(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom *= 2
    End Sub

    Private Sub Onrevert(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveDocument.revertSelectedImageSize()
    End Sub

    Private Sub onDeleteGuide(ByVal s As Object, ByVal e As EventArgs)

        MDIMain.ActiveDocument.Guides.Remove(App.ToolManager.SelectedGuide)
        App.ToolManager.SelectedGuide = Nothing
    End Sub


    Private Sub OnSelectGuide(ByVal s As Object, ByVal e As EventArgs)
        App.ToolManager.ActiveToolMode = ToolManager.TM.eMovingGuide
    End Sub
    Private Sub OnZoomOut(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom /= 2
    End Sub


    Private Sub OnDelete(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveDocument.Selected.DeleteAll()
    End Sub

    Private Sub OnSendToBack(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveDocument.Selected.SendToBack()
    End Sub


    Private Sub OnSendBackward(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveDocument.Selected.SendBackward()
    End Sub


    Private Sub OnBringForward(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveDocument.Selected.BringForward()
    End Sub

    Private Sub OnBringToFront(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveDocument.Selected.BringToFront()
    End Sub

    Private Sub OnCenterVert(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.AlignVertCenter()
    End Sub

    Private Sub OnCenterHoriz(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.AlignHorizCenter()
    End Sub

    Private Sub OnAlignLeft(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.AlignLeft()
    End Sub

    Private Sub OnAlignTop(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.AlignTop()
    End Sub

    Private Sub OnAlignRight(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.AlignRight()
    End Sub

    Private Sub OnAlignBottom(ByVal s As Object, ByVal e As EventArgs)
        App.AlignManager.AlignBottom()
    End Sub


    Private Sub OnMag25(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom = 0.25
    End Sub

    Private Sub OnMag50(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom = 0.5
    End Sub

    Private Sub OnMag75(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom = 0.75
    End Sub

    Private Sub OnMag100(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom = 1
    End Sub

    Private Sub OnMag150(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom = 1.5
    End Sub

    Private Sub OnMag200(ByVal s As Object, ByVal e As EventArgs)
        App.MDIMain.ActiveZoom = 2
    End Sub

    Private Sub OnMag500(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveZoom = 5
    End Sub

    Private Sub OnDeselectAll(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveDocument.Selected.DeselectAll()
    End Sub


    Private Sub OnSelectAll(ByVal s As Object, ByVal e As EventArgs)
        MDIMain.ActiveDocument.SelectAll()
    End Sub

#End Region



End Class


