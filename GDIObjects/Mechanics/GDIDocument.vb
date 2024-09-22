

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIDocument
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' The primary GDI+ Architect document class.
''' This class is where all of the various GDI+ information winds up when it is saved 
''' to a file.  This is also the class that is rendered to the drawing surface.
''' </summary>
''' <remarks>
''' Notice that the GDIDocument inherits from CollectionBase.  Each GDIDocument contains a collection of pages.
''' For the Graphics Class variant of the GDIDocument, there is a single page.
''' </remarks>
'''  -----------------------------------------------------------------------------
<Serializable()> _
Public Class GDIDocument
    Inherits CollectionBase
    Implements IDisposable

#Region "NonSerialized Fields"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether the GDIDocument instance has been disposed
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _Disposed As Boolean = False

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Holds the set of selected objects in the GDIDocument.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
      Private _SelectedSet As SelectedSet

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A PrintDocument used when rendering the GDIDocument in print preview mode.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
        Private _PrintDocument As New System.Drawing.Printing.PrintDocument



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' When print previewing with, the current page number
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
     Private _PrintPageNumber As Int32 = 0



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Used to maintain history for undo and redo operations.
    ''' </summary>
    ''' -----------------------------------------------------------------------------   
    <NonSerialized()> _
    Private _History As History

#End Region

#Region "Local Fields"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The current page in the GDIDocument.  For graphics objects this is always the
    ''' same page since they contain only one page.  
    ''' </summary>
    ''' <remarks>This is serialized only for refreshing history in undo and redo.
    ''' When a GDIDocument is opened, it is always started at the first page.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private _CurrentPage As GDIPage



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Boolean indicating for the PrintDocument type GDIDocument where to display and  
    ''' print in Portrait (true) or landscape (false).
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Portrait As Boolean



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The GDIDocument's ExportSettings (see the ExportSettings class for more information)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _ExportSettings As ExportSettings

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The current Text hint of the GDIDocument. This maps to a 
    ''' Drawing.Text.TextRenderingHint.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _TextHint As Drawing.Text.TextRenderingHint

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The SmoothingMode of the GDIDocument.  This maps to a Drawing2D.SmoothingMode.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _SmoothingMode As Drawing2D.SmoothingMode


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The name of the binary GDI+ Architect file.  Set after a save or load.
    ''' Initially empty for new files.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _FileName As String = String.Empty
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The path of the binary GDI+ Architect file. Set after a save or load.  
    ''' Initially empty for new files.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Path As String = String.Empty

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A Boolean indicating if the GDIDocument is dirty or not.  This is set to true every 
    ''' time history is changed, and is set to false after a save operation.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Dirty As Boolean = False
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether the GDIDocument has ever been saved or not.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Saved As Boolean = False

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The next valid page number for new pages added to this GDIDocument.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _NextPageIndex As Int32 = 1

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A rectangle representing the entire page size of the GDIDocument.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _PageArea As Rectangle

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A rectangle representing the printable area of the page (portion inside margins)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _PrintArea As Rectangle



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The GDIDocument's back color.  This isn't actually exported, but is cosmetic to help the design process.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _BackColor As Color = Color.White

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An arraylist containing the GDIGuides that belong to this GDIDocument. 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Guides As New ArrayList(10)





#End Region

#Region "Delegate Declarations and Instances"
    'Invoked when the content of a page is changed

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invoked when a new page becomes the current page.  Used to 
    ''' notify interested subscribers that a new page has become this 
    ''' GDIDocument's current page.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Delegate Sub OnPageChanged(ByVal sender As Object, ByVal e As EventArgs)
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invoked when the GDIDocument history changes.  Used to notify interested 
    ''' subscribers of changes to the state of the GDIDocument.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Delegate Sub OnHistoryUpdated(ByVal s As Object, ByVal e As EventArgs)

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invoked when the set of selected objects (The SelectedSet) changes.  Used to 
    ''' notify subscribers of changes to what is selected within a GDIDocument.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Delegate Sub OnSelectionChanged(ByVal sender As Object, ByVal e As EventArgs)

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Holds the callbacks for when a new page becomes the current page. 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
      Private _PageChangedCallbacks As OnPageChanged


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Holds the callbacks for when history changes
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _HistoryChangedCallbacks As OnHistoryUpdated
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Holds the callbacks for when the current selection changes
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _SelectionChangedCallbacks As OnSelectionChanged

#End Region


#Region "Constructors"

    ' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructor for the graphics class style of GDIDocument
    ''' </summary>
    ''' <param name="rectPixSize">The size requested for the graphics class object</param>
    ''' <param name="ExportSettings">The current GDI+ export settings defined in the User interface</param>
    ''' <param name="TextHint">The current GDI+ text hint settings defined in the user interface</param>
    ''' <param name="SmoothingMode">The current GDI+ smoothing mode defined in the user interface</param>
    ''' <remarks>This constructor variant is used when the user elects to create a new graphics class instead of a 
    ''' page style GDIDocument.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal rectPixSize As Rectangle, ByVal ExportSettings As ExportSettings, ByVal TextHint As System.Drawing.Text.TextRenderingHint, ByVal SmoothingMode As Drawing2D.SmoothingMode)
        Trace.WriteLineIf(Session.Tracing, "Creating new graphics class - " & rectPixSize.ToString)


        _PageArea = rectPixSize
        _PrintArea = rectPixSize
        _SmoothingMode = SmoothingMode
        _TextHint = TextHint
        _ExportSettings = ExportSettings
        _SelectedSet = New SelectedSet(Me)

        AddHandler _SelectedSet.SelectionChanged, AddressOf SelectedSet_OnSelectionChanged

        Me.addNewPage()
        Me.invokePageChanged()

        ''create a history entry at the start point 
        _History = New History
        Me.recordHistory("New File")

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructor for the PrintDocument style of GDIDocument
    ''' </summary>
    ''' <param name="oprintersettings">Current printer settings defined in the user interface.</param>
    ''' <param name="opagesettings">The current page settings defined in the user interface.</param>
    ''' <param name="exportsettings">The current GDI+ export settings defined in the User interface</param>
    ''' <param name="TextHint">The current GDI+ text hint settings defined in the user interface</param>
    ''' <param name="SmoothingMode">The current GDI+ smoothing mode defined in the user interface</param>
    ''' <remarks>This constructor is for when the user elects to create a new print style GDIDocument 
    ''' (repeating pages of content).
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal oprintersettings As Drawing.Printing.PrinterSettings, _
                 ByVal opagesettings As Drawing.Printing.PageSettings, _
                 ByVal exportsettings As ExportSettings, _
                 ByVal TextHint As System.Drawing.Text.TextRenderingHint, _
                 ByVal SmoothingMode As Drawing2D.SmoothingMode)

        Trace.WriteLineIf(Session.Tracing, "Creating new print document")
        Trace.Indent()
        Trace.WriteLineIf(Session.Tracing, "Margins:" & opagesettings.Margins.ToString)
        Trace.Unindent()

        _SelectedSet = New SelectedSet(Me)
        AddHandler _SelectedSet.SelectionChanged, AddressOf SelectedSet_OnSelectionChanged
        _SmoothingMode = SmoothingMode
        _TextHint = TextHint

        _ExportSettings = exportsettings

        SetResoByPrinterSettings(oprintersettings)
        SetResoByPageSettings(opagesettings)

        Me.addNewPage()
        Me.invokePageChanged()

        ''create a history entry at the start point 
        _History = New History
        Me.recordHistory("New File")

    End Sub

#End Region

#Region "Misc Methods"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the resolution of the GDIDocument based on PageSettings.
    ''' </summary>
    ''' <param name="opagesettings">A PageSettings object to base settings off of.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub SetResoByPageSettings(ByVal opagesettings As Printing.PageSettings)
        Trace.WriteLineIf(Session.Tracing, "Setting Page Resolution")

        'Bounds is in 100th's of an inch.  So 8.5 x 11 would be 850 x 11000
        'This is the total area of the Instance, including margin settings
        Dim bounds As Rectangle = opagesettings.Bounds

        'Number of dots per inch
        Dim horizRes As Integer = opagesettings.PrinterResolution.X
        Dim vertRes As Integer = opagesettings.PrinterResolution.Y


        'TotalPageSize is set to the complete bounds of the pagesettings
        _PageArea = bounds
        _Dirty = True

        'The PrintableArea is set to the total Instance minus the margin sizes
        PrintableArea = _
            New Rectangle(bounds.Left + opagesettings.Margins.Left, _
                bounds.Top + opagesettings.Margins.Top, _
                bounds.Width - opagesettings.Margins.Left - opagesettings.Margins.Right, _
                bounds.Height - opagesettings.Margins.Top - opagesettings.Margins.Bottom)

        _PrintDocument.OriginAtMargins = False

        If opagesettings.Landscape Then
            _Portrait = False
            _PrintDocument.DefaultPageSettings.Landscape = True
        Else
            _Portrait = True
            _PrintDocument.DefaultPageSettings.Landscape = False
        End If


        _PrintDocument.DefaultPageSettings = opagesettings

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the GDIDOcument properties based on a PrinterSettings object.
    ''' </summary>
    ''' <param name="oPrinterSettings">The printer settings to apply the GDIDocument.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub SetResoByPrinterSettings(ByVal oPrinterSettings As Printing.PrinterSettings)
        _PrintDocument.PrinterSettings = oPrinterSettings
        Me.SetResoByPageSettings(_PrintDocument.DefaultPageSettings)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a guide to the current GDIDocument
    ''' </summary>
    ''' <param name="direction">The direction of the guide (horizontal or vertical)</param>
    ''' <param name="xy">The XY position.  Whether this is an x or y value depends on 
    ''' the value of the direction param.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub InsertGuide(ByVal direction As GDIGuide.EnumGuideDirection, ByVal xy As Int32)
        Dim gd As GDIObjects.GDIGuide = New GDIObjects.GDIGuide

        gd.Direction = direction
        gd.XY = xy

        Me._Guides.Add(gd)

    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Reverts a selected image on the surface to its original sized dimensions.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub revertSelectedImageSize()
        Trace.WriteLineIf(Session.Tracing, "Reverting image size")
        If TypeOf Selected.LastSelected Is GDIObjects.GDIImage Then
            Dim gdiImageItem As GDIObjects.GDIImage = DirectCast(Selected.LastSelected, GDIObjects.GDIImage)
            gdiImageItem.RevertSize()
            Me.recordHistory("Image size reverted")
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if a name is already in use.  Used to avoid name conflicts.
    ''' </summary>
    ''' <param name="selectedobj">The object being compared for name violations</param>
    ''' <param name="sSuggestedName">The potential name to give this object.</param>
    ''' <returns>A Boolean indicating whether a collision has occurred</returns>
    ''' <remarks>Notice the "If not selectedobj" clause in the code.  This verifies that a self comparison
    ''' is not taking place.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Function NameExists(ByVal selectedobj As GDIObject, ByVal sSuggestedName As String) As Boolean
        Trace.WriteLineIf(Session.Tracing, "Checking if name exists: " & sSuggestedName)

        For Each pg As GDIPage In Me.InnerList
            For Each obj As GDIObject In pg.GDIObjects
                If Not selectedobj Is obj Then
                    If LCase(obj.Name.Trim) = LCase(sSuggestedName.Trim) Then
                        Return True
                    End If
                End If
            Next
        Next

        Return False
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if a name is in use anywhere inside a particular GDIDocument.
    ''' </summary>
    ''' <param name="sSuggestedName">The potential name to give this object.</param>
    ''' <returns>A Boolean indicating whether the name is in use anywhere.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function NameExists(ByVal sSuggestedName As String) As Boolean
        Trace.WriteLineIf(Session.Tracing, "Checking if name exists: " & sSuggestedName)

        For Each pg As GDIPage In Me.InnerList
            For Each obj As GDIObject In pg.GDIObjects
                If LCase(obj.Name.Trim) = LCase(sSuggestedName.Trim) Then
                    Return True
                End If
            Next
        Next

        Return False
    End Function

#End Region

#Region "Delegate implementation"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Notifies interested listeners that the set of selected objects has changed.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub SelectedSet_OnSelectionChanged(ByVal s As Object, ByVal e As EventArgs)
        If Not _SelectionChangedCallbacks Is Nothing Then
            _SelectionChangedCallbacks(Me, EventArgs.Empty)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a listener to the OnSelectionChanged delegate.
    ''' </summary>
    ''' <param name="cb">The callback to add</param>
    ''' -----------------------------------------------------------------------------
    Public Sub setSelectionChangedCallBack(ByVal cb As OnSelectionChanged)
        If _SelectionChangedCallbacks Is Nothing Then
            _SelectionChangedCallbacks = cb
        Else
            _SelectionChangedCallbacks = CType(System.Delegate.Combine(_SelectionChangedCallbacks, cb), OnSelectionChanged)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Removes a listener from the OnSelectionChanged delegate.
    ''' </summary>
    ''' <param name="cb">The callback to remove</param>
    ''' -----------------------------------------------------------------------------
    Public Sub removeSelectionChangedCallBack(ByVal cb As OnSelectionChanged)
        If Not _SelectionChangedCallbacks Is Nothing Then
            _SelectionChangedCallbacks = CType(System.Delegate.Remove(_SelectionChangedCallbacks, cb), OnSelectionChanged)
        End If
    End Sub
#End Region


#Region "Code Generation related methods"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the code represented by this GDIDocument
    ''' </summary>
    ''' <returns>A string of code in the current language that will render this GDIDocument</returns>
    ''' <remarks>There are two separate parent code generation classes, ExportPrint and ExportClass.
    ''' ExportPrint is for the multipage PrintDocument scenario.  
    ''' ExportClass is for the graphics class style of document.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Function GenerateCode() As String

        Dim export As BaseExport
        Select Case Me.ExportSettings.DocumentType
            Case EnumDocumentTypes.ePrintDocument
                export = New ExportPrint(Me)
            Case EnumDocumentTypes.eClass
                export = New ExportClass(Me)
        End Select

        Return export.Generate()

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the quick code for the selected object
    ''' </summary>
    ''' <returns>A string of code in the current language for the selected object.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function QuickCode() As String

        If Me.Selected.Count > 0 Then
            Return Me.Selected.QuickCode()
        Else
            Return String.Empty
        End If
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the SVG code for the currently selected object.
    ''' </summary>
    ''' <returns>A string of SVG code for the currently selected object.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function QuickSVG() As String

        If Me.Selected.Count > 0 Then
            Return Me.Selected.QuickSVG()
        Else
            Return String.Empty
        End If
    End Function

#End Region

#Region "Drawing Related Methods"


 

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders guides to the current view at the current scale.
    ''' </summary>
    ''' <param name="g">The current graphics object</param>
    ''' <param name="fScale">The current scaled view of the surface.</param>
    ''' <remarks>Notice that each guide is asked to draw itself.  This technique is used
    ''' frequently inside GDI+ Architect.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub DrawGuides(ByVal g As Graphics, ByVal fScale As Single)
        Session.GraphicsManager.BeginScaleMode(g, fScale)
        'Ask each item to draw itself to the graphic container
        If _Guides.Count > 0 Then
            For i As Int32 = 0 To _Guides.Count - 1
                DirectCast(_Guides(i), GDIGuide).Draw(g, fScale)
            Next
        End If

        Session.GraphicsManager.EndScaleMode(g)

    End Sub

#End Region


#Region "Delegate and Callback Related methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets up a call back when history has been changed.
    ''' </summary>
    ''' <param name="cb">The call back to add.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub setHistoryChangedCallBack(ByVal cb As OnHistoryUpdated)
        If _HistoryChangedCallbacks Is Nothing Then
            _HistoryChangedCallbacks = cb
        Else
            _HistoryChangedCallbacks = CType(System.Delegate.Combine(_HistoryChangedCallbacks, cb), OnHistoryUpdated)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Removes a listener from the history changed event
    ''' </summary>
    ''' <param name="cb">The call back to remove</param>
    ''' -----------------------------------------------------------------------------
    Public Sub removeHistoryChangedCallBack(ByVal cb As OnHistoryUpdated)
        If Not _HistoryChangedCallbacks Is Nothing Then
            _HistoryChangedCallbacks = CType(System.Delegate.Remove(_HistoryChangedCallbacks, cb), OnHistoryUpdated)
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Wraps a call to the invokeListPageChanged delegate.  This delegate lets listeners
    ''' know that a page change event has occurred.
    ''' </summary>
    ''' <remarks>Any number of the GDI+ Architect user interface panels may care about this 
    ''' event.  The delegate architecture lets any number of interested listeners know
    ''' about this event.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Sub invokePageChanged()
        Trace.WriteLineIf(Session.Tracing, "Invoking Page Change")

        If Not _PageChangedCallbacks Is Nothing Then
            _PageChangedCallbacks(Me, New EventArgs)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a listener to the invokeListPageChanged event.
    ''' </summary>
    ''' <param name="cb">The call back function that should be notified when this event occurs.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub setPageChangeCallBack(ByVal cb As OnPageChanged)
        If _PageChangedCallbacks Is Nothing Then
            _PageChangedCallbacks = cb
        Else
            _PageChangedCallbacks = CType(System.Delegate.Combine(_PageChangedCallbacks, cb), OnPageChanged)
        End If

    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Removes a listener from the Pagechanged callback set
    ''' </summary>
    ''' <param name="cb">The call back to remove.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub removePageChangeCallBack(ByVal cb As OnPageChanged)
        If Not _PageChangedCallbacks Is Nothing Then
            _PageChangedCallbacks = CType(System.Delegate.Remove(_PageChangedCallbacks, cb), OnPageChanged)
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Provides a wrapper for invoking call backs to listeners waiting for a history changed event.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub invokeHistoryChangedCallBacks()
        Trace.WriteLineIf(Session.Tracing, "Invoking History Callbacks")
        If Not _HistoryChangedCallbacks Is Nothing Then
            _HistoryChangedCallbacks(Me, EventArgs.Empty)
        End If

        _Dirty = True
    End Sub

#End Region



#Region "GDIPage Operations"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a new page to the GDIDocument.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub addNewPage()
        Trace.WriteLineIf(Session.Tracing, "Adding new page")
        Dim pg As New GDIObjects.GDIPage(_NextPageIndex)

        Me.InnerList.Add(pg)
        Me.CurrentPage = pg
        refreshPageNumbers()

        _NextPageIndex += 1

        If Not _History Is Nothing Then
            Me.recordHistory(pg.Name & " Added")
        End If
    End Sub





    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the current page given a page number.
    ''' </summary>
    ''' <param name="iPageNum">The page number of the page to make the current page.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub SetCurrentPage(ByVal iPageNum As Int32)
        For Each pg As GDIPage In Me.InnerList
            If pg.PageNum = iPageNum Then
                Me.CurrentPage = pg
                Return
            End If
        Next
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renumbers pages when a page is added or removed.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub refreshPageNumbers()
        For i As Int32 = 0 To Me.Count - 1
            Me(i).PageNum = i
        Next
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the name of a page to a new name.  This should be used when a page is being renamed 
    ''' rather than setting it explicitly on the page object so history is recorded.
    ''' </summary>
    ''' <param name="pg">The page to apply the new name to.</param>
    ''' <param name="name">The new name to apply to the page.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub renamePage(ByVal pg As GDIPage, ByVal name As String)
        Trace.WriteLineIf(Session.Tracing, "Renaming page")
        For Each opg As GDIPage In Me.InnerList
            If pg Is opg Then
                Dim tempname As String = pg.Name
                pg.Name = name
                Me.recordHistory(tempname & " renamed to " & name)

                Return

            End If
        Next


    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deletes a page from a GDIDocument.
    ''' </summary>
    ''' <param name="pg">The page to remove from set</param>
    ''' -----------------------------------------------------------------------------
    Public Sub deletePage(ByVal pg As GDIPage)
        Trace.WriteLineIf(Session.Tracing, "Deleting page")



        Me.InnerList.Remove(pg)
        If Me.CurrentPage Is pg Then
            Me.CurrentPage = DirectCast(Me.InnerList.Item(0), GDIPage)
        End If

        Me.recordHistory(pg.Name & " deleted")

        Me.refreshPageNumbers()

    End Sub

#End Region

#Region "Clipboard paste related methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a paste operation to the GDIDocument.
    ''' </summary>
    ''' <param name="pasteCol">A collection of GDIObjects to paste to the current page.</param>
    ''' <remarks>Notice that rather than copy objects, each goes through a deserialize process.
    ''' This is to detect name collisions and verify that external resources exist.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Sub HandlePaste(ByVal pasteCol As GDIObjCol)
        Trace.WriteLineIf(Session.Tracing, "Handling paste of objects")


        For Each gdiobj As GDIObject In pasteCol
            gdiobj.deserialize(Me, Me.CurrentPage)
        Next

        Me.CurrentPage.GDIObjects.AddRange(pasteCol)

        Selected.SetSelection(pasteCol)

        Me.recordHistory("Paste")
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a Paste command to this GDIDocument.  (from a paste or control V)
    ''' </summary>
    ''' <param name="d">The object being pasted to the GDIDocument</param>
    ''' <remarks>The first thing this method does is determine if it's one GDI+ Architect can 
    ''' handle.  This could be expanded to allow all sorts of pasting to the surface 
    ''' from various sources, but at this time it only allows for pasting of GDIObjects 
    ''' to the surface.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub Paste(ByVal d As System.Windows.Forms.IDataObject)
        Trace.WriteLineIf(Session.Tracing, "Begining Paste")
        Dim pasteSet As Object = d.GetData(GDIObjects.GDIObjCol.Format.Name, False)

        If Not pasteSet Is Nothing Then
            Dim gdiCol As GDIObjects.GDIObjCol = _
             DirectCast(pasteSet, GDIObjects.GDIObjCol)
            Me.HandlePaste(gdiCol)
        End If
    End Sub

#End Region

#Region "Selection related methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Select a specific single item in the GDIDocument.  This is only used by the property panel.
    ''' </summary>
    ''' <param name="obj">The object to select</param>
    ''' -----------------------------------------------------------------------------
    Public Sub SelectOneItem(ByVal obj As GDIObject)
        Trace.WriteLineIf(Session.Tracing, "Selecting one item")

        If Not obj Is Nothing Then
            For Each pg As GDIPage In Me.InnerList
                For Each tempobj As GDIObject In pg.GDIObjects
                    If tempobj Is obj Then
                        Me.CurrentPage = pg
                        Me.Selected.SelectOneItem(obj)
                        Return
                    End If
                Next
            Next
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Selects all objects on the current page.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub SelectAll()
        Trace.WriteLineIf(Session.Tracing, "Selecting all")
        Selected.SetSelection(Me.CurrentPage.GDIObjects)
    End Sub

#End Region


#Region "History, Undo, and redo methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populates a windows list with this GDIDocument's history information.  By providing 
    ''' this method, it allows history to stay in scope of the GDIObjects project.
    ''' </summary>
    ''' <param name="lst">The list to populate with history information</param>
    ''' -----------------------------------------------------------------------------
    Public Sub PopulateHistoryList(ByVal lst As System.Windows.Forms.ListBox)
        For i As Int32 = 0 To _History.Count - 1
            Dim histitem As HistoryItem = _History(i)
            lst.Items.Add(histitem.ToString)
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Moves one step forward in history if that step exists (a redo)
    ''' </summary>
    ''' <remarks>Notice that the history changed call backs are invoked in response to a redo.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub redo()
        Trace.WriteLineIf(Session.Tracing, "Invoking Redo")

        _History.redo()
        Me.refreshHistoricView()
        Me.invokeHistoryChangedCallBacks()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a value indicating if this GDIDocument has redo information
    ''' </summary>
    ''' <value>A Boolean indicating if redo information is available.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property HasRedo() As Boolean
        Get
            Return _History.CurPos < _History.Count - 1
        End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a value indicating if this GDIDocument has undo information
    ''' </summary>
    ''' <value>A Boolean indicating if undo information is available.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property HasUndo() As Boolean
        Get
            Return _History.CurPos > 0
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the current history position.  This is intended to be used instead of directly 
    ''' setting the history object's position so callbacks are invoked.
    ''' </summary>
    ''' <param name="iIndex">The new index history to make current.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub setHistoryPos(ByVal iIndex As Int32)
        _History.CurPos = iIndex
        Me.refreshHistoricView()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Records an instance of history information to the history object.
    ''' </summary>
    ''' <param name="sLabel">A textual description of the history point.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub recordHistory(ByVal sLabel As String)
        _History.recordHistory(Me, sLabel)
        Me.invokeHistoryChangedCallBacks()
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Loads the current GDIDocument from history and makes each page of the historic GDIDocument 
    ''' belong to the current GDIDocument, in effect, undoing or redoing operations.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub refreshHistoricView()
        Dim curItem As HistoryItem = _History.CurHistItem

        Dim tempdoc As GDIDocument = curItem.HistoricDocument


        If Not tempdoc Is Nothing Then

            Me.Clear()

            For Each pg As GDIPage In tempdoc

                For i As Int32 = pg.GDIObjects.Count - 1 To 0 Step -1
                    Dim obj As GDIObject = pg.GDIObjects.Item(i)
                    Dim bOk As Boolean = obj.deserialize(Me, pg)
                    If Not bOk Then
                        MsgBox("Could not load object " & obj.Name)
                        pg.GDIObjects.Remove(obj)
                    End If
                Next

                Me.InnerList.Add(pg)

                If pg.PageNum = tempdoc.CurrentPage.PageNum Then
                    Me.CurrentPage = pg
                End If
            Next

        End If


        Me.invokeHistoryChangedCallBacks()

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Undoes a single history step.
    ''' </summary>
    ''' <remarks>This call wraps the history object's undo method.  It then refreshes this 
    ''' history and redraws the GDIDocument to the current surface.  Finally it invokes 
    ''' the historychanged callback for listeners.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub undo()
        Trace.WriteLineIf(Session.Tracing, "Undoing step")
        _History.undo()
        Me.refreshHistoricView()
        Me.invokeHistoryChangedCallBacks()
    End Sub
#End Region

#Region "Print preview methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Used to render print preview of the GDIDocument.  
    ''' </summary>
    ''' <param name="sender">invoker</param>
    ''' <param name="e">PrintPageEventArgs</param>
    ''' <remarks>This is attached to the onPrintPage 
    ''' handler of the printdocument object contained in this project.  Each time it is called 
    ''' it  asks the current page to print itself and then increments if there are more pages.
    ''' When it is out of pages this method sets the PrintPageEventArgs.HasMorePages argument 
    ''' to false indicating printing has been completed.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Sub prDoc_PrintPage(ByVal sender As System.Object, _
              ByVal e As System.Drawing.Printing.PrintPageEventArgs)

        Trace.WriteLineIf(Session.Tracing, "Invoking print page")
        Dim g As Graphics = e.Graphics

        g.TextRenderingHint = Me.TextRenderingHint
        g.SmoothingMode = Me.SmoothingMode

        Me(_PrintPageNumber).Print(g)

        If _PrintPageNumber < Me.InnerList.Count - 1 Then

            e.HasMorePages = True
            _PrintPageNumber += 1
        Else
            _PrintPageNumber = 0
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invokes a print preview of the current GDIDocument to a preview window.
    ''' </summary>
    ''' <remarks>Most of this rather long method is setting up a Preview dialogue to show 
    ''' the print preview.  There is one thing to note, however.  Immediately before the 
    ''' preview, this method turns off the draw borders and stores their original values.
    ''' 
    ''' This technique allows the application to use virtually the same code 
    ''' for drawing to print surface as everywhere else in the application.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub Preview()

        Trace.WriteLineIf(Session.Tracing, "Invoking print preview")

        Dim dgPreview As System.Windows.Forms.PrintPreviewDialog = New System.Windows.Forms.PrintPreviewDialog

        With dgPreview
            .AutoScrollMargin = New System.Drawing.Size(0, 0)
            .AutoScrollMinSize = New System.Drawing.Size(0, 0)
            .ClientSize = New System.Drawing.Size(400, 300)
            .Enabled = True
            .Location = New System.Drawing.Point(0, 0)
            .MinimumSize = New System.Drawing.Size(375, 250)
            .Name = "dgPreview"
            .TransparencyKey = System.Drawing.Color.Empty
            .Opacity = 1
            .Visible = False
        End With

        'set session options for preview mode
        Dim tmpfldborders As Boolean = Session.Settings.DrawTextFieldBorders

        Session.Settings.DrawTextFieldBorders = False

        'AddHandler _PrintDocument.BeginPrint, AddressOf prDoc_BeginPrint
        AddHandler _PrintDocument.PrintPage, AddressOf prDoc_PrintPage

        If _Portrait Then
            _PrintDocument.DefaultPageSettings.Landscape = False
        Else
            _PrintDocument.DefaultPageSettings.Landscape = True
        End If

        dgPreview.Document = _PrintDocument

        dgPreview.ShowDialog()

        RemoveHandler _PrintDocument.PrintPage, AddressOf prDoc_PrintPage

        Session.Settings.DrawTextFieldBorders = tmpfldborders

    End Sub




#End Region



#Region "Persistence Related Methods"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Recreates a GDI+ Architect GDIDocument from a previously saved file.
    ''' </summary>
    ''' <param name="sfilename">A full path to the file.</param>
    ''' <remarks>
    ''' This method is only called from "LoadFromFile".  The reason this is separate is that LoadFromFile 
    ''' is a shared (static) method. Once the generic LoadFromFile is done with its work, this method sets up
    ''' the instance related properties.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Sub deserialize(ByVal sfilename As String)
        Dim fInfo As New System.IO.FileInfo(sfilename)

        _FileName = fInfo.Name
        _Path = fInfo.DirectoryName


        _PrintDocument = New System.Drawing.Printing.PrintDocument
        _SelectedSet = New SelectedSet(Me)
        AddHandler _SelectedSet.SelectionChanged, AddressOf SelectedSet_OnSelectionChanged
        _History = New History
        Me.CurrentPage = Me.Item(0)

        Me.recordHistory("Load File")
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Loads a previously saved GDI+ Architect GDIDocument from a saved file.
    ''' </summary>
    ''' <param name="fileName">A full path to the file</param>
    ''' <returns>A deserialized GDI+ Architect GDIDocument</returns>
    ''' -----------------------------------------------------------------------------
    Public Shared Function LoadFromFile(ByVal fileName As String) As GDIDocument
        'load objects from a binary file
        Trace.WriteLineIf(Session.Tracing, "Attempting to load file:" & fileName)

        Dim ioStream As System.IO.Stream
        Try

            Dim binSerial As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter

            ioStream = New IO.FileStream(fileName, IO.FileMode.Open)


            Dim tempdoc As GDIDocument = DirectCast(binSerial.Deserialize(ioStream), GDIDocument)

            If Not tempdoc Is Nothing Then

                For Each pg As GDIPage In tempdoc

                    For i As Int32 = pg.GDIObjects.Count - 1 To 0 Step -1
                        Dim obj As GDIObject = pg.GDIObjects.Item(i)
                        Dim bOk As Boolean = obj.deserialize(tempdoc, pg)
                        If Not bOk Then
                            MsgBox("Could not load object " & obj.Name)
                            pg.GDIObjects.Remove(obj)
                        End If
                    Next
                Next

            End If

            tempdoc.deserialize(fileName)

            tempdoc._Saved = True


            Return tempdoc


        Catch e As Exception
            Trace.WriteLineIf(Session.Tracing, "Failed to load file" & e.Message)
            Throw New ApplicationException("File failed to load.", e)

        Finally
            If Not ioStream Is Nothing Then
                ioStream.Close()
                ioStream = Nothing
            End If

        End Try

    End Function



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Saves an existing GDIDocument to its original location.  Sets dirty to false.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub Save()
        Trace.WriteLineIf(Session.Tracing, "Beginning save")
        Me.SaveToFile()
        _Dirty = False
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Saves the GDIDocument to a file, overwriting any existing file at the location.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub SaveToFile()
        'save objects to a binary file
        Dim ioStream As System.IO.Stream

        Try
            ioStream = New IO.FileStream(_Path & "\" & _FileName, IO.FileMode.Create)
            Dim myBinarySerializer As New System.Runtime.Serialization.formatters.Binary.BinaryFormatter

            myBinarySerializer.Serialize(ioStream, Me)

        Catch e As Exception
            Throw e
        Finally
            If Not ioStream Is Nothing Then
                ioStream.Close()
                ioStream = Nothing
            End If

        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Saves the GDIDocument to the given path with the given file name.
    ''' </summary>
    ''' <param name="sPath">The path to save the file to.</param>
    ''' <param name="sFileName">The filename to save the file as.</param>
    ''' <remarks>Notice this sets _Saved to true, indicating if the GDIDocument has ever been saved.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub SaveAs(ByVal sPath As String, ByVal sFileName As String)
        _FileName = sFileName
        _Path = sPath
        Me.SaveToFile()

        _Dirty = False
        _Saved = True
    End Sub
#End Region


#Region "Property Accessors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the current set of GDIGuides belonging to this document as an arraylist.
    ''' </summary>
    ''' <value>An arraylist containing this document's guides</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Guides() As ArrayList
        Get
            Return _Guides
        End Get

    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the GDIDocument's back color.
    ''' </summary>
    ''' <value>The GDIDocument's back color.</value>
    ''' <remarks>
    ''' BackColor is only used at design time and is not exported to code.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Property BackColor() As Color
        Get
            Return _BackColor
        End Get
        Set(ByVal Value As Color)
            _BackColor = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the GDIDocument's current ExportSettings.  ExportSettings are properties 
    ''' relevant to the export process.
    ''' </summary>
    ''' <value>The current ExportSettings defined for the GDIDocument.</value>
    ''' <remarks>Export settings are initially set in the constructor.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ExportSettings() As ExportSettings
        Get
            Return _ExportSettings
        End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the printable area of the GDIDocument (the area inside the margins) 
    ''' </summary>
    ''' <value>The printable area of the page expressed as a rectangle</value>
    ''' -----------------------------------------------------------------------------
    Public Property PrintableArea() As Rectangle
        'bounds in 1/100ths of an inch, just like the printer objects use
        Get
            Return _PrintArea
        End Get
        Set(ByVal Value As Rectangle)
            _PrintArea = Value
            _Dirty = True
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the filename that this GDIDocument will be saved with.
    ''' </summary>
    ''' <value>A string containing the file name.</value>
    ''' <remarks>This also sets the exportsetting's property class name if it has not previously 
    ''' been set.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Property FileName() As String
        Get
            Return _FileName
        End Get
        Set(ByVal Value As String)

            Trace.WriteLineIf(Session.Tracing, "Assigning file name")
            _FileName = Value
            If ExportSettings.ClassName.Length = 0 Then
                ExportSettings.ClassName = _FileName
            End If

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the GDIPage at index within the GDIDocument's collection.
    ''' </summary>
    ''' <param name="index">The index at which to retrieve a page.</param>
    ''' <value>A GDIPage at the corresponding index.</value>
    ''' -----------------------------------------------------------------------------
    Default Public ReadOnly Property Item(ByVal index As Int32) As GDIPage
        Get
            Return DirectCast(Me.InnerList(index), GDIPage)
        End Get
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the full path to the file that this GDIDocument has been saved as / loaded from
    ''' </summary>
    ''' <value>A string containing the path to the file</value>
    ''' <remarks>Notice that path and the filename are concatenated to get this.  
    ''' Also note this is readonly. This enforces that this information can only be changed 
    ''' in response to a Save.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property FullPath() As String
        Get
            Return _Path & "\" & _FileName
        End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a value indicating whether the GDIDocument is dirty.
    ''' </summary>
    ''' <value>A Boolean indicating whether the GDIDocument is dirty</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Dirty() As Boolean
        Get
            Return _Dirty
        End Get
       
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a value indicating whether the GDIDocument has ever been saved.
    ''' </summary>
    ''' <value>A Boolean indicating whether the GDIDocument has been saved</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Saved() As Boolean
        Get
            Return _Saved
        End Get

    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets this GDIDocument's set of currently selected objects.
    ''' </summary>
    ''' <value>A set of GDIObjects that are currently selected.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Selected() As SelectedSet
        Get
            Return _SelectedSet
        End Get

    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Appends a GDIObject to the current page.  Records history for this append with the 
    ''' label specified in historylabel.
    ''' </summary>
    ''' <param name="obj">The GDIObject inheritor to append to the page</param>
    ''' <param name="historyLabel">A label to append to the history marker.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub AddObjectToPage(ByVal obj As GDIObject, ByVal historyLabel As String)
        _CurrentPage.GDIObjects.Add(obj)
        recordHistory(historyLabel)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the current page.
    ''' </summary>
    ''' <value>A GDIPage object (see the GDI Page class for more information)</value>
    ''' -----------------------------------------------------------------------------
    Public Property CurrentPage() As GDIPage
        Get
            Return _CurrentPage
        End Get
        Set(ByVal Value As GDIPage)
            Selected.DeselectAll()
            _CurrentPage = Value

            Me.invokePageChanged()
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the GDIDocument's smoothing mode
    ''' </summary>
    ''' <value>A smoothing mode used for pretty rendering.</value>
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
    ''' Gets or sets the GDIDocument's text rendering hint.
    ''' </summary>
    ''' <value>A TextRenderingHint used to render text to the surface.</value>
    ''' -----------------------------------------------------------------------------
    Public Property TextRenderingHint() As Drawing.Text.TextRenderingHint
        Get
            Return _TextHint
        End Get
        Set(ByVal Value As Drawing.Text.TextRenderingHint)
            _TextHint = Value
        End Set
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the history objects current position
    ''' </summary>
    ''' <value>An integer representing the history object's current position.</value>
    ''' <remarks>Notice when this function is "Set", it invokes 
    ''' the historychanged callbacks so that listeners will be informed.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Property CurHistoryPos() As Int32
        Get
            Return _History.CurPos
        End Get
        Set(ByVal Value As Int32)
            _History.CurPos = Value
            Me.invokeHistoryChangedCallBacks()
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets a value indicating the current page size.  In graphic class style 
    ''' documents this is equivalent to the entire drawing surface.
    ''' </summary>
    ''' <value>A rectangle representing the current page size</value>
    ''' -----------------------------------------------------------------------------
    Public Property RectPageSize() As Rectangle
        Get
            Return _PageArea
        End Get
        Set(ByVal Value As Rectangle)
            _PageArea = Value
        End Set
    End Property


#End Region

#Region "Disposal and Cleanup"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the GDIDocument, removing handlers and releasing resources.
    ''' </summary>
    ''' <param name="disposing">Whether unmanaged resources are being disposed or not.</param>
    ''' -----------------------------------------------------------------------------
    Private Overloads Sub Dispose(ByVal disposing As Boolean)
        If Not _Disposed Then
            If disposing Then

                RemoveHandler _SelectedSet.SelectionChanged, AddressOf SelectedSet_OnSelectionChanged
                _CurrentPage = Nothing
                _SelectedSet = Nothing

                If Not _PrintDocument Is Nothing Then
                    _PrintDocument.Dispose()
                End If

                If Not _History Is Nothing Then
                    _History.Dispose()
                End If

                For Each pg As GDIPage In Me.InnerList
                    pg.Dispose()
                Next
            End If


        End If

        _Disposed = True

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Calls the custom dispose method.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Overloads Sub Dispose() Implements System.IDisposable.Dispose
        Me.Dispose(True)
    End Sub
#End Region

End Class
