Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : GDIDOC
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' This is the actual GDI+ document window.  When the user opens or closes a document,
''' this is what is recreated.  The drawing related functionality is contained in this 
''' forms surface field.  The GDIDocument instance is contained in this 
''' forms document field
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class GDIDOC
    Inherits System.Windows.Forms.Form


#Region "Local Fields"


    Private _GDISurface As GDISurface

#End Region

#Region "Document specific callback declarations"
    Private _CBSelectionChanged As New GDIDocument.OnSelectionChanged(AddressOf onSelectionChanged)
    Private _CBSelectedPageChanged As New GDIDocument.OnPageChanged(AddressOf onPageChanged)
    Private _CBDocHistoryChanged As New GDIDocument.OnHistoryUpdated(AddressOf onHistoryChanged)
#End Region

#Region "Constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a GDIDoc window.
    ''' </summary>
    ''' <param name="doc">The GDIDocument that this window is managing.</param>
    ''' <remarks>When the local document property is set in the constructor, a series of 
    ''' delegates are attached.  Make sure to look at the Document property to understand 
    ''' what is happening.
    ''' </remarks>
        ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal doc As GDIDocument)
        Trace.WriteLineIf(App.TraceOn, "GDIDoc.New")

        _GDISurface = New GDISurface(doc)
        Me.Document = doc

        _GDISurface.Dock = DockStyle.Fill

        Me.AutoScaleBaseSize = New Size(5, 13)
        Me.ClientSize = New Size(432, 573)
        Me.ShowInTaskbar = False


        Me.Controls.Add(_GDISurface)

        refreshHeader()

        AddHandler _GDISurface.ZoomChanged, AddressOf GDISurface_ZoomChanged

        Me.MdiParent = App.mdimain
        Me.Icon = IconManager.GetIconResource("GDIPlusArchitect.app16.ico")

    End Sub

#End Region

#Region "Public Interface"
    Public Sub InsertHGuide()
        Me._GDISurface.InsertHGuide()
    End Sub


    Public Sub InsertVGuide()
        Me._GDISurface.InsertVGuide()
    End Sub


    Public Function snapPoint(ByVal ptInitial As Point, ByRef bXSnapped As Boolean, ByRef bYSnapped As Boolean) As Point
        Return Me._GDISurface.snapPoint(ptInitial, bXSnapped, bYSnapped)
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Begins a saveAs process, returning a boolean indicating if the user actually did 
    ''' save the document.  The user may cancel the save as dialog in which case this 
    ''' method returns false
    ''' </summary>
    ''' <param name="sName">The initial name of the document to use in a save as operation</param>
    ''' <returns>A boolean indicating if the document was actually saved.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function saveAs(ByVal sName As String) As Boolean
        Trace.WriteLineIf(App.TraceOn, "GDIDoc.saveAs")

        Dim dgsave As New SaveFileDialog
        dgsave.Filter = "GDI Doc | *.gdd"
        dgsave.FileName = MDIMain.ActiveDocument.FileName

        Dim res As DialogResult = dgsave.ShowDialog(Me)

        If res = DialogResult.OK Then

            Dim fInfo As System.IO.FileInfo = New System.IO.FileInfo(dgsave.FileName)

            If App.MDIMain.windowFromFileName(dgsave.FileName) Is Nothing Then
                Try

                    MDIMain.ActiveDocument.SaveAs(fInfo.DirectoryName.ToString, _
                                            fInfo.Name.ToString)
                    App.MDIMain.RefreshMenu()


                    App.PanelManager.OnDocumentChanged()

                    refreshHeader()
                    App.Options.addRecentFile(dgsave.FileName)
                    App.MDIMain.onDocumentChanged()
                    App.MDIMain.StatusText = "Item(s) saved"
                    Return True
                Catch ex As Exception

                End Try

            Else
                MsgBox("This file is currently open.  Please close it or select a different name", MsgBoxStyle.Information)
                Return False
            End If
        Else
            Return False
        End If
        dgsave.Dispose()
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Asks the document to save itself.  Returns a boolean indicating if the user has 
    ''' successfully saved the document.
    ''' </summary>
    ''' <returns>A boolean indicating if the document has been saved</returns>
    ''' <remarks>This method can return false if the document has never been saved and 
    ''' requires a name, and the user cancels the "saveas" process.
    ''' </remarks>
        ''' -----------------------------------------------------------------------------
    Public Function saveDocument() As Boolean
        Trace.WriteLineIf(App.TraceOn, "GDIDoc.handleSave")

        If Me.Document.Saved = False Then
            Return saveAs(Me.Document.FileName)
        Else
            Me.Document.Save()
            App.MDIMain.RefreshMenu()

            App.MDIMain.StatusText = "Item(s) saved"
            Return True
        End If
    End Function
#End Region


#Region "Delegate and Event Handlers"
 
 
    Private Sub GDISurface_ZoomChanged(ByVal s As Object, ByVal e As EventArgs)
        refreshHeader()
    End Sub

    Private Sub onSelectionChanged(ByVal s As Object, ByVal e As EventArgs)
        App.MDIMain.RefreshMenu()
        App.PanelManager.OnSelectionChanged()
    End Sub

    Private Sub onHistoryChanged(ByVal s As Object, ByVal e As EventArgs)
        Trace.WriteLineIf(App.TraceOn, "GDIDoc.HistoryChanged")

        'Notify all listeners that the History has changed
        App.PanelManager.OnHistoryChanged()

        App.MDIMain.RefreshMenu()

        _GDISurface.Invalidate()
    End Sub

    Private Sub onPageChanged(ByVal s As Object, ByVal e As EventArgs)
        'Notify all related windows that the page has changed
        Trace.WriteLineIf(App.TraceOn, "GDIDoc.PageChanged")

        refreshHeader()

        App.PanelManager.OnPageChanged()

        _GDISurface.Invalidate()

        'At this time, the only way the page could be changed is if the PageWin changes it, 
        'but in the future, there could be other ways the page could be set.
        'So we want to give this window a chance to update itself.

    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Interrupts the standard form close to check if the document needs to be saved and 
    ''' prompts the user as necessary.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub GDIDOC_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Me.Document.Dirty Then


            Dim iresp As MsgBoxResult = MsgBox("Save Changes to " & Me.Document.FileName & "?", MsgBoxStyle.YesNoCancel, "Save Changes?")

            Select Case iresp
                'Saves the document and closes it if the user doesn't cancel the saveas dialog.

            Case MsgBoxResult.Yes
                    e.Cancel = Not Me.saveDocument()
                Case MsgBoxResult.Cancel
                    e.Cancel = True
                Case MsgBoxResult.No
                    e.Cancel = False


            End Select

        End If
    End Sub
#End Region



#Region "Property Accessors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the current zoom of the surface.
    ''' </summary>
    ''' <value></value>
    ''' -----------------------------------------------------------------------------
    Public Property Zoom() As Single
        Get
            Return Me._GDISurface.Zoom
        End Get
        Set(ByVal Value As Single)
            Me._GDISurface.Zoom = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or Sets the document this GDIDoc window is displaying.
    ''' </summary>
    ''' <value>A GDIDocument instance to display within the GDIDoc window</value>
    ''' <remarks>Notice that callbacks are removed prior to setting this method 
    ''' if the window previously contained a document, and then are reassigned 
    ''' for the new document.  Finally, notice that this method raises the OnPageChanged 
    ''' event since with a new document, the page that was being displayed has changed
    ''' </remarks>
        ''' -----------------------------------------------------------------------------
    Public Property Document() As GDIDocument
        Get
            Return _GDISurface.Document
        End Get
        Set(ByVal Value As GDIDocument)

            If Not _GDISurface.Document Is Nothing Then

                _GDISurface.Document.removeSelectionChangedCallBack(_CBSelectionChanged)
                _GDISurface.Document.removeHistoryChangedCallBack(_CBDocHistoryChanged)
                _GDISurface.Document.removePageChangeCallBack(_CBSelectedPageChanged)

                _GDISurface.Document = Nothing
            End If

            _GDISurface.Document = Value
            _GDISurface.Document.setSelectionChangedCallBack(_CBSelectionChanged)
            _GDISurface.Document.setHistoryChangedCallBack(_CBDocHistoryChanged)
            _GDISurface.Document.setPageChangeCallBack(_CBSelectedPageChanged)

            _GDISurface.Invalidate()

        End Set
    End Property




#End Region

#Region "Implementation Details"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the text displayed at the top of the window to display the document 
    ''' name and the current zoom as a percentage
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub refreshHeader()
        Me.Text = _GDISurface.Document.FileName & " - " & _
      _GDISurface.Document.CurrentPage.Name & " Zoom: " & FormatPercent(_GDISurface.Zoom, 0)
    End Sub


#End Region



#Region "Disposal and Cleanup"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the GDIDocument window.  Notice that this method removes the window 
    ''' from delegates it was subscribing to (it would be a bad thing to call these on 
    ''' a disposed window)
    ''' </summary>
    ''' <param name="disposing></param>
    ''' -----------------------------------------------------------------------------
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then

            If Not _GDISurface.Document Is Nothing Then
                _GDISurface.Document.removeSelectionChangedCallBack(_CBSelectionChanged)
                _GDISurface.Document.removeHistoryChangedCallBack(_CBDocHistoryChanged)
                _GDISurface.Document.removePageChangeCallBack(_CBSelectedPageChanged)

            End If

            _GDISurface.Dispose()

        End If

        MyBase.Dispose(disposing)
    End Sub


#End Region


End Class
