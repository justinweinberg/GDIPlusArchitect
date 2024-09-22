Imports Crownwood.Magic.Docking
Imports Crownwood.Magic.Controls
Imports System.ComponentModel
Imports GDIObjects


 Friend Class MDIMain
    Inherits System.Windows.Forms.Form


#Region "Local Fields"


    Private _NextDocumentNumber As Int32 = 1

    Private _PrinterSettings As Printing.PrinterSettings
    Private _PageSettings As Printing.PageSettings

    Private _FullScreenView As Boolean = False


    Private _GDIMenu As GDIMenu

    Private Shared _activeWindow As GDIDOC

#End Region


#Region "Application Startup"
    Public Sub New()
        MyBase.New()
    End Sub


    Public Sub HandleCommand(ByVal strArgs() As String)
        Dim strCmd As String
        If Not strArgs Is Nothing AndAlso Not strArgs.Length = 0 Then
            For Each strCmd In strArgs
                openfile(System.IO.Path.GetFullPath(strCmd))
            Next
        End If
    End Sub

    Public Sub startApplication(ByVal cmdArgs() As String)



        SetupDefaultDocuments()

        InitializeComponent()



        Trace.WriteLineIf(App.TraceOn, "MDIMain.ApplicationStarting")


        Me._GDIMenu = New GDIMenu

        Dim outerControls() As Control = {_GDIMenu.TopLevelMenu, statusBar}


        App.PanelManager.InitializePanels(Me, outerControls)


        Me._GDIMenu.TopLevelMenu.MdiContainer = Me
        Me.Controls.Add(_GDIMenu.TopLevelMenu)

        Me.RefreshMenu()


        Me.Icon = IconManager.GetIconResource("GDIPlusArchitect.app16.ico")

        If Not cmdArgs Is Nothing AndAlso cmdArgs.Length > 0 Then
            HandleCommand(cmdArgs)
        End If

        Me.StatusText = "Ready"





    End Sub



    Private Sub SetupDefaultDocuments()
        Dim printtdoc As New Printing.PrintDocument

        _PrinterSettings = printtdoc.PrinterSettings
        _PageSettings = printtdoc.DefaultPageSettings
        printtdoc.Dispose()
        printtdoc = Nothing
    End Sub

#End Region

#Region "Public Members"



    Public Sub RefreshMenu()
        _GDIMenu.Refresh()
    End Sub


    Public Sub InsertHGuide()
        If Not Me.ActiveDocument Is Nothing Then
            _activeWindow.InsertHGuide()
        End If
    End Sub



    Public Sub InsertVGuide()
        If Not Me.ActiveDocument Is Nothing Then
            _activeWindow.InsertVGuide()
        End If
    End Sub


    Public Shared Sub SaveActiveDocument(ByVal name As String)
        _activeWindow.saveAs(name)
    End Sub


    Public Shared Sub CloseActiveWindow()
        _activeWindow.Close()
    End Sub



    Public Shared Sub RefreshActiveDocument()
        _activeWindow.Refresh()
    End Sub


    Public Function snapPoint(ByVal ptInitial As Point, ByRef bXSnapped As Boolean, ByRef bYSnapped As Boolean) As Point
        Return Me._activeWindow.snapPoint(ptInitial, bXSnapped, bYSnapped)
    End Function




    Public Sub InvokeNewGraphicClass()

        Dim dgNew As New dgNewGraphicClass
        Dim iresp As DialogResult = dgNew.ShowDialog()

        If iresp = DialogResult.OK Then
            Dim sname As String = "Untitled" & _NextDocumentNumber
            Dim exportsettings As New ExportSettings("MyNameSpace", EnumDocumentTypes.eClass, sname, App.Options.CodeLanguage)
            Dim newDoc As New GDIDocument(New Rectangle(New Point(0, 0), dgNew.ControlSize), exportsettings, App.Options.TextRenderingHint, App.Options.SmoothingMode)
            newDoc.BackColor = dgNew.SelectedBackColor
            newDoc.FileName = sname

            _NextDocumentNumber += 1

            Dim gdidoc As New GDIDOC(newDoc)
            gdidoc.Show()

        End If

    End Sub

    Public Sub InvokeDocumentSettings()
        Dim dgDocProperties As New dgDocProperties(_PageSettings)
        dgDocProperties.ShowDialog()
    End Sub




    Public Sub InvokeNewPrintDocument()
        Try

            Dim sname As String = "Untitled" & _NextDocumentNumber

            Dim exportsettings As New ExportSettings("MynameSpace", EnumDocumentTypes.ePrintDocument, sname, App.Options.CodeLanguage)

            Dim newDoc As New GDIDocument(_PrinterSettings, _PageSettings, exportsettings, App.Options.TextRenderingHint, App.Options.SmoothingMode)


            newDoc.FileName = sname
 
            _NextDocumentNumber += 1
 
            Dim gdidoc As New GDIDOC(newDoc)
            gdidoc.Show()
 

            gdidoc.MdiParent = Me
            gdidoc.Show()

            _NextDocumentNumber += 1
        Catch ex As System.Drawing.printing.InvalidPrinterException
            MsgBox("GDI+ Architect had trouble accessing your default printer.  Make sure you have a printer installed that is accessible and have selected a default printer")
        End Try
    End Sub




    Public Sub InvokePrinterSettings()
        Dim pSettings As New PrintDialog
        Dim iresp As DialogResult

        With pSettings
            .PrinterSettings = _PrinterSettings
            iresp = .ShowDialog()
        End With

        If Not iresp = DialogResult.Cancel Then
            _PrinterSettings = pSettings.PrinterSettings()
            If Me.ActiveDocument.ExportSettings.DocumentType = _
            EnumDocumentTypes.ePrintDocument Then

                ActiveDocument.SetResoByPrinterSettings(_PrinterSettings)

            End If
        End If
    End Sub


    Public Sub InvokeOpenrecent(ByVal sPath As String)
        Try
            openfile(sPath)
        Catch ex As Exception
            MsgBox("Could not open file or file does not exist")

        End Try

    End Sub



    Public Sub InvokeOpen()
        Trace.WriteLineIf(App.TraceOn, "MDIMain.InvokeOpen")
        Dim dgOpen As New OpenFileDialog


        dgOpen.Filter = "GDI Doc | *.gdd"
        dgOpen.DefaultExt = "gdd"

        Dim iResp As DialogResult = dgOpen.ShowDialog(Me)

        If iResp = DialogResult.OK Then
            Try
                openfile(dgOpen.FileName)
                App.Options.addRecentFile(dgOpen.FileName)
                onDocumentChanged()

            Catch ex As Exception
                MsgBox("Invalid gdd file")
            Finally
                dgOpen.Dispose()
            End Try

        End If

    End Sub



    Public Sub InvokeRevert()
        Trace.WriteLineIf(App.TraceOn, "MDIMain.InvokeRevert")

        Dim iresp As MsgBoxResult = MsgBox("Are you sure you wish to revert to the last saved version?", MsgBoxStyle.OKCancel)
        If iresp = MsgBoxResult.OK Then
            Dim doc As GDIDocument = GDIDocument.LoadFromFile(MDIMain.ActiveDocument.FullPath)
            _activeWindow.Document = doc
            HandleDocumentStateChanged(_activeWindow)
        End If
    End Sub

    Public Sub invokeExport()

        Trace.WriteLineIf(App.TraceOn, "MDIMain.InvokeExport")
        Dim _originalLang As EnumCodeTypes = MDIMain.ActiveDocument.ExportSettings.Language
        Dim _selectedLang As EnumCodeTypes
        Dim dgsave As New SaveFileDialog
        Dim ioOut As IO.StreamWriter

        With dgsave
            .CheckPathExists = True
            .AddExtension = True
            Select Case MDIMain.ActiveDocument.ExportSettings.Language
                Case EnumCodeTypes.eCSharp
                    .FilterIndex = 2
                    .Filter = "Visual basic (*.vb)|*.vb|C Sharp (*.cs)|*.cs|All files (*.*)|*.*"
                Case EnumCodeTypes.eVB
                    .DefaultExt = "vb"
                    .FilterIndex = 1
                    .Filter = "Visual basic (*.vb)|*.vb|C Sharp (*.cs)|*.cs|All files (*.*)|*.*"
            End Select

        End With

        Dim fileStream As System.IO.Stream
        Try
            Dim iresp As DialogResult = dgsave.ShowDialog()

            If iresp = DialogResult.OK Then
                Select Case dgsave.FilterIndex
                    Case 1
                        _selectedLang = EnumCodeTypes.eVB
                    Case 2
                        _selectedLang = EnumCodeTypes.eCSharp
                End Select

                MDIMain.ActiveDocument.ExportSettings.Language = _selectedLang
                'temporarily change the code base if different 

                Dim sCode As String = _activeWindow.Document.GenerateCode
                fileStream = New System.IO.FileStream(dgsave.FileName, IO.FileMode.Create)
                ioOut = New System.IO.StreamWriter(fileStream)
                ioOut.Write(sCode)
            End If

            MDIMain.ActiveDocument.ExportSettings.Language = _originalLang
        Catch ex As Exception
            MsgBox("Could not access the export destination or other errors occurred")
        Finally

            If Not ioOut Is Nothing Then
                ioOut.Close()
                ioOut = Nothing
            End If

            fileStream = Nothing
            dgsave.Dispose()

        End Try
    End Sub







#End Region


#Region "Delegates and Events"

    Public Sub onOptionsChanged()
        Me._GDIMenu.OnOptionsChanged()
    End Sub

    'Public Shared Sub OnFillChanged()
    '    _activeWindow.OnFillChanged()
    'End Sub

    'Public Shared Sub OnStrokeChanged()
    '    _activeWindow.OnStrokeChanged()
    'End Sub


    Public Sub onDocumentChanged()
        _GDIMenu.onDocumentChanged()
    End Sub


    Private Sub HandleDocumentStateChanged(ByVal newChild As GDIDOC)
        Trace.WriteLineIf(App.TraceOn, "MDIMain.HandleDocumentStateChanged")

        Dim newDocument As GDIDocument

        If newChild Is Nothing Then
            newDocument = Nothing
        Else
            newDocument = newChild.Document
        End If

        App.MDIMain.RefreshMenu()

        App.PanelManager.OnDocumentChanged()

        Session.DocumentManager.CurrentDocument = newDocument
 

    End Sub

#End Region



#Region "Property Accessors"
    Public Shared Property ActiveZoom() As Single
        Get
            Return _activeWindow.Zoom
        End Get
        Set(ByVal Value As Single)
            _activeWindow.Zoom = Value
        End Set
    End Property

    Public Sub InvokeSave()
        Trace.WriteLineIf(App.TraceOn, "MDIMain.save")

        If Not _activeWindow Is Nothing Then
            _activeWindow.saveDocument()
        End If
    End Sub



    Public Shared ReadOnly Property ActiveDocument() As GDIDocument
        Get
            If _activeWindow Is Nothing Then
                Return Nothing
            Else
                Return _activeWindow.Document
            End If
        End Get
    End Property



    Public Property StatusText() As String
        Get
            Return statusBarPanel.Text
        End Get
        Set(ByVal Value As String)
            statusBarPanel.Text = Value
        End Set
    End Property
#End Region


#Region "Cleanup and Disposal"
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then

            If Not App.Options Is Nothing Then
                App.Options.EndTracing()
                App.Options.Dispose()
            End If

            App.PanelManager.DisposePanels()


            App.IconManager.Dispose()


            Me._activeWindow = Nothing

        End If

        MyBase.Dispose(disposing)
    End Sub
#End Region

#Region "Windows Generated Code"



    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend statusBar As System.Windows.Forms.StatusBar
    Private statusBarPanel As System.Windows.Forms.StatusBarPanel

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        statusBar = New StatusBar
        statusBarPanel = New StatusBarPanel
        CType(statusBarPanel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        statusBar.Parent = Me
        '
        'mdiStatus
        '
        statusBar.Bounds = New Rectangle(0, 447, 560, 22)
        statusBar.Panels.AddRange(New StatusBarPanel() {statusBarPanel})
        statusBar.ShowPanels = True
        statusBar.SizingGrip = False
        statusBar.TabIndex = 1
        '
        'statusMain
        '
        statusBarPanel.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        statusBarPanel.Width = 560
        '
        'MDIMain
        '
        Me.AllowDrop = True
        Me.AutoScaleBaseSize = New Size(5, 13)
        Me.ClientSize = New Size(560, 469)
        Me.IsMdiContainer = True
        Me.Text = "GDI+ Architect"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(statusBarPanel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub


#End Region


    Public Function windowFromFileName(ByVal sFullPath As String) As GDIDOC
        For Each frm As Form In Me.MdiChildren
            If TypeOf frm Is GDIDOC Then
                Dim GDIDOC As GDIDOC = DirectCast(frm, GDIDOC)
                If GDIDOC.Document.FullPath = sFullPath Then
                    Return GDIDOC
                End If
            End If

            Return Nothing
        Next
    End Function



#Region "Implementation Members"


    Private Sub openfile(ByVal sFilename As String)
        Trace.WriteLineIf(App.TraceOn, "MDIMain.openfile")

        Dim fInfo As System.IO.FileInfo = New System.IO.FileInfo(sFilename)
        Dim gdiDoc As GDIDOC = windowFromFileName(fInfo.FullName)

        If gdiDoc Is Nothing Then
            Dim doc As GDIDocument = GDIDocument.LoadFromFile(fInfo.FullName)
            ' Dim createdDoc As New GDIDocument(doc)
            Dim newDoc As New GDIDOC(doc)
            HandleDocumentStateChanged(newDoc)

            newDoc.MdiParent = Me

            newDoc.Show()
        Else
            gdiDoc.Focus()
        End If
        App.MDIMain.StatusText = "Item Opened"

    End Sub

    Private Sub invokeFullScreen()
        Trace.WriteLineIf(App.TraceOn, "MDIMain.InvokeFullScreen")

        If _FullScreenView Then
            App.PanelManager.HidePanels()
        Else
            App.PanelManager.ShowPanels()
        End If
    End Sub
#End Region




#Region "Event Handlers"

    Private Sub MDIMain_MdiChildActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.MdiChildActivate
        Trace.WriteLineIf(App.TraceOn, "MDIMain.ChildActivate")

        Dim child As GDIDOC
        If TypeOf ActiveMdiChild Is GDIDOC Then
            child = DirectCast(Me.ActiveMdiChild, GDIDOC)
        Else
            child = Nothing
        End If

        Me._activeWindow = child

        HandleDocumentStateChanged(child)

    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message, ByVal keyData As System.Windows.Forms.Keys) As Boolean
        If keyData = Keys.Tab AndAlso Not App.ToolManager.GlobalMode = EnumTools.eText Then
            _FullScreenView = Not _FullScreenView
            invokeFullScreen()
        End If


        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function


    Private Sub MDIMain_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed

        App.PanelManager.SaveLayout()
        App.Options.SaveOptions()
    End Sub

#End Region


#Region "Drag Drop Implementation"


    Private Sub MDIMain_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragDrop

        ' Handle FileDrop data.
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' Assign the file names to a string array, in 
            ' case the user has selected multiple files.
            Dim files As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Try
                For Each strFile As String In files
                    Me.openfile(strFile)
                Next

            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Return
            End Try
        End If

    End Sub

    Private Sub MDIMain_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

#End Region
End Class
