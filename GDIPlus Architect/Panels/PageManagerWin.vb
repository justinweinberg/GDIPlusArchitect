Imports GDIObjects
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : PageManagerWin
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Responsible for managing the series of pages for the PrintDocument type of output.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class PageManagerWin
    Inherits UserControl


#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a PageManagerWin
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
#End Region

#Region " Windows Form Designer generated code "


    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Private WithEvents lstPages As System.Windows.Forms.ListBox
    Private WithEvents btnAddPage As System.Windows.Forms.Button
    Private WithEvents btnDelete As System.Windows.Forms.Button
    Private WithEvents btnName As System.Windows.Forms.Button

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lstPages = New System.Windows.Forms.ListBox
        Me.btnAddPage = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnName = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lstPages
        '
        Me.lstPages.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstPages.Location = New System.Drawing.Point(0, 0)
        Me.lstPages.Name = "lstPages"
        Me.lstPages.Size = New System.Drawing.Size(150, 108)
        Me.lstPages.TabIndex = 0
        '
        'btnAddPage
        '
        Me.btnAddPage.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAddPage.Location = New System.Drawing.Point(32, 120)
        Me.btnAddPage.Name = "btnAddPage"
        Me.btnAddPage.Size = New System.Drawing.Size(24, 23)
        Me.btnAddPage.TabIndex = 1
        Me.btnAddPage.Text = "+"
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Location = New System.Drawing.Point(64, 120)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(24, 24)
        Me.btnDelete.TabIndex = 2
        Me.btnDelete.Text = "-"
        '
        'btnName
        '
        Me.btnName.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnName.Location = New System.Drawing.Point(96, 120)
        Me.btnName.Name = "btnName"
        Me.btnName.Size = New System.Drawing.Size(48, 24)
        Me.btnName.TabIndex = 3
        Me.btnName.Text = "Name"
        '
        'PageManagerWin
        '
        Me.Controls.Add(Me.lstPages)
        Me.Controls.Add(Me.btnAddPage)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnName)
        Me.Name = "PageManagerWin"
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Refresh"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles history change events that bubble up from the current document.  This 
    ''' handler is invoked from the window that contains the current document.
    ''' </summary>
    ''' <param name="doc">Document which has had its history changed</param>
    ''' -----------------------------------------------------------------------------
    Public Sub refreshPanel()


        Trace.WriteLineIf(App.TraceOn, "PageWin.HistoryChanged")

        If MDIMain.ActiveDocument Is Nothing Then
            lstPages.Items.Clear()
            lstPages.SelectedItem = Nothing
        Else
            RefreshListContents()

            If MDIMain.ActiveDocument.ExportSettings.DocumentType = EnumDocumentTypes.ePrintDocument Then
                Me.Enabled = True
                lstPages.SelectedItem = MDIMain.ActiveDocument.CurrentPage

            Else
                Me.Enabled = False
            End If
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Rebuilds the list of pages given a document to represent
    ''' </summary>
    ''' <param name="doc">The document to represent</param>
    ''' -----------------------------------------------------------------------------
    Private Sub RefreshListContents()
        Trace.WriteLineIf(App.TraceOn, "PageWin.RefreshListContents")

        With lstPages
            .BeginUpdate()
            .Items.Clear()

            For Each pg As GDIPage In MDIMain.ActiveDocument
                .Items.Add(pg)
            Next

            .SelectedItem = MDIMain.ActiveDocument.CurrentPage
            .EndUpdate()

        End With



    End Sub

#End Region


#Region "Event Handlers"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a new page to the current document
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnAddPage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddPage.Click
        MDIMain.ActiveDocument.addNewPage()
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Confirms and then removes a page from the current document
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If Not lstPages.SelectedIndex = -1 Then
            If lstPages.Items.Count > 1 Then
                Dim iresp As MsgBoxResult = MsgBox("Are you sure you wish to delete this page?", MsgBoxStyle.OKCancel)
                If iresp = DialogResult.OK Then
                    Trace.WriteLineIf(App.TraceOn, "GDIDoc.DeletePage")

                    Dim pg As GDIPage = DirectCast(lstPages.SelectedItem, GDIPage)
                    MDIMain.ActiveDocument.deletePage(pg)


                End If
            Else
                MsgBox("You cannot delete the last page in a document")
            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the selected page when the list of pages is clicked.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub lstPages_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstPages.Click
        If Not lstPages.SelectedIndex = -1 Then

            Dim pageNum As Int32 = (DirectCast(lstPages.SelectedItem, GDIPage).PageNum)
            MDIMain.ActiveDocument.SetcurrentPage(pageNum)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Prompts the user for a new name for a page and assigns it if a value is received.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnName_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnName.Click
        If Not lstPages.SelectedIndex = -1 Then

            Dim pg As GDIPage = DirectCast(lstPages.SelectedItem, GDIPage)
            Dim sNewName As String = InputBox("Enter the new name of this page", "Enter new page name", pg.Name)
            If Trim(sNewName).Length > 0 Then
                MDIMain.ActiveDocument.renamePage(pg, sNewName)
            End If
        End If
    End Sub

#End Region

End Class
