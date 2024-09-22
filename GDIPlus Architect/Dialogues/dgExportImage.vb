Imports System.Runtime.Serialization
Imports System.IO



''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : dgExportImage
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Initiates an image export dialog for the current drawn page.
''' </summary>
''' <remarks>Credit to Michael gold for his scrollable picture box example!
''' </remarks>
''' -----------------------------------------------------------------------------
Friend Class dgExportImage
    Inherits System.Windows.Forms.Form

#Region "Invoker"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Launches the image export dialog.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Shared Sub GO()
        Dim dgexportImage As New dgExportImage
        dgexportImage.ShowDialog()
    End Sub

#End Region

#Region "Local Fields"
    '''<summary>X Offset of the image in the picture box</summary>
    Private _OffsetX As Int32 = 0
    '''<summary>Y OFfset of the image in the picture box</summary>
    Private _OffsetY As Int32 = 0
    '''<summary>Current zoom of the picture box</summary>
    Private _Zoom As Int32 = 100



    '''<summary>Container for the various export formats GDI+ Architect can produce.</summary>
    Private _ExportFormats As New ExportFormats
    '''<summary>The currently selected export format</summary>
    Private _SelectedExport As Exp


    '''<summary>Image to render into the scrollable picture box</summary>
    Private _ImagePreview As Image


    '''<summary>Whether to update the image preview or not. </summary> 
    Private _UpdatePending As Boolean = False


#End Region

#Region "Contructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the Export Image dialog.  
    ''' Populate the dialog with zoom information and fills the format drop down.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Populates the zoom drop down
        populateZoom()
        'Populates the available export formats
        populateFormats()

        'Assigns the ExportImage set of tooltips to the form
        App.ToolTipManager.PopulatePopupTip(Tip, "ExportImage", Me)
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
    Private WithEvents btnOk As System.Windows.Forms.Button
    Private WithEvents btnCancel As System.Windows.Forms.Button
    Private WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents cboExp As System.Windows.Forms.ComboBox
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents nudJPEGQuality As System.Windows.Forms.NumericUpDown
    Private WithEvents lblOptions As System.Windows.Forms.Label
    Private WithEvents pnJPGOptions As System.Windows.Forms.Panel
    Private WithEvents picPreview As System.Windows.Forms.PictureBox
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents hscPreview As System.Windows.Forms.HScrollBar
    Private WithEvents vscPreview As System.Windows.Forms.VScrollBar
    Private WithEvents Label4 As System.Windows.Forms.Label
    Private WithEvents cboZoom As System.Windows.Forms.ComboBox
    Private WithEvents lblfile As System.Windows.Forms.Label
    Private WithEvents lblFileSize As System.Windows.Forms.Label
    Private WithEvents Tip As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.btnOk = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.pnJPGOptions = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.nudJPEGQuality = New System.Windows.Forms.NumericUpDown
        Me.cboExp = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.lblOptions = New System.Windows.Forms.Label
        Me.picPreview = New System.Windows.Forms.PictureBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.hscPreview = New System.Windows.Forms.HScrollBar
        Me.vscPreview = New System.Windows.Forms.VScrollBar
        Me.cboZoom = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.lblfile = New System.Windows.Forms.Label
        Me.lblFileSize = New System.Windows.Forms.Label
        Me.Tip = New System.Windows.Forms.ToolTip(Me.components)
        Me.pnJPGOptions.SuspendLayout()
        CType(Me.nudJPEGQuality, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(16, 440)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.TabIndex = 0
        Me.btnOk.Text = "Ok"
        '
        'btnCancel
        '
        Me.btnCancel.CausesValidation = False
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(96, 440)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        '
        'pnJPGOptions
        '
        Me.pnJPGOptions.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnJPGOptions.Controls.Add(Me.Label2)
        Me.pnJPGOptions.Controls.Add(Me.nudJPEGQuality)
        Me.pnJPGOptions.Location = New System.Drawing.Point(16, 112)
        Me.pnJPGOptions.Name = "pnJPGOptions"
        Me.pnJPGOptions.Size = New System.Drawing.Size(232, 224)
        Me.pnJPGOptions.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Quality:"
        '
        'nudJPEGQuality
        '
        Me.nudJPEGQuality.Location = New System.Drawing.Point(64, 8)
        Me.nudJPEGQuality.Name = "nudJPEGQuality"
        Me.nudJPEGQuality.TabIndex = 0
        '
        'cboExp
        '
        Me.cboExp.DisplayMember = "Name"
        Me.cboExp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboExp.Location = New System.Drawing.Point(96, 32)
        Me.cboExp.Name = "cboExp"
        Me.cboExp.Size = New System.Drawing.Size(192, 21)
        Me.cboExp.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 16)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Export Format:"
        '
        'lblOptions
        '
        Me.lblOptions.Location = New System.Drawing.Point(16, 88)
        Me.lblOptions.Name = "lblOptions"
        Me.lblOptions.Size = New System.Drawing.Size(232, 23)
        Me.lblOptions.TabIndex = 6
        Me.lblOptions.Text = "Label3"
        '
        'picPreview
        '
        Me.picPreview.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picPreview.Location = New System.Drawing.Point(296, 48)
        Me.picPreview.Name = "picPreview"
        Me.picPreview.Size = New System.Drawing.Size(475, 475)
        Me.picPreview.TabIndex = 7
        Me.picPreview.TabStop = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(296, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Preview"
        '
        'hscPreview
        '
        Me.hscPreview.Location = New System.Drawing.Point(296, 528)
        Me.hscPreview.Name = "hscPreview"
        Me.hscPreview.Size = New System.Drawing.Size(472, 16)
        Me.hscPreview.TabIndex = 9
        '
        'vscPreview
        '
        Me.vscPreview.Location = New System.Drawing.Point(776, 48)
        Me.vscPreview.Name = "vscPreview"
        Me.vscPreview.Size = New System.Drawing.Size(16, 472)
        Me.vscPreview.TabIndex = 10
        '
        'cboZoom
        '
        Me.cboZoom.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboZoom.Location = New System.Drawing.Point(344, 8)
        Me.cboZoom.Name = "cboZoom"
        Me.cboZoom.Size = New System.Drawing.Size(136, 21)
        Me.cboZoom.TabIndex = 11
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(296, 8)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 23)
        Me.Label4.TabIndex = 12
        Me.Label4.Text = "Zoom:"
        '
        'lblfile
        '
        Me.lblfile.Location = New System.Drawing.Point(16, 64)
        Me.lblfile.Name = "lblfile"
        Me.lblfile.Size = New System.Drawing.Size(64, 16)
        Me.lblfile.TabIndex = 13
        Me.lblfile.Text = "File Size:"
        '
        'lblFileSize
        '
        Me.lblFileSize.Location = New System.Drawing.Point(104, 64)
        Me.lblFileSize.Name = "lblFileSize"
        Me.lblFileSize.Size = New System.Drawing.Size(160, 23)
        Me.lblFileSize.TabIndex = 14
        '
        'dgExportImage
        '
        Me.AcceptButton = Me.btnOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(794, 559)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.pnJPGOptions)
        Me.Controls.Add(Me.cboExp)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblOptions)
        Me.Controls.Add(Me.picPreview)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.hscPreview)
        Me.Controls.Add(Me.vscPreview)
        Me.Controls.Add(Me.cboZoom)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lblfile)
        Me.Controls.Add(Me.lblFileSize)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "dgExportImage"
        Me.ShowInTaskbar = False
        Me.Text = "dgExport"
        Me.pnJPGOptions.ResumeLayout(False)
        CType(Me.nudJPEGQuality, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Event Handlers"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Initiates image export.  This method asks the _SelectedExport to export itself 
    ''' to the path returned by a save dialog.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Trace.WriteLineIf(App.TraceOn, "ExportImage.Ok")

        If Not _SelectedExport Is Nothing Then
            initiateImageExport()
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Request a save location and exports to the selected format unless the user 
    ''' cancels.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub initiateImageExport()
        'The format knows its file filter and assigns it to the saveFileDialog's filter property
        Dim dgSave As New SaveFileDialog
        dgSave.Filter = _SelectedExport.FileFilter

        Dim dgResult As DialogResult = dgSave.ShowDialog()

        If dgResult = DialogResult.OK Then
            Try
                'Asks the selected export format to perform the export process
                _SelectedExport.Export(MDIMain.ActiveDocument, _
                MDIMain.ActiveDocument.CurrentPage, _
                dgSave.FileName)

                App.MDIMain.StatusText = "Image exported"
            Catch ex As Exception
                MsgBox("Failed to Export")
            End Try
        End If

        dgSave.Dispose()

        Me.DialogResult = DialogResult.OK
    End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Cancels image export
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Trace.WriteLineIf(App.TraceOn, "ExportImage.Cancel")
        Me.DialogResult = DialogResult.Cancel
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles changes in the export options type drop down.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' <remarks>Notice that if the export is JPEG, we bring a special quality panel 
    ''' to the the front.  Otherwise we hide it.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Sub cboExp_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboExp.SelectedValueChanged
        Trace.WriteLineIf(App.TraceOn, "ExportImage.FormatChange")

        'Updatepending = True stops repaint during the change
         _UpdatePending = True
        _SelectedExport = DirectCast(cboExp.SelectedItem, Exp)

        If TypeOf _SelectedExport Is ExpJPEG Then
            showJPEGOptions()
        Else
            hideJPEGOptions()
        End If

        _UpdatePending = False

        updatePreview()

    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Hides the JPEG options panel
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub hideJPEGOptions()
        pnJPGOptions.Visible = False
        lblOptions.Text = String.Empty
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Shows the JPEG options panel
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub showJPEGOPtions()
        pnJPGOptions.Visible = True
        pnJPGOptions.BringToFront()
        Dim tempjpeg As ExpJPEG = DirectCast(_SelectedExport, ExpJPEG)
        nudJPEGQuality.Value = tempjpeg.Quality
        lblOptions.Text = _SelectedExport.Name & " Options"
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the image to the picture box for the current offset and zoom.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub picPreview_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles picPreview.Paint

        Dim g As Graphics = e.Graphics
        Dim pOutline As New Pen(Color.Black)

        Dim fScale As Single = _Zoom / 100.0!

        g.Clear(App.Options.ColorNonPrintArea)

        App.GraphicsManager.BeginScaledView(g, fScale)
 
 
        pOutline.Width = 1.0! / fScale
        g.DrawRectangle(pOutline, -_OffsetX, -_OffsetY, _ImagePreview.Width, _ImagePreview.Height)

        g.DrawImageUnscaled(_ImagePreview, -_OffsetX, -_OffsetY, _ImagePreview.Width, _ImagePreview.Height)


        App.GraphicsManager.EndScaledView(g)

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles changes to the pic preview horizontal scrollbar
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub hscPreview_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles hscPreview.Scroll
        Dim nVal As Int32 = e.NewValue
        _OffsetX = nVal
        picPreview.Invalidate()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles changes to the pic preview vertical scrollbar
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub vscPreview_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles vscPreview.Scroll
        Dim nVal As Int32 = e.NewValue
        _OffsetY = nVal
        picPreview.Invalidate()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles changes to the zoom drop down
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboZoom_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboZoom.SelectedIndexChanged
        Dim iZoom As Int32 = CInt(cboZoom.Text)
        _Zoom = iZoom
        sizeBars()
        picPreview.Invalidate()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles changes to the JPEG quality up down
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub nudJPEGQuality_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nudJPEGQuality.ValueChanged
        If TypeOf _SelectedExport Is ExpJPEG Then
            Dim ExpJPeg As ExpJPEG = DirectCast(_SelectedExport, ExpJPEG)
            ExpJPeg.Quality = CLng(nudJPEGQuality.Value)
        End If

        updatePreview()
    End Sub


#End Region

#Region "Helper methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exports the page in the current format to a temporary file
    ''' </summary>
    ''' <returns>The name of the temporary file</returns>
    ''' -----------------------------------------------------------------------------
    Private Function exportTemporaryFile() As String
        'Get a temporary file
        Dim sTempName As String = Path.GetTempFileName

        'Export the document to the temporary file
        _SelectedExport.Export(MDIMain.ActiveDocument, MDIMain.ActiveDocument.CurrentPage, sTempName)


        Return sTempName
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the file size information and assigns it to lblFileSize
    ''' </summary>
    ''' <param name="sTempName">The file to get file length for.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub getExportImageSize(ByVal sTempName As String)
        Dim finfo As New FileInfo(sTempName)

        Dim lLength As Long = finfo.Length
        Dim dLengthInK As Double = finfo.Length / 1000

        lblFileSize.Text = FormatNumber(dLengthInK, 2, TriState.True) & "K"
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Recreates the local image resource by loading the image specified in sImagePath
    ''' </summary>
    ''' <param name="sImagePath">A path t othe image to retrieve</param>
    ''' -----------------------------------------------------------------------------
    Private Sub recreatePreviewImage(ByVal sImagePath As String)
        If Not _ImagePreview Is Nothing Then
            _ImagePreview.Dispose()
        End If

        _ImagePreview = Image.FromFile(sImagePath)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the image preview in the picture box.  Creates a new temporary image and 
    ''' places it inside the box.
    ''' </summary>
    ''' <remarks>Notice the use of temporary files here.  First, if you have never seen the 
    ''' very nice method Path.GetTempFileName, it's useful for doing exactly what we're doing 
    ''' here which is creating a temp file.  In order to get the export image size, we actually 
    ''' do the export and check the size of the exported file.  
    ''' 
    ''' This method also short circuits if the export format drop down is in the process of 
    ''' being changed
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Sub updatePreview()

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        If Not _UpdatePending AndAlso Not _SelectedExport Is Nothing Then

            Dim sImagePath As String = exportTemporaryFile()

            recreatePreviewImage(sImagePath)

            getExportImageSize(sImagePath)


            sizeBars()
            picPreview.Invalidate()
        End If

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populates the list of drop down boxes
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub populateFormats()
        For Each format As Exp In _ExportFormats
            cboExp.Items.Add(format)
        Next

        cboExp.SelectedIndex = 0
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populates the zoom drop down with appropriate values.  Sets the 
    ''' initial zoom to 100%
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub populateZoom()
        With cboZoom
            .BeginUpdate()
            .Items.AddRange(New String() {"25", "50", "75", "100", "150", "200", "250", "500"})
            .EndUpdate()
            .SelectedIndex = 3
        End With

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set the amount of scroll based on the size of the image at the current zoom
    ''' </summary>
      ''' -----------------------------------------------------------------------------
    Private Sub sizeScrollToPreview()
        hscPreview.Minimum = 0
        vscPreview.Minimum = 0


        Dim iScrollWidth As Int32 = CInt((_Zoom / 100) * _ImagePreview.Width - picPreview.Width)
        Dim iScrollHeight As Int32 = CInt((_Zoom / 100) * _ImagePreview.Height - picPreview.Height)

        If iScrollWidth > 0 Then
            hscPreview.Maximum = iScrollWidth
            hscPreview.Visible = True
        Else
            hscPreview.Visible = False
        End If

        If iScrollHeight > 0 Then
            vscPreview.Maximum = iScrollHeight
            vscPreview.Visible = True
        Else
            vscPreview.Visible = False
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Calculates the sizes of the up and down veritcal image scroll bars based upon 
    ''' the image being displayed and the size of the picPreview box. This is called 
    ''' when the preview image changes or the current zoom changes.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub sizeBars()

        vscPreview.SetBounds(picPreview.Right, picPreview.Top, vscPreview.Width, picPreview.Height)
        hscPreview.SetBounds(picPreview.Left, picPreview.Bottom, picPreview.Width, hscPreview.Height)


        If Not _ImagePreview Is Nothing Then
            'Set the amount of allowed scroll based on the size of the image at the current zoom
            sizeScrollToPreview()
        Else
            'No scroll necessary, hide the scroll bars
            hscPreview.Visible = False
            vscPreview.Visible = False
        End If

        'reset scroll values to 0
        hscPreview.Value = 0
        vscPreview.Value = 0

        _OffsetX = 0
        _OffsetY = 0

        picPreview.Invalidate()
    End Sub

#End Region

End Class
