
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : dgQuickCode
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a window for displaying quick code (Generated code that is not to be 
''' exported)
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class dgQuickCode
    Inherits System.Windows.Forms.Form


#Region "Constructors"


    '''<summary>Constant holding the file filter for saving quick code</summary>
    Private Const CONST_TEXT_FILTER As String = "Text File | *.txt | All Files | *.*"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the quick code dialog given the code to display.
    ''' </summary>
    ''' <param name="sCode">The code to display</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal sCode As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        txtQuickCode.Text = sCode
        'Add any initialization after the InitializeComponent() call
        App.ToolTipManager.PopulatePopupTip(Tip, "QuickCode", Me)
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
    Private WithEvents txtQuickCode As System.Windows.Forms.TextBox
    Private WithEvents btnCopy As System.Windows.Forms.Button
    Private WithEvents btnClose As System.Windows.Forms.Button
    Private WithEvents btnSave As System.Windows.Forms.Button
    Private WithEvents btnHelp As System.Windows.Forms.Button
    Private WithEvents Tip As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.txtQuickCode = New System.Windows.Forms.TextBox
        Me.btnCopy = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnHelp = New System.Windows.Forms.Button
        Me.Tip = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'txtQuickCode
        '
        Me.txtQuickCode.AcceptsReturn = True
        Me.txtQuickCode.AcceptsTab = True
        Me.txtQuickCode.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtQuickCode.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtQuickCode.CausesValidation = False
        Me.txtQuickCode.Location = New System.Drawing.Point(0, 0)
        Me.txtQuickCode.Multiline = True
        Me.txtQuickCode.Name = "txtQuickCode"
        Me.txtQuickCode.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtQuickCode.Size = New System.Drawing.Size(568, 352)
        Me.txtQuickCode.TabIndex = 0
        Me.txtQuickCode.Text = ""
        '
        'btnCopy
        '
        Me.btnCopy.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCopy.Location = New System.Drawing.Point(360, 376)
        Me.btnCopy.Name = "btnCopy"
        Me.btnCopy.Size = New System.Drawing.Size(80, 23)
        Me.btnCopy.TabIndex = 4
        Me.btnCopy.Text = "Copy"
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.CausesValidation = False
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Location = New System.Drawing.Point(448, 376)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(80, 23)
        Me.btnClose.TabIndex = 3
        Me.btnClose.Text = "Close"
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(120, 376)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(104, 23)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "Save To Text File"
        '
        'btnHelp
        '
        Me.btnHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnHelp.Location = New System.Drawing.Point(272, 376)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.Size = New System.Drawing.Size(80, 23)
        Me.btnHelp.TabIndex = 5
        Me.btnHelp.Text = "Help"
        '
        'dgQuickCode
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(568, 405)
        Me.Controls.Add(Me.txtQuickCode)
        Me.Controls.Add(Me.btnCopy)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnHelp)
        Me.Name = "dgQuickCode"
        Me.Text = "Code View"
        Me.ResumeLayout(False)

    End Sub
    '    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    '        Me.components = New System.ComponentModel.Container
    '        Me.txtQuickCode = New System.Windows.Forms.TextBox
    '        Me.btnCopy = New System.Windows.Forms.Button
    '        Me.btnClose = New System.Windows.Forms.Button
    '        Me.btnSave = New System.Windows.Forms.Button
    '        Me.btnHelp = New System.Windows.Forms.Button
    '        Me.Tip = New System.Windows.Forms.ToolTip(Me.components)
    '        Me.SuspendLayout()
    '        '
    '        'txtQuickCode
    '        '
    '        Me.txtQuickCode.AcceptsReturn = True
    '        Me.txtQuickCode.AcceptsTab = True
    '        Me.txtQuickCode.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
    '                    Or System.Windows.Forms.AnchorStyles.Left) _
    '                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    '        Me.txtQuickCode.BorderStyle = System.Windows.Forms.BorderStyle.None
    '        Me.txtQuickCode.CausesValidation = False
    '        Me.txtQuickCode.Multiline = True
    '        Me.txtQuickCode.Name = "txtQuickCode"
    '        Me.txtQuickCode.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
    '        Me.txtQuickCode.Bounds = New System.Drawing.Rectangle(0, 0, 568, 352)
    '        Me.txtQuickCode.TabIndex = 0
    '        Me.txtQuickCode.Text = ""
    '        '
    '        'btnCopy
    '        '
    '        Me.btnCopy.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    '        Me.btnCopy.Name = "btnCopy"
    '        Me.btnCopy.Bounds = New System.Drawing.Rectangle(360, 376, 80, 23)
    '        Me.btnCopy.TabIndex = 4
    '        Me.btnCopy.Text = "Copy"
    '        '
    '        'btnClose
    '        '
    '        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    '        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
    '        Me.btnClose.Name = "btnClose"
    '        Me.btnClose.Bounds = New System.Drawing.Rectangle(448, 376, 80, 23)
    '        Me.btnClose.TabIndex = 3
    '        Me.btnClose.Text = "Close"
    '        '
    '        'btnSave
    '        '
    '        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    '        Me.btnSave.Name = "btnSave"
    '        Me.btnSave.Bounds = New System.Drawing.Rectangle(120, 376, 104, 23)
    '        Me.btnSave.TabIndex = 1
    '        Me.btnSave.Text = "Save To Text File"
    '        '
    '        'btnHelp
    '        '
    '        Me.btnHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    '        Me.btnHelp.Name = "btnHelp"
    '        Me.btnHelp.Bounds = New System.Drawing.Rectangle(272, 376, 80, 23)
    '        Me.btnHelp.TabIndex = 5
    '        Me.btnHelp.Text = "Help"
    '        '
    '        'dgQuickCode
    '        '
    '        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    '        Me.CancelButton = Me.btnClose
    '        Me.ClientSize = New System.Drawing.Size(568, 405)
    '        Me.Controls.Add(Me.txtQuickCode)
    '        Me.Controls.Add(Me.btnCopy)
    '        Me.Controls.Add(Me.btnClose)
    '        Me.Controls.Add(Me.btnSave)
    '        Me.Controls.Add(Me.btnHelp)
    '        Me.Name = "dgQuickCode"
    '        Me.Text = "Code View"
    '        Me.ResumeLayout(False)
    '
    '    End Sub
    '

#End Region

#Region "Event Handlers"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Closes the quick code window
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Copies quick code to the clip board.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCopy.Click
        System.Windows.Forms.Clipboard.SetDataObject(txtQuickCode.Text, True)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Initiates a quick code save.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim dgSave As New SaveFileDialog
        Dim sw As System.IO.StreamWriter

        dgSave.Filter = Me.const_text_Filter

        Dim dgResult As DialogResult = dgSave.ShowDialog()

        If dgResult = DialogResult.OK Then
            Try
                sw = New System.IO.StreamWriter(dgSave.FileName)
                sw.Write(txtQuickCode.Text)

            Catch ex As Exception
                MsgBox("Could not access the file specified in file name")

            Finally
                If Not sw Is Nothing Then
                    sw.Close()
                End If
            End Try


        End If

        dgSave.Dispose()

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invokes help for the quick code window.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHelp.Click
        App.HelpManager.InvokeHelpContents("QuickCode.htm")
    End Sub

#End Region
End Class
