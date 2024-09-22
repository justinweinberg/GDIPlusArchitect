Imports GDIObjects

Friend Class CodeWin
    Inherits System.Windows.Forms.UserControl


#Region "Type Declarations"

    Private Enum EnumPanelView As Integer
        eDocument = 0
        eSelection = 1
        eSVGSelection = 2
    End Enum

#End Region

#Region "Local Fields"


    '''<summary>What mode to generate code for.  The options are eDocument 
    ''' (entire document), eSelection (the selected objects on the surface),
    '''  or eSVGSelection (generate SVG for the selection)</summary>
    Private _PanelView As EnumPanelView = EnumPanelView.eDocument
#End Region

#Region "Refresh"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Refreshes the currently generated code.  This method first checks if code 
    ''' generation has been disabled under user options since this is an expensive 
    ''' bit.  It then determines the mode to generate code in and asks the GDIDocument
    ''' specified in the doc argument to generate the code.
    ''' </summary>
    ''' <param name="doc">The document to generate code for</param>
    ''' -----------------------------------------------------------------------------
    Public Sub refreshPanel()
        Dim sCode As String

        If Not App.Options.DisableCodePanel And Not MDIMain.ActiveDocument Is Nothing Then
            Select Case _PanelView

                Case EnumPanelView.eDocument
                    sCode = MDIMain.ActiveDocument.GenerateCode
                Case EnumPanelView.eSelection
                    sCode = MDIMain.ActiveDocument.QuickCode()
                Case EnumPanelView.eSVGSelection
                    sCode = MDIMain.ActiveDocument.QuickSVG()
            End Select

        End If

        txtQuickCode.Text = sCode
    End Sub

#End Region

#Region "Constructors"


    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        cboView.Items.Add(New NVPair(EnumPanelView.eDocument, "Entire Document"))
        cboView.Items.Add(New NVPair(EnumPanelView.eSelection, "Selected"))
        cboView.Items.Add(New NVPair(EnumPanelView.eSVGSelection, "Selected SVG"))
        'Add any initialization after the InitializeComponent() call
        cboView.SelectedIndex = 0
    End Sub

#End Region

#Region " Windows Form Designer generated code "


    'UserControl overrides dispose to clean up the component list.
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
    Private WithEvents cboView As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtQuickCode = New System.Windows.Forms.TextBox
        Me.cboView = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'txtQuickCode
        '
        Me.txtQuickCode.AcceptsReturn = True
        Me.txtQuickCode.AcceptsTab = True
        Me.txtQuickCode.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtQuickCode.CausesValidation = False
        Me.txtQuickCode.Location = New System.Drawing.Point(0, 24)
        Me.txtQuickCode.Multiline = True
        Me.txtQuickCode.Name = "txtQuickCode"
        Me.txtQuickCode.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtQuickCode.Size = New System.Drawing.Size(528, 328)
        Me.txtQuickCode.TabIndex = 1
        Me.txtQuickCode.Text = ""
        '
        'cboView
        '
        Me.cboView.Dock = System.Windows.Forms.DockStyle.Top
        Me.cboView.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboView.Location = New System.Drawing.Point(0, 0)
        Me.cboView.Name = "cboView"
        Me.cboView.Size = New System.Drawing.Size(528, 21)
        Me.cboView.TabIndex = 2
        '
        'CodeWin
        '
        Me.Controls.Add(Me.txtQuickCode)
        Me.Controls.Add(Me.cboView)
        Me.Name = "CodeWin"
        Me.Size = New System.Drawing.Size(528, 360)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Event Handlers"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the drop down view combo, changing what type of code is generated.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboView_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboView.SelectedIndexChanged
        _PanelView = CType(cboView.SelectedIndex, EnumPanelView)
        refreshPanel()
    End Sub


#End Region


End Class
