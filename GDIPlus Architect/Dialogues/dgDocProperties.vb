Imports GDIObjects


''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : dgDocProperties
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides an interface for changing properties on GDIDocuments as well as 
''' gathering statistics on the document. This dialog is shown when the 
''' properties command is selected from the main menu.
''' Most of the options here have to do with code export.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class dgDocProperties
    Inherits System.Windows.Forms.Form

#Region "Type Declarations"


    ''' -----------------------------------------------------------------------------
    ''' Project	 : GDIPlus Architect
    ''' Class	 : dgDocProperties.ColumnSort
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Implements a sortable column for a list view.  The sortable list view columns 
    ''' are used when the list of image exports are rendered 
    ''' </summary>
    '''-----------------------------------------------------------------------------
    Protected Class ColumnSort
        Implements IComparer

        '''<summary>The sorter's column index</summary>
        Private _Column As Int32

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Creates a new instance of a sortable column given a column number
        ''' </summary>
        ''' <param name="columnSorted">The column index of this sorter</param>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal columnSorted As Int32)
            _Column = columnSorted
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Compares two distinct items in a list view. 
        '''  Implements IComparable to indicate which list view item 
        ''' should be before another list view item.
        ''' </summary>
        ''' <param name="a">The first item to compare</param>
        ''' <param name="b">The second item to compare</param>
        ''' <returns>A boolean indicating if column A is before column B</returns>
        ''' <remarks>All comparisons in this version of a ColumnSort are string based.  The 
        ''' example can be expanded upon to do custom sorting based on data types.
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Function Compare(ByVal a As Object, ByVal b As Object) As Int32 Implements IComparer.Compare
            Dim Al As ListViewItem = DirectCast(a, ListViewItem)
            Dim Bl As ListViewItem = DirectCast(b, ListViewItem)

            If Al.ListView.Sorting = System.Windows.Forms.SortOrder.Ascending Then

                '//here i grab the enum value for what type of sorting the user has
                'specified...i really don't know how

                '//i am able to access the listView properties (the parent of the
                'listitem)....but it works

                Return String.Compare(Al.SubItems(_Column).Text, Bl.SubItems(_Column).Text) * -1
            Else

                Return String.Compare(Al.SubItems(_Column).Text, Bl.SubItems(_Column).Text)

            End If
        End Function
    End Class


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to a column click, initiating sorting 
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard ColumnClickEventArgs</param>
    '''-----------------------------------------------------------------------------
    Private Sub lvImageDetails_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvImageDetails.ColumnClick
        If lvImageDetails.Sorting = SortOrder.Descending Then
            lvImageDetails.Sorting = System.Windows.Forms.SortOrder.Ascending
        Else
            lvImageDetails.Sorting = System.Windows.Forms.SortOrder.Descending

        End If

        lvImageDetails.ListViewItemSorter = New ColumnSort(e.Column)
        lvImageDetails.Sort()

    End Sub

#End Region


#Region "Local Fields"


    '''<summary>Counter for how many rectangles are in the document</summary>
    Private _RectangleCount As Int32 = 0
    '''<summary>Counter for how many ellipses are in the document</summary>
    Private _EllipseCount As Int32 = 0
    '''<summary>Counter for how many lines are in the document</summary>
    Private _LineCount As Int32 = 0
    '''<summary>Counter for how many images are in the document</summary>
    Private _ImageCount As Int32 = 0
    '''<summary>Counter for how many Text objects are in the document</summary>
    Private _TextCount As Int32 = 0
    '''<summary>Counter for how many GDIFields are in the document</summary>
    Private _FieldCount As Int32 = 0
    '''<summary>Counter for how many Paths (open or closed) are in the document</summary>
    Private _PathCount As Int32 = 0
    '''<summary>Counter for how many solid fills are in the document </summary>
    Private _SolidBrushCount As Int32 = 0
    '''<summary>Counter for how many gradient fills are in the document </summary>
    Private _GradientBrushCount As Int32 = 0
    '''<summary>Counter for how texture fills are in the document </summary>
    Private _TextureBrushCount As Int32 = 0
    '''<summary>Counter for how many hatch fills are in the document </summary>
    Private _HatchBrushCount As Int32 = 0

    '''<summary>Counter for how many strokes are in the document </summary>
    Private _PenCount As Int32 = 0
    '''<summary>Holds the transient page settings which can be applied to 
    ''' the repeating page style of document.</summary>
    Private _PageSettings As System.Drawing.Printing.PageSettings


#End Region

#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populates the text hint drop down and assigns the current value from the 
    ''' GDIDocument.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub populateHintDropDown()


        With cboTextRenderingHint
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(System.Drawing.Text.TextRenderingHint)))
            .EndUpdate()
            .SelectedItem = MDIMain.ActiveDocument.TextRenderingHint.ToString

        End With

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populates the smoothing drop down and assigns the current value from 
    ''' the GDIDocument
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub populateSmoothingDropDown()

        With cboSmoothingMode
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(Drawing2D.SmoothingMode)))
            .Items.Remove(System.Enum.GetName(GetType(Drawing2D.SmoothingMode), Drawing2D.SmoothingMode.Invalid))
            .EndUpdate()
            .SelectedItem = MDIMain.ActiveDocument.SmoothingMode.ToString
        End With


    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Showes and hides as well as initialized controls based off of whether the 
    ''' current GDIDocument is a PrintDocument type or a GraphicsClass type.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub showTypeSpecifics()

        'Toggle the view based on whether it's displaying a printodcument or a graphics class.
        If MDIMain.ActiveDocument.ExportSettings.DocumentType = EnumDocumentTypes.ePrintDocument Then
            lblPageSize.Text = MDIMain.ActiveDocument.RectPageSize.ToString
            pnPrintSpecific.Visible = True
            pnControlSpecific.Visible = False
            btnPageSetup.Enabled = True
        Else
            btnPageSetup.Enabled = False
            picBackColor.BackColor = MDIMain.ActiveDocument.BackColor
            txtPixelsX.Text = MDIMain.ActiveDocument.RectPageSize.Width.ToString
            txtPixelsY.Text = MDIMain.ActiveDocument.RectPageSize.Height.ToString
            pnPrintSpecific.Visible = False
            pnControlSpecific.Visible = True
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populates the form with the current export settings from the active document
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub populateExportSettings()

        With MDIMain.ActiveDocument.ExportSettings
            'Populate text boxes
            txtClassName.Text = .ClassName
            txtRootNamespace.Text = .RootnameSpace
            txtMemberVariablePrefix.Text = .MemberPrefix
            txtFieldPropertyPrefix.Text = .FieldPrefix

            'Populate check boxes
            chkFills.Checked = .OverrideConsolidateFill
            chkStrokes.Checked = .OverrideConsolidateStroke
            chkFonts.Checked = .OverrideConsolidateFont
            chkStringFormats.Checked = .OverrideConsolidateStringFormats

        End With
     
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the document properties object given a document and 
    ''' the MDIMain's current user indicated page settings.
    ''' page se
    ''' </summary>
    ''' <param name="pgsettings">Current page settings from MDIMain.</param>
    '''-----------------------------------------------------------------------------
    Public Sub New(ByVal pgsettings As System.Drawing.Printing.PageSettings)
        MyBase.New()

        InitializeComponent()


        'Set the local page settings property to the page settings from MDIMain.
        _PageSettings = pgsettings

        'fill dropdown lists and set current values
        populateHintDropDown()
        populateSmoothingDropDown()


        'Set the language option buttons
        Select Case MDIMain.ActiveDocument.ExportSettings.Language
            Case EnumCodeTypes.eCSharp
                optCSharp.Checked = True

            Case EnumCodeTypes.eVB
                optVB.Checked = True
        End Select

        'Set elements specific to whether the document is a PrintDocument or a graphics class
        showTypeSpecifics()

        populateExportSettings()

        'gather statistics on the GDIDocument
        GatherStatistics()

        App.ToolTipManager.PopulatePopupTip(Tip, "DocumentProperties", Me)

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
    Private WithEvents clSource As System.Windows.Forms.ColumnHeader
    Private WithEvents tabSummary As System.Windows.Forms.TabControl
    Private WithEvents tbProperties As System.Windows.Forms.TabPage
    Private WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents txtRootNamespace As System.Windows.Forms.TextBox
    Private WithEvents pnControlSpecific As System.Windows.Forms.Panel
    Private WithEvents pnPrintSpecific As System.Windows.Forms.Panel
    Private WithEvents tbStatistics As System.Windows.Forms.TabPage
    Private WithEvents Label16 As System.Windows.Forms.Label
    Private WithEvents Label14 As System.Windows.Forms.Label
    Private WithEvents lblImages As System.Windows.Forms.Label
    Private WithEvents Label15 As System.Windows.Forms.Label
    Private WithEvents lblPaths As System.Windows.Forms.Label
    Private WithEvents lblFields As System.Windows.Forms.Label
    Private WithEvents lblRectangles As System.Windows.Forms.Label
    Private WithEvents lblEllipses As System.Windows.Forms.Label
    Private WithEvents lblLines As System.Windows.Forms.Label
    Private WithEvents lblText As System.Windows.Forms.Label
    Private WithEvents lblTotalPens As System.Windows.Forms.Label
    Private WithEvents Label13 As System.Windows.Forms.Label
    Private WithEvents lblTotalBrushes As System.Windows.Forms.Label
    Private WithEvents Label20 As System.Windows.Forms.Label
    Private WithEvents Label19 As System.Windows.Forms.Label
    Private WithEvents Label18 As System.Windows.Forms.Label
    Private WithEvents Label17 As System.Windows.Forms.Label
    Private WithEvents lblHatchBrushes As System.Windows.Forms.Label
    Private WithEvents lblGradientBrushes As System.Windows.Forms.Label
    Private WithEvents lblTextureBrushes As System.Windows.Forms.Label
    Private WithEvents lblSolidBrushes As System.Windows.Forms.Label
    Private WithEvents Label8 As System.Windows.Forms.Label
    Private WithEvents Label10 As System.Windows.Forms.Label
    Private WithEvents Label11 As System.Windows.Forms.Label
    Private WithEvents Label12 As System.Windows.Forms.Label
    Private WithEvents Label6 As System.Windows.Forms.Label
    Private WithEvents Label5 As System.Windows.Forms.Label
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents tbImages As System.Windows.Forms.TabPage
    Private WithEvents lvImageDetails As System.Windows.Forms.ListView
    Private WithEvents clName As System.Windows.Forms.ColumnHeader
    Private WithEvents clType As System.Windows.Forms.ColumnHeader
    Private WithEvents clAction As System.Windows.Forms.ColumnHeader
    Private WithEvents clPath As System.Windows.Forms.ColumnHeader
    Private WithEvents btnPageSetup As System.Windows.Forms.Button
    Private WithEvents btnCancel As System.Windows.Forms.Button
    Private WithEvents txtPixelsX As System.Windows.Forms.TextBox
    Private WithEvents txtPixelsY As System.Windows.Forms.TextBox
    Private WithEvents picBackColor As System.Windows.Forms.PictureBox
    Private WithEvents Label4 As System.Windows.Forms.Label
    Private WithEvents Label9 As System.Windows.Forms.Label
    Private WithEvents Label7 As System.Windows.Forms.Label
    Private WithEvents Label21 As System.Windows.Forms.Label
    Private WithEvents lblPageSize As System.Windows.Forms.Label
    Private WithEvents btnViewCode As System.Windows.Forms.Button
    Private WithEvents grpQuickCode As System.Windows.Forms.GroupBox
    Private WithEvents optCSharp As System.Windows.Forms.RadioButton
    Private WithEvents optVB As System.Windows.Forms.RadioButton
    Private WithEvents txtClassName As System.Windows.Forms.TextBox
    Private WithEvents Label22 As System.Windows.Forms.Label
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents Label23 As System.Windows.Forms.Label
    Private WithEvents cboSmoothingMode As System.Windows.Forms.ComboBox
    Private WithEvents cboTextRenderingHint As System.Windows.Forms.ComboBox
    Private WithEvents btnHelp As System.Windows.Forms.Button
    Private WithEvents txtFieldPropertyPrefix As System.Windows.Forms.TextBox
    Private WithEvents Label24 As System.Windows.Forms.Label
    Private WithEvents txtMemberVariablePrefix As System.Windows.Forms.TextBox
    Private WithEvents Label25 As System.Windows.Forms.Label
    Private WithEvents grpConsolidation As System.Windows.Forms.GroupBox
    Private WithEvents chkFills As System.Windows.Forms.CheckBox
    Private WithEvents chkStrokes As System.Windows.Forms.CheckBox
    Private WithEvents chkFonts As System.Windows.Forms.CheckBox
    Private WithEvents chkStringFormats As System.Windows.Forms.CheckBox
    Private WithEvents btnApply As System.Windows.Forms.Button
    Private WithEvents Tip As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.btnOk = New System.Windows.Forms.Button
        Me.clSource = New System.Windows.Forms.ColumnHeader
        Me.tabSummary = New System.Windows.Forms.TabControl
        Me.tbProperties = New System.Windows.Forms.TabPage
        Me.grpConsolidation = New System.Windows.Forms.GroupBox
        Me.chkStringFormats = New System.Windows.Forms.CheckBox
        Me.chkFonts = New System.Windows.Forms.CheckBox
        Me.chkStrokes = New System.Windows.Forms.CheckBox
        Me.chkFills = New System.Windows.Forms.CheckBox
        Me.txtFieldPropertyPrefix = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.txtMemberVariablePrefix = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.cboSmoothingMode = New System.Windows.Forms.ComboBox
        Me.cboTextRenderingHint = New System.Windows.Forms.ComboBox
        Me.txtClassName = New System.Windows.Forms.TextBox
        Me.grpQuickCode = New System.Windows.Forms.GroupBox
        Me.optCSharp = New System.Windows.Forms.RadioButton
        Me.optVB = New System.Windows.Forms.RadioButton
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtRootNamespace = New System.Windows.Forms.TextBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label23 = New System.Windows.Forms.Label
        Me.pnControlSpecific = New System.Windows.Forms.Panel
        Me.txtPixelsX = New System.Windows.Forms.TextBox
        Me.txtPixelsY = New System.Windows.Forms.TextBox
        Me.picBackColor = New System.Windows.Forms.PictureBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.pnPrintSpecific = New System.Windows.Forms.Panel
        Me.lblPageSize = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.tbStatistics = New System.Windows.Forms.TabPage
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.lblImages = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.lblPaths = New System.Windows.Forms.Label
        Me.lblFields = New System.Windows.Forms.Label
        Me.lblRectangles = New System.Windows.Forms.Label
        Me.lblEllipses = New System.Windows.Forms.Label
        Me.lblLines = New System.Windows.Forms.Label
        Me.lblText = New System.Windows.Forms.Label
        Me.lblTotalPens = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.lblTotalBrushes = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.lblHatchBrushes = New System.Windows.Forms.Label
        Me.lblGradientBrushes = New System.Windows.Forms.Label
        Me.lblTextureBrushes = New System.Windows.Forms.Label
        Me.lblSolidBrushes = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.tbImages = New System.Windows.Forms.TabPage
        Me.lvImageDetails = New System.Windows.Forms.ListView
        Me.clName = New System.Windows.Forms.ColumnHeader
        Me.clType = New System.Windows.Forms.ColumnHeader
        Me.clAction = New System.Windows.Forms.ColumnHeader
        Me.clPath = New System.Windows.Forms.ColumnHeader
        Me.btnPageSetup = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnViewCode = New System.Windows.Forms.Button
        Me.btnHelp = New System.Windows.Forms.Button
        Me.btnApply = New System.Windows.Forms.Button
        Me.Tip = New System.Windows.Forms.ToolTip(Me.components)
        Me.tabSummary.SuspendLayout()
        Me.tbProperties.SuspendLayout()
        Me.grpConsolidation.SuspendLayout()
        Me.grpQuickCode.SuspendLayout()
        Me.pnControlSpecific.SuspendLayout()
        Me.pnPrintSpecific.SuspendLayout()
        Me.tbStatistics.SuspendLayout()
        Me.tbImages.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnOk
        '
        Me.btnOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnOk.Location = New System.Drawing.Point(376, 472)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.TabIndex = 22
        Me.btnOk.Text = "Ok"
        '
        'clSource
        '
        Me.clSource.Text = "Source"
        Me.clSource.Width = 150
        '
        'tabSummary
        '
        Me.tabSummary.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabSummary.Controls.Add(Me.tbProperties)
        Me.tabSummary.Controls.Add(Me.tbStatistics)
        Me.tabSummary.Controls.Add(Me.tbImages)
        Me.tabSummary.Location = New System.Drawing.Point(4, 8)
        Me.tabSummary.Name = "tabSummary"
        Me.tabSummary.SelectedIndex = 0
        Me.tabSummary.Size = New System.Drawing.Size(540, 440)
        Me.tabSummary.TabIndex = 21
        '
        'tbProperties
        '
        Me.tbProperties.Controls.Add(Me.grpConsolidation)
        Me.tbProperties.Controls.Add(Me.txtFieldPropertyPrefix)
        Me.tbProperties.Controls.Add(Me.Label24)
        Me.tbProperties.Controls.Add(Me.txtMemberVariablePrefix)
        Me.tbProperties.Controls.Add(Me.Label25)
        Me.tbProperties.Controls.Add(Me.cboSmoothingMode)
        Me.tbProperties.Controls.Add(Me.cboTextRenderingHint)
        Me.tbProperties.Controls.Add(Me.txtClassName)
        Me.tbProperties.Controls.Add(Me.grpQuickCode)
        Me.tbProperties.Controls.Add(Me.Label1)
        Me.tbProperties.Controls.Add(Me.txtRootNamespace)
        Me.tbProperties.Controls.Add(Me.Label22)
        Me.tbProperties.Controls.Add(Me.Label2)
        Me.tbProperties.Controls.Add(Me.Label23)
        Me.tbProperties.Controls.Add(Me.pnControlSpecific)
        Me.tbProperties.Controls.Add(Me.pnPrintSpecific)
        Me.tbProperties.Location = New System.Drawing.Point(4, 22)
        Me.tbProperties.Name = "tbProperties"
        Me.tbProperties.Size = New System.Drawing.Size(532, 414)
        Me.tbProperties.TabIndex = 3
        Me.tbProperties.Text = "Properties"
        '
        'grpConsolidation
        '
        Me.grpConsolidation.Controls.Add(Me.chkStringFormats)
        Me.grpConsolidation.Controls.Add(Me.chkFonts)
        Me.grpConsolidation.Controls.Add(Me.chkStrokes)
        Me.grpConsolidation.Controls.Add(Me.chkFills)
        Me.grpConsolidation.Location = New System.Drawing.Point(24, 208)
        Me.grpConsolidation.Name = "grpConsolidation"
        Me.grpConsolidation.Size = New System.Drawing.Size(432, 56)
        Me.grpConsolidation.TabIndex = 26
        Me.grpConsolidation.TabStop = False
        Me.grpConsolidation.Text = "Override Consolidation Settings For"
        '
        'chkStringFormats
        '
        Me.chkStringFormats.Location = New System.Drawing.Point(264, 24)
        Me.chkStringFormats.Name = "chkStringFormats"
        Me.chkStringFormats.TabIndex = 5
        Me.chkStringFormats.Text = "String Formats"
        '
        'chkFonts
        '
        Me.chkFonts.Location = New System.Drawing.Point(176, 24)
        Me.chkFonts.Name = "chkFonts"
        Me.chkFonts.Size = New System.Drawing.Size(56, 24)
        Me.chkFonts.TabIndex = 3
        Me.chkFonts.Text = "Fonts"
        '
        'chkStrokes
        '
        Me.chkStrokes.Location = New System.Drawing.Point(88, 24)
        Me.chkStrokes.Name = "chkStrokes"
        Me.chkStrokes.Size = New System.Drawing.Size(64, 24)
        Me.chkStrokes.TabIndex = 2
        Me.chkStrokes.Text = "Strokes"
        '
        'chkFills
        '
        Me.chkFills.Location = New System.Drawing.Point(16, 24)
        Me.chkFills.Name = "chkFills"
        Me.chkFills.Size = New System.Drawing.Size(48, 24)
        Me.chkFills.TabIndex = 1
        Me.chkFills.Text = "Fills"
        '
        'txtFieldPropertyPrefix
        '
        Me.txtFieldPropertyPrefix.Location = New System.Drawing.Point(336, 184)
        Me.txtFieldPropertyPrefix.Name = "txtFieldPropertyPrefix"
        Me.txtFieldPropertyPrefix.Size = New System.Drawing.Size(40, 20)
        Me.txtFieldPropertyPrefix.TabIndex = 25
        Me.txtFieldPropertyPrefix.Text = ""
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(224, 184)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(120, 23)
        Me.Label24.TabIndex = 24
        Me.Label24.Text = "Field property prefix:"
        '
        'txtMemberVariablePrefix
        '
        Me.txtMemberVariablePrefix.Location = New System.Drawing.Point(160, 184)
        Me.txtMemberVariablePrefix.Name = "txtMemberVariablePrefix"
        Me.txtMemberVariablePrefix.Size = New System.Drawing.Size(40, 20)
        Me.txtMemberVariablePrefix.TabIndex = 22
        Me.txtMemberVariablePrefix.Text = ""
        '
        'Label25
        '
        Me.Label25.Location = New System.Drawing.Point(32, 184)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(128, 23)
        Me.Label25.TabIndex = 23
        Me.Label25.Text = "Member variable prefix:"
        '
        'cboSmoothingMode
        '
        Me.cboSmoothingMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSmoothingMode.Location = New System.Drawing.Point(168, 152)
        Me.cboSmoothingMode.Name = "cboSmoothingMode"
        Me.cboSmoothingMode.Size = New System.Drawing.Size(176, 21)
        Me.cboSmoothingMode.TabIndex = 19
        '
        'cboTextRenderingHint
        '
        Me.cboTextRenderingHint.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTextRenderingHint.Location = New System.Drawing.Point(168, 128)
        Me.cboTextRenderingHint.Name = "cboTextRenderingHint"
        Me.cboTextRenderingHint.Size = New System.Drawing.Size(176, 21)
        Me.cboTextRenderingHint.TabIndex = 18
        '
        'txtClassName
        '
        Me.txtClassName.Location = New System.Drawing.Point(104, 104)
        Me.txtClassName.Name = "txtClassName"
        Me.txtClassName.Size = New System.Drawing.Size(160, 20)
        Me.txtClassName.TabIndex = 16
        Me.txtClassName.Text = ""
        '
        'grpQuickCode
        '
        Me.grpQuickCode.Controls.Add(Me.optCSharp)
        Me.grpQuickCode.Controls.Add(Me.optVB)
        Me.grpQuickCode.Location = New System.Drawing.Point(32, 40)
        Me.grpQuickCode.Name = "grpQuickCode"
        Me.grpQuickCode.Size = New System.Drawing.Size(168, 56)
        Me.grpQuickCode.TabIndex = 15
        Me.grpQuickCode.TabStop = False
        Me.grpQuickCode.Text = "Export Language"
        '
        'optCSharp
        '
        Me.optCSharp.Location = New System.Drawing.Point(112, 24)
        Me.optCSharp.Name = "optCSharp"
        Me.optCSharp.Size = New System.Drawing.Size(40, 24)
        Me.optCSharp.TabIndex = 1
        Me.optCSharp.Text = "C#"
        '
        'optVB
        '
        Me.optVB.Location = New System.Drawing.Point(16, 24)
        Me.optVB.Name = "optVB"
        Me.optVB.Size = New System.Drawing.Size(64, 24)
        Me.optVB.TabIndex = 0
        Me.optVB.Text = "VB.NET"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(32, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(168, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Default embedded namespace:"
        '
        'txtRootNamespace
        '
        Me.txtRootNamespace.Location = New System.Drawing.Point(200, 16)
        Me.txtRootNamespace.Name = "txtRootNamespace"
        Me.txtRootNamespace.Size = New System.Drawing.Size(264, 20)
        Me.txtRootNamespace.TabIndex = 0
        Me.txtRootNamespace.Text = ""
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(32, 104)
        Me.Label22.Name = "Label22"
        Me.Label22.TabIndex = 17
        Me.Label22.Text = "Class name:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(32, 128)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(112, 23)
        Me.Label2.TabIndex = 20
        Me.Label2.Text = "Text Rendering Hint:"
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(32, 152)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(96, 23)
        Me.Label23.TabIndex = 21
        Me.Label23.Text = "Smoothing Mode:"
        '
        'pnControlSpecific
        '
        Me.pnControlSpecific.Controls.Add(Me.txtPixelsX)
        Me.pnControlSpecific.Controls.Add(Me.txtPixelsY)
        Me.pnControlSpecific.Controls.Add(Me.picBackColor)
        Me.pnControlSpecific.Controls.Add(Me.Label4)
        Me.pnControlSpecific.Controls.Add(Me.Label9)
        Me.pnControlSpecific.Controls.Add(Me.Label7)
        Me.pnControlSpecific.Location = New System.Drawing.Point(24, 272)
        Me.pnControlSpecific.Name = "pnControlSpecific"
        Me.pnControlSpecific.Size = New System.Drawing.Size(432, 128)
        Me.pnControlSpecific.TabIndex = 2
        '
        'txtPixelsX
        '
        Me.txtPixelsX.Location = New System.Drawing.Point(96, 16)
        Me.txtPixelsX.Name = "txtPixelsX"
        Me.txtPixelsX.Size = New System.Drawing.Size(40, 20)
        Me.txtPixelsX.TabIndex = 34
        Me.txtPixelsX.Text = ""
        '
        'txtPixelsY
        '
        Me.txtPixelsY.Location = New System.Drawing.Point(96, 40)
        Me.txtPixelsY.Name = "txtPixelsY"
        Me.txtPixelsY.Size = New System.Drawing.Size(40, 20)
        Me.txtPixelsY.TabIndex = 35
        Me.txtPixelsY.Text = ""
        '
        'picBackColor
        '
        Me.picBackColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picBackColor.Location = New System.Drawing.Point(96, 80)
        Me.picBackColor.Name = "picBackColor"
        Me.picBackColor.Size = New System.Drawing.Size(48, 40)
        Me.picBackColor.TabIndex = 38
        Me.picBackColor.TabStop = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 16)
        Me.Label4.TabIndex = 39
        Me.Label4.Text = "Backcolor:"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(16, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(48, 16)
        Me.Label9.TabIndex = 36
        Me.Label9.Text = "Width:"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(16, 40)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(56, 16)
        Me.Label7.TabIndex = 37
        Me.Label7.Text = "Height:"
        '
        'pnPrintSpecific
        '
        Me.pnPrintSpecific.Controls.Add(Me.lblPageSize)
        Me.pnPrintSpecific.Controls.Add(Me.Label21)
        Me.pnPrintSpecific.Location = New System.Drawing.Point(32, 272)
        Me.pnPrintSpecific.Name = "pnPrintSpecific"
        Me.pnPrintSpecific.Size = New System.Drawing.Size(432, 128)
        Me.pnPrintSpecific.TabIndex = 3
        '
        'lblPageSize
        '
        Me.lblPageSize.Location = New System.Drawing.Point(128, 16)
        Me.lblPageSize.Name = "lblPageSize"
        Me.lblPageSize.Size = New System.Drawing.Size(272, 23)
        Me.lblPageSize.TabIndex = 2
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(16, 16)
        Me.Label21.Name = "Label21"
        Me.Label21.TabIndex = 0
        Me.Label21.Text = "Page Size:"
        '
        'tbStatistics
        '
        Me.tbStatistics.Controls.Add(Me.Label16)
        Me.tbStatistics.Controls.Add(Me.Label14)
        Me.tbStatistics.Controls.Add(Me.lblImages)
        Me.tbStatistics.Controls.Add(Me.Label15)
        Me.tbStatistics.Controls.Add(Me.lblPaths)
        Me.tbStatistics.Controls.Add(Me.lblFields)
        Me.tbStatistics.Controls.Add(Me.lblRectangles)
        Me.tbStatistics.Controls.Add(Me.lblEllipses)
        Me.tbStatistics.Controls.Add(Me.lblLines)
        Me.tbStatistics.Controls.Add(Me.lblText)
        Me.tbStatistics.Controls.Add(Me.lblTotalPens)
        Me.tbStatistics.Controls.Add(Me.Label13)
        Me.tbStatistics.Controls.Add(Me.lblTotalBrushes)
        Me.tbStatistics.Controls.Add(Me.Label20)
        Me.tbStatistics.Controls.Add(Me.Label19)
        Me.tbStatistics.Controls.Add(Me.Label18)
        Me.tbStatistics.Controls.Add(Me.Label17)
        Me.tbStatistics.Controls.Add(Me.lblHatchBrushes)
        Me.tbStatistics.Controls.Add(Me.lblGradientBrushes)
        Me.tbStatistics.Controls.Add(Me.lblTextureBrushes)
        Me.tbStatistics.Controls.Add(Me.lblSolidBrushes)
        Me.tbStatistics.Controls.Add(Me.Label8)
        Me.tbStatistics.Controls.Add(Me.Label10)
        Me.tbStatistics.Controls.Add(Me.Label11)
        Me.tbStatistics.Controls.Add(Me.Label12)
        Me.tbStatistics.Controls.Add(Me.Label6)
        Me.tbStatistics.Controls.Add(Me.Label5)
        Me.tbStatistics.Controls.Add(Me.Label3)
        Me.tbStatistics.Location = New System.Drawing.Point(4, 22)
        Me.tbStatistics.Name = "tbStatistics"
        Me.tbStatistics.Size = New System.Drawing.Size(532, 414)
        Me.tbStatistics.TabIndex = 0
        Me.tbStatistics.Text = "Statistics"
        Me.tbStatistics.Visible = False
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.SystemColors.Control
        Me.Label16.Location = New System.Drawing.Point(296, 8)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(64, 23)
        Me.Label16.TabIndex = 31
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.SystemColors.Control
        Me.Label14.Location = New System.Drawing.Point(200, 16)
        Me.Label14.Name = "Label14"
        Me.Label14.TabIndex = 30
        Me.Label14.Text = "Objects:"
        '
        'lblImages
        '
        Me.lblImages.BackColor = System.Drawing.SystemColors.Control
        Me.lblImages.Location = New System.Drawing.Point(296, 232)
        Me.lblImages.Name = "lblImages"
        Me.lblImages.Size = New System.Drawing.Size(64, 23)
        Me.lblImages.TabIndex = 29
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.SystemColors.Control
        Me.Label15.Location = New System.Drawing.Point(224, 232)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(64, 23)
        Me.Label15.TabIndex = 28
        Me.Label15.Text = "Images:"
        '
        'lblPaths
        '
        Me.lblPaths.BackColor = System.Drawing.SystemColors.Control
        Me.lblPaths.Location = New System.Drawing.Point(296, 200)
        Me.lblPaths.Name = "lblPaths"
        Me.lblPaths.Size = New System.Drawing.Size(64, 23)
        Me.lblPaths.TabIndex = 25
        '
        'lblFields
        '
        Me.lblFields.BackColor = System.Drawing.SystemColors.Control
        Me.lblFields.Location = New System.Drawing.Point(296, 40)
        Me.lblFields.Name = "lblFields"
        Me.lblFields.Size = New System.Drawing.Size(64, 23)
        Me.lblFields.TabIndex = 24
        '
        'lblRectangles
        '
        Me.lblRectangles.BackColor = System.Drawing.SystemColors.Control
        Me.lblRectangles.Location = New System.Drawing.Point(296, 104)
        Me.lblRectangles.Name = "lblRectangles"
        Me.lblRectangles.Size = New System.Drawing.Size(64, 23)
        Me.lblRectangles.TabIndex = 23
        '
        'lblEllipses
        '
        Me.lblEllipses.BackColor = System.Drawing.SystemColors.Control
        Me.lblEllipses.Location = New System.Drawing.Point(296, 136)
        Me.lblEllipses.Name = "lblEllipses"
        Me.lblEllipses.Size = New System.Drawing.Size(64, 23)
        Me.lblEllipses.TabIndex = 22
        '
        'lblLines
        '
        Me.lblLines.BackColor = System.Drawing.SystemColors.Control
        Me.lblLines.Location = New System.Drawing.Point(296, 168)
        Me.lblLines.Name = "lblLines"
        Me.lblLines.Size = New System.Drawing.Size(64, 23)
        Me.lblLines.TabIndex = 21
        '
        'lblText
        '
        Me.lblText.BackColor = System.Drawing.SystemColors.Control
        Me.lblText.Location = New System.Drawing.Point(296, 72)
        Me.lblText.Name = "lblText"
        Me.lblText.Size = New System.Drawing.Size(64, 23)
        Me.lblText.TabIndex = 20
        '
        'lblTotalPens
        '
        Me.lblTotalPens.BackColor = System.Drawing.SystemColors.Control
        Me.lblTotalPens.Location = New System.Drawing.Point(112, 232)
        Me.lblTotalPens.Name = "lblTotalPens"
        Me.lblTotalPens.Size = New System.Drawing.Size(64, 23)
        Me.lblTotalPens.TabIndex = 19
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Location = New System.Drawing.Point(16, 232)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(56, 23)
        Me.Label13.TabIndex = 18
        Me.Label13.Text = "Pens:"
        '
        'lblTotalBrushes
        '
        Me.lblTotalBrushes.BackColor = System.Drawing.SystemColors.Control
        Me.lblTotalBrushes.Location = New System.Drawing.Point(112, 16)
        Me.lblTotalBrushes.Name = "lblTotalBrushes"
        Me.lblTotalBrushes.Size = New System.Drawing.Size(64, 23)
        Me.lblTotalBrushes.TabIndex = 17
        '
        'Label20
        '
        Me.Label20.BackColor = System.Drawing.SystemColors.Control
        Me.Label20.Location = New System.Drawing.Point(40, 144)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(56, 23)
        Me.Label20.TabIndex = 16
        Me.Label20.Text = "Texture:"
        '
        'Label19
        '
        Me.Label19.BackColor = System.Drawing.SystemColors.Control
        Me.Label19.Location = New System.Drawing.Point(40, 112)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(56, 23)
        Me.Label19.TabIndex = 15
        Me.Label19.Text = "Gradient:"
        '
        'Label18
        '
        Me.Label18.BackColor = System.Drawing.SystemColors.Control
        Me.Label18.Location = New System.Drawing.Point(40, 80)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(56, 23)
        Me.Label18.TabIndex = 14
        Me.Label18.Text = "Hatch:"
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.SystemColors.Control
        Me.Label17.Location = New System.Drawing.Point(40, 48)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(56, 23)
        Me.Label17.TabIndex = 13
        Me.Label17.Text = "Solid:"
        '
        'lblHatchBrushes
        '
        Me.lblHatchBrushes.BackColor = System.Drawing.SystemColors.Control
        Me.lblHatchBrushes.Location = New System.Drawing.Point(112, 80)
        Me.lblHatchBrushes.Name = "lblHatchBrushes"
        Me.lblHatchBrushes.Size = New System.Drawing.Size(64, 23)
        Me.lblHatchBrushes.TabIndex = 12
        '
        'lblGradientBrushes
        '
        Me.lblGradientBrushes.BackColor = System.Drawing.SystemColors.Control
        Me.lblGradientBrushes.Location = New System.Drawing.Point(112, 112)
        Me.lblGradientBrushes.Name = "lblGradientBrushes"
        Me.lblGradientBrushes.Size = New System.Drawing.Size(64, 23)
        Me.lblGradientBrushes.TabIndex = 11
        '
        'lblTextureBrushes
        '
        Me.lblTextureBrushes.BackColor = System.Drawing.SystemColors.Control
        Me.lblTextureBrushes.Location = New System.Drawing.Point(112, 144)
        Me.lblTextureBrushes.Name = "lblTextureBrushes"
        Me.lblTextureBrushes.Size = New System.Drawing.Size(64, 23)
        Me.lblTextureBrushes.TabIndex = 10
        '
        'lblSolidBrushes
        '
        Me.lblSolidBrushes.BackColor = System.Drawing.SystemColors.Control
        Me.lblSolidBrushes.Location = New System.Drawing.Point(112, 48)
        Me.lblSolidBrushes.Name = "lblSolidBrushes"
        Me.lblSolidBrushes.Size = New System.Drawing.Size(64, 23)
        Me.lblSolidBrushes.TabIndex = 9
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Location = New System.Drawing.Point(224, 200)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(64, 23)
        Me.Label8.TabIndex = 8
        Me.Label8.Text = "Paths:"
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Location = New System.Drawing.Point(224, 168)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(64, 23)
        Me.Label10.TabIndex = 7
        Me.Label10.Text = "Lines:"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Location = New System.Drawing.Point(224, 136)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(64, 23)
        Me.Label11.TabIndex = 6
        Me.Label11.Text = "Ellipses:"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.SystemColors.Control
        Me.Label12.Location = New System.Drawing.Point(224, 104)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(64, 23)
        Me.Label12.TabIndex = 5
        Me.Label12.Text = "Rectangles:"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Location = New System.Drawing.Point(224, 72)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 23)
        Me.Label6.TabIndex = 3
        Me.Label6.Text = "Text:"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Location = New System.Drawing.Point(224, 40)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(64, 23)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Fields:"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Location = New System.Drawing.Point(16, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 23)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Brushes:"
        '
        'tbImages
        '
        Me.tbImages.Controls.Add(Me.lvImageDetails)
        Me.tbImages.Location = New System.Drawing.Point(4, 22)
        Me.tbImages.Name = "tbImages"
        Me.tbImages.Size = New System.Drawing.Size(532, 414)
        Me.tbImages.TabIndex = 2
        Me.tbImages.Text = "Images"
        Me.tbImages.Visible = False
        '
        'lvImageDetails
        '
        Me.lvImageDetails.AllowColumnReorder = True
        Me.lvImageDetails.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.clName, Me.clType, Me.clAction, Me.clPath, Me.clSource})
        Me.lvImageDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.lvImageDetails.HideSelection = False
        Me.lvImageDetails.Location = New System.Drawing.Point(0, 0)
        Me.lvImageDetails.Name = "lvImageDetails"
        Me.lvImageDetails.Size = New System.Drawing.Size(532, 374)
        Me.lvImageDetails.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lvImageDetails.TabIndex = 16
        Me.lvImageDetails.View = System.Windows.Forms.View.Details
        '
        'clName
        '
        Me.clName.Text = "Name"
        '
        'clType
        '
        Me.clType.Text = "Type"
        '
        'clAction
        '
        Me.clAction.Text = "Action"
        '
        'clPath
        '
        Me.clPath.Text = "Path"
        Me.clPath.Width = 150
        '
        'btnPageSetup
        '
        Me.btnPageSetup.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnPageSetup.Location = New System.Drawing.Point(24, 472)
        Me.btnPageSetup.Name = "btnPageSetup"
        Me.btnPageSetup.TabIndex = 25
        Me.btnPageSetup.Text = "Page Setup"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.CausesValidation = False
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(464, 472)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 23
        Me.btnCancel.Text = "Cancel"
        '
        'btnViewCode
        '
        Me.btnViewCode.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnViewCode.Location = New System.Drawing.Point(200, 472)
        Me.btnViewCode.Name = "btnViewCode"
        Me.btnViewCode.TabIndex = 26
        Me.btnViewCode.Text = "View Code"
        '
        'btnHelp
        '
        Me.btnHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnHelp.Location = New System.Drawing.Point(112, 472)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.TabIndex = 27
        Me.btnHelp.Text = "Help"
        '
        'btnApply
        '
        Me.btnApply.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnApply.Location = New System.Drawing.Point(288, 472)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.TabIndex = 28
        Me.btnApply.Text = "Apply"
        '
        'dgDocProperties
        '
        Me.AcceptButton = Me.btnOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(544, 517)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.tabSummary)
        Me.Controls.Add(Me.btnPageSetup)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnViewCode)
        Me.Controls.Add(Me.btnHelp)
        Me.Controls.Add(Me.btnApply)
        Me.Name = "dgDocProperties"
        Me.Text = "Summary"
        Me.tabSummary.ResumeLayout(False)
        Me.tbProperties.ResumeLayout(False)
        Me.grpConsolidation.ResumeLayout(False)
        Me.grpQuickCode.ResumeLayout(False)
        Me.pnControlSpecific.ResumeLayout(False)
        Me.pnPrintSpecific.ResumeLayout(False)
        Me.tbStatistics.ResumeLayout(False)
        Me.tbImages.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Helper Methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Append an image summary to the image list for a textured object
    ''' </summary>
    ''' <param name="obj">Object that the fill belongs to.</param>
    ''' <param name="fill">The GDITexturedFill used to paint the object</param>
    '''-----------------------------------------------------------------------------
    Private Sub addTextureFillToImageSummary(ByVal obj As GDIObject, ByVal fill As GDITexturedFill)
        'Add relevant properties to the  listview

        Dim lvitem As ListViewItem = lvImageDetails.Items.Add(New ListViewItem(obj.Name))

        With lvitem.SubItems
            .Add("Texture")

            Select Case fill.RuntimeSource
                Case EnumLinkType.AbsolutePath
                    .Add("Absolute Path")

                Case EnumLinkType.Embedded
                    .Add("Embedded")

                Case EnumLinkType.RelativePath
                    .Add("Relative Path")

            End Select

            .Add(fill.Path)
            .Add(fill.ImageSource)

        End With
    
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds an image to the list of summarized images
    ''' </summary>
    ''' <param name="gdiImage">The GDILinkedImage object to append to the list</param>
    '''-----------------------------------------------------------------------------
    Private Sub addImageSummary(ByVal gdiImage As GDILinkedImage)

        'Add relevant properties to the  listview

        Dim lvitem As ListViewItem = lvImageDetails.Items.Add(New ListViewItem(gdiImage.Name))

        With lvitem.SubItems
            .Add("Image")

            Select Case gdiImage.RuntimeSource
                Case EnumLinkType.AbsolutePath
                    .Add("Absolute Path")

                Case EnumLinkType.Embedded
                    .Add("Embedded")

                Case EnumLinkType.RelativePath
                    .Add("Relative Path")
            End Select

            .Add(gdiImage.Path)
            .Add(gdiImage.ImageSource)
        End With


    End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Calculates the number of times various fills have been used.  Also adds an 
    ''' image summary for textured fills
    ''' </summary>
    ''' <param name="obj">Object filled by the fill</param>
    ''' <param name="fill">The GDIFill used to fill the object</param>
    '''-----------------------------------------------------------------------------
    Private Sub addToFillTotal(ByVal obj As GDIObject, ByVal fill As GDIFill)

        'Interrogate the type of fill and increment appropriate counters

        If TypeOf fill Is GDISolidFill Then
            _SolidBrushCount += 1

        ElseIf TypeOf fill Is GDIGradientFill Then
            _GradientBrushCount += 1

        ElseIf TypeOf fill Is GDITexturedFill Then
            _TextureBrushCount += 1
            addTextureFillToImageSummary(obj, DirectCast(fill, GDITexturedFill))

        ElseIf TypeOf fill Is GDIHatchFill Then
            _HatchBrushCount += 1
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Initiates the process of gathering statistics from a GDIDocument.  This method 
    ''' adds up the number of times objects have been used on a particular document and 
    ''' summarized the images used in a document.
    ''' </summary>
    '''-----------------------------------------------------------------------------
    Private Sub GatherStatistics()
        Trace.WriteLineIf(App.TraceOn, "DocProperties.loadStatistics")

        For Each pg As GDIPage In MDIMain.ActiveDocument

            For Each obj As GDIObject In pg.GDIObjects

                'Request statistics for each object
                gatherStatisticsOnObject(obj)

            Next
        Next


        'Fill in the statistic fields
        populateStatisticFields()

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gathers sundry statistics on GDIObjects
    ''' </summary>
    ''' <param name="obj">The object to gather statistics on</param>
    ''' -----------------------------------------------------------------------------
    Private Sub gatherStatisticsOnObject(ByVal obj As GDIObject)
        Dim fill As GDIFill
        Dim stk As GDIStroke


        If TypeOf obj Is GDILinkedImage Then
            _ImageCount += 1
            Dim gdiimg As GDILinkedImage = DirectCast(obj, GDILinkedImage)
            addImageSummary(gdiimg)
            Return
        End If

        If TypeOf obj Is GDIRect Then
            _RectangleCount += 1
            Dim rect As GDIRect = DirectCast(obj, GDIRect)

            If rect.DrawFill Then
                fill = rect.Fill
            Else
                fill = Nothing
            End If

            If rect.DrawStroke Then
                stk = rect.Stroke
            Else
                stk = Nothing
            End If

        ElseIf TypeOf obj Is GDIEllipse Then
            _EllipseCount += 1
            Dim ellipse As GDIEllipse = DirectCast(obj, GDIEllipse)

            If ellipse.DrawFill Then
                fill = ellipse.Fill
            Else
                fill = Nothing
            End If

            If ellipse.DrawStroke Then
                stk = ellipse.Stroke
            Else
                stk = Nothing
            End If


        ElseIf TypeOf obj Is GDIField Then
            _FieldCount += 1
            Dim field As GDIField = DirectCast(obj, GDIField)

            If field.DrawFill Then
                fill = field.Fill
            Else
                fill = Nothing
            End If
            stk = Nothing


        ElseIf TypeOf obj Is GDIText Then
            _TextCount += 1
            Dim txt As GDIText = DirectCast(obj, GDIText)

            If txt.DrawFill Then
                fill = txt.Fill
            Else
                fill = Nothing
            End If
            stk = Nothing



        ElseIf TypeOf obj Is GDILine Then
            _LineCount += 1
            Dim line As GDILine = DirectCast(obj, GDILine)

            fill = Nothing

            If line.DrawStroke Then
                stk = line.Stroke
            Else
                stk = Nothing
            End If


        ElseIf TypeOf obj Is GDIClosedPath Then
            _PathCount += 1
            Dim pth As GDIClosedPath = DirectCast(obj, GDIClosedPath)

            If pth.DrawFill Then
                fill = pth.Fill
            Else
                fill = Nothing
            End If

            If pth.DrawStroke Then
                stk = pth.Stroke
            Else
                stk = Nothing
            End If


        ElseIf TypeOf obj Is GDIOpenPath Then
            _PathCount += 1
            Dim pth As GDIOpenPath = DirectCast(obj, GDIOpenPath)

            fill = Nothing

            If pth.DrawStroke Then
                stk = pth.Stroke
            Else
                stk = Nothing
            End If

        End If

        If Not fill Is Nothing Then
            addToFillTotal(obj, fill)
        End If

        If Not stk Is Nothing Then
            _PenCount += 1
        End If


    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fills out statistic field on the the form
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub populateStatisticFields()
        lblEllipses.Text = _EllipseCount.ToString
        lblFields.Text = _FieldCount.ToString
        lblGradientBrushes.Text = _GradientBrushCount.ToString
        lblHatchBrushes.Text = _HatchBrushCount.ToString
        lblImages.Text = _ImageCount.ToString
        lblLines.Text = _LineCount.ToString
        lblPaths.Text = _PathCount.ToString
        lblRectangles.Text = _RectangleCount.ToString
        lblSolidBrushes.Text = _SolidBrushCount.ToString
        lblTextureBrushes.Text = _TextureBrushCount.ToString
        lblTotalBrushes.Text = (_TextureBrushCount + _SolidBrushCount + _GradientBrushCount + _HatchBrushCount).ToString
        lblTotalPens.Text = _PenCount.ToString
        lblText.Text = _TextCount.ToString
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' applies the current export settings on the form to the export settings of the 
    ''' document.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub applyExportSettings()
        With MDIMain.ActiveDocument.ExportSettings


            If txtClassName.Text.Length > 0 Then
                .ClassName = txtClassName.Text
            End If

            If optCSharp.Checked Then
                .Language = EnumCodeTypes.eCSharp
            ElseIf optVB.Checked Then
                .Language = EnumCodeTypes.eVB
            End If

            .RootnameSpace = txtRootNamespace.Text
            .MemberPrefix = txtMemberVariablePrefix.Text
            .FieldPrefix = txtFieldPropertyPrefix.Text

            .OverrideConsolidateFill = chkFills.Checked
            .OverrideConsolidateStroke = chkStrokes.Checked
            .OverrideConsolidateFont = chkFonts.Checked
            .OverrideConsolidateStringFormats = chkStringFormats.Checked

        End With
    End Sub


    Private Sub applyPageSize()
        Select Case MDIMain.ActiveDocument.ExportSettings.DocumentType
            Case EnumDocumentTypes.ePrintDocument
                MDIMain.ActiveDocument.SetResoByPageSettings(_PageSettings)
            Case Else
                With MDIMain.ActiveDocument
                    .RectPageSize = New Rectangle(New Point(0, 0), Me.GDIGraphicsClassSize)
                    .PrintableArea = New Rectangle(New Point(0, 0), Me.GDIGraphicsClassSize)
                    .BackColor = Me.SelectedBackColor
                End With
        End Select
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Applies changes based on the dgDocProperties value to the current GDIDocument.
    ''' </summary>
    '''-----------------------------------------------------------------------------
    Private Sub applyChanges()
        Trace.WriteLineIf(App.TraceOn, "DocProperties.applyChanges")


        'Applies the export settings on the form to the document.
        applyExportSettings()

        applyPageSize()

        With MDIMain.ActiveDocument
            .SmoothingMode = DirectCast(System.Enum.GetValues(GetType(Drawing2D.SmoothingMode)).GetValue(cboSmoothingMode.SelectedIndex), Drawing2D.SmoothingMode)
            .TextRenderingHint = DirectCast(System.Enum.GetValues(GetType(Drawing.text.TextRenderingHint)).GetValue(cboTextRenderingHint.SelectedIndex), Drawing.text.TextRenderingHint)
        End With

    End Sub


#End Region


#Region "Private properties"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the size for the GraphicsClass style of GDIDocument
    ''' </summary>
    ''' <value>A size structure to make the current size of the GDIDocument</value>
    ''' -----------------------------------------------------------------------------
    Private ReadOnly Property GDIGraphicsClassSize() As Size
        Get
            Try
                Return New Size(CInt(txtPixelsX.Text), CInt(txtPixelsY.Text))
            Catch ex As Exception
                Return New Size(200, 200)
            End Try
        End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the final selected back color
    ''' </summary>
    ''' <value>A System.Drawing.Color to make the back color</value
    ''' -----------------------------------------------------------------------------
    Private ReadOnly Property SelectedBackColor() As Color
        Get
            Return picBackColor.BackColor
        End Get
    End Property

#End Region


#Region "Event Handlers and Overrides"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to a click on the back color dialog box.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub picBackColor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picBackColor.Click

        Dim dgColor As New ColorDialog

        Dim iresp As DialogResult = dgColor.ShowDialog()

        If iresp = DialogResult.OK Then
            picBackColor.BackColor = dgColor.Color
        End If

        dgColor.Dispose()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Applies changed on an OK click
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        applyChanges()

        Me.DialogResult = DialogResult.OK
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Initiates page setup for PrintDocument type of exports.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    '''-----------------------------------------------------------------------------
    Private Sub btnPageSetup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPageSetup.Click
        Trace.WriteLineIf(App.TraceOn, "DocProperties.invokePageSetup")

        Dim dgpagesettings As New PageSetupDialog
        dgpagesettings.PageSettings = _PageSettings

        Dim iresp As DialogResult = dgpagesettings.ShowDialog

        If iresp = DialogResult.OK Then
            _PageSettings = dgpagesettings.PageSettings
            lblPageSize.Text = _PageSettings.Bounds.ToString
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' This bit checks if code changes on the document need to be applied prior to 
    ''' displaying code.   
    ''' </summary>
    ''' <returns>True if the process should continue, false otherwise</returns>
    ''' -----------------------------------------------------------------------------
    Private Function preCheckViewCode() As Boolean


        If (optCSharp.Checked AndAlso _
        Not MDIMain.ActiveDocument.ExportSettings.Language = EnumCodeTypes.eCSharp) OrElse _
        (optVB.Checked AndAlso _
        Not MDIMain.ActiveDocument.ExportSettings.Language = EnumCodeTypes.eVB) Then

            Dim iresp As MsgBoxResult = MsgBox("Apply changes before viewing code?", MsgBoxStyle.YesNoCancel)

            If iresp = MsgBoxResult.Yes Then
                applyChanges()
                Return True
            ElseIf iresp = MsgBoxResult.Cancel Then
                Return False
            End If
        Else
            Return True
        End If

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Initiates a code view for the document in the current export language
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' <remarks>Checks if the settings on the form are current prior to showing code.</remarks>
    '''-----------------------------------------------------------------------------
    Private Sub btnViewCode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewCode.Click
        Trace.WriteLineIf(App.TraceOn, "DocProperties.invokeViewCode")

        Dim bContinue As Boolean = preCheckViewCode()

        Dim sCode As String = MDIMain.ActiveDocument.GenerateCode

        Dim dgQuickCode As New dgQuickCode(sCode)
        dgQuickCode.ShowDialog()


    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invokes help on document properties
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHelp.Click
        App.HelpManager.InvokeHelpContents("DocumentProperties.htm")
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Cancels the dialog.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Applies pending changes
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard event args</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnApply_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnApply.Click
        applyChanges()
    End Sub
#End Region



End Class
