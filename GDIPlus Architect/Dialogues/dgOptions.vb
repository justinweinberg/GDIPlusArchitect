Imports GDIObjects
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : dgOptions
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a wrapper for setting application wide options and settings
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class dgOptions
    Inherits System.Windows.Forms.Form

#Region "Local Properties"

    ''' <summary>indicates if the dialog may have changed options.  
    ''' Rather than tracking each option for changes, this is set to true 
    ''' whenever apply or ok occurs.</summary>
    Private _Dirty As Boolean = False
#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the Options dialog.  Populates the dialog with 
    ''' the current application wide options.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Attempts to load options from isolated storage.  Sets to defaults 
        'if this fails.
        loadoptions()

        'Sets the options tooltips
        App.ToolTipManager.PopulatePopupTip(Tip, "Options", Me)
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
    Private WithEvents Label5 As System.Windows.Forms.Label
    Private WithEvents Label8 As System.Windows.Forms.Label
    Private WithEvents Label11 As System.Windows.Forms.Label
    Private WithEvents btnOk As System.Windows.Forms.Button
    Private WithEvents btnCancel As System.Windows.Forms.Button
    Private WithEvents nudGridSize As System.Windows.Forms.NumericUpDown
    Private WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents nudMinorGridSize As System.Windows.Forms.NumericUpDown
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents Label12 As System.Windows.Forms.Label
    Private WithEvents txtSampleText As System.Windows.Forms.TextBox
    Private WithEvents chkSampleText As System.Windows.Forms.CheckBox
    Private WithEvents chkHighlightObjects As System.Windows.Forms.CheckBox
    Private WithEvents picNonPrintAreaColor As System.Windows.Forms.PictureBox
    Private WithEvents picOutlineColor As System.Windows.Forms.PictureBox
    Private WithEvents btnReset As System.Windows.Forms.Button
    Private WithEvents nudUndoSteps As System.Windows.Forms.NumericUpDown
    Private WithEvents btnApply As System.Windows.Forms.Button
    Private WithEvents chkMajorGrid As System.Windows.Forms.CheckBox
    Private WithEvents chkMinorGrid As System.Windows.Forms.CheckBox
    Private WithEvents grpQuickCode As System.Windows.Forms.GroupBox
    Private WithEvents optVB As System.Windows.Forms.RadioButton
    Private WithEvents optCSharp As System.Windows.Forms.RadioButton
    Private WithEvents Label13 As System.Windows.Forms.Label
    Private WithEvents Label14 As System.Windows.Forms.Label
    Private WithEvents nudSmallNudge As System.Windows.Forms.NumericUpDown
    Private WithEvents nudLargeNudge As System.Windows.Forms.NumericUpDown
    Private WithEvents Label15 As System.Windows.Forms.Label
    Private WithEvents Label17 As System.Windows.Forms.Label
    Private WithEvents TabOptions As System.Windows.Forms.TabControl
    Private WithEvents tbGeneral As System.Windows.Forms.TabPage
    Private WithEvents tbColors As System.Windows.Forms.TabPage
    Private WithEvents tbIncrements As System.Windows.Forms.TabPage
    Private WithEvents nudGridElasticity As System.Windows.Forms.NumericUpDown
    Private WithEvents nudGuideElasticity As System.Windows.Forms.NumericUpDown
    Private WithEvents chkSnapGuides As System.Windows.Forms.CheckBox
    Private WithEvents chkSnapMargins As System.Windows.Forms.CheckBox
    Private WithEvents grpSnap As System.Windows.Forms.GroupBox
    Private WithEvents chkSnapMinorGrid As System.Windows.Forms.CheckBox
    Private WithEvents chkSnapMajorGrid As System.Windows.Forms.CheckBox
    Private WithEvents chkMargins As System.Windows.Forms.CheckBox
    Private WithEvents chkGuides As System.Windows.Forms.CheckBox
    Private WithEvents Label19 As System.Windows.Forms.Label
    Private WithEvents txtMemberVariablePrefix As System.Windows.Forms.TextBox
    Private WithEvents Label20 As System.Windows.Forms.Label
    Private WithEvents tbImageCode As System.Windows.Forms.TabPage
    Private WithEvents tbCodeGen As System.Windows.Forms.TabPage
    Private WithEvents Label24 As System.Windows.Forms.Label
    Private WithEvents Label23 As System.Windows.Forms.Label
    Private WithEvents Label22 As System.Windows.Forms.Label
    Private WithEvents Label21 As System.Windows.Forms.Label
    Private WithEvents txtImageAbsPath As System.Windows.Forms.TextBox
    Private WithEvents cboImageRuntime As System.Windows.Forms.ComboBox
    Private WithEvents Label25 As System.Windows.Forms.Label
    Private WithEvents txtTextureAbsPath As System.Windows.Forms.TextBox
    Private WithEvents cboTextureRuntime As System.Windows.Forms.ComboBox
    Private WithEvents txtImageRelPath As System.Windows.Forms.TextBox
    Private WithEvents Label26 As System.Windows.Forms.Label
    Private WithEvents btnAbsImage As System.Windows.Forms.Button
    Private WithEvents btnAbsTexture As System.Windows.Forms.Button
    Private WithEvents txtTextureRelPath As System.Windows.Forms.TextBox
    Private WithEvents txtFieldPropertyPrefix As System.Windows.Forms.TextBox
    Private WithEvents Label28 As System.Windows.Forms.Label
    Private WithEvents Label27 As System.Windows.Forms.Label
    Private WithEvents cboMemberScope As System.Windows.Forms.ComboBox
    Private WithEvents cboFieldScope As System.Windows.Forms.ComboBox
    Private WithEvents tbEnvironment As System.Windows.Forms.TabPage
    Private WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Private WithEvents optEnhancedMenus As System.Windows.Forms.RadioButton
    Private WithEvents optPlainMenus As System.Windows.Forms.RadioButton
    Private WithEvents chkMenuIcons As System.Windows.Forms.CheckBox
    Private WithEvents grpColorPicker As System.Windows.Forms.GroupBox
    Private WithEvents optCustom As System.Windows.Forms.RadioButton
    Private WithEvents optWindows As System.Windows.Forms.RadioButton
    Private WithEvents btnHelp As System.Windows.Forms.Button
    Private WithEvents grpHandles As System.Windows.Forms.GroupBox
    Private WithEvents picCurveHandleColor As System.Windows.Forms.PictureBox
    Private WithEvents Label10 As System.Windows.Forms.Label
    Private WithEvents picMouseOverColor As System.Windows.Forms.PictureBox
    Private WithEvents Label9 As System.Windows.Forms.Label
    Private WithEvents picSelhandleColor As System.Windows.Forms.PictureBox
    Private WithEvents Label6 As System.Windows.Forms.Label
    Private WithEvents grpAlign As System.Windows.Forms.GroupBox
    Private WithEvents Label16 As System.Windows.Forms.Label
    Private WithEvents picGuideColor As System.Windows.Forms.PictureBox
    Private WithEvents picMarginColor As System.Windows.Forms.PictureBox
    Private WithEvents Label4 As System.Windows.Forms.Label
    Private WithEvents picMinorGridColor As System.Windows.Forms.PictureBox
    Private WithEvents picGridColor As System.Windows.Forms.PictureBox
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents Label7 As System.Windows.Forms.Label
    Private WithEvents cboSmoothingMode As System.Windows.Forms.ComboBox
    Private WithEvents cboTextRenderingHint As System.Windows.Forms.ComboBox
    Private WithEvents Label30 As System.Windows.Forms.Label
    Private WithEvents Label29 As System.Windows.Forms.Label
    Private WithEvents chkDrawTextFieldBorders As System.Windows.Forms.CheckBox
    Private WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Private WithEvents tbConsolidation As System.Windows.Forms.TabPage
    Private WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Private WithEvents chkConsolidateFills As System.Windows.Forms.CheckBox
    Private WithEvents chkConsolidateStrokes As System.Windows.Forms.CheckBox
    Private WithEvents chkConsolidateFonts As System.Windows.Forms.CheckBox
    Private WithEvents chkConsolidateFormats As System.Windows.Forms.CheckBox
    Private WithEvents grpConsolidation As System.Windows.Forms.GroupBox
    Private WithEvents chkTrace As System.Windows.Forms.CheckBox
    Private WithEvents chkToolTips As System.Windows.Forms.CheckBox
    Private WithEvents Tip As System.Windows.Forms.ToolTip
    Private WithEvents chkDisableCodePanel As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.TabOptions = New System.Windows.Forms.TabControl
        Me.tbGeneral = New System.Windows.Forms.TabPage
        Me.txtSampleText = New System.Windows.Forms.TextBox
        Me.nudUndoSteps = New System.Windows.Forms.NumericUpDown
        Me.Label8 = New System.Windows.Forms.Label
        Me.chkSampleText = New System.Windows.Forms.CheckBox
        Me.chkDrawTextFieldBorders = New System.Windows.Forms.CheckBox
        Me.chkHighlightObjects = New System.Windows.Forms.CheckBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.tbCodeGen = New System.Windows.Forms.TabPage
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.grpQuickCode = New System.Windows.Forms.GroupBox
        Me.optCSharp = New System.Windows.Forms.RadioButton
        Me.optVB = New System.Windows.Forms.RadioButton
        Me.Label19 = New System.Windows.Forms.Label
        Me.txtFieldPropertyPrefix = New System.Windows.Forms.TextBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.txtMemberVariablePrefix = New System.Windows.Forms.TextBox
        Me.cboMemberScope = New System.Windows.Forms.ComboBox
        Me.Label28 = New System.Windows.Forms.Label
        Me.cboFieldScope = New System.Windows.Forms.ComboBox
        Me.Label27 = New System.Windows.Forms.Label
        Me.cboTextRenderingHint = New System.Windows.Forms.ComboBox
        Me.Label29 = New System.Windows.Forms.Label
        Me.cboSmoothingMode = New System.Windows.Forms.ComboBox
        Me.Label30 = New System.Windows.Forms.Label
        Me.tbConsolidation = New System.Windows.Forms.TabPage
        Me.grpConsolidation = New System.Windows.Forms.GroupBox
        Me.chkConsolidateFormats = New System.Windows.Forms.CheckBox
        Me.chkConsolidateFonts = New System.Windows.Forms.CheckBox
        Me.chkConsolidateStrokes = New System.Windows.Forms.CheckBox
        Me.chkConsolidateFills = New System.Windows.Forms.CheckBox
        Me.tbImageCode = New System.Windows.Forms.TabPage
        Me.txtTextureRelPath = New System.Windows.Forms.TextBox
        Me.Label26 = New System.Windows.Forms.Label
        Me.txtImageRelPath = New System.Windows.Forms.TextBox
        Me.Label25 = New System.Windows.Forms.Label
        Me.btnAbsTexture = New System.Windows.Forms.Button
        Me.txtTextureAbsPath = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.btnAbsImage = New System.Windows.Forms.Button
        Me.txtImageAbsPath = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.cboTextureRuntime = New System.Windows.Forms.ComboBox
        Me.cboImageRuntime = New System.Windows.Forms.ComboBox
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.tbColors = New System.Windows.Forms.TabPage
        Me.grpAlign = New System.Windows.Forms.GroupBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.picGuideColor = New System.Windows.Forms.PictureBox
        Me.picMarginColor = New System.Windows.Forms.PictureBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.picMinorGridColor = New System.Windows.Forms.PictureBox
        Me.picGridColor = New System.Windows.Forms.PictureBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.grpHandles = New System.Windows.Forms.GroupBox
        Me.picCurveHandleColor = New System.Windows.Forms.PictureBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.picMouseOverColor = New System.Windows.Forms.PictureBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.picSelhandleColor = New System.Windows.Forms.PictureBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.picNonPrintAreaColor = New System.Windows.Forms.PictureBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.picOutlineColor = New System.Windows.Forms.PictureBox
        Me.tbIncrements = New System.Windows.Forms.TabPage
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.chkMajorGrid = New System.Windows.Forms.CheckBox
        Me.chkMinorGrid = New System.Windows.Forms.CheckBox
        Me.chkMargins = New System.Windows.Forms.CheckBox
        Me.chkGuides = New System.Windows.Forms.CheckBox
        Me.grpSnap = New System.Windows.Forms.GroupBox
        Me.chkSnapMinorGrid = New System.Windows.Forms.CheckBox
        Me.chkSnapMajorGrid = New System.Windows.Forms.CheckBox
        Me.chkSnapGuides = New System.Windows.Forms.CheckBox
        Me.chkSnapMargins = New System.Windows.Forms.CheckBox
        Me.nudGuideElasticity = New System.Windows.Forms.NumericUpDown
        Me.Label17 = New System.Windows.Forms.Label
        Me.nudGridElasticity = New System.Windows.Forms.NumericUpDown
        Me.Label15 = New System.Windows.Forms.Label
        Me.nudLargeNudge = New System.Windows.Forms.NumericUpDown
        Me.Label14 = New System.Windows.Forms.Label
        Me.nudSmallNudge = New System.Windows.Forms.NumericUpDown
        Me.Label13 = New System.Windows.Forms.Label
        Me.nudMinorGridSize = New System.Windows.Forms.NumericUpDown
        Me.Label3 = New System.Windows.Forms.Label
        Me.nudGridSize = New System.Windows.Forms.NumericUpDown
        Me.Label1 = New System.Windows.Forms.Label
        Me.tbEnvironment = New System.Windows.Forms.TabPage
        Me.chkDisableCodePanel = New System.Windows.Forms.CheckBox
        Me.chkToolTips = New System.Windows.Forms.CheckBox
        Me.chkTrace = New System.Windows.Forms.CheckBox
        Me.grpColorPicker = New System.Windows.Forms.GroupBox
        Me.optCustom = New System.Windows.Forms.RadioButton
        Me.optWindows = New System.Windows.Forms.RadioButton
        Me.chkMenuIcons = New System.Windows.Forms.CheckBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.optPlainMenus = New System.Windows.Forms.RadioButton
        Me.optEnhancedMenus = New System.Windows.Forms.RadioButton
        Me.btnOk = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnReset = New System.Windows.Forms.Button
        Me.btnApply = New System.Windows.Forms.Button
        Me.btnHelp = New System.Windows.Forms.Button
        Me.Tip = New System.Windows.Forms.ToolTip(Me.components)
        Me.TabOptions.SuspendLayout()
        Me.tbGeneral.SuspendLayout()
        CType(Me.nudUndoSteps, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbCodeGen.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.grpQuickCode.SuspendLayout()
        Me.tbConsolidation.SuspendLayout()
        Me.grpConsolidation.SuspendLayout()
        Me.tbImageCode.SuspendLayout()
        Me.tbColors.SuspendLayout()
        Me.grpAlign.SuspendLayout()
        Me.grpHandles.SuspendLayout()
        Me.tbIncrements.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.grpSnap.SuspendLayout()
        CType(Me.nudGuideElasticity, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudGridElasticity, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudLargeNudge, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudSmallNudge, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudMinorGridSize, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.nudGridSize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbEnvironment.SuspendLayout()
        Me.grpColorPicker.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabOptions
        '
        Me.TabOptions.Controls.Add(Me.tbGeneral)
        Me.TabOptions.Controls.Add(Me.tbCodeGen)
        Me.TabOptions.Controls.Add(Me.tbConsolidation)
        Me.TabOptions.Controls.Add(Me.tbImageCode)
        Me.TabOptions.Controls.Add(Me.tbColors)
        Me.TabOptions.Controls.Add(Me.tbIncrements)
        Me.TabOptions.Controls.Add(Me.tbEnvironment)
        Me.TabOptions.Location = New System.Drawing.Point(0, 0)
        Me.TabOptions.Name = "TabOptions"
        Me.TabOptions.SelectedIndex = 0
        Me.TabOptions.Size = New System.Drawing.Size(552, 368)
        Me.TabOptions.TabIndex = 0
        '
        'tbGeneral
        '
        Me.tbGeneral.Controls.Add(Me.txtSampleText)
        Me.tbGeneral.Controls.Add(Me.nudUndoSteps)
        Me.tbGeneral.Controls.Add(Me.Label8)
        Me.tbGeneral.Controls.Add(Me.chkSampleText)
        Me.tbGeneral.Controls.Add(Me.chkDrawTextFieldBorders)
        Me.tbGeneral.Controls.Add(Me.chkHighlightObjects)
        Me.tbGeneral.Controls.Add(Me.Label5)
        Me.tbGeneral.Location = New System.Drawing.Point(4, 22)
        Me.tbGeneral.Name = "tbGeneral"
        Me.tbGeneral.Size = New System.Drawing.Size(544, 342)
        Me.tbGeneral.TabIndex = 1
        Me.tbGeneral.Text = "General"
        '
        'txtSampleText
        '
        Me.txtSampleText.Location = New System.Drawing.Point(24, 176)
        Me.txtSampleText.Multiline = True
        Me.txtSampleText.Name = "txtSampleText"
        Me.txtSampleText.Size = New System.Drawing.Size(248, 64)
        Me.txtSampleText.TabIndex = 8
        Me.txtSampleText.Text = ""
        '
        'nudUndoSteps
        '
        Me.nudUndoSteps.Location = New System.Drawing.Point(104, 24)
        Me.nudUndoSteps.Name = "nudUndoSteps"
        Me.nudUndoSteps.Size = New System.Drawing.Size(48, 20)
        Me.nudUndoSteps.TabIndex = 11
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(24, 160)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(176, 16)
        Me.Label8.TabIndex = 9
        Me.Label8.Text = "Default sample text for new fields"
        '
        'chkSampleText
        '
        Me.chkSampleText.Location = New System.Drawing.Point(24, 120)
        Me.chkSampleText.Name = "chkSampleText"
        Me.chkSampleText.Size = New System.Drawing.Size(152, 24)
        Me.chkSampleText.TabIndex = 7
        Me.chkSampleText.Text = "Draw sample text in fields"
        '
        'chkDrawTextFieldBorders
        '
        Me.chkDrawTextFieldBorders.Location = New System.Drawing.Point(24, 88)
        Me.chkDrawTextFieldBorders.Name = "chkDrawTextFieldBorders"
        Me.chkDrawTextFieldBorders.Size = New System.Drawing.Size(296, 24)
        Me.chkDrawTextFieldBorders.TabIndex = 6
        Me.chkDrawTextFieldBorders.Text = "Draw helper borders around fields and text"
        '
        'chkHighlightObjects
        '
        Me.chkHighlightObjects.Location = New System.Drawing.Point(24, 56)
        Me.chkHighlightObjects.Name = "chkHighlightObjects"
        Me.chkHighlightObjects.Size = New System.Drawing.Size(120, 24)
        Me.chkHighlightObjects.TabIndex = 2
        Me.chkHighlightObjects.Text = "Highlight objects"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(24, 24)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 16)
        Me.Label5.TabIndex = 1
        Me.Label5.Text = "Undo steps:"
        '
        'tbCodeGen
        '
        Me.tbCodeGen.Controls.Add(Me.GroupBox4)
        Me.tbCodeGen.Location = New System.Drawing.Point(4, 22)
        Me.tbCodeGen.Name = "tbCodeGen"
        Me.tbCodeGen.Size = New System.Drawing.Size(544, 342)
        Me.tbCodeGen.TabIndex = 4
        Me.tbCodeGen.Text = "Code Defaults"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.grpQuickCode)
        Me.GroupBox4.Controls.Add(Me.Label19)
        Me.GroupBox4.Controls.Add(Me.txtFieldPropertyPrefix)
        Me.GroupBox4.Controls.Add(Me.Label20)
        Me.GroupBox4.Controls.Add(Me.txtMemberVariablePrefix)
        Me.GroupBox4.Controls.Add(Me.cboMemberScope)
        Me.GroupBox4.Controls.Add(Me.Label28)
        Me.GroupBox4.Controls.Add(Me.cboFieldScope)
        Me.GroupBox4.Controls.Add(Me.Label27)
        Me.GroupBox4.Controls.Add(Me.cboTextRenderingHint)
        Me.GroupBox4.Controls.Add(Me.Label29)
        Me.GroupBox4.Controls.Add(Me.cboSmoothingMode)
        Me.GroupBox4.Controls.Add(Me.Label30)
        Me.GroupBox4.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(528, 320)
        Me.GroupBox4.TabIndex = 47
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "New Document Defaults"
        '
        'grpQuickCode
        '
        Me.grpQuickCode.Controls.Add(Me.optCSharp)
        Me.grpQuickCode.Controls.Add(Me.optVB)
        Me.grpQuickCode.Location = New System.Drawing.Point(16, 16)
        Me.grpQuickCode.Name = "grpQuickCode"
        Me.grpQuickCode.Size = New System.Drawing.Size(184, 56)
        Me.grpQuickCode.TabIndex = 14
        Me.grpQuickCode.TabStop = False
        Me.grpQuickCode.Text = "Code language"
        '
        'optCSharp
        '
        Me.optCSharp.Location = New System.Drawing.Point(136, 24)
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
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(16, 112)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(128, 23)
        Me.Label19.TabIndex = 16
        Me.Label19.Text = "Member variable prefix:"
        '
        'txtFieldPropertyPrefix
        '
        Me.txtFieldPropertyPrefix.Location = New System.Drawing.Point(328, 112)
        Me.txtFieldPropertyPrefix.Name = "txtFieldPropertyPrefix"
        Me.txtFieldPropertyPrefix.Size = New System.Drawing.Size(40, 20)
        Me.txtFieldPropertyPrefix.TabIndex = 18
        Me.txtFieldPropertyPrefix.Text = ""
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(208, 112)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(120, 23)
        Me.Label20.TabIndex = 17
        Me.Label20.Text = "Field property prefix:"
        '
        'txtMemberVariablePrefix
        '
        Me.txtMemberVariablePrefix.Location = New System.Drawing.Point(152, 112)
        Me.txtMemberVariablePrefix.Name = "txtMemberVariablePrefix"
        Me.txtMemberVariablePrefix.Size = New System.Drawing.Size(40, 20)
        Me.txtMemberVariablePrefix.TabIndex = 15
        Me.txtMemberVariablePrefix.Text = ""
        '
        'cboMemberScope
        '
        Me.cboMemberScope.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMemberScope.Location = New System.Drawing.Point(16, 184)
        Me.cboMemberScope.Name = "cboMemberScope"
        Me.cboMemberScope.TabIndex = 40
        '
        'Label28
        '
        Me.Label28.Location = New System.Drawing.Point(16, 168)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(112, 16)
        Me.Label28.TabIndex = 44
        Me.Label28.Text = "Key Property Scope"
        '
        'cboFieldScope
        '
        Me.cboFieldScope.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFieldScope.Location = New System.Drawing.Point(152, 184)
        Me.cboFieldScope.Name = "cboFieldScope"
        Me.cboFieldScope.TabIndex = 39
        '
        'Label27
        '
        Me.Label27.Location = New System.Drawing.Point(152, 168)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(72, 16)
        Me.Label27.TabIndex = 43
        Me.Label27.Text = "Field Scope"
        '
        'cboTextRenderingHint
        '
        Me.cboTextRenderingHint.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTextRenderingHint.ItemHeight = 13
        Me.cboTextRenderingHint.Location = New System.Drawing.Point(144, 256)
        Me.cboTextRenderingHint.Name = "cboTextRenderingHint"
        Me.cboTextRenderingHint.TabIndex = 42
        '
        'Label29
        '
        Me.Label29.Location = New System.Drawing.Point(144, 240)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(112, 23)
        Me.Label29.TabIndex = 44
        Me.Label29.Text = "Text Rendering Hint"
        '
        'cboSmoothingMode
        '
        Me.cboSmoothingMode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSmoothingMode.ItemHeight = 13
        Me.cboSmoothingMode.Location = New System.Drawing.Point(16, 256)
        Me.cboSmoothingMode.Name = "cboSmoothingMode"
        Me.cboSmoothingMode.TabIndex = 43
        '
        'Label30
        '
        Me.Label30.Location = New System.Drawing.Point(16, 240)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(96, 23)
        Me.Label30.TabIndex = 45
        Me.Label30.Text = "Smoothing Mode"
        '
        'tbConsolidation
        '
        Me.tbConsolidation.Controls.Add(Me.grpConsolidation)
        Me.tbConsolidation.Location = New System.Drawing.Point(4, 22)
        Me.tbConsolidation.Name = "tbConsolidation"
        Me.tbConsolidation.Size = New System.Drawing.Size(544, 342)
        Me.tbConsolidation.TabIndex = 9
        Me.tbConsolidation.Text = "Object Defaults"
        '
        'grpConsolidation
        '
        Me.grpConsolidation.Controls.Add(Me.chkConsolidateFormats)
        Me.grpConsolidation.Controls.Add(Me.chkConsolidateFonts)
        Me.grpConsolidation.Controls.Add(Me.chkConsolidateStrokes)
        Me.grpConsolidation.Controls.Add(Me.chkConsolidateFills)
        Me.grpConsolidation.Location = New System.Drawing.Point(16, 24)
        Me.grpConsolidation.Name = "grpConsolidation"
        Me.grpConsolidation.Size = New System.Drawing.Size(482, 59)
        Me.grpConsolidation.TabIndex = 2
        Me.grpConsolidation.TabStop = False
        Me.grpConsolidation.Text = "Default new object Consolidation values"
        '
        'chkConsolidateFormats
        '
        Me.chkConsolidateFormats.Location = New System.Drawing.Point(355, 22)
        Me.chkConsolidateFormats.Name = "chkConsolidateFormats"
        Me.chkConsolidateFormats.TabIndex = 5
        Me.chkConsolidateFormats.Text = "String Formats"
        '
        'chkConsolidateFonts
        '
        Me.chkConsolidateFonts.Location = New System.Drawing.Point(246, 22)
        Me.chkConsolidateFonts.Name = "chkConsolidateFonts"
        Me.chkConsolidateFonts.Size = New System.Drawing.Size(64, 24)
        Me.chkConsolidateFonts.TabIndex = 4
        Me.chkConsolidateFonts.Text = "Fonts"
        '
        'chkConsolidateStrokes
        '
        Me.chkConsolidateStrokes.Location = New System.Drawing.Point(12, 22)
        Me.chkConsolidateStrokes.Name = "chkConsolidateStrokes"
        Me.chkConsolidateStrokes.TabIndex = 2
        Me.chkConsolidateStrokes.Text = "Strokes"
        '
        'chkConsolidateFills
        '
        Me.chkConsolidateFills.Location = New System.Drawing.Point(136, 22)
        Me.chkConsolidateFills.Name = "chkConsolidateFills"
        Me.chkConsolidateFills.TabIndex = 1
        Me.chkConsolidateFills.Text = "Fills"
        '
        'tbImageCode
        '
        Me.tbImageCode.Controls.Add(Me.txtTextureRelPath)
        Me.tbImageCode.Controls.Add(Me.Label26)
        Me.tbImageCode.Controls.Add(Me.txtImageRelPath)
        Me.tbImageCode.Controls.Add(Me.Label25)
        Me.tbImageCode.Controls.Add(Me.btnAbsTexture)
        Me.tbImageCode.Controls.Add(Me.txtTextureAbsPath)
        Me.tbImageCode.Controls.Add(Me.Label24)
        Me.tbImageCode.Controls.Add(Me.btnAbsImage)
        Me.tbImageCode.Controls.Add(Me.txtImageAbsPath)
        Me.tbImageCode.Controls.Add(Me.Label23)
        Me.tbImageCode.Controls.Add(Me.cboTextureRuntime)
        Me.tbImageCode.Controls.Add(Me.cboImageRuntime)
        Me.tbImageCode.Controls.Add(Me.Label22)
        Me.tbImageCode.Controls.Add(Me.Label21)
        Me.tbImageCode.Location = New System.Drawing.Point(4, 22)
        Me.tbImageCode.Name = "tbImageCode"
        Me.tbImageCode.Size = New System.Drawing.Size(544, 342)
        Me.tbImageCode.TabIndex = 7
        Me.tbImageCode.Text = "Image Code Generation"
        '
        'txtTextureRelPath
        '
        Me.txtTextureRelPath.Location = New System.Drawing.Point(16, 272)
        Me.txtTextureRelPath.Name = "txtTextureRelPath"
        Me.txtTextureRelPath.Size = New System.Drawing.Size(304, 20)
        Me.txtTextureRelPath.TabIndex = 41
        Me.txtTextureRelPath.Text = ""
        '
        'Label26
        '
        Me.Label26.Location = New System.Drawing.Point(16, 256)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(144, 23)
        Me.Label26.TabIndex = 42
        Me.Label26.Text = "Default relative image path"
        '
        'txtImageRelPath
        '
        Me.txtImageRelPath.Location = New System.Drawing.Point(16, 120)
        Me.txtImageRelPath.Name = "txtImageRelPath"
        Me.txtImageRelPath.Size = New System.Drawing.Size(304, 20)
        Me.txtImageRelPath.TabIndex = 39
        Me.txtImageRelPath.Text = ""
        '
        'Label25
        '
        Me.Label25.Location = New System.Drawing.Point(16, 104)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(192, 23)
        Me.Label25.TabIndex = 40
        Me.Label25.Text = "Default relative image path"
        '
        'btnAbsTexture
        '
        Me.btnAbsTexture.Location = New System.Drawing.Point(328, 224)
        Me.btnAbsTexture.Name = "btnAbsTexture"
        Me.btnAbsTexture.Size = New System.Drawing.Size(32, 23)
        Me.btnAbsTexture.TabIndex = 37
        Me.btnAbsTexture.Text = "..."
        '
        'txtTextureAbsPath
        '
        Me.txtTextureAbsPath.Location = New System.Drawing.Point(16, 224)
        Me.txtTextureAbsPath.Name = "txtTextureAbsPath"
        Me.txtTextureAbsPath.Size = New System.Drawing.Size(304, 20)
        Me.txtTextureAbsPath.TabIndex = 36
        Me.txtTextureAbsPath.Text = ""
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(16, 208)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(152, 23)
        Me.Label24.TabIndex = 38
        Me.Label24.Text = "Default absolute texture path"
        '
        'btnAbsImage
        '
        Me.btnAbsImage.Location = New System.Drawing.Point(328, 72)
        Me.btnAbsImage.Name = "btnAbsImage"
        Me.btnAbsImage.Size = New System.Drawing.Size(32, 23)
        Me.btnAbsImage.TabIndex = 34
        Me.btnAbsImage.Text = "..."
        '
        'txtImageAbsPath
        '
        Me.txtImageAbsPath.Location = New System.Drawing.Point(16, 72)
        Me.txtImageAbsPath.Name = "txtImageAbsPath"
        Me.txtImageAbsPath.Size = New System.Drawing.Size(304, 20)
        Me.txtImageAbsPath.TabIndex = 33
        Me.txtImageAbsPath.Text = ""
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(16, 56)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(240, 23)
        Me.Label23.TabIndex = 35
        Me.Label23.Text = "Default absolute image path"
        '
        'cboTextureRuntime
        '
        Me.cboTextureRuntime.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTextureRuntime.Location = New System.Drawing.Point(192, 176)
        Me.cboTextureRuntime.Name = "cboTextureRuntime"
        Me.cboTextureRuntime.TabIndex = 32
        '
        'cboImageRuntime
        '
        Me.cboImageRuntime.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboImageRuntime.Location = New System.Drawing.Point(192, 24)
        Me.cboImageRuntime.Name = "cboImageRuntime"
        Me.cboImageRuntime.TabIndex = 31
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(16, 176)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(184, 23)
        Me.Label22.TabIndex = 30
        Me.Label22.Text = "Default texture runtime type:"
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(16, 24)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(184, 23)
        Me.Label21.TabIndex = 29
        Me.Label21.Text = "Default Image runtime type:"
        '
        'tbColors
        '
        Me.tbColors.Controls.Add(Me.grpAlign)
        Me.tbColors.Controls.Add(Me.grpHandles)
        Me.tbColors.Controls.Add(Me.Label12)
        Me.tbColors.Controls.Add(Me.picNonPrintAreaColor)
        Me.tbColors.Controls.Add(Me.Label11)
        Me.tbColors.Controls.Add(Me.picOutlineColor)
        Me.tbColors.Location = New System.Drawing.Point(4, 22)
        Me.tbColors.Name = "tbColors"
        Me.tbColors.Size = New System.Drawing.Size(544, 342)
        Me.tbColors.TabIndex = 2
        Me.tbColors.Text = "Colors"
        '
        'grpAlign
        '
        Me.grpAlign.Controls.Add(Me.Label16)
        Me.grpAlign.Controls.Add(Me.picGuideColor)
        Me.grpAlign.Controls.Add(Me.picMarginColor)
        Me.grpAlign.Controls.Add(Me.Label4)
        Me.grpAlign.Controls.Add(Me.picMinorGridColor)
        Me.grpAlign.Controls.Add(Me.picGridColor)
        Me.grpAlign.Controls.Add(Me.Label2)
        Me.grpAlign.Controls.Add(Me.Label7)
        Me.grpAlign.Location = New System.Drawing.Point(16, 16)
        Me.grpAlign.Name = "grpAlign"
        Me.grpAlign.Size = New System.Drawing.Size(432, 88)
        Me.grpAlign.TabIndex = 27
        Me.grpAlign.TabStop = False
        Me.grpAlign.Text = "Guides and Grids"
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(336, 16)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(56, 16)
        Me.Label16.TabIndex = 33
        Me.Label16.Text = "Guides"
        '
        'picGuideColor
        '
        Me.picGuideColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picGuideColor.Location = New System.Drawing.Point(336, 32)
        Me.picGuideColor.Name = "picGuideColor"
        Me.picGuideColor.Size = New System.Drawing.Size(48, 40)
        Me.picGuideColor.TabIndex = 32
        Me.picGuideColor.TabStop = False
        '
        'picMarginColor
        '
        Me.picMarginColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picMarginColor.Location = New System.Drawing.Point(232, 32)
        Me.picMarginColor.Name = "picMarginColor"
        Me.picMarginColor.Size = New System.Drawing.Size(48, 40)
        Me.picMarginColor.TabIndex = 30
        Me.picMarginColor.TabStop = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(128, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 16)
        Me.Label4.TabIndex = 29
        Me.Label4.Text = "Minor grid"
        '
        'picMinorGridColor
        '
        Me.picMinorGridColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picMinorGridColor.Location = New System.Drawing.Point(128, 32)
        Me.picMinorGridColor.Name = "picMinorGridColor"
        Me.picMinorGridColor.Size = New System.Drawing.Size(48, 40)
        Me.picMinorGridColor.TabIndex = 28
        Me.picMinorGridColor.TabStop = False
        '
        'picGridColor
        '
        Me.picGridColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picGridColor.Location = New System.Drawing.Point(24, 32)
        Me.picGridColor.Name = "picGridColor"
        Me.picGridColor.Size = New System.Drawing.Size(48, 40)
        Me.picGridColor.TabIndex = 26
        Me.picGridColor.TabStop = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 23)
        Me.Label2.TabIndex = 27
        Me.Label2.Text = "Major grid"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(232, 16)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(40, 16)
        Me.Label7.TabIndex = 31
        Me.Label7.Text = "Margins"
        '
        'grpHandles
        '
        Me.grpHandles.Controls.Add(Me.picCurveHandleColor)
        Me.grpHandles.Controls.Add(Me.Label10)
        Me.grpHandles.Controls.Add(Me.picMouseOverColor)
        Me.grpHandles.Controls.Add(Me.Label9)
        Me.grpHandles.Controls.Add(Me.picSelhandleColor)
        Me.grpHandles.Controls.Add(Me.Label6)
        Me.grpHandles.Location = New System.Drawing.Point(16, 112)
        Me.grpHandles.Name = "grpHandles"
        Me.grpHandles.Size = New System.Drawing.Size(432, 88)
        Me.grpHandles.TabIndex = 26
        Me.grpHandles.TabStop = False
        Me.grpHandles.Text = "Drag Handle Colors"
        '
        'picCurveHandleColor
        '
        Me.picCurveHandleColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picCurveHandleColor.Location = New System.Drawing.Point(232, 32)
        Me.picCurveHandleColor.Name = "picCurveHandleColor"
        Me.picCurveHandleColor.Size = New System.Drawing.Size(48, 40)
        Me.picCurveHandleColor.TabIndex = 28
        Me.picCurveHandleColor.TabStop = False
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(120, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(64, 16)
        Me.Label10.TabIndex = 27
        Me.Label10.Text = "Mouse over"
        '
        'picMouseOverColor
        '
        Me.picMouseOverColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picMouseOverColor.Location = New System.Drawing.Point(128, 32)
        Me.picMouseOverColor.Name = "picMouseOverColor"
        Me.picMouseOverColor.Size = New System.Drawing.Size(48, 40)
        Me.picMouseOverColor.TabIndex = 26
        Me.picMouseOverColor.TabStop = False
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(16, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(96, 16)
        Me.Label9.TabIndex = 25
        Me.Label9.Text = "Selected"
        '
        'picSelhandleColor
        '
        Me.picSelhandleColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picSelhandleColor.Location = New System.Drawing.Point(16, 32)
        Me.picSelhandleColor.Name = "picSelhandleColor"
        Me.picSelhandleColor.Size = New System.Drawing.Size(48, 40)
        Me.picSelhandleColor.TabIndex = 24
        Me.picSelhandleColor.TabStop = False
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(224, 16)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 16)
        Me.Label6.TabIndex = 29
        Me.Label6.Text = "Curve points"
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(24, 216)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(80, 16)
        Me.Label12.TabIndex = 21
        Me.Label12.Text = "Non print area"
        '
        'picNonPrintAreaColor
        '
        Me.picNonPrintAreaColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picNonPrintAreaColor.Location = New System.Drawing.Point(32, 232)
        Me.picNonPrintAreaColor.Name = "picNonPrintAreaColor"
        Me.picNonPrintAreaColor.Size = New System.Drawing.Size(48, 40)
        Me.picNonPrintAreaColor.TabIndex = 20
        Me.picNonPrintAreaColor.TabStop = False
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(112, 216)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(80, 16)
        Me.Label11.TabIndex = 19
        Me.Label11.Text = "Action outline"
        '
        'picOutlineColor
        '
        Me.picOutlineColor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picOutlineColor.Location = New System.Drawing.Point(120, 232)
        Me.picOutlineColor.Name = "picOutlineColor"
        Me.picOutlineColor.Size = New System.Drawing.Size(48, 40)
        Me.picOutlineColor.TabIndex = 18
        Me.picOutlineColor.TabStop = False
        '
        'tbIncrements
        '
        Me.tbIncrements.Controls.Add(Me.GroupBox2)
        Me.tbIncrements.Controls.Add(Me.grpSnap)
        Me.tbIncrements.Controls.Add(Me.nudGuideElasticity)
        Me.tbIncrements.Controls.Add(Me.Label17)
        Me.tbIncrements.Controls.Add(Me.nudGridElasticity)
        Me.tbIncrements.Controls.Add(Me.Label15)
        Me.tbIncrements.Controls.Add(Me.nudLargeNudge)
        Me.tbIncrements.Controls.Add(Me.Label14)
        Me.tbIncrements.Controls.Add(Me.nudSmallNudge)
        Me.tbIncrements.Controls.Add(Me.Label13)
        Me.tbIncrements.Controls.Add(Me.nudMinorGridSize)
        Me.tbIncrements.Controls.Add(Me.Label3)
        Me.tbIncrements.Controls.Add(Me.nudGridSize)
        Me.tbIncrements.Controls.Add(Me.Label1)
        Me.tbIncrements.Location = New System.Drawing.Point(4, 22)
        Me.tbIncrements.Name = "tbIncrements"
        Me.tbIncrements.Size = New System.Drawing.Size(544, 342)
        Me.tbIncrements.TabIndex = 5
        Me.tbIncrements.Text = "Snaps and Grids"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.chkMajorGrid)
        Me.GroupBox2.Controls.Add(Me.chkMinorGrid)
        Me.GroupBox2.Controls.Add(Me.chkMargins)
        Me.GroupBox2.Controls.Add(Me.chkGuides)
        Me.GroupBox2.Location = New System.Drawing.Point(24, 224)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(384, 56)
        Me.GroupBox2.TabIndex = 33
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Show"
        '
        'chkMajorGrid
        '
        Me.chkMajorGrid.Location = New System.Drawing.Point(192, 16)
        Me.chkMajorGrid.Name = "chkMajorGrid"
        Me.chkMajorGrid.Size = New System.Drawing.Size(80, 24)
        Me.chkMajorGrid.TabIndex = 24
        Me.chkMajorGrid.Text = "Major grid"
        '
        'chkMinorGrid
        '
        Me.chkMinorGrid.Location = New System.Drawing.Point(288, 16)
        Me.chkMinorGrid.Name = "chkMinorGrid"
        Me.chkMinorGrid.Size = New System.Drawing.Size(80, 24)
        Me.chkMinorGrid.TabIndex = 25
        Me.chkMinorGrid.Text = "Minor grid"
        '
        'chkMargins
        '
        Me.chkMargins.Location = New System.Drawing.Point(112, 16)
        Me.chkMargins.Name = "chkMargins"
        Me.chkMargins.Size = New System.Drawing.Size(72, 24)
        Me.chkMargins.TabIndex = 26
        Me.chkMargins.Text = "Margins"
        '
        'chkGuides
        '
        Me.chkGuides.Location = New System.Drawing.Point(16, 16)
        Me.chkGuides.Name = "chkGuides"
        Me.chkGuides.Size = New System.Drawing.Size(64, 24)
        Me.chkGuides.TabIndex = 27
        Me.chkGuides.Text = "Guides"
        '
        'grpSnap
        '
        Me.grpSnap.Controls.Add(Me.chkSnapMinorGrid)
        Me.grpSnap.Controls.Add(Me.chkSnapMajorGrid)
        Me.grpSnap.Controls.Add(Me.chkSnapGuides)
        Me.grpSnap.Controls.Add(Me.chkSnapMargins)
        Me.grpSnap.Location = New System.Drawing.Point(24, 160)
        Me.grpSnap.Name = "grpSnap"
        Me.grpSnap.Size = New System.Drawing.Size(384, 56)
        Me.grpSnap.TabIndex = 32
        Me.grpSnap.TabStop = False
        Me.grpSnap.Text = "Snap to"
        '
        'chkSnapMinorGrid
        '
        Me.chkSnapMinorGrid.Location = New System.Drawing.Point(288, 16)
        Me.chkSnapMinorGrid.Name = "chkSnapMinorGrid"
        Me.chkSnapMinorGrid.Size = New System.Drawing.Size(80, 24)
        Me.chkSnapMinorGrid.TabIndex = 35
        Me.chkSnapMinorGrid.Text = "Minor grid"
        '
        'chkSnapMajorGrid
        '
        Me.chkSnapMajorGrid.Location = New System.Drawing.Point(192, 16)
        Me.chkSnapMajorGrid.Name = "chkSnapMajorGrid"
        Me.chkSnapMajorGrid.Size = New System.Drawing.Size(88, 24)
        Me.chkSnapMajorGrid.TabIndex = 34
        Me.chkSnapMajorGrid.Text = "Major grid"
        '
        'chkSnapGuides
        '
        Me.chkSnapGuides.Location = New System.Drawing.Point(16, 16)
        Me.chkSnapGuides.Name = "chkSnapGuides"
        Me.chkSnapGuides.Size = New System.Drawing.Size(64, 24)
        Me.chkSnapGuides.TabIndex = 33
        Me.chkSnapGuides.Text = "Guides"
        '
        'chkSnapMargins
        '
        Me.chkSnapMargins.Location = New System.Drawing.Point(104, 16)
        Me.chkSnapMargins.Name = "chkSnapMargins"
        Me.chkSnapMargins.Size = New System.Drawing.Size(64, 24)
        Me.chkSnapMargins.TabIndex = 32
        Me.chkSnapMargins.Text = "Margins"
        '
        'nudGuideElasticity
        '
        Me.nudGuideElasticity.Location = New System.Drawing.Point(152, 120)
        Me.nudGuideElasticity.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.nudGuideElasticity.Name = "nudGuideElasticity"
        Me.nudGuideElasticity.Size = New System.Drawing.Size(64, 20)
        Me.nudGuideElasticity.TabIndex = 7
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(24, 120)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(128, 23)
        Me.Label17.TabIndex = 6
        Me.Label17.Text = "Snap to guide elasticity:"
        '
        'nudGridElasticity
        '
        Me.nudGridElasticity.Location = New System.Drawing.Point(152, 88)
        Me.nudGridElasticity.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.nudGridElasticity.Name = "nudGridElasticity"
        Me.nudGridElasticity.Size = New System.Drawing.Size(64, 20)
        Me.nudGridElasticity.TabIndex = 5
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(24, 88)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(120, 23)
        Me.Label15.TabIndex = 4
        Me.Label15.Text = "Snap to grid elasticity:"
        '
        'nudLargeNudge
        '
        Me.nudLargeNudge.Location = New System.Drawing.Point(152, 56)
        Me.nudLargeNudge.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.nudLargeNudge.Name = "nudLargeNudge"
        Me.nudLargeNudge.Size = New System.Drawing.Size(64, 20)
        Me.nudLargeNudge.TabIndex = 3
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(24, 56)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(128, 23)
        Me.Label14.TabIndex = 2
        Me.Label14.Text = "Large nudge increment:"
        '
        'nudSmallNudge
        '
        Me.nudSmallNudge.Location = New System.Drawing.Point(152, 24)
        Me.nudSmallNudge.Name = "nudSmallNudge"
        Me.nudSmallNudge.Size = New System.Drawing.Size(64, 20)
        Me.nudSmallNudge.TabIndex = 1
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(24, 24)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(128, 23)
        Me.Label13.TabIndex = 0
        Me.Label13.Text = "Small nudge increment:"
        '
        'nudMinorGridSize
        '
        Me.nudMinorGridSize.Location = New System.Drawing.Point(328, 96)
        Me.nudMinorGridSize.Name = "nudMinorGridSize"
        Me.nudMinorGridSize.Size = New System.Drawing.Size(72, 20)
        Me.nudMinorGridSize.TabIndex = 22
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(328, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.TabIndex = 23
        Me.Label3.Text = " Minor grid size"
        '
        'nudGridSize
        '
        Me.nudGridSize.Location = New System.Drawing.Point(328, 40)
        Me.nudGridSize.Name = "nudGridSize"
        Me.nudGridSize.Size = New System.Drawing.Size(72, 20)
        Me.nudGridSize.TabIndex = 20
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(328, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 16)
        Me.Label1.TabIndex = 21
        Me.Label1.Text = "Major grid size"
        '
        'tbEnvironment
        '
        Me.tbEnvironment.Controls.Add(Me.chkDisableCodePanel)
        Me.tbEnvironment.Controls.Add(Me.chkToolTips)
        Me.tbEnvironment.Controls.Add(Me.chkTrace)
        Me.tbEnvironment.Controls.Add(Me.grpColorPicker)
        Me.tbEnvironment.Controls.Add(Me.chkMenuIcons)
        Me.tbEnvironment.Controls.Add(Me.GroupBox3)
        Me.tbEnvironment.Location = New System.Drawing.Point(4, 22)
        Me.tbEnvironment.Name = "tbEnvironment"
        Me.tbEnvironment.Size = New System.Drawing.Size(544, 342)
        Me.tbEnvironment.TabIndex = 8
        Me.tbEnvironment.Text = "Environment"
        '
        'chkDisableCodePanel
        '
        Me.chkDisableCodePanel.Location = New System.Drawing.Point(16, 296)
        Me.chkDisableCodePanel.Name = "chkDisableCodePanel"
        Me.chkDisableCodePanel.Size = New System.Drawing.Size(144, 24)
        Me.chkDisableCodePanel.TabIndex = 12
        Me.chkDisableCodePanel.Text = "Disable Code Panel"
        '
        'chkToolTips
        '
        Me.chkToolTips.Location = New System.Drawing.Point(144, 80)
        Me.chkToolTips.Name = "chkToolTips"
        Me.chkToolTips.TabIndex = 11
        Me.chkToolTips.Text = "Show ToolTips"
        '
        'chkTrace
        '
        Me.chkTrace.Location = New System.Drawing.Point(16, 264)
        Me.chkTrace.Name = "chkTrace"
        Me.chkTrace.Size = New System.Drawing.Size(296, 24)
        Me.chkTrace.TabIndex = 10
        Me.chkTrace.Text = "Enable debug tracing"
        '
        'grpColorPicker
        '
        Me.grpColorPicker.Controls.Add(Me.optCustom)
        Me.grpColorPicker.Controls.Add(Me.optWindows)
        Me.grpColorPicker.Location = New System.Drawing.Point(16, 120)
        Me.grpColorPicker.Name = "grpColorPicker"
        Me.grpColorPicker.Size = New System.Drawing.Size(336, 64)
        Me.grpColorPicker.TabIndex = 6
        Me.grpColorPicker.TabStop = False
        Me.grpColorPicker.Text = "Color picker"
        '
        'optCustom
        '
        Me.optCustom.Location = New System.Drawing.Point(16, 24)
        Me.optCustom.Name = "optCustom"
        Me.optCustom.Size = New System.Drawing.Size(88, 24)
        Me.optCustom.TabIndex = 1
        Me.optCustom.Text = "Use custom"
        '
        'optWindows
        '
        Me.optWindows.Location = New System.Drawing.Point(224, 24)
        Me.optWindows.Name = "optWindows"
        Me.optWindows.Size = New System.Drawing.Size(96, 24)
        Me.optWindows.TabIndex = 0
        Me.optWindows.Text = "Use Windows"
        '
        'chkMenuIcons
        '
        Me.chkMenuIcons.Location = New System.Drawing.Point(24, 80)
        Me.chkMenuIcons.Name = "chkMenuIcons"
        Me.chkMenuIcons.TabIndex = 3
        Me.chkMenuIcons.Text = "Icons on menus"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.optPlainMenus)
        Me.GroupBox3.Controls.Add(Me.optEnhancedMenus)
        Me.GroupBox3.Location = New System.Drawing.Point(16, 16)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(208, 56)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Menu Style"
        '
        'optPlainMenus
        '
        Me.optPlainMenus.Location = New System.Drawing.Point(96, 16)
        Me.optPlainMenus.Name = "optPlainMenus"
        Me.optPlainMenus.Size = New System.Drawing.Size(64, 24)
        Me.optPlainMenus.TabIndex = 1
        Me.optPlainMenus.Text = "Plain"
        '
        'optEnhancedMenus
        '
        Me.optEnhancedMenus.Location = New System.Drawing.Point(8, 16)
        Me.optEnhancedMenus.Name = "optEnhancedMenus"
        Me.optEnhancedMenus.Size = New System.Drawing.Size(80, 24)
        Me.optEnhancedMenus.TabIndex = 0
        Me.optEnhancedMenus.Text = "Enhanced"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(392, 376)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.TabIndex = 1
        Me.btnOk.Text = "Ok"
        '
        'btnCancel
        '
        Me.btnCancel.CausesValidation = False
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(472, 376)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "Cancel"
        '
        'btnReset
        '
        Me.btnReset.Location = New System.Drawing.Point(8, 376)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.TabIndex = 4
        Me.btnReset.Text = "Reset"
        '
        'btnApply
        '
        Me.btnApply.Location = New System.Drawing.Point(312, 376)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.TabIndex = 5
        Me.btnApply.Text = "Apply"
        '
        'btnHelp
        '
        Me.btnHelp.Location = New System.Drawing.Point(224, 376)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.TabIndex = 6
        Me.btnHelp.Text = "Help"
        '
        'dgOptions
        '
        Me.AcceptButton = Me.btnOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(562, 415)
        Me.Controls.Add(Me.TabOptions)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnReset)
        Me.Controls.Add(Me.btnApply)
        Me.Controls.Add(Me.btnHelp)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "dgOptions"
        Me.ShowInTaskbar = False
        Me.Text = "Options"
        Me.TabOptions.ResumeLayout(False)
        Me.tbGeneral.ResumeLayout(False)
        CType(Me.nudUndoSteps, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbCodeGen.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.grpQuickCode.ResumeLayout(False)
        Me.tbConsolidation.ResumeLayout(False)
        Me.grpConsolidation.ResumeLayout(False)
        Me.tbImageCode.ResumeLayout(False)
        Me.tbColors.ResumeLayout(False)
        Me.grpAlign.ResumeLayout(False)
        Me.grpHandles.ResumeLayout(False)
        Me.tbIncrements.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.grpSnap.ResumeLayout(False)
        CType(Me.nudGuideElasticity, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudGridElasticity, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudLargeNudge, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudSmallNudge, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudMinorGridSize, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.nudGridSize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbEnvironment.ResumeLayout(False)
        Me.grpColorPicker.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Event Handlers"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invokes help on Options
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHelp.Click
        App.HelpManager.InvokeHelpContents("options.htm")
    End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles changes to tool tips settings. 
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' <remarks>Unlike other option settings, this one is updated immediately 
    ''' so that the Options dilaogue can reflect changes in tool tips.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Sub chkToolTips_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkToolTips.CheckedChanged
        If Me.Created Then
            If Not Tip Is Nothing Then
                Tip.Dispose()
            End If
            If chkToolTips.Checked Then
                Tip = New ToolTip
                App.ToolTipManager.PopulatePopupTip(Tip, "Options", Me)
            End If
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Causes options to reset to their defaults.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Trace.WriteLineIf(App.TraceOn, "Optionsdialog.Reset")

        Dim iResp As MsgBoxResult = MsgBox("Are you sure you wish to revert to default options?", MsgBoxStyle.OKCancel, "Confirm Revert Options")
        If iResp = MsgBoxResult.OK Then
            App.Options.Revert()
        End If

        loadoptions()
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Applies options on an OK click.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        applyOptions()
        Me.DialogResult = DialogResult.OK
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles the various color box press events.  This initiates color selection 
    ''' for application wide colors such as the color of grids and drag handles.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub popColor(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles picCurveHandleColor.Click, picGridColor.Click, picGuideColor.Click, _
    picMarginColor.Click, picMinorGridColor.Click, picMouseOverColor.Click, _
    picNonPrintAreaColor.Click, picOutlineColor.Click, picSelhandleColor.Click
        Dim dgColor As New ColorDialog
        Dim picbox As PictureBox = DirectCast(sender, PictureBox)
        dgColor.Color = picbox.BackColor
        Dim iresp As DialogResult = dgColor.ShowDialog()

        If iresp = DialogResult.OK Then
            picbox.BackColor = dgColor.Color
        End If

        picbox = Nothing

        dgColor.Dispose()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Applies options.
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApply.Click
        applyOptions()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Invokes a dialog for selecting a default texture path
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnAbsTexture_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAbsTexture.Click
        Dim x As New OpenFolderDialog
        Dim sPath As String = x.GetFolder()
        If sPath.Length > 0 Then
            txtTextureAbsPath.Text = sPath
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Prompts for a new absolute image path for textures
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' ----------------------------------------------------------------------------- 
    Private Sub btnAbsImage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAbsImage.Click
        Dim x As New OpenFolderDialog
        Dim sPath As String = x.GetFolder()
        If sPath.Length > 0 Then
            txtImageAbsPath.Text = sPath
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Cancels the dialog
    ''' </summary>
    ''' <param name="sender">Standard sender</param>
    ''' <param name="e">Standard eventargs</param>
    ''' ----------------------------------------------------------------------------- 
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = DialogResult.Cancel
    End Sub

#End Region


#Region "helper Methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retrieves the current application wide options 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub loadoptions()
        Trace.WriteLineIf(App.TraceOn, "Optionsdialog.Load")
        With cboTextRenderingHint
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(System.Drawing.Text.TextRenderingHint)))
            .EndUpdate()
        End With
        With cboSmoothingMode
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(Drawing2D.SmoothingMode)))
            .Items.Remove(System.Enum.GetName(GetType(Drawing2D.SmoothingMode), Drawing2D.SmoothingMode.Invalid))
            .EndUpdate()
        End With
        With cboTextureRuntime
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(EnumLinkType)))
            .EndUpdate()
        End With

        With cboImageRuntime
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(EnumLinkType)))
            .EndUpdate()
        End With

        With cboMemberScope
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(EnumScope)))
            .EndUpdate()
        End With

        With cboFieldScope
            .BeginUpdate()
            .Items.Clear()
            .Items.AddRange(System.Enum.GetNames(GetType(EnumScope)))
            .EndUpdate()
        End With


        With App.Options

            If .MenuStyle = Crownwood.Magic.Common.VisualStyle.IDE Then
                optEnhancedMenus.Checked = True
            Else
                optPlainMenus.Checked = True
            End If

            chkTrace.Checked = .Tracing

            chkConsolidateStrokes.Checked = .ConsolidateStrokes
            chkConsolidateFills.Checked = .ConsolidateFills
            chkConsolidateFormats.Checked = .ConsolidateStringFormats
            chkConsolidateFonts.Checked = .ConsolidateFonts

            chkMenuIcons.Checked = .ShowMenuIcons
            chkToolTips.Checked = .ShowToolTips
            cboTextRenderingHint.SelectedItem = .TextRenderingHint.ToString
            cboSmoothingMode.SelectedItem = .SmoothingMode.ToString

            cboTextureRuntime.SelectedItem = .TextureLinkType.ToString
            cboImageRuntime.SelectedItem = .ImageLinkType.ToString
            cboFieldScope.SelectedItem = .FieldScope.ToString
            cboMemberScope.SelectedItem = .MemberScope.ToString


            txtTextureAbsPath.Text = .TextureAbsolutePath
            txtTextureRelPath.Text = .TextureRelativePath

            txtImageAbsPath.Text = .ImageAbsolutePath
            txtImageRelPath.Text = .ImageRelativePath

            txtMemberVariablePrefix.Text = .MemberPrefix
            txtFieldPropertyPrefix.Text = .FieldPrefix

            picCurveHandleColor.BackColor = .ColorCurve
            nudGridElasticity.Value = .GridElasticity
            nudGuideElasticity.Value = .GuideElasticity
            nudSmallNudge.Value = .SmallNudge
            nudLargeNudge.Value = .LargeNudge
            picGridColor.BackColor = .ColorMajorGrid
            picMarginColor.BackColor = .ColorMargin
            picMinorGridColor.BackColor = .ColorMinorGrid
            picNonPrintAreaColor.BackColor = .ColorNonPrintArea
            picOutlineColor.BackColor = .ColorOutline
            picMouseOverColor.BackColor = .ColorMouseOver
            picSelhandleColor.BackColor = .ColorSelected
            picGuideColor.BackColor = .ColorGuide
            nudMinorGridSize.Value = CDec(.MinorGridSize)
            nudGridSize.Value = CDec(.MajorGridSize)

            chkMajorGrid.Checked = .ShowGrid
            chkMinorGrid.Checked = .ShowMinorGrid
            chkMargins.Checked = .ShowMargins
            chkGuides.Checked = .ShowGuides
            chkDrawTextFieldBorders.Checked = .ShowTextBorders
            chkDisableCodePanel.Checked = .DisableCodePanel

            chkSampleText.Checked = .ShowSampleText
            chkHighlightObjects.Checked = .ShowMouseOverHandles
            txtSampleText.Text = App.Options.SampleText


            nudUndoSteps.Value = .UndoSteps

            chkSnapGuides.Checked = .SnapToGuides
            chkSnapMajorGrid.Checked = .SnapToMajorGrid
            chkSnapMinorGrid.Checked = .SnapToMinorGrid
            chkSnapMargins.Checked = .SnapToMargins


            If .UseCustomColorPicker Then
                optCustom.Checked = True
            Else
                optWindows.Checked = True
            End If

            Select Case .CodeLanguage

                Case EnumCodeTypes.eCSharp
                    optCSharp.Checked = True

                Case EnumCodeTypes.eVB
                    optVB.Checked = True
            End Select



        End With

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Applies the options as currently selected to the 
    ''' application wide Options object.  Marks the dialog's _dirty field to true.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub applyOptions()
        Trace.WriteLineIf(App.TraceOn, "Optionsdialog.Apply")

        With App.Options
            .FieldScope = DirectCast(System.Enum.GetValues(GetType(EnumScope)).GetValue(cboFieldScope.SelectedIndex), EnumScope)
            .MemberScope = DirectCast(System.Enum.GetValues(GetType(EnumScope)).GetValue(cboMemberScope.SelectedIndex), EnumScope)

            .TextureLinkType = DirectCast(System.Enum.GetValues(GetType(EnumLinkType)).GetValue(cboTextureRuntime.SelectedIndex), EnumLinkType)
            .ImageLinkType = DirectCast(System.Enum.GetValues(GetType(EnumLinkType)).GetValue(cboImageRuntime.SelectedIndex), EnumLinkType)
            .SmoothingMode = DirectCast(System.Enum.GetValues(GetType(Drawing2D.SmoothingMode)).GetValue(cboSmoothingMode.SelectedIndex), Drawing2D.SmoothingMode)
            .TextRenderingHint = DirectCast(System.Enum.GetValues(GetType(Drawing.text.TextRenderingHint)).GetValue(cboTextRenderingHint.SelectedIndex), Drawing.text.TextRenderingHint)

            If optCustom.Checked Then
                .UseCustomColorPicker = True
            Else
                .UseCustomColorPicker = False
            End If

            If Not chkTrace.Checked = .Tracing AndAlso chkTrace.Checked = True Then
                Dim fInfo As New System.IO.FileInfo(App.Options.RuntimePath & "\GDIInterface.log")
                If fInfo.Exists Then
                    Dim iresp As MsgBoxResult = MsgBox("Delete existing trace?", MsgBoxStyle.YesNo)
                    If iresp = MsgBoxResult.Yes Then
                        fInfo.Delete()
                    End If
                End If

            End If
            .Tracing = chkTrace.Checked

            .ShowMenuIcons = chkMenuIcons.Checked
            .ShowToolTips = chkToolTips.Checked
            .TextureAbsolutePath = txtTextureAbsPath.Text
            .TextureRelativePath = txtTextureRelPath.Text

            .ImageAbsolutePath = txtImageAbsPath.Text
            .ImageRelativePath = txtImageRelPath.Text
            .MemberPrefix = txtMemberVariablePrefix.Text
            .FieldPrefix = txtFieldPropertyPrefix.Text


            .ConsolidateStrokes = chkConsolidateStrokes.Checked
            .ConsolidateFills = chkConsolidateFills.Checked
            .ConsolidateStringFormats = chkConsolidateFormats.Checked
            .ConsolidateFonts = chkConsolidateFonts.Checked


            .ColorMajorGrid = picGridColor.BackColor
            .ColorMargin = picMarginColor.BackColor
            .ColorMinorGrid = picMinorGridColor.BackColor
            .ColorNonPrintArea = picNonPrintAreaColor.BackColor
            .ColorOutline = picOutlineColor.BackColor
            .ColorMouseOver = picMouseOverColor.BackColor
            .MajorGridSize = CSng(nudGridSize.Value)
            .MinorGridSize = CSng(nudMinorGridSize.Value)
            .ColorSelected = picSelhandleColor.BackColor
            .ColorCurve = picCurveHandleColor.BackColor
            .ColorGuide = picGuideColor.BackColor
            .SampleText = txtSampleText.Text
            .ShowGrid = chkMajorGrid.Checked
            .ShowMinorGrid = chkMinorGrid.Checked
            .ShowMargins = chkMargins.Checked
            .ShowGuides = chkGuides.Checked
            .ShowMouseOverHandles = chkHighlightObjects.Checked
            .ShowTextBorders = chkDrawTextFieldBorders.Checked
            .DisableCodePanel = chkDisableCodePanel.Checked
            .ShowSampleText = chkSampleText.Checked
            .LargeNudge = CInt(nudLargeNudge.Value)
            .SmallNudge = CInt(nudSmallNudge.Value)
            .GridElasticity = CInt(nudGridElasticity.Value)
            .GuideElasticity = CInt(nudGuideElasticity.Value)

            If .UndoSteps <> nudUndoSteps.Value Then
                .UndoSteps = CInt(nudUndoSteps.Value)
                MsgBox("Undo step changes will not take place for current documents until you close them.")
            End If

            .SnapToGuides = chkSnapGuides.Checked
            .SnapToMajorGrid = chkSnapMajorGrid.Checked
            .SnapToMinorGrid = chkSnapMinorGrid.Checked
            .SnapToMargins = chkSnapMargins.Checked

            If optPlainMenus.Checked = True Then
                .MenuStyle = Crownwood.Magic.Common.VisualStyle.Plain

            Else
                .MenuStyle = Crownwood.Magic.Common.VisualStyle.IDE

            End If


            If optVB.Checked Then
                .CodeLanguage = EnumCodeTypes.eVB
            ElseIf optCSharp.Checked Then
                .CodeLanguage = EnumCodeTypes.eCSharp

            End If

        End With


        _Dirty = True
    End Sub



#End Region

#Region "Property Accessors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a value indicating if the OptionsManager should broadcast an options 
    ''' changed event
    ''' </summary>
    ''' <value>A boolean indicating if the options have been changed</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property OptionsChanged() As Boolean
        Get
            Return _Dirty
        End Get

    End Property

#End Region

End Class
