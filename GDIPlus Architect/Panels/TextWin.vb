Imports GDIObjects
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : TextWin
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides an interface for manipulating the properties of text on the design 
''' surface as well as setting the application wide font properties
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class TextWin
    Inherits System.Windows.Forms.UserControl

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new instance of the Text panel given a starting font to use.
    ''' </summary>
    ''' <param name="ft">The initial font to use</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal ft As Font)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Populates the options available in the drop down fonts dialog
        populatefontnames()
        'Populates the units available 
        populateFontUnits()

        createInitialFont(ft)
        'Add any initialization after the InitializeComponent() call

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populates the list of valid fonts into cboFontName.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub populatefontnames()
        Trace.WriteLineIf(App.TraceOn, "textwin.populatefonts")
        cboFontname.BeginUpdate()

        For Each ft As FontFamily In FontFamily.Families
            Dim nv As New NVPair(-1, ft.Name)
            cboFontname.Items.Add(nv)

        Next

        cboFontname.EndUpdate()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Populates cboUnits with a series of NVPairs (read this as name value pair) containing 
    ''' the name of the unit and a Drawing.GraphicsUnit which corresponds to the name of the 
    ''' unit.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub populateFontUnits()
        Trace.WriteLineIf(App.TraceOn, "textwin.populateunits")

        cboUnits.BeginUpdate()

        With cboUnits
            .BeginUpdate()
            .Items.Clear()
            .Items.Add(New NVPair(Drawing.GraphicsUnit.Point, "Point"))
            .Items.Add(New NVPair(Drawing.GraphicsUnit.Pixel, "Pixel"))
            .Items.Add(New NVPair(Drawing.GraphicsUnit.Inch, "Inch"))
            .Items.Add(New NVPair(Drawing.GraphicsUnit.Millimeter, "Millimeter"))
            .EndUpdate()
        End With


        cboUnits.EndUpdate()

    End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Given a font, assigns the text window to reflect the settings of the font
    ''' the font.
    ''' </summary>
    ''' <param name="font">The font to base settings off of</param>
    ''' -----------------------------------------------------------------------------
    Private Sub createInitialFont(ByVal font As Font)


        For Each nv As NVPair In cboFontname.Items

            If nv.Name = font.FontFamily.Name Then
                cboFontname.SelectedItem = nv
            End If
        Next

        For Each nv As NVPair In cboUnits.Items
            If nv.ID = font.Unit Then
                cboUnits.SelectedItem = nv
            End If
        Next


        nudSize.Text = font.Size.ToString
        chkBold.Checked = font.Bold
        chkItalic.Checked = font.Italic
        chkUnderline.Checked = font.Underline
        chkStrike.Checked = font.Strikeout

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
    Private WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents cboUnits As System.Windows.Forms.ComboBox
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents cboFontname As System.Windows.Forms.ComboBox
    Private WithEvents btnDialog As System.Windows.Forms.Button
    Private WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Private WithEvents chkStrike As System.Windows.Forms.CheckBox
    Private WithEvents chkUnderline As System.Windows.Forms.CheckBox
    Private WithEvents chkItalic As System.Windows.Forms.CheckBox
    Private WithEvents chkBold As System.Windows.Forms.CheckBox
    Private WithEvents nudSize As System.Windows.Forms.NumericUpDown
    Private WithEvents MrgNud1 As GDIPlusArchitect.MRGNud
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cboFontname = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.cboUnits = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnDialog = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkStrike = New System.Windows.Forms.CheckBox
        Me.chkUnderline = New System.Windows.Forms.CheckBox
        Me.chkItalic = New System.Windows.Forms.CheckBox
        Me.chkBold = New System.Windows.Forms.CheckBox
        Me.nudSize = New System.Windows.Forms.NumericUpDown
        Me.MrgNud1 = New GDIPlusArchitect.MRGNud
        Me.GroupBox1.SuspendLayout()
        CType(Me.nudSize, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cboFontname
        '
        Me.cboFontname.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFontname.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboFontname.Location = New System.Drawing.Point(8, 13)
        Me.cboFontname.Name = "cboFontname"
        Me.cboFontname.Size = New System.Drawing.Size(128, 22)
        Me.cboFontname.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(7, 1)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 23)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Font"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(6, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 23)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Size"
        '
        'cboUnits
        '
        Me.cboUnits.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboUnits.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboUnits.Location = New System.Drawing.Point(58, 48)
        Me.cboUnits.Name = "cboUnits"
        Me.cboUnits.Size = New System.Drawing.Size(80, 22)
        Me.cboUnits.TabIndex = 11
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(56, 36)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 23)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Units"
        '
        'btnDialog
        '
        Me.btnDialog.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDialog.Location = New System.Drawing.Point(8, 120)
        Me.btnDialog.Name = "btnDialog"
        Me.btnDialog.Size = New System.Drawing.Size(128, 23)
        Me.btnDialog.TabIndex = 15
        Me.btnDialog.Text = "Font Dialog..."
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkStrike)
        Me.GroupBox1.Controls.Add(Me.chkUnderline)
        Me.GroupBox1.Controls.Add(Me.chkItalic)
        Me.GroupBox1.Controls.Add(Me.chkBold)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 73)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(128, 40)
        Me.GroupBox1.TabIndex = 16
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Style"
        '
        'chkStrike
        '
        Me.chkStrike.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkStrike.Location = New System.Drawing.Point(94, 14)
        Me.chkStrike.Name = "chkStrike"
        Me.chkStrike.Size = New System.Drawing.Size(32, 24)
        Me.chkStrike.TabIndex = 18
        Me.chkStrike.Text = "S"
        '
        'chkUnderline
        '
        Me.chkUnderline.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUnderline.Location = New System.Drawing.Point(64, 14)
        Me.chkUnderline.Name = "chkUnderline"
        Me.chkUnderline.Size = New System.Drawing.Size(32, 24)
        Me.chkUnderline.TabIndex = 17
        Me.chkUnderline.Text = "U"
        '
        'chkItalic
        '
        Me.chkItalic.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkItalic.Location = New System.Drawing.Point(36, 14)
        Me.chkItalic.Name = "chkItalic"
        Me.chkItalic.Size = New System.Drawing.Size(24, 24)
        Me.chkItalic.TabIndex = 16
        Me.chkItalic.Text = "I"
        '
        'chkBold
        '
        Me.chkBold.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBold.Location = New System.Drawing.Point(5, 14)
        Me.chkBold.Name = "chkBold"
        Me.chkBold.Size = New System.Drawing.Size(27, 24)
        Me.chkBold.TabIndex = 15
        Me.chkBold.Text = "B"
        '
        'nudSize
        '
        Me.nudSize.Location = New System.Drawing.Point(9, 48)
        Me.nudSize.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.nudSize.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.nudSize.Name = "nudSize"
        Me.nudSize.Size = New System.Drawing.Size(40, 20)
        Me.nudSize.TabIndex = 17
        Me.nudSize.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'MrgNud1
        '
        Me.MrgNud1.AllowBlank = False
        Me.MrgNud1.DecimalPlaces = 0
        Me.MrgNud1.Location = New System.Drawing.Point(24, 120)
        Me.MrgNud1.MaxValue = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.MrgNud1.MinValue = New Decimal(New Integer() {0, 0, 0, 0})
        Me.MrgNud1.Name = "MrgNud1"
        Me.MrgNud1.Size = New System.Drawing.Size(112, 24)
        Me.MrgNud1.TabIndex = 18
        '
        'TextWin
        '
        Me.Controls.Add(Me.cboFontname)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboUnits)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnDialog)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.nudSize)
        Me.Controls.Add(Me.MrgNud1)
        Me.Controls.Add(Me.Label2)
        Me.Name = "TextWin"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.nudSize, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region




#Region "Event Handlers"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the bold check box.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>Notice that not all fonts support bold, so we have to try to create a 
    ''' font that does and handle cases where it does not.  
    ''' </remarks>
        ''' -----------------------------------------------------------------------------
    Private Sub chkBold_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBold.CheckedChanged
        If chkBold.Checked = True AndAlso App.Options.Font.Bold = False Then
            Try
                Dim tempfont As Font = New Font(App.Options.Font, App.Options.Font.Style Or FontStyle.Bold)
                App.Options.Font = tempfont
                tempfont.Dispose()
            Catch ex As Exception
                MsgBox("Font does not support this combination of styles")
                chkBold.Checked = False
            End Try
        ElseIf chkBold.Checked = False AndAlso App.Options.Font.Bold = True Then
            Try
                Dim tempfont As Font = New Font(App.Options.Font, App.Options.Font.Style And Not FontStyle.Bold)
                App.Options.Font = tempfont
                tempfont.Dispose()
            Catch ex As Exception
                MsgBox("Font does not support this combination of styles")
                chkUnderline.Checked = False
            End Try
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the italic check box.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>Notice that not all fonts support italics, so we have to try to create a 
    ''' font that does and handle cases where it does not.  
    ''' </remarks>
        ''' -----------------------------------------------------------------------------
    Private Sub chkItalic_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkItalic.CheckedChanged
        If chkItalic.Checked = True AndAlso App.Options.Font.Italic = False Then
            Try
                Dim tempfont As Font = New Font(App.Options.Font, App.Options.Font.Style Or FontStyle.Italic)
                App.Options.Font = tempfont
                tempfont.Dispose()
            Catch ex As Exception
                MsgBox("Font does not support this combination of styles")
                chkItalic.Checked = False
            End Try
        ElseIf chkItalic.Checked = False AndAlso App.Options.Font.Italic = True Then
            Try
                Dim tempfont As Font = New Font(App.Options.Font, App.Options.Font.Style And Not FontStyle.Italic)
                App.Options.Font = tempfont
                tempfont.Dispose()
            Catch ex As Exception
                MsgBox("Font does not support this combination of styles")
                chkUnderline.Checked = False
            End Try
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the underline check box.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>Notice that not all fonts support underlining, so we have to try to create a 
    ''' font that does and handle cases where it does not.  
    ''' </remarks>
        ''' -----------------------------------------------------------------------------
    Private Sub chkUnderline_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUnderline.CheckedChanged
        If chkUnderline.Checked = True AndAlso App.Options.Font.Underline = False Then
            Try
                Dim tempfont As Font = New Font(App.Options.Font, App.Options.Font.Style Or FontStyle.Underline)
                App.Options.Font = tempfont
                tempfont.Dispose()
            Catch ex As Exception
                MsgBox("Font does not support this combination of styles")
                chkUnderline.Checked = False
            End Try
        ElseIf chkUnderline.Checked = False AndAlso App.Options.Font.Underline = True Then
            Try
                Dim tempfont As Font = New Font(App.Options.Font, App.Options.Font.Style And Not FontStyle.Underline)
                App.Options.Font = tempfont
                tempfont.Dispose()
            Catch ex As Exception
                MsgBox("Font does not support this combination of styles")
                chkUnderline.Checked = False
            End Try
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the stroke through check box.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>Notice that not all fonts strike through, so we have to try to create a 
    ''' font that does and handle cases where it does not.  
    ''' </remarks>
        ''' -----------------------------------------------------------------------------
    Private Sub chkStrike_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkStrike.CheckedChanged
        If chkStrike.Checked = True AndAlso App.Options.Font.Strikeout = False Then
            Try
                Dim tempfont As Font = New Font(App.Options.Font, App.Options.Font.Style Or FontStyle.Strikeout)
                App.Options.Font = tempfont
                tempfont.Dispose()
            Catch ex As Exception
                MsgBox("Font does not support this combination of styles")
                chkStrike.Checked = False
            End Try
        ElseIf chkStrike.Checked = False AndAlso App.Options.Font.Strikeout = True Then
            Try
                Dim tempfont As Font = New Font(App.Options.Font, App.Options.Font.Style And Not FontStyle.Strikeout)
                App.Options.Font = tempfont
                tempfont.Dispose()
            Catch ex As Exception
                MsgBox("Font does not support this combination of styles")
                chkUnderline.Checked = False
            End Try
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the selected font drop down box.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>Believe it or not, not all fonts support "Regular" type face, for example
    ''' Arahoni.  The initial GDI+ Architect release shipped with a bug that assumed they did which 
    ''' which caused a top level exception to occur.
    ''' 
    ''' In the code below, the font is attempted to be created.  If it fails, it degrades to the previously 
    ''' selected font and informs the user they will need to use the font dialog to get at the font.
    ''' </remarks>
        ''' -----------------------------------------------------------------------------
    Private Sub cboFontname_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboFontname.SelectedValueChanged

        If Not cboFontname.SelectedItem Is Nothing Then
            Dim nv As NVPair = DirectCast(cboFontname.SelectedItem, NVPair)
            If String.op_Equality(nv.Name, App.Options.Font.Name) = False Then
                Try
                    Dim tempfont As Font = New Font(cboFontname.Text, App.Options.Font.Size, App.Options.Font.Style, App.Options.Font.Unit)
                    App.Options.Font = tempfont
                    tempfont.Dispose()
                Catch ex As Exception
                    Try
                        Dim tempfont2 As Font = New Font(cboFontname.Text, App.Options.Font.Unit)
                        App.Options.Font = tempfont2
                        tempfont2.Dispose()
                    Catch ex2 As Exception
                        MsgBox("Font does not support style regular.  Please set this font using the font dialog")
                        cboFontname.Text = App.Options.Font.Name
                    End Try
                End Try
            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the units value
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub cboUnits_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboUnits.SelectedValueChanged
        If Not cboUnits.SelectedItem Is Nothing Then
            Dim selUnit As NVPair = DirectCast(cboUnits.SelectedItem, NVPair)
            If Not selUnit.ID = App.Options.Font.Unit Then
                Try
                    Dim tempfont As Font = New Font(App.Options.Font.Name, App.Options.Font.Size, App.Options.Font.Style, CType(selUnit.ID, GraphicsUnit))
                    App.Options.Font = tempfont
                    tempfont.Dispose()
                Catch ex As Exception
                    MsgBox("Unable to set the units on this font")
                End Try
            End If
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Launches the .NET  font selection dialog box.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub btnDialog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDialog.Click
        Dim dgFont As New FontDialog
        dgFont.Font = App.Options.Font
        Dim iresp As DialogResult = dgFont.ShowDialog()

        If iresp = DialogResult.OK Then
            App.Options.Font = dgFont.Font
        End If
    End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the value of the font size. The NUD control used on this window 
    ''' is a custom nud included in GDI+ Architect.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub nudSize_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles nudSize.TextChanged, nudSize.ValueChanged

        If nudSize.Text.Length > 0 AndAlso IsNumeric(nudSize.Text) Then

            Dim fSize As Single = CSng(nudSize.Value)

            If Not fSize = App.Options.Font.Size Then
                Try
                    Dim tempfont As Font = New Font(App.Options.Font.Name, fSize, App.Options.Font.Style, App.Options.Font.Unit)
                    App.Options.Font = tempfont
                    tempfont.Dispose()
                Catch ex As Exception
                    MsgBox("Unable to set the units on this font")
                End Try
            End If

        End If

    End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Verifies that something valid is in the NUD control if it looses focus with an 
    ''' empty string.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub nudSize_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles nudSize.LostFocus
        If nudSize.Text.Length = 0 Then
            nudSize.Text = "10"
        End If
    End Sub


#End Region
End Class
