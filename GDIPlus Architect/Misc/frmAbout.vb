Imports GDIObjects

Friend Class frmAbout
    Inherits System.Windows.Forms.Form

#Region "Constructors"
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        '  GetLicenseInfo()
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
    Private WithEvents Label2 As System.Windows.Forms.Label
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents lblDotNetMAgic As System.Windows.Forms.LinkLabel
    Private WithEvents pbAbout As System.Windows.Forms.PictureBox
    Private WithEvents Label1 As System.Windows.Forms.Label
    Private WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Private WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Private WithEvents btnOk As System.Windows.Forms.Button
    Private WithEvents btnVersion As System.Windows.Forms.Button
    Private WithEvents lnkSource As System.Windows.Forms.LinkLabel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAbout))
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblDotNetMAgic = New System.Windows.Forms.LinkLabel
        Me.pbAbout = New System.Windows.Forms.PictureBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.btnOk = New System.Windows.Forms.Button
        Me.btnVersion = New System.Windows.Forms.Button
        Me.lnkSource = New System.Windows.Forms.LinkLabel
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 248)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(368, 23)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Copyright 2003 MRGSoft.  All rights Reserved"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 288)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(200, 16)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Portions of User Interface curteousy of"
        '
        'lblDotNetMAgic
        '
        Me.lblDotNetMAgic.Location = New System.Drawing.Point(208, 288)
        Me.lblDotNetMAgic.Name = "lblDotNetMAgic"
        Me.lblDotNetMAgic.Size = New System.Drawing.Size(128, 23)
        Me.lblDotNetMAgic.TabIndex = 2
        Me.lblDotNetMAgic.TabStop = True
        Me.lblDotNetMAgic.Text = "www.dotnetmagic.com"
        '
        'pbAbout
        '
        Me.pbAbout.Image = CType(resources.GetObject("pbAbout.Image"), System.Drawing.Image)
        Me.pbAbout.Location = New System.Drawing.Point(0, 0)
        Me.pbAbout.Name = "pbAbout"
        Me.pbAbout.Size = New System.Drawing.Size(496, 80)
        Me.pbAbout.TabIndex = 4
        Me.pbAbout.TabStop = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 312)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(472, 64)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Special thanks to: Duncan Mackenzie for his fabulous article on how to create a G" & _
        "DI+ design surface, Crownwood Consulting the architects of Dotnet Magic, Christi" & _
        "an Graus  for his series of articles on GDI+, Michael Gold for his scrollable pi" & _
        "cturebox example, Ken Getz for his GDI+ color picker example, Ben Peterson for h" & _
        "is SVG example projects, and Bob Powell for his ever useful GDI+ FAQ. "
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(0, 80)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(496, 160)
        Me.PictureBox1.TabIndex = 6
        Me.PictureBox1.TabStop = False
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Location = New System.Drawing.Point(8, 264)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(128, 23)
        Me.LinkLabel1.TabIndex = 7
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "www.mrgsoft.com"
        '
        'btnOk
        '
        Me.btnOk.BackColor = System.Drawing.SystemColors.Control
        Me.btnOk.Location = New System.Drawing.Point(400, 392)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.TabIndex = 8
        Me.btnOk.Text = "Ok"
        '
        'btnVersion
        '
        Me.btnVersion.BackColor = System.Drawing.SystemColors.Control
        Me.btnVersion.Location = New System.Drawing.Point(312, 392)
        Me.btnVersion.Name = "btnVersion"
        Me.btnVersion.TabIndex = 9
        Me.btnVersion.Text = "Version Info"
        '
        'lnkSource
        '
        Me.lnkSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lnkSource.Location = New System.Drawing.Point(16, 400)
        Me.lnkSource.Name = "lnkSource"
        Me.lnkSource.Size = New System.Drawing.Size(200, 23)
        Me.lnkSource.TabIndex = 10
        Me.lnkSource.TabStop = True
        Me.lnkSource.Text = "Want the source?  Buy It here."
        '
        'frmAbout
        '
        Me.AcceptButton = Me.btnOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(496, 466)
        Me.Controls.Add(Me.lnkSource)
        Me.Controls.Add(Me.btnVersion)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lblDotNetMAgic)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.pbAbout)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmAbout"
        Me.ShowInTaskbar = False
        Me.Text = "About GDI+ Architect"
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Event Handlers"

    Private Sub lnkSource_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkSource.LinkClicked
        System.Diagnostics.Process.Start(Constants.URL_BUYSOURCE)

    End Sub
    Private Sub lblDotNetMAgic_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lblDotNetMAgic.LinkClicked
        lblDotNetMAgic.LinkVisited = True
        ' Call the Process.Start method to open the default browser 
        ' with a URL:z
        System.Diagnostics.Process.Start(Constants.URL_DOTNETMAGIC)
    End Sub

    Protected Overrides Sub OnClick(ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub pbAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pbAbout.Click
        Me.Close()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        lblDotNetMAgic.LinkVisited = True
        ' Call the Process.Start method to open the default browser 
        ' with a URL:
        System.Diagnostics.Process.Start("http://www.mrgsoft.com")

    End Sub

    Private Sub btnOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOk.Click
        Me.Close()
    End Sub

    Private Sub btnVersion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVersion.Click
        Trace.WriteLineIf(App.TraceOn, "About.Version")

        Dim curAssem As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
        Dim objAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetAssembly(GetType(GDIDocument))

        Dim vUIMajor As Int32 = curAssem.GetName.Version.Major()
        Dim vUIMinor As Int32 = curAssem.GetName.Version.Minor()
        Dim vUIBuild As Int32 = curAssem.GetName.Version.Build()
        Dim vUIRevision As Int32 = curAssem.GetName.Version.Revision()

        Dim vDocumentMajor As Int32 = curAssem.GetName.Version.Major()
        Dim vDocumentMinor As Int32 = curAssem.GetName.Version.Minor()
        Dim vDocumentBuild As Int32 = curAssem.GetName.Version.Build()
        Dim vDocumentRevision As Int32 = curAssem.GetName.Version.Revision()

        Dim sUIVersion As String = "GDI+ Architect Interface Version: " & vUIMajor & "." & vUIMinor & "." & vUIBuild
        Dim sDataVersion As String = "GDI+ Architect Document Version: " & vDocumentMajor & "." & vDocumentMinor & "." & vDocumentBuild

        MsgBox(sUIVersion & vbLf & vbLf & sDataVersion)

    End Sub

#End Region

  
End Class
