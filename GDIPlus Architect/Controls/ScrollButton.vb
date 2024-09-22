

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : ScrollButton
''' 
''' -----------------------------------------------------------------------------
''' <summary>The ScrollButton implements a scroll on click for the ToolBox Panel.</summary>
''' <remarks>
''' The scrollbutton is used exclusively on the ToolBox Panel to mimic 
''' Visual Studio's scroll buttons.  The concept is that 
''' on the first mouse down, the control should register a single scroll.  
''' If the mouse held down, the button should raise scroll events at a normal 
''' increment. 
''' </remarks>
''' -----------------------------------------------------------------------------
<System.ComponentModel.DefaultEvent("Scroll")> _
Friend Class ScrollButton
    Inherits System.Windows.Forms.UserControl


#Region "Local Fields"


    '''<summary>Whether the mouse is down on the button or not </summary>
    Private _MouseDown As Boolean = False

    '''<summary>How long the mouse has been held down.</summary>
    Private _Ticks As Int32 = 0

    '''<summary>Image list that holds the image needed to
    ''' paint the scroll button</summary>
    Private _ImageList As ImageList

    '''<summary>Index within the image list of the scroll button's image.</summary>
    Private _ImageIndex As Int32 = 0
#End Region

#Region "Event Declarations"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A custom event raised when the button is initially clicked and when 
    ''' it's being held down.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Event Scroll(ByVal sender As Object, ByVal e As EventArgs)


#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a scroll button.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
        SetStyle(ControlStyles.DoubleBuffer, True)


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
    Private WithEvents tmrScroll As System.Windows.Forms.Timer

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
        tmrScroll = New Timer(components)
        '
        'tmrScroll
        '
        tmrScroll.Interval = 25
        '
        'ScrollButton
        '
        Me.Size = New Size(26, 26)

    End Sub

#End Region


#Region "Property Accessors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the image list that renders the scroll button to the surface.
    ''' </summary>
    ''' <value>An image list used that contains the image needed to paint the button</value>
    ''' -----------------------------------------------------------------------------
    Public Property ImageList() As ImageList
        Get
            Return _ImageList
        End Get
        Set(ByVal Value As ImageList)
            _ImageList = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the index within the image list used to paint the button
    ''' </summary>
    ''' <value>The position in the ImageList to retrieve the button's image from.</value>
    ''' -----------------------------------------------------------------------------
    Public Property ImageIndex() As Int32
        Get
            Return _ImageIndex
        End Get
        Set(ByVal Value As Int32)
            _ImageIndex = Value

        End Set
    End Property

#End Region


#Region "Event Handlers and Overrides"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Overrides mouse down to start the timer, reset ticks, and set  to true.
    ''' </summary>
    ''' <param name="e">Standard MouseEventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        _MouseDown = True
        Invalidate()
        _Ticks = 0
        RaiseEvent Scroll(Me, New EventArgs)
        tmrScroll.Start()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Overrides mouse up to stop the timer. and set  to false.
    ''' </summary>
    ''' <param name="e">Standard MouseEventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        _MouseDown = False
        Invalidate()
        tmrScroll.Stop()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles ticks from the timer and raises the scroll event repeatedly if the 
    ''' tick count is over 20
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub tmrScroll_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrScroll.Tick
        _Ticks += 1
        If _Ticks > 20 Then
            RaiseEvent Scroll(Me, New EventArgs)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Stops the timer on mouse leave and sets mouse down to false
    ''' </summary>
    ''' <param name="e">Standard EventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseLeave(ByVal e As System.EventArgs)
        _MouseDown = False
        Invalidate()
        tmrScroll.Stop()
    End Sub


#End Region


#Region "Paint Override"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the scroll button.  The scroll button has a different appearance
    ''' if it is disabled or if the mouse is down.
    ''' </summary>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        Dim g As Graphics = e.Graphics

        Dim rectBounds As New Rectangle(0, 0, Me.Bounds.Width, Me.Bounds.Height)
        Dim bHighlight As New SolidBrush(ControlPaint.Light(System.Drawing.SystemColors.Control))
        Dim bNormal As New SolidBrush(System.Drawing.SystemColors.Control)

        If Not _ImageList Is Nothing Then
            If Me.Enabled Then
                If _MouseDown Then
                    g.FillRectangle(bHighlight, rectBounds)
                    ControlPaint.DrawBorder3D(g, New Rectangle(0, 0, Me.Bounds.Width, Me.Bounds.Height), Border3DStyle.SunkenInner)
                    g.DrawImage(_ImageList.Images(_ImageIndex), 6, 6)
                Else
                    g.FillRectangle(bNormal, rectBounds)
                    g.DrawImage(_ImageList.Images(_ImageIndex), 5, 5)
                End If

            Else
                ControlPaint.DrawImageDisabled(g, _ImageList.Images(_ImageIndex), 5, 5, System.Drawing.SystemColors.Control)
            End If
        End If

        bHighlight.Dispose()
        bNormal.Dispose()

    End Sub
#End Region

End Class
