
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : ToolButton
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a simple button for use on the toolbar panel.
''' </summary>
''' <remarks>
''' ToolButtons act in conjuction with each other to give the ToolBox the desired 
''' look and feel.  Each ToolButton is responsible for being Selectable and 
''' responding to mouse events.
''' </remarks>
''' -----------------------------------------------------------------------------
Friend Class ToolButton
    Inherits System.Windows.Forms.Control

#Region "Local Fields"


    '''<summary>Image list that holds the image used by this button</summary>
    Private _ImageList As ImageList

    '''<summary>Index of this button's image in the image list</summary>
    Private _ImageIndex As Int32 = 0


    '''<summary>Formatter used to draw text onto the control surface</summary>
    Private _StringFormat As StringFormat

    '''<summary>A text label to render on the tool button</summary>
    Private _Caption As String = String.Empty


    '''<summary>Whether the toolbutton is selected or not </summary>
    Private _Selected As Boolean = False

    '''<summary>Whether the mouse is over the toolbutton or not</summary>
    Private _MouseOver As Boolean = False

    '''<summary>Whether the mouse is held down on the toolbutton or not</summary>
    Private _MouseDown As Boolean = False

#End Region

#Region "Event Declarations"
    '''<summary>Custom Event Raised on mouse up if the mouse
    ''' is still within the button area</summary>
    Public Event ButtonSelected(ByVal s As Object, ByVal e As EventArgs)

#End Region

#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a Toolbutton
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

        'All GDI+ Architect tool buttons are  26 x 26 at this time.
        Me.Size = New Size(26, 26)

        'Indicates to .NET that the button will handle its own mouse processing
        SetStyle(ControlStyles.UserMouse, True)

        'Sets up the string format used to draw the text
        _StringFormat = System.Drawing.StringFormat.GenericDefault
        _StringFormat.Trimming = StringTrimming.EllipsisCharacter
        _StringFormat.Alignment = StringAlignment.Near


    End Sub

#End Region


#Region "Property Accessors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets a value indicating if this button is selected.
    ''' </summary>
    ''' <value>Whether the button is selected or not</value>
    ''' -----------------------------------------------------------------------------
    Public Property Selected() As Boolean
        Get
            Return _Selected
        End Get
        Set(ByVal Value As Boolean)
            _Selected = Value
            Invalidate()
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the tool button's caption property
    ''' </summary>
    ''' <value>String to display on the button</value>
    ''' -----------------------------------------------------------------------------
    Public Property Caption() As String
        Get
            Return _Caption
        End Get
        Set(ByVal Value As String)
            _Caption = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the tool button's image list
    ''' </summary>
    ''' <value>Index of the image this ToolButton should use</value>
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
    ''' Gets or sets the image index of the button in the imagelist
    ''' </summary>
    ''' <value>Int32 representing the image index of the button</value>
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
    ''' Reacts to a mouse up on the tool button.  Raises the custom ToolButtonClick
    ''' if the mouse was released within the bounds of the button and makes it the 
    ''' selected button.
    ''' </summary>
    ''' <param name="e">Standard MouseEventArgs</param>
    '''-----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        _MouseDown = False

        If New Rectangle(0, 0, Me.Bounds.Width, Me.Bounds.Height).Contains(e.X, e.Y) Then
            _Selected = True
            RaiseEvent ButtonSelected(Me, EventArgs.Empty)
        Else
            _Selected = False
        End If

        Invalidate()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Reacts to a mouse down on the tool buttoin
    ''' </summary>
    ''' <param name="e">Standard MouseEventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        _MouseDown = True
        Invalidate()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles mouse over
    ''' </summary>
    ''' <param name="e">Standard EventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseEnter(ByVal e As System.EventArgs)
        _MouseOver = True
        Invalidate()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles mouse leave
    ''' </summary>
    ''' <param name="e">Standard EventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseLeave(ByVal e As System.EventArgs)
        _MouseOver = False
        Invalidate()
    End Sub

#End Region

#Region "Paint Override"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the tool button when OnPaint is initiated.
    ''' </summary>
    ''' <param name="e">Standard PaintEventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        Dim g As Graphics = e.Graphics

        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SystemDefault

        'Rectangle representing the client area
        Dim rectBounds As New Rectangle(0, 0, Me.Bounds.Width, Me.Bounds.Height)

        'Brush used to paint the highlighted background look
        Dim bHighlight As New SolidBrush(ControlPaint.Light(System.Drawing.SystemColors.Control))

        'Brush used to paint the normal look
        Dim bNormal As New SolidBrush(System.Drawing.SystemColors.Control)

        'If we have an image 
        If Not _ImageList Is Nothing Then

            Dim _image As Image = _ImageList.Images(_ImageIndex)


            'If the control is disabled, draw it disabled
            If Not Me.Enabled Then
                ControlPaint.DrawImageDisabled(g, _image, 5, 5, System.Drawing.SystemColors.Control)

                ControlPaint.DrawStringDisabled(g, _Caption, Me.Font, Color.LightGray, New RectangleF(_image.Width + 9, 7, Me.Width, Me.Height), _StringFormat)

            Else
                'If the mouse is down on the button, paint it sunken, higlighted, and with the image 
                'and caption slightly offset to give it a pressed look.
                If _MouseDown Then
                    g.FillRectangle(bHighlight, rectBounds)
                    ControlPaint.DrawBorder3D(g, New Rectangle(0, 0, Me.Bounds.Width, Me.Bounds.Height), Border3DStyle.SunkenInner)
                    g.DrawImage(_image, 6, 6)
                    g.DrawString(_Caption, Me.Font, Brushes.Black, _image.Width + 10, 8)
                Else
                    'If the button is selected, paint sunken and highlighted
                    If _Selected Then
                        g.FillRectangle(bHighlight, rectBounds)
                        ControlPaint.DrawBorder3D(g, New Rectangle(0, 0, Me.Bounds.Width, Me.Bounds.Height), Border3DStyle.SunkenInner)
                        g.DrawImage(_image, 5, 5)
                        g.DrawString(_Caption, Me.Font, Brushes.Black, _image.Width + 9, 7)
                        'If the mouse is over the button, paint it raised
                    ElseIf _MouseOver Then
                        g.FillRectangle(bNormal, rectBounds)
                        ControlPaint.DrawBorder3D(g, New Rectangle(0, 0, Me.Bounds.Width, Me.Bounds.Height), Border3DStyle.RaisedInner)
                        g.DrawImage(_image, 5, 5)
                        g.DrawString(_Caption, Me.Font, Brushes.Black, _image.Width + 9, 7)
                    Else

                        'Paint it normal
                        g.FillRectangle(bNormal, rectBounds)
                        g.DrawImage(_image, 5, 5)
                        g.DrawString(_Caption, Me.Font, Brushes.Black, _image.Width + 9, 7)
                    End If

                End If

                bNormal.Dispose()
                bHighlight.Dispose()
            End If

        End If
    End Sub


#End Region

#Region "Public Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' User to help the tool buttons act as a group of buttons.  Each button is asked 
    ''' to check if it is selected and update its selection property and UI as necessary.  
    ''' </summary>
    ''' <param name="s">The selected ToolButton.  THe toolbutton is expected to check 
    ''' if its reference is the same as this button and select itself as necessary.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub ToggleSelection(ByVal s As ToolButton)
        If Not s Is Me Then
            Me.Selected = False
        Else
            Me.Selected = True
        End If

        Invalidate()

    End Sub
#End Region

#Region "Cleanup and Disposal"
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)

        If disposing Then
            If Not _StringFormat Is Nothing Then
                _StringFormat.Dispose()
            End If
        End If

        MyBase.Dispose(disposing)
    End Sub
#End Region
   
End Class
