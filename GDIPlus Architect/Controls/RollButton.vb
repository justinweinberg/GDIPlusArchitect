
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : RollButton
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A roll button is very similar to a The ToolButton control. 
''' The goal of the roll button is to provide a roll over popup look.  A roll button 
''' is assumed to not have a textual component.</summary>
''' <remarks>
''' The button raises its own custom click event called ButtonClick.</remarks>
''' 
''' -----------------------------------------------------------------------------
Friend Class RollButton
    Inherits System.Windows.Forms.Control

#Region "Local Fields"

    '''<summary>The source from which the roll button obtains its images</summary>
    Private _ImageList As ImageList

    '''<summary>Index position in the image list that has this button's image.</summary>
    Private _ImageIndex As Int32 = 0



    '''<summary>Whether the mouse is hovering over the button</summary>
    Private _MouseOver As Boolean = False
    '''<summary>Whether the mouse is held down on the button</summary>
    Private _MouseDown As Boolean = False
#End Region

#Region "Event Declarations"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Click event raised from the button.  This event is only
    ''' raised upon a mouse up within the bounds of the button
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Event ButtonClick(ByVal s As Object, ByVal e As EventArgs)

#End Region

#Region "Constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a RollButton.  Sets the style properties to double 
    ''' buffer and usermouse.</summary>
    ''' <remarks>
    ''' All RollButtons are assumed to be 20 x 20.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

         Me.Size = New Size(20, 20)

        SetStyle(ControlStyles.UserMouse, True)
        SetStyle(ControlStyles.DoubleBuffer, True)
    End Sub
#End Region

 
 

#Region "Property Accessors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the image list that has the image that is painted on the button
    ''' </summary>
    ''' <value>The imagelist to assign to the rollbutton</value>
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
    ''' Gets or sets the image index from which the button should retrieve its 
    ''' image in the image list.
    ''' </summary>
    ''' <value>The image index to assign to this roll button</value>
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
    ''' Overrides mouse up on the control in order to raise the custom buttonclick 
    ''' event given that the click was within bounds.
    ''' </summary>
    ''' <param name="e">Standard MouseEventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)
        _MouseDown = False

        If New Rectangle(0, 0, Me.Bounds.Width, Me.Bounds.Height).Contains(e.X, e.Y) Then
            RaiseEvent ButtonClick(Me, New EventArgs)
        End If

        Invalidate()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Overrides mouse down, recording that the mouse is down for when painting
    ''' occurs
    ''' </summary>
    ''' <param name="e">Standard MouseEventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
        _MouseDown = True
        Invalidate()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Initiates the roll over effect by setting the mouse over property.
    ''' </summary>
    ''' <param name="e">Standard MouseEventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseEnter(ByVal e As System.EventArgs)
        _MouseOver = True
        Invalidate()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends the roll over effect by settings the mouse over property to false
    ''' </summary>
    ''' <param name="e">Standard MouseEventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnMouseLeave(ByVal e As System.EventArgs)
        _MouseOver = False
        Invalidate()
    End Sub

#End Region

#Region "Paint Override"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the roll button depending on mouse state.</summary>
    ''' <remarks>First this method checks  whether the button is enabled or not.
    ''' If it's not enabled, it paints the button disabled.
    ''' If it is enabled, it checks if the mouse is down on the button or not.  If 
    ''' it is, it draws the down state.  If not, it checks 
    ''' if the mouse is over the button. If it is, renders the up look.  
    ''' Otherwise if none of the above happened, it renders it normally.
    ''' </remarks>
    ''' <param name="e">Standard PaintEventArgs</param>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
        Dim g As Graphics = e.Graphics

        Dim rectBounds As New Rectangle(0, 0, Me.Bounds.Width, Me.Bounds.Height)
        Dim bHighlight As New SolidBrush(ControlPaint.Light(System.Drawing.SystemColors.Control))
        Dim bNormal As New SolidBrush(System.Drawing.SystemColors.Control)

        If Not _ImageList Is Nothing Then
            Dim _image As Image = _ImageList.Images(_ImageIndex)

            If Me.Enabled Then
                If _MouseDown Then
                    g.FillRectangle(bHighlight, rectBounds)
                    ControlPaint.DrawBorder3D(g, New Rectangle(0, 0, Me.Bounds.Width, Me.Bounds.Height), Border3DStyle.SunkenInner)
                    g.DrawImage(_image, 4, 4)
                ElseIf _MouseOver Then
                    g.FillRectangle(bNormal, rectBounds)
                    ControlPaint.DrawBorder3D(g, New Rectangle(0, 0, Me.Bounds.Width, Me.Bounds.Height), Border3DStyle.RaisedInner)
                    g.DrawImage(_image, 3, 3)
                Else
                    g.FillRectangle(bNormal, rectBounds)
                    g.DrawImage(_image, 3, 3)
                End If


                bNormal.Dispose()
                bHighlight.Dispose()
                bNormal = Nothing
                bHighlight = Nothing

            Else
                ControlPaint.DrawImageDisabled(g, _image, 3, 3, System.Drawing.SystemColors.Control)
            End If

        Else
            ControlPaint.DrawBorder3D(g, Me.Bounds, Border3DStyle.Flat)

        End If
    End Sub
#End Region
End Class
