
''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDIGuide
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Represents a guide (alignment line) on the GDI+ Architect drawing surface.
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public Class GDIGuide

    '''<summary>Some extra pixels for hit testing.</summary>
    Private Const HIT_TEST_RANGE As Int32 = 3



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Performs a hit test against a guide.  Returns a boolean 
    ''' indicating if the 
    ''' point specified crosses the guide.
    ''' </summary>
    ''' <param name="pt">A point to hit test against in object coordinates</param>
    ''' <returns>A boolean indicating if the guide was hit or not.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function HitTest(ByVal pt As Point) As Boolean
        If Me.Direction = GDIGuide.EnumGuideDirection.eHoriz Then
            If Math.Abs(Me.XY - pt.Y) <= HIT_TEST_RANGE Then
                Return True
            End If
        Else
            If Math.Abs(Me.XY - pt.X) <= HIT_TEST_RANGE Then
                Return True
            End If
        End If

        Return False
    End Function

#Region "Type Declarations"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether the guide is horizontal or vertical.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Enum EnumGuideDirection
        ''' <summary>A horizontal guide</summary>
        eHoriz
        ''' <summary>A vertical guide</summary>
        eVert
    End Enum
#End Region

#Region "Local Fields"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether the guide is horizontal or vertical.
    ''' </summary>
    ''' 
    '''  -----------------------------------------------------------------------------
    Private _Direction As EnumGuideDirection = EnumGuideDirection.eHoriz

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The guide's position.  (Y value for a vertical guide, X for a horizontal)
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _XY As Int32 = 0

#End Region

#Region "Drawing related methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders a guide to the drawing surface.
    ''' </summary>
    ''' <param name="g">The current graphics context</param>
    ''' <param name="fScale">the current zoom</param>
    ''' -----------------------------------------------------------------------------
    Public Sub Draw(ByVal g As Graphics, ByVal fScale As Single)

        Dim pPen As New Pen(Session.Settings.GuideColor, 1 / fScale)
        Select Case _Direction
            Case EnumGuideDirection.eHoriz
                g.DrawLine(pPen, -30000, _XY, 30000, _XY)
            Case EnumGuideDirection.eVert
                g.DrawLine(pPen, _XY, -30000, _XY, 30000)
        End Select

        pPen.Dispose()
        pPen = Nothing
    End Sub
#End Region

#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the direction of the Guide.
    ''' </summary>
    ''' <value>An enumeration representing whether the guide is drawn horizontal or vertical.</value>
    ''' -----------------------------------------------------------------------------
    Public Property Direction() As EnumGuideDirection
        Get
            Return _Direction
        End Get
        Set(ByVal Value As EnumGuideDirection)
            _Direction = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the XY position of the guide (Y for vertical, X for horizontal).
    ''' </summary>
    ''' <value>An integer representing the position of the guide</value>
    ''' -----------------------------------------------------------------------------
    Public Property XY() As Int32
        Get
            Return _XY
        End Get
        Set(ByVal Value As Int32)
            _XY = Value
        End Set
    End Property
#End Region
End Class
