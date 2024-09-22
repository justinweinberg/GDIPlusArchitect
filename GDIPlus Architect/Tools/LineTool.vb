Imports GDIObjects

Friend Class LineTool
    Inherits GDITool

#Region "Local Fields"
    '''<summary>The end point of the line</summary>
    Private _PTEnd As Point

#End Region


#Region "Constructors"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the LineTool given a starting point
    ''' </summary>
    ''' <param name="pt">The point at which to begin drawing the line</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal pt As Point)
        MyBase.New(pt)
        _PTEnd = pt
    End Sub

#End Region

#Region "Helper Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a GDILine from the current state of the LineTool
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub createLine()
        Dim newLine As New GDILine(_PTOrigin, _PTEnd)
        MDIMain.ActiveDocument.AddObjectToPage(newLine, "Draw Line")
     End Sub

#End Region


#Region "Base Class Overrides"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the line tool to the surface
    ''' </summary>
    ''' <param name="g">Graphics context to draw against</param>
    ''' <param name="fScale">The current zoom factor of the surface</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub Draw(ByVal g As Graphics, ByVal fScale As Single)
        _PenOutline.Width = 1 / fScale
        App.GraphicsManager.BeginScaledView(g)

        g.DrawLine(_PenOutline, _PTOrigin, _PTEnd)

        App.GraphicsManager.EndScaledView(g)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends the line tool, creating a new line if start and end points are different
    ''' </summary>
    ''' <param name="bShiftDown"></param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub EndTool(ByVal bShiftDown As Boolean)
        If Point.op_Inequality(_PTEnd, _PTOrigin) Then
            createLine()
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the LineTool based on the latest mouse point
    ''' </summary>
    ''' <param name="ptPoint">Last recorded mouse point</param>
    ''' <param name="btn">Which buttons are being held.</param>
    ''' <param name="bShiftDown">Whether shift is down or not</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub UpdateTool(ByVal ptPoint As System.Drawing.Point, ByVal btn As MouseButtons, ByVal bShiftDown As Boolean, ByVal zoom As Single)
        _PTLastPoint = ptPoint
        _PTEnd = getEndPoint(ptPoint, bShiftDown)
    End Sub

#End Region
End Class
