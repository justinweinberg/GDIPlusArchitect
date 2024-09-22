Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : DrawingTool
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Providers a wraper for drawing shapes into the surface (ellipses and circles)
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class ShapeTool
    Inherits GDITool

#Region "Local Fields"
    '''<summary>Current drawing rectangle bounds</summary>
    Private _Bounds As Rectangle
    '''<summary>Last mouse point recorded</summary>
    Private _PTEnd As Point

#End Region


#Region "Constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new instance of the ShapeTool given an initial point where 
    ''' drawing should bei
    ''' </summary>
    ''' <param name="pt"></param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal pt As Point)
        MyBase.New(pt)
        _PTEnd = pt


    End Sub
#End Region

#Region "Base Class Implementors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the ShapeTool based upon the latest information from the client 
    ''' surface
    ''' </summary>
    ''' <param name="ptPoint">The latest mouse point recorded on the surface</param>
    ''' <param name="btn">Which mouse button is being held down</param>
    ''' <param name="bShiftDown">Whether the shift key is held down or not</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub UpdateTool(ByVal ptPoint As System.Drawing.Point, ByVal btn As MouseButtons, ByVal bShiftDown As Boolean, ByVal zoom As Single)
        _PTEnd = MyBase.getEndPoint(ptPoint, bShiftDown)
        _Bounds = getDrawingRect(ptPoint, bShiftDown)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the ShapeTool to the client surface
    ''' </summary>
    ''' <param name="g">Graphics context to draw against</param>
    ''' <param name="fScale">Current zoom factor on the surfaceparam>
    ''' -----------------------------------------------------------------------------
    Friend Overloads Overrides Sub draw(ByVal g As System.Drawing.Graphics, ByVal fScale As Single)
        _PenOutline.Width = 1 / fScale

        App.GraphicsManager.BeginScaledView(g)

        Select Case App.ToolManager.GlobalMode
            Case enumTools.eCircle
                g.DrawEllipse(_PenOutline, _Bounds)

            Case enumTools.eSquare
                g.DrawRectangle(_PenOutline, _Bounds)


        End Select


        App.GraphicsManager.EndScaledView(g)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends a ShapeTool operation
    ''' </summary>
    ''' <param name="bShiftDown">Whether shift is being held down or not</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub EndTool(ByVal bShiftDown As Boolean)
        If Point.op_Inequality(_PTEnd, _PTOrigin) Then

            Select Case App.ToolManager.GlobalMode

                Case enumTools.eCircle
                    If _Bounds.Width > 0 AndAlso _Bounds.Height > 0 Then
                        CreateCircle()
                    End If

                Case enumTools.eSquare
                    If _Bounds.Width > 0 AndAlso _Bounds.Height > 0 Then
                        createRectangle()
                    End If


            End Select
        End If
    End Sub

#End Region


#Region "Helper Methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a GDIRectangle
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub createRectangle()
        Dim newSquare As New GDIRect(_Bounds)

        MDIMain.ActiveDocument.AddObjectToPage(newSquare, "Draw Rectangle")
     End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a GDIEllipse
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub CreateCircle()

        Dim newCircle As New GDIEllipse

        With newCircle
            .Stroke.Color = Session.Stroke.Color
            .Bounds = _Bounds
        End With

        MDIMain.ActiveDocument.AddObjectToPage(newCircle, "Draw Circle")
     End Sub



#End Region


End Class
