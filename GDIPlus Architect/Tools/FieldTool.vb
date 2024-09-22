Imports GDIObjects
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : FieldTool
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Wraps the functionality needed to created GDI+ fields.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class FieldTool
    Inherits GDITool


#Region "Local Fields"
    '''<summary>Current rectangle defined for this tool on the surface</summary>
    Private _Bounds As Rectangle
    '''<summary>The last point recorded by the tool</summary>
    Private _PTEnd As Point


#End Region


#Region "Constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the GDIFieldTool given an initial point to start 
    ''' bounding the field box from.
    ''' </summary>
    ''' <param name="pt">The point to begin drawing the field from.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal pt As Point)
        MyBase.New(pt)
        _PTEnd = pt
    End Sub

#End Region


#Region "Helper Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new GDI_ Field based upon the FieldTool's state.  All GDIFields are 
    ''' initially created as wrapping GDIFields.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub createField()
        Dim newField As New GDIField(App.Options.SampleText, App.Options.Font, _Bounds, True)
        MDIMain.ActiveDocument.AddObjectToPage(newField, "Field Created")
    End Sub

#End Region

#Region "Base Class Implementors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Renders the field tool to a drawing surface
    ''' </summary>
    ''' <param name="g">Graphics context to draw against </param>
    ''' <param name="fscale">The current zoom factor of the surface</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub Draw(ByVal g As Graphics, ByVal fscale As Single)
        _PenOutline.Width = 1 / fscale
   
        App.GraphicsManager.BeginScaledView(g)
        'Set scale and units for our draw
        'g.PageUnit = GraphicsUnit.Pixel
        g.ScaleTransform(fscale, fscale)

        g.DrawRectangle(_PenOutline, _Bounds.X, _Bounds.Y, _Bounds.Width, _Bounds.Height)

        App.GraphicsManager.EndScaledView(g)

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends the FieldTool operation
    ''' </summary>
    ''' <param name="bShiftDown">Whether shift is being held down or not</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub EndTool(ByVal bShiftDown As Boolean)
        If Point.op_Inequality(_PTEnd, _PTOrigin) Then

            'Verify the field has width and height, otherwise don't create the field.
            If _Bounds.Width > 0 AndAlso _Bounds.Height > 0 Then
                createField()
            End If

        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates a field tool with the latest state information from the surface
    ''' </summary>
    ''' <param name="ptPoint">Last mouse point recorded on the surface</param>
    ''' <param name="btn">Whether the mouse button is down or not</param>
    ''' <param name="bShiftDown">Whether shift is being held down or not</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub UpdateTool(ByVal ptPoint As System.Drawing.Point, ByVal btn As MouseButtons, ByVal bShiftDown As Boolean, ByVal zoom As Single)
        _PTEnd = MyBase.getEndPoint(ptPoint, bShiftDown)

        _Bounds = getDrawingRect(ptPoint, bShiftDown)
    End Sub

#End Region
End Class
