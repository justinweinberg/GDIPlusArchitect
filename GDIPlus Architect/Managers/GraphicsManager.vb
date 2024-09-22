Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : GraphicsManager
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Responsible for managing sundry graphics operations that span the object model 
''' and don't have a fixed point in the application.  This is basically a graphics 
''' utility class.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class GraphicsManager


    '''<summary>Holds transformation information for the surface's graphics context for zoom</summary>
    Private _GraphicsContainer As Drawing.Drawing2D.GraphicsContainer

    '''<summary>Holds the original matrix prior to transformations for zoom</summary>
    Private _MatrixInitial As Drawing2D.Matrix

#Region "Public Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a graphics container for arbitary scaled drawing.
    ''' </summary>
    ''' <param name="g">A System.Drawing Graphics context to scale.</param>
    ''' <param name="fScale">Amount to scale the container by.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub BeginScaledView(ByVal g As Graphics, ByVal fScale As Single)

        'Get any current transformations applied to the incoming graphics object
        'so we can restore them later
        _MatrixInitial = g.Transform()

        'Create a graphics container for this series of output.  This prevents
        'The transformation from applying to all previous graphic transformation
        _GraphicsContainer = g.BeginContainer

        g.TextRenderingHint = MDIMain.ActiveDocument.TextRenderingHint
        g.SmoothingMode = MDIMain.ActiveDocument.SmoothingMode
        'Set scale and units for our draw
        g.ScaleTransform(fScale, fScale)

    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a graphics container scaled to the ActiveDocument's scale
    ''' </summary>
    ''' <param name="g">A System.Drawing Graphics context to scale.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub BeginScaledView(ByVal g As Graphics)

        'Get any current transformations applied to the incoming graphics object
        'so we can restore them later
        _MatrixInitial = g.Transform()

        'Create a graphics container for this series of output.  This prevents
        'The transformation from applying to all previous graphic transformation
        _GraphicsContainer = g.BeginContainer

        g.TextRenderingHint = MDIMain.ActiveDocument.TextRenderingHint
        g.SmoothingMode = MDIMain.ActiveDocument.SmoothingMode
        'Set scale and units for our draw
        g.ScaleTransform(MDIMain.ActiveZoom, MDIMain.ActiveZoom)

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends a previously created graphics container .
    ''' </summary>
    ''' <param name="g">A System.Drawing Graphics context to scale.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub EndScaledView(ByVal g As Graphics)

        'reset the graphics object to its original value
        g.EndContainer(_GraphicsContainer)
        g.Transform = _MatrixInitial
        _MatrixInitial.Dispose()
    End Sub


#End Region
End Class
