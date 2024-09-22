
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : GDITool
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Base class for all of the "tools" that can be used on the surface.  Tools are
''' things such as the pen tool and line tool but also includes less obvious 
''' tools such as the "drag handle tool".
''' </summary>
''' -----------------------------------------------------------------------------
Public MustInherit Class GDITool
    Implements IDisposable

#Region "Local Fields"


    '''<summary>Whether the tool has been disposed or not</summary>
    Protected _Disposed As Boolean = False

    '''<summary>A pen to use when outlining shapes </summary>
    Protected _PenOutline As Pen


    '''<summary>Origin point of the tool </summary>
    Protected _PTOrigin As Point

    '''<summary>Last point the tool used.</summary>
    Protected _PTLastPoint As Point
#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new instance of a tool
    ''' </summary>
    ''' <param name="ptOrigin">The initial point, snapped as appropriate, where the tool was 
    ''' invoked.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal ptOrigin As Point)
        _PTOrigin = ptOrigin
        _PenOutline = New Pen(App.Options.ColorOutline, 1)
    End Sub
#End Region



#Region "Requires Implementation"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Required for inheritors to implement.  Ends the operation of the current tool.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend MustOverride Sub EndTool(ByVal bShiftDown As Boolean)

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Required for implementors to implement.  Renders the tool to the surface
    ''' </summary>
    ''' <param name="g">Graphics context to draw to</param>
    ''' <param name="fScale">Current zoom factor of the surface</param>
    ''' -----------------------------------------------------------------------------
    Friend MustOverride Sub draw(ByVal g As Graphics, ByVal fScale As Single)

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Required for implementors to implement.  Given a point and a mouse button, 
    ''' tools should update themselves as needed.
    ''' </summary>
    ''' <param name="ptPoint">The last point, snapped appropriately if needed.</param>
    ''' <param name="ebutton">The mouse button being reacted to.</param>
    ''' -----------------------------------------------------------------------------
    Friend MustOverride Overloads Sub UpdateTool(ByVal ptPoint As Point, ByVal ebutton As MouseButtons, ByVal bShiftDown As Boolean, ByVal Zoom As Single)
#End Region

#Region "Helper Methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the bounding rectangle appropriately depending on if shift is 
    ''' being held.
    ''' </summary>
    ''' <param name="ptsnapped">Last mouse point, snapped appropriately</param>
    ''' <param name="bShiftDown">Whether shift is being held down or not</param>
    ''' -----------------------------------------------------------------------------
    Protected Function getDrawingRect(ByVal ptSnapped As Point, ByVal bShiftDown As Boolean) As Rectangle
        Dim temporigin As New Point
        Dim tempextent As New Size

        temporigin.X = Math.Min(_PTOrigin.X, ptSnapped.X)
        temporigin.Y = Math.Min(_PTOrigin.Y, ptSnapped.Y)

        tempextent.Width = Math.Abs(ptSnapped.X - _PTOrigin.X)
        tempextent.Height = Math.Abs(ptSnapped.Y - _PTOrigin.Y)

        If bShiftDown Then
            If tempextent.Height > tempextent.Width Then
                tempextent.Width = tempextent.Height
            Else
                tempextent.Height = tempextent.Width
            End If
        End If

        Return New Rectangle(temporigin, tempextent)

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the end point of the line based on whether shift is being held down.
    ''' </summary>
    ''' <param name="ptsnapped">The end point of the line, snapped appropriately</param>
    ''' <param name="bShiftDown">Whether shift is being held or not</param>
    ''' -----------------------------------------------------------------------------
    Protected Function getEndPoint(ByVal ptsnapped As Point, ByVal bShiftDown As Boolean) As Point


        If bShiftDown Then
            Return Utility.ShiftedDownPoint(_PTOrigin, ptsnapped)
        Else
            Return ptsnapped
        End If
    End Function

#End Region




#Region "Cleanup and disposal"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of a GDIStroke, specifically disposing of the pen.
    ''' </summary>
    ''' <param name="disposing">Whether unmanaged resources are being disposed or not.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overridable Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            _PenOutline.Dispose()
            _Disposed = True
        End If

        GC.SuppressFinalize(Me)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Calls the custom dispose method
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Sub Dispose() Implements System.IDisposable.Dispose
        Me.Dispose(True)
    End Sub
#End Region
End Class
