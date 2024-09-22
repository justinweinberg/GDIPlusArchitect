

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : HitHandle
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A simple container for handle related information.
''' </summary>
''' <remarks>Hit handles can be either part of a curvature or an edge to a rectangle.
''' </remarks>
''' -----------------------------------------------------------------------------
Public Class HitHandle

#Region "Type Declarations"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Represents the type of a hit handle.  A normal handle is a typical edge point on 
    ''' a rectangle, whereas a curve handle is part of a Bezier curve on a path.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Enum EnumHandletypes
        '''<summary>eNormal represents a typical handle such as the edge of a rectangle 
        ''' or a standard point along the path.</summary>
        eNormal
        '''<summary>eCurvePoint represents a handle that is part of a Bezier curve</summary>
        eCurvePoint
    End Enum

#End Region

#Region "Local Fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A rectangleF (a rectangle composed of floats) representing the area of the handle
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _HandleArea As RectangleF
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The type of handle (part of a curvature or a normal edge handle). This is used 
    ''' to paint points differently depending on the purpose of the point.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _HandleType As EnumHandletypes

#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructor for a new hit handle
    ''' </summary>
    ''' <param name="rHandleRect">The rectangle representing the hit area</param>
    ''' <param name="eHandleType">The type of the handle (part of a curvature or normal edge)</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal rHandleRect As RectangleF, ByVal eHandleType As EnumHandletypes)
        _HandleArea = rHandleRect
        _HandleType = eHandleType
    End Sub

#End Region


#Region "Misc Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a value indicating whether this handle contains a specified point.  This 
    ''' is where the actual hit test is performed.
    ''' </summary>
    ''' <param name="pt">The handle point to test.</param>
    ''' <returns>A </returns>
    ''' -----------------------------------------------------------------------------
    Public Function Contains(ByVal pt As PointF) As Boolean
        Return _HandleArea.Contains(pt)
    End Function
#End Region

#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The rectangle this hit handle is valid within
    ''' </summary>
    ''' <value>A rectangleF structure containing the area of the handle</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property HandleRect() As RectangleF
        Get
            Return _HandleArea
        End Get

    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The type of the hit handle (curvature point or a normal edge style handle)
    ''' </summary>
    ''' <value>EnumHandleTypes which is the type of handle.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property HandleType() As EnumHandletypes
        Get
            Return _HandleType
        End Get

    End Property

#End Region
End Class
