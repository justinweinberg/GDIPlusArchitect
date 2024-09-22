Imports System.CodeDom



''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDISolidFill
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A solid fill for fillable objects.  The solid fill is composed only of a color
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public Class GDISolidFill
    Inherits GDIFill

#Region "Local Fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The color to fill this parent object with.  This is the only interesting 
    ''' property for solid fills.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _FillColor As Color = Color.White
#End Region

#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a solid fill given a parent shape and a color.
    ''' </summary>
    ''' <param name="parent">A filled shape instance this fill is associated with.</param>
    ''' <param name="ccolor">The color to assign to the solid fill.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal cColor As Color)
        MyBase.New()
        Trace.WriteLineIf(Session.Tracing, "SolidFill.New")

        Me.Color = cColor

        UpdateFill()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a solid fill given a parent shape and another fill to base this one on.
    ''' </summary>
    ''' <param name="fill">The other fill to use as a basis for this fill's properties.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal fill As GDISolidFill)
        MyBase.New()
        Trace.WriteLineIf(Session.Tracing, "SolidFill.New.Clone")
        With Me
            .Color = fill.Color
        End With

        UpdateFill()

    End Sub


#End Region

#Region "Code Export Related Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts the solid fill to an XML source for SVG display.
    ''' </summary>
    ''' <param name="xmlDoc">See base class</param>
    ''' <param name="defs">See base class</param>
    ''' <param name="parent">See base class</param>
    ''' -----------------------------------------------------------------------------
    Public Overrides Sub toXML(ByVal xmlDoc As Xml.XmlDocument, ByVal defs As Xml.XmlNode, ByVal parent As Xml.XmlNode)
        Dim arg As String
        Dim attr As Xml.XmlAttribute
        attr = xmlDoc.CreateAttribute("fill")


        If Me.Color.IsNamedColor Then
            attr.Value = Me.Color.Name
        Else
            attr.Value = String.Format("rgb({0},{1},{2})", Me.Color.R, Me.Color.G, Me.Color.B)
        End If

        parent.Attributes.Append(attr)

        If Me.Color.A < 255 Then
            attr = xmlDoc.CreateAttribute("opacity")
            arg = FormatNumber(Me.Color.A / 255, 3)
            attr.Value = arg
            parent.Attributes.Append(attr)
        End If




    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a consolidated solid fill to code.
    ''' </summary>
    ''' <param name="sSharedName">See base class</param>
    ''' <param name="declarations">See base class</param>
    ''' <param name="InitGraphics">See base class</param>
    ''' <param name="RenderGDI">See base class</param>
    ''' <param name="DisposeGDI">See base class</param>
    ''' <param name="ExportSettings">See base class</param>
    ''' <returns>See base class</returns>
    ''' -----------------------------------------------------------------------------
    Friend Overloads Overrides Function emit(ByVal sSharedName As String, _
  ByVal declarations As CodeDom.CodeTypeMemberCollection, _
  ByVal InitGraphics As CodeDom.CodeMemberMethod, _
  ByVal RenderGDI As CodeDom.CodeStatementCollection, _
  ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
  ByVal ExportSettings As ExportSettings) As String

        Trace.WriteLineIf(Session.Tracing, "SolidFill.emit")

        Dim fillDeclaration As New CodeMemberField
        Dim fillInit As New CodeObjectCreateExpression
        Dim fillDispose As New CodeDom.CodeMethodInvokeExpression

        With fillInit
            .CreateType = New CodeTypeReference(GetType(System.Drawing.SolidBrush))
            .Parameters.Add(getColorAssignment(Me.Color))
        End With

        With fillDeclaration
            .Name = sSharedName
            .Type = New CodeTypeReference(GetType(System.Drawing.SolidBrush))
            .Attributes = MemberAttributes.Private
            .InitExpression = fillInit
        End With

        With fillDispose
            .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, fillDeclaration.Name)
            .Method.MethodName = "Dispose"
        End With

        With DisposeGDI
            .Statements.Add(fillDispose)
        End With

        With declarations
            .Add(fillDeclaration)
        End With


        Return fillDeclaration.Name

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a solid fill to code.
    ''' </summary>
    ''' <param name="obj">See base class</param>
    ''' <param name="declarations">See base class</param>
    ''' <param name="InitGraphics">See base class</param>
    ''' <param name="RenderGDI">See base class</param>
    ''' <param name="DisposeGDI">See base class</param>
    ''' <param name="ExportSettings">See base class</param>
    ''' <param name="Consolidated">See base class</param>
    ''' <returns>See base class</returns>
        '''  -----------------------------------------------------------------------------
    Friend Overloads Overrides Function emit(ByVal obj As GDIObject, _
  ByVal declarations As CodeDom.CodeTypeMemberCollection, _
  ByVal InitGraphics As CodeDom.CodeMemberMethod, _
  ByVal RenderGDI As CodeDom.CodeStatementCollection, _
  ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
  ByVal ExportSettings As ExportSettings, _
  ByVal Consolidated As ExportConsolidate) As String

        Trace.WriteLineIf(Session.Tracing, "SolidFill.emit.Consolidated")

        If Consolidated.hasFillMatch(Me) And Me.Consolidate Then
            Return Consolidated.getFillName(Me)
        Else

            Dim fillDeclaration As New CodeMemberField
            Dim fillInit As New CodeObjectCreateExpression
            Dim fillDispose As New CodeDom.CodeMethodInvokeExpression

            With fillInit
                .CreateType = New CodeTypeReference(GetType(System.Drawing.SolidBrush))
                .Parameters.Add(getColorAssignment(Me.Color))
            End With

            With fillDeclaration
                .Name = obj.ExportName & "Brush"
                .Type = New CodeTypeReference(GetType(System.Drawing.SolidBrush))
                .Attributes = MemberAttributes.Private
                .InitExpression = fillInit
            End With

            With fillDispose
                .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, fillDeclaration.Name)
                .Method.MethodName = "Dispose"
            End With

            With DisposeGDI
                .Statements.Add(fillDispose)
            End With

            With declarations
                .Add(fillDeclaration)
            End With


            Return fillDeclaration.Name
        End If
    End Function
#End Region


#Region "Misc Methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a string representation of this object.
    ''' </summary>
    ''' <returns>A string representation of the object</returns>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function ToString() As String
        Return "Solid Fill"
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Compares two solid fills to see if they are equivalent
    ''' </summary>
    ''' <param name="fill1">The first solid fill</param>
    ''' <param name="fill2">The second solid fill</param>
    ''' <returns>A Boolean indicating if the fills are equivalent.</returns>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function op_Equality(ByVal fill1 As GDISolidFill, ByVal fill2 As GDISolidFill) As Boolean
        Return System.Drawing.Color.op_Equality(fill1.Color, fill2.Color)
    End Function


#End Region

#Region "Property Accessors"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the color of the solid fill.
    ''' </summary>
    ''' <value>A System.Drawing color</value>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.Description("The color of a solid fill brush")> _
       Public Property Color() As Color
        Get
            Return _FillColor
        End Get
        Set(ByVal Value As Color)
            Trace.WriteLineIf(Session.Tracing, "SolidFill.Color.Set" & Value.ToString)
            _FillColor = Value
            UpdateFill()
        End Set
    End Property
#End Region

#Region "Base Class Implementers"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Implements the base class's (GDIFill) onparentupdated function.  In this case 
    ''' it does nothing since solid fills do not need to know anything about their 
    ''' parent.
    ''' </summary>
    ''' <param name="obj">The parent object being updated</param>
    ''' -----------------------------------------------------------------------------
    Public Overrides Sub OnParentUpdated(ByVal obj As GDIFilledShape)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Implements the base class's (GDIFill) updateFill method.  Creates a solid brush from 
    ''' the current settings and assigns it to the _Brush member.
    ''' </summary>
    ''' <remarks>Raises the base class FillUpdated event
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub UpdateFill()
        If Not _Brush Is Nothing Then
            _Brush.Dispose()
        End If


        _Brush = New SolidBrush(_FillColor)
        DirectCast(_Brush, SolidBrush).Color = _FillColor

        NotifyFillUpdated()

    End Sub

#End Region
End Class
