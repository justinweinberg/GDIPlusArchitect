Imports System.ComponentModel
Imports System.CodeDom

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GDITextBaseClass
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a base class for shared functionality for text objects on the surface.
''' At this time this is the GDIField and GDIText classes.
''' </summary>
''' -----------------------------------------------------------------------------
<Serializable()> _
Public MustInherit Class GDITextBaseClass
    Inherits GDIFilledShape

#Region "Local Fields"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The string to draw to the surface
    ''' </summary>
    ''' <remarks>This maps to sample text property for GDIFields and to text for GDIText.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Protected _Text As String
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The fond the text should be drawn with
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _Font As System.Drawing.Font

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The text alignment property of the text.  This specifies, given a bounds,
    ''' how to align text within the rectangular bounds.  This is only relevant when 
    ''' the bounds on text are fixed.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _Alignment As System.Drawing.StringAlignment = StringAlignment.Near

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A Boolean indicating whether the GDIText object has a fixed width.  If the 
    ''' object does have a fixed width, text wraps around the fixed width bounds and 
    ''' the string alignment property comes into play.  Otherwise, text is drawn straight 
    ''' from the top left horizontally until out of text to draw.

    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _DefinedWidth As Boolean


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to consolidate fonts for the  object with the
    ''' fonts of other objects.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _ConsolidateFont As Boolean
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to consolidate string formats for this text object with similar string formats 
    ''' in the document.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _ConsolidateFormat As Boolean

#End Region

#Region "non serialized fields"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A string format object used to render text at runtime.  The actual behavior and 
    ''' properties of this object are determined by settings.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Protected _StringFormat As StringFormat


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A series of character ranges used to hit test character positions when the 
    ''' Text tool is used.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Protected _RectCharRanges As New ArrayList(100)



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  'Whether the object is currently in the middle of a drag operation or not.
    ''' </summary>
    ''' <remarks> This is used to short circuit some of the more expensive 
    ''' functionality to give a smoother drag as text is being moved until
    '''  its done being positioned.
    ''' 
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Protected _InDrag As Boolean = False

#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new GDITextBaseClass object.
    ''' </summary>
    ''' <param name="Text">The string to draw</param>
    ''' <param name="Font">The font to draw the text with</param>
    ''' <param name="rectBounds">A rectangle that bounds the text.  Note that the text 
    ''' can extend vertically beyond the bounds of this rectangle.  </param>
    ''' <param name="wrap">Whether to wrap to bounds (fixed width)or extend horizontally</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal Text As String, ByVal Font As Font, ByVal rectBounds As Rectangle, ByVal wrap As Boolean)
        MyBase.New()

        Trace.WriteLineIf(Session.Tracing, "Text.New")
        Trace.WriteLineIf(Session.Tracing, "Text:" & Text)
        Trace.WriteLineIf(Session.Tracing, "Font:" & Font.ToString)
        Trace.WriteLineIf(Session.Tracing, "rectBounds:" & rectBounds.ToString)


        _Text = Text
        _ConsolidateFont = Session.Settings.ConsolidateFonts
        _ConsolidateFormat = Session.Settings.ConsolidateStringFormats

        _Font = DirectCast(Font.Clone(), Font)
        _Bounds = rectBounds
        _DefinedWidth = wrap
        setupFormat()
        UpdateCharRanges()
    End Sub

#End Region

#Region " Property Accessors"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the point at which text is rotated.
    ''' </summary>
    ''' <value>A System.Drawing.PointF about which text is rotated</value>
    ''' <remarks>
    ''' Whereas for most GDIObjects, the rotation point is simply the center point of the object, 
    ''' this point depends on whether the wrap mode is set to fixed width or 
    ''' variable width.  If it's fixed width, it is handled it 
    ''' like every other GDIObject, rotating from the center.  If it's variable width, 
    ''' the text is rotated from its top left point instead.</remarks>
    ''' -----------------------------------------------------------------------------
    Public Overrides ReadOnly Property RotationPoint() As System.Drawing.PointF
        Get
            If _DefinedWidth Then
                Return MyBase.RotationPoint
            Else
                Return New PointF(Me.Bounds.X, Me.Bounds.Y)
            End If
        End Get
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets a value indicating if Text should wrap.  Wrap is a friendlier way to say 
    ''' restrict to a fixed width bounds.
    ''' </summary>
    ''' <value>A Boolean indicating whether to wrap.</value>
    ''' <remarks>On true, defaults to alignment "near".
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    <DefaultValue(False)> _
 Public Property Wrap() As Boolean
        Get
            Return _DefinedWidth
        End Get
        Set(ByVal Value As Boolean)
            Trace.WriteLineIf(Session.Tracing, "Text.Wrap.Set: " & Value)

            _DefinedWidth = Value

            If Not _DefinedWidth Then
                Alignment = StringAlignment.Near
                calcbounds()
            Else

            End If


        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or set the rectangle that bounds the text object.  If this property is Set, 
    ''' it will change a variable length text block to a fixed width text block.
    ''' </summary>
    ''' <value>The new bounds of the text inheritor specified as a rectangle.</value>
    ''' <remarks>Among other things, when set, this updates the character range array used to 
    ''' highlight text on the surface.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Overrides Property Bounds() As Rectangle
        Get
            Return _Bounds
        End Get
        Set(ByVal Value As Rectangle)
            Trace.WriteLineIf(Session.Tracing, "Text.Bounds.Set: " & Value.ToString)

            'If the rectangle got a new width
            If _Bounds.Width <> CInt(Value.Width) Then
                'go to fixed width mode
                _DefinedWidth = True
            End If

            If _DefinedWidth Then
                _Bounds = Value
            Else
                'recalc the bounds
                _Bounds.X = Value.X
                _Bounds.Y = Value.Y

                calcbounds()
                'else just assign it to the new value

            End If

            If Not _InDrag Then
                UpdateCharRanges()
            End If

            _Fill.OnParentUpdated(Me)

        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines whether to attempt to consolidate fonts at export time or not for this 
    ''' object with other objects.
    ''' </summary>
    ''' <value>A Boolean indicating whether to export fonts or not.</value>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.Description("Whether, on export, to consolidate this font with like text fonts.  See help for details")> _
    Public Property ConsolidateFont() As Boolean
        Get
            Return _ConsolidateFont
        End Get
        Set(ByVal Value As Boolean)
            Trace.WriteLineIf(Session.Tracing, "Text.ConsolidateFont.Set: " & Value.ToString)

            _ConsolidateFont = Value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines whether to attempt to consolidate string formats at export time or not for this 
    ''' object with other objects.
    ''' </summary>
    ''' <value>A Boolean indicating whether to export string formats or not.</value>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.Description("Whether, on export, to consolidate the string format with like text formats.  Only applies to fixed width text.  See help for details")> _
    Public Property ConsolidateStringFormat() As Boolean
        Get
            Return _ConsolidateFormat
        End Get
        Set(ByVal Value As Boolean)
            Trace.WriteLineIf(Session.Tracing, "Text.ConsolidateStringFormat.Set: " & Value.ToString)

            _ConsolidateFormat = Value
        End Set
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the string of text that is drawn with this GDIText object.
    ''' </summary>
    ''' <value>The string to draw.</value>
    ''' <remarks>Notice that this forces a recalc of the bounds of the text object 
    ''' when the property is Set.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.Description("The text to be drawn"), ComponentModel.DefaultValue(GetType(String), "")> _
  Public Property Text() As String
        Get
            Return _Text
        End Get
        Set(ByVal Value As String)
            Trace.WriteLineIf(Session.Tracing, "Text.Text.Set: " & Value.ToString)

            _Text = Value
            calcbounds()
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the string format used to draw the GDIText object.
    ''' </summary>
    ''' <value>A StringFormat object used to render the GDIText object.</value>
    ''' -----------------------------------------------------------------------------
    '''  
    <Browsable(False)> _
     Friend ReadOnly Property StringFormat() As StringFormat
        Get
            Return _StringFormat
        End Get

    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the stroke used to stroke text, which is always nothing (null) since 
    ''' text is only filled.  The property is here in order to make it not browsable inside 
    ''' the property grid.
    ''' </summary>
    ''' <value>A GDIStroke which is always nothing.</value>
    ''' ----------------------------------------------------------------------------- 
    <Browsable(False)> _
          Public Overrides Property Stroke() As GDIStroke
        Get
            Return Nothing
        End Get
        Set(ByVal Value As GDIStroke)

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the alignment of text.  This property only has a purpose with fixed 
    ''' width text.  Otherwise, text is always aligned top left.
    ''' </summary>
    ''' <value>A valid StringAlignment</value>
    ''' -----------------------------------------------------------------------------
    <ComponentModel.DefaultValue(0)> _
        Public Property Alignment() As StringAlignment
        Get
            Return _Alignment
        End Get
        Set(ByVal Value As StringAlignment)
            Trace.WriteLineIf(Session.Tracing, "Text.Alignment.Set: " & Value.ToString)

            If _DefinedWidth Then
                _Alignment = Value
                _StringFormat.Alignment = _Alignment
            Else
                _Alignment = StringAlignment.Near
                _StringFormat.Alignment = StringAlignment.Near
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the font to draw the text with. 
    ''' </summary>
    ''' <value>The Font to draw text with.</value>
    ''' <remarks>Notice that changing the font causes bounds to recalculate as well as 
    ''' raising a fill updated event to interested listeners.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    <Description("The Font used to draw the text")> _
 Public Property Font() As Font
        Get
            Return _Font
        End Get
        Set(ByVal Value As Font)
            Trace.WriteLineIf(Session.Tracing, "Text.Font.Set: " & Value.ToString)

            _Font = DirectCast(Value.Clone, Font)
            calcbounds()

            If Not _Fill Is Nothing Then
                _Fill.OnParentUpdated(Me)
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to draw text "filled".  This is always true, since if it was false it 
    ''' could leave blank text on the surface. To enforce this, it is marked with the 
    ''' Browsable(False) attribute.
    ''' </summary>
    ''' <value>A Boolean indicating whether to fill the text or not.</value>
    ''' -----------------------------------------------------------------------------
    ''' 
    <Browsable(False)> _
    Public Overrides Property DrawFill() As Boolean
        Get
            Return MyBase.DrawFill
        End Get
        Set(ByVal Value As Boolean)

            _DrawFill = True
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether to draw text stroked or not, which is irrelevant since 
    ''' the text objects do not have a stroke.  To enforce this, it is marked with the 
    ''' Browsable(False) attribute.
    ''' </summary>
    ''' <value>A Boolean indicating whether to stroke text.</value>
    ''' -----------------------------------------------------------------------------
    <Browsable(False)> _
        Public Overrides Property DrawStroke() As Boolean
        Get
            Return False
        End Get
        Set(ByVal Value As Boolean)
            _DrawStroke = False
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the length of the string drawn to the surface as an int32.
    ''' </summary>
    ''' <value>An integer containing the length of the string.</value>
    ''' -----------------------------------------------------------------------------
    <Browsable(False)> _
    Public ReadOnly Property Length() As Int32
        Get
            Return _Text.Length
        End Get
    End Property



#End Region

#Region "Base class Implementers"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deserializes a GDIText object for a specific document on a specific page.
    ''' </summary>
    ''' <param name="doc">The GDIdocument the GDIText exists on.</param>
    ''' <param name="pg">The GDIPage the GDIText exists on.</param>
    ''' <returns>True.  Deserialization is always expected to be successful 
    ''' for text inheritors.</returns>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Function deserialize(ByVal doc As GDIDocument, ByVal pg As GDIPage) As Boolean
        Trace.WriteLineIf(Session.Tracing, "Text.Deserialize:")

        MyBase.deserialize(doc, pg)
        _RectCharRanges = New ArrayList(100)
        setupFormat()
        Wrap = _DefinedWidth
        UpdateCharRanges()

        Return True
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a name used when displaying the class in the property browser.
    ''' </summary>
    ''' <value>A string containing the word "Text Object"</value>
    ''' -----------------------------------------------------------------------------   
    Public Overrides ReadOnly Property ClassName() As String
        Get
            Return "Text Object"
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Overrided to shadow functionality.  Text objects do not use graphics paths to 
    ''' render their content.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected Overrides Sub createPath()
        '   MyBase.resetPath()
    End Sub

#End Region


#Region "Drag related functionality"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Ends a drag operation on GDIText.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub endDrag()
        _InDrag = False
        UpdateCharRanges()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Starts a new drag operation on GDIText.
    ''' </summary>
    ''' <param name="ptObject">The initial mouse point dragging started from.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub startDrag(ByVal ptObject As Point)
        _InDrag = True
        _DragoffSet = New Point(Bounds.X - ptObject.X, Bounds.Y - ptObject.Y)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the position of the GDIText object during a drag operation.
    ''' </summary>
    ''' <param name="dragPoint">The last mouse position where dragging occurred at.</param>
    ''' -----------------------------------------------------------------------------
    Friend Overrides Sub updateDrag(ByVal dragPoint As Point)

        dragPoint.Offset(_DragoffSet.X, _DragoffSet.Y)
        Me.Bounds = New Rectangle(dragPoint.X, dragPoint.Y, Bounds.Width, Bounds.Height)

    End Sub

#End Region


#Region "Code generation related methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a font creation statement to code.
    ''' </summary>
    ''' <param name="exportSettings">Current export settings for the document.</param>
    ''' <returns>A CodeObjectCreateExpression that constructs the font.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function getFontInitializer(ByVal exportSettings As ExportSettings) As CodeObjectCreateExpression
        Dim fontinit As New CodeObjectCreateExpression

        Dim fontcast As New CodeDom.CodeCastExpression

        Trace.WriteLineIf(Session.Tracing, "TextBaseClass.getFontInitializer")
        fontcast.TargetType = New CodeTypeReference(GetType(System.Drawing.FontStyle))
        fontcast.Expression = New CodePrimitiveExpression(CInt(Me.Font.Style))

        Dim fSize As Single

        If exportSettings.DocumentType = EnumDocumentTypes.ePrintDocument Then
            If _Font.Unit = GraphicsUnit.Pixel OrElse _Font.Unit = GraphicsUnit.World Then
                fSize = _Font.Size

                'Have to convert to appropriate units for printing DPI
            Else
                fSize = _Font.Size * (Session.Settings.DPIY / 100)
            End If

        Else
            fSize = _Font.Size
        End If


        With fontinit
            .CreateType = New CodeTypeReference(GetType(System.Drawing.Font))
            .Parameters.Add(New CodePrimitiveExpression(Me.Font.Name))
            .Parameters.Add(New CodePrimitiveExpression(fSize))
            .Parameters.Add(fontcast)
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(System.Drawing.GraphicsUnit)), _Font.Unit.ToString))
            .Parameters.Add(New CodePrimitiveExpression(Me._Font.GdiCharSet))
        End With

        Return fontinit
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the Font Disposal statement for code emits.
    ''' </summary>
    ''' <param name="fontdeclare">The name of the font declared in code.</param>
    ''' <returns>A CodeMethodInvokeExpression that disposes the font.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function getFontDisposal(ByVal fontdeclare As String) As CodeMethodInvokeExpression

        Dim disposeFont As New CodeMethodInvokeExpression

        With disposeFont
            .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, fontdeclare)
            .Method.MethodName = "Dispose"
        End With

        Return disposeFont
    End Function

#End Region

#Region "Formatting"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles getting the bounds on variable length GDIText objects.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub setVariableBoundingRect()
        Dim rightmost As Single = 0
        Dim bottommost As Single = 0

        Dim newlinefound As Boolean = False

        For i As Int32 = 0 To _RectCharRanges.Count - 1

            Dim rect As RectangleF = DirectCast(_RectCharRanges(i), RectangleF)
            If rect.Bottom > bottommost Then
                bottommost = rect.Bottom
            End If
            If rect.Right > rightmost Then
                rightmost = rect.Right
            End If

        Next

        _Bounds = New Rectangle(Bounds.X, Bounds.Y, _
        CInt(Math.Ceiling(rightmost)) - Bounds.X, CInt(Math.Ceiling(bottommost)) - Bounds.Y)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Recalculates the bounding box as needed as well as updates the character 
    ''' ranges used when highlighting characters on the surface.
    ''' </summary>
    ''' <remarks>MeasureString has some issues, but it seems to work ok for the majority of 
    ''' situations.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Sub calcbounds()
        Dim tempstring As String = _Text
        Dim g As Graphics = Session.GraphicsManager.getTempGraphics()

        If Not _DefinedWidth Then

            Dim baseSize As SizeF = g.MeasureString(Me.Text, Me.Font, New PointF(Me.Bounds.X, Me.Bounds.Y), _StringFormat)

            UpdateCharRanges()
            setVariableBoundingRect()
        Else

            UpdateCharRanges()
        End If

        g.Dispose()
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Builds the initial StringFormat object used to render text to the drawing surface.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Sub setupFormat()
        _StringFormat = DirectCast(StringFormat.GenericTypographic.Clone(), StringFormat)

        _StringFormat.FormatFlags = StringFormatFlags.MeasureTrailingSpaces
        _StringFormat.Trimming = StringTrimming.None
        _StringFormat.Alignment = _Alignment

    End Sub

#End Region


#Region "Character sizing and position related methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a series of rectangleF structures that bound the characters from miTextStart index to 
    ''' miTextEnd 
    ''' </summary>
    ''' <param name="miTextStart">The position to begin bounding characters at.</param>
    ''' <param name="miTextEnd">The position to end bounding characters at.</param>
    ''' <returns>An array holding a series of rectangles that bound the selected characters.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function getSelectedCharRanges(ByVal miTextStart As Int32, ByVal miTextEnd As Int32) As RectangleF()

        If miTextStart > _Text.Length - 1 Then
            miTextStart = _Text.Length - 1
        End If

        If miTextStart > miTextEnd Then
            Dim temppos As Int32 = miTextStart
            miTextStart = miTextEnd
            miTextEnd = temppos
        End If

        Dim rectSelChars(miTextEnd - miTextStart) As RectangleF


        For i As Int32 = 0 To miTextEnd - miTextStart - 1
            rectSelChars(i) = DirectCast(_RectCharRanges(i + miTextStart), RectangleF)
        Next

        Return rectSelChars

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a bounding rectangle for a character at position miTextIndex within the 
    ''' string.  This bounding rectangle is used to highlight that character at the given 
    ''' position.
    ''' </summary>
    ''' <param name="miTextIndex">The index of the character to retrieve a bounding rectangle 
    ''' for.</param>
    ''' <returns>A rectangle bounding the character at the specific position.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function getRectangleForCharPos(ByVal miTextIndex As Int32) As RectangleF
        If _RectCharRanges.Count > 0 Then
            Return DirectCast(_RectCharRanges(miTextIndex), RectangleF)
        Else
            Return RectangleF.Empty
        End If
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Checks whether a character in the drawn string intersects with a specific point.
    ''' If it does, returns the index of the character at the point.
    ''' </summary>
    ''' <param name="pt">The pointF at which to test for a character at.</param>
    ''' <returns>The index of the character at the point, the last character if the structure 
    ''' is not empty, or 0 if there are no character ranges. </returns>
    ''' -----------------------------------------------------------------------------
    Public Function CharIndexAtPoint(ByVal pt As PointF) As Int32
        If _RectCharRanges.Count > 0 Then

            For i As Int32 = 0 To _Text.Length - 1
                '_RectCharRanges is an arraylist of rectangleF structures.
                If CType(_RectCharRanges(i), RectangleF).Contains(pt) Then
                    Return i
                End If
            Next

            Return _Text.Length - 1

        Else
            Return 0
        End If
    End Function



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Update the array of character ranges that are used when highlighting and 
    ''' manipulating individual characters on the drawing surface.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected Sub UpdateCharRanges()
        Dim regions() As Region

        'assume all chars visible
        _RectCharRanges.Clear()

        Dim charRanges() As CharacterRange

        Dim g As Graphics = Session.GraphicsManager.getTempGraphics

        Dim iCharPos As Int32 = 0
        Dim tempbounds As RectangleF
        If _DefinedWidth Then
            tempbounds = RectangleF.op_Implicit(_Bounds)
            tempbounds.Inflate(0, 1)
        Else
            tempbounds = New RectangleF(_Bounds.X, _Bounds.Y, Int32.MaxValue, Int32.MaxValue)
        End If


        For i As Int32 = 0 To _Text.Length - 1 Step 30
            Dim iRemainder As Int32 = _Text.Length - i

            ReDim charRanges(Math.Min(30, iRemainder) - 1)

            'populate current set of char ranges from i to 30 
            For j As Int32 = 0 To (Math.Min(30, iRemainder) - 1)
                charRanges(j).First = j + i
                charRanges(j).Length = 1
            Next

            _StringFormat.SetMeasurableCharacterRanges(charRanges)
            regions = g.MeasureCharacterRanges(_Text, _Font, tempbounds, _StringFormat)

            'populate current set of char ranges
            For j As Int32 = 0 To Math.Min(30, iRemainder) - 1
                Dim rectf As RectangleF = regions(j).GetBounds(g)

                _RectCharRanges.Add(rectf)
                regions(j).Dispose()
            Next

            If i >= _Text.Length - 1 Then
                Exit For
            End If

        Next

        g.Dispose()


    End Sub


#End Region


End Class
