Imports System.CodeDom




''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : ExportConsolidate
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Responsible for handling consolidation scenarios. 
''' </summary>
''' <remarks>
''' Consolidation is when a single declarationcan be created instead of multiple declarations,
'''  thus creating less code.  For example, two objects on the surface may be filled 
''' with a solid blue brush.  Instead of emitting this declaration twice, the user 
''' has the option to consolidate these objects into a single brush declaration.
''' </remarks>
''' -----------------------------------------------------------------------------
Public Class ExportConsolidate
    Inherits System.ComponentModel.Component

#Region "Type Declarations"
    ''' -----------------------------------------------------------------------------
    ''' Project	 : GDIObjects
    ''' Class	 : ExportConsolidate.ConsolidatedItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A simple wrapper for consolidated objects.  Contains the common name that will be 
    ''' emitted for consolidated items, an item which is valid as an emit target for the 
    ''' consolidated set, and a matches count.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Class ConsolidatedItem
    ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Name that will be emitted for the set of consolidated items (Brush1, etc)
        ''' </summary>
            ''' -----------------------------------------------------------------------------
        Public Name As String
    ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' A reference to one of the identical items in the consolidation set used to 
        ''' generate code.
        ''' </summary>
            ''' -----------------------------------------------------------------------------
        Public item As Object
    ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Total number of matches.
        ''' </summary>
            ''' -----------------------------------------------------------------------------
        Public matches As Int32 = 1

    ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Creates a new instance of a ConsolidatedItem.
        ''' </summary>
        ''' <param name="oitem">A key object that is identical to all objects in that this item 
        ''' will consolidate.</param>
        ''' <param name="sname">The name that will be used when exporting to code for all 
        ''' items in the set.</param>
            ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal oitem As Object, ByVal sname As String)
            Name = sname
            item = oitem
        End Sub


    End Class

#End Region


#Region "Local Fields"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''  Next key to use for consolidated fills
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _NextFillID As Int32 = 1

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Next key to use for consolidated strokes
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _NextStrokeID As Int32 = 1
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Next key to use for consolidated Fonts
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _NextFontID As Int32 = 1

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Next key to use for consolidated formats
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _NextFormatID As Int32 = 1



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An arraylist containing common (consolidated) fills
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Fills As New ArrayList(10)
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An arraylist containing common (consolidated) stokes
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Strokes As New ArrayList(10)
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An arraylist containing common (consolidated) fonts
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Fonts As New ArrayList(10)
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An arraylist containing common (consolidated) string format objects.
    ''' String formats are used to render text to the drawing surface
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Formats As New ArrayList(10)
#End Region


#Region "Code Generation"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a color assignment statement for a specific color value.
    ''' </summary>
    ''' <param name="val">the color to get a statement for.</param>
    ''' <returns>A CodeExpression which describes the color assignment.</returns>
    ''' <remarks>Notice this code attempts to use named colors first.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Function getColorAssignment(ByVal val As Color) As CodeExpression
        Dim outparamInvoke As New CodeDom.CodeMethodInvokeExpression
        Dim simpleOut As New CodeDom.CodeFieldReferenceExpression

        If val.IsNamedColor AndAlso val.A = 255 Then
            With simpleOut
                .TargetObject = New CodeTypeReferenceExpression(GetType(Drawing.Color))
                .FieldName = val.Name
                Return simpleOut
            End With

        Else
            With outparamInvoke
                .Method.TargetObject = New CodeTypeReferenceExpression(GetType(Drawing.Color))
                .Method.MethodName = "FromArgb"
                .Parameters.Add(New CodePrimitiveExpression(val.A))
                .Parameters.Add(New CodePrimitiveExpression(val.R))
                .Parameters.Add(New CodePrimitiveExpression(val.G))
                .Parameters.Add(New CodePrimitiveExpression(val.B))
            End With

            Return outparamInvoke

        End If


    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a consolidated fill.
    ''' </summary>
    ''' <param name="sharedFill">The fill shared by the various consolidated objects.</param>
    ''' <param name="Declarations">Declarations section of the code class being created.</param>
    ''' <param name="InitGraphics">InitGraphics section of the code class being created.</param>
    ''' <param name="RenderGDI">The RenderGDI method of the code class being created.</param>
    ''' <param name="DisposeGDI">The DisposeGDI section of the code class being created.</param>
    ''' <param name="ExportSettings">The current export settings.</param>
    ''' <remarks>Notice that the consolidated item is asked to emit itself.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Sub emitFill(ByVal sharedFill As ConsolidatedItem, _
     ByVal Declarations As CodeDom.CodeTypeMemberCollection, _
    ByVal InitGraphics As CodeDom.CodeMemberMethod, _
    ByVal RenderGDI As CodeDom.CodeStatementCollection, _
    ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
    ByVal ExportSettings As ExportSettings)

        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.emitFill")

        Dim fill As GDIFill = DirectCast(sharedFill.item, GDIFill)
        fill.emit(sharedFill.Name, Declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings)

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a consolidated stroke.
    ''' </summary>
    ''' <param name="sharedStroke">The stroke shared by the various consolidated objects.</param>
    ''' <param name="Declarations">Declarations section of the code class being created.</param>
    ''' <param name="InitGraphics">InitGraphics section of the code class being created.</param>
    ''' <param name="RenderGDI">The RenderGDI method of the code class being created.</param>
    ''' <param name="DisposeGDI">The DisposeGDI section of the code class being created.</param>
    ''' <param name="ExportSettings">The current export settings.</param>
    ''' <remarks>All of the "IF" statements below are used to determine if the stroke is using 
    ''' default values.  If they are, these properties are not emitted explicitly.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Sub emitStroke(ByVal sharedStroke As ConsolidatedItem, _
    ByVal Declarations As CodeDom.CodeTypeMemberCollection, _
    ByVal InitGraphics As CodeDom.CodeMemberMethod, _
    ByVal RenderGDI As CodeDom.CodeStatementCollection, _
    ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
    ByVal ExportSettings As ExportSettings)

        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.emitStroke")


        Dim strk As GDIStroke = DirectCast(sharedStroke.item, GDIStroke)

        Dim declareStroke As New CodeMemberField
        Dim createStatement As New CodeObjectCreateExpression
        Dim disposeStroke As New CodeDom.CodeMethodInvokeExpression

        'Create the initialization args and assignment for a stroke
        With createStatement
            .CreateType = New CodeTypeReference(GetType(Pen))
            .Parameters.Add(getColorAssignment(strk.Color))
            .Parameters.Add(New CodePrimitiveExpression(strk.Width))
        End With

        'Create the member variable for the stroke
        With declareStroke
            .InitExpression = createStatement
            .Name = sharedStroke.Name
            .Attributes = MemberAttributes.Private
            .Type = New CodeTypeReference(GetType(System.Drawing.Pen))
        End With

        If Not strk.Alignment = CType(0, Drawing2D.PenAlignment) Then
            Dim AssignAlignment As New CodeAssignStatement
            AssignAlignment.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "Alignment")
            AssignAlignment.Right = New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(Drawing2D.PenAlignment)), strk.Alignment.ToString)
            InitGraphics.Statements.Add(AssignAlignment)
        End If

        If Not strk.DashCap = CType(0, Drawing2D.DashCap) Then
            Dim AssignDashCap As New CodeAssignStatement
            AssignDashCap.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "DashCap")
            AssignDashCap.Right = New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(Drawing2D.DashCap)), strk.DashCap.ToString)
            InitGraphics.Statements.Add(AssignDashCap)
        End If

        If Not strk.DashOffset = 0.0! Then
            Dim AssignDashOffset As New CodeAssignStatement

            AssignDashOffset.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "DashOffset")
            AssignDashOffset.Right = New CodePrimitiveExpression(strk.DashOffset)
            InitGraphics.Statements.Add(AssignDashOffset)
        End If


        If Not strk.DashStyle = CType(0, Drawing2D.DashStyle) Then
            Dim AssignDashStyle As New CodeAssignStatement
            AssignDashStyle.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "DashStyle")
            AssignDashStyle.Right = New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(Drawing2D.DashStyle)), strk.DashStyle.ToString)
            InitGraphics.Statements.Add(AssignDashStyle)
        End If

        If Not strk.Startcap = CType(0, Drawing2D.DashCap) Then
            Dim AssignStartcap As New CodeAssignStatement
            AssignStartcap.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "StartCap")
            AssignStartcap.Right = New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(Drawing2D.LineCap)), strk.Startcap.ToString)
            InitGraphics.Statements.Add(AssignStartcap)
        End If

        If Not strk.Endcap = CType(0, Drawing2D.DashCap) Then
            Dim AssignEndcap As New CodeAssignStatement
            AssignEndcap.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "EndCap")
            AssignEndcap.Right = New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(Drawing2D.LineCap)), strk.Endcap.ToString)
            InitGraphics.Statements.Add(AssignEndcap)
        End If

        If Not strk.LineJoin = CType(0, Drawing2D.LineJoin) Then
            Dim AssignJoin As New CodeAssignStatement
            AssignJoin.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "LineJoin")
            AssignJoin.Right = New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(Drawing2D.LineJoin)), strk.LineJoin.ToString)
            InitGraphics.Statements.Add(AssignJoin)
        End If

        If Not strk.MiterLimit = 10.0! Then
            Dim MiterLimit As New CodeAssignStatement
            MiterLimit.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression(declareStroke.Name), "MiterLimit")
            MiterLimit.Right = New CodePrimitiveExpression(strk.MiterLimit)
            InitGraphics.Statements.Add(MiterLimit)
        End If

        'create dispose call
        With disposeStroke
            .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, declareStroke.Name)
            .Method.MethodName = "Dispose"
        End With

        'add object disposal
        With DisposeGDI
            .Statements.Add(disposeStroke)
        End With

        With Declarations
            .Add(declareStroke)
        End With


    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a consolidated string format.
    ''' </summary>
    ''' <param name="sharedFormat">The string format shared by the consolidated objects.</param>
    ''' <param name="Declarations">Declarations section of the code class being created.</param>
    ''' <param name="InitGraphics">InitGraphics section of the code class being created.</param>
    ''' <param name="RenderGDI">The RenderGDI method of the code class being created.</param>
    ''' <param name="DisposeGDI">The DisposeGDI section of the code class being created.</param>
    ''' <param name="ExportSettings">The current export settings.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub emitFormat(ByVal sharedFormat As ConsolidatedItem, ByVal Declarations As CodeDom.CodeTypeMemberCollection, _
  ByVal InitGraphics As CodeDom.CodeMemberMethod, _
  ByVal RenderGDI As CodeDom.CodeStatementCollection, _
  ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
  ByVal ExportSettings As ExportSettings)

        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.emitFormat")

        Dim _DataFormat As StringFormat = DirectCast(sharedFormat.item, StringFormat)

        Dim formatDeclaration As New CodeMemberField
        Dim formatInit As New CodeObjectCreateExpression(GetType(System.Drawing.StringFormat))
        formatInit.Parameters.Add(New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(System.Drawing.StringFormat)), "GenericTypographic"))


        With formatDeclaration
            .Name = sharedFormat.Name
            .Type = New CodeTypeReference(GetType(System.Drawing.StringFormat))
            .Attributes = MemberAttributes.Private
            .InitExpression = formatInit
        End With
        declarations.Add(formatDeclaration)
        Dim formatDispose As New CodeMethodInvokeExpression


        'create dispose call
        With formatDispose
            .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, formatDeclaration.Name)
            .Method.MethodName = "Dispose"
        End With

        With DisposeGDI
            .Statements.Add(formatDispose)
        End With

        Dim sFormatName As String = formatDeclaration.Name


        With InitGraphics

            Dim AlignAssignment As New CodeDom.CodeAssignStatement

            AlignAssignment.Left = New CodeFieldReferenceExpression(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, sFormatName), "Alignment")
            AlignAssignment.Right = New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(System.Drawing.StringAlignment)), _DataFormat.Alignment.ToString)

            .Statements.Add(AlignAssignment)
        End With
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a statement that sets up a font.
    ''' </summary>
    ''' <param name="fnt">The font to emit </param>
    ''' <param name="exportSettings">Current export settings.</param>
    ''' <returns>An expression that creates a font.</returns>
    ''' <remarks>Notice that if the document is a print document,  some 
    ''' conversion is performed based on the DPI in order to get a matching display output font.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Function getFontInitializer(ByVal fnt As Font, ByVal exportSettings As ExportSettings) As CodeObjectCreateExpression
        Dim fontinit As New CodeObjectCreateExpression

        Dim fontcast As New CodeDom.CodeCastExpression
        fontcast.TargetType = New CodeTypeReference(GetType(System.Drawing.FontStyle))
        fontcast.Expression = New CodePrimitiveExpression(CInt(fnt.Style))

        Dim fSize As Single

        If exportSettings.DocumentType = EnumDocumentTypes.ePrintDocument Then
            If fnt.Unit = GraphicsUnit.Pixel OrElse fnt.Unit = GraphicsUnit.World Then
                fSize = fnt.Size

                'Have to convert to appropriate units for printing DPI
            Else
                fSize = fnt.Size * (Session.Settings.DPIY / 100)
            End If

        Else
            fSize = fnt.Size
        End If


        With fontinit
            .CreateType = New CodeTypeReference(GetType(System.Drawing.Font))
            .Parameters.Add(New CodePrimitiveExpression(fnt.Name))
            .Parameters.Add(New CodePrimitiveExpression(fSize))
            .Parameters.Add(fontcast)
            .Parameters.Add(New CodeFieldReferenceExpression(New CodeTypeReferenceExpression(GetType(System.Drawing.GraphicsUnit)), fnt.Unit.ToString))
            .Parameters.Add(New CodePrimitiveExpression(fnt.GdiCharSet))
        End With

        Return fontinit
    End Function



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Generates a font dispose statement.
    ''' </summary>
    ''' <param name="fontdeclare">The name of the font created</param>
    ''' <returns>An invocation that will dispose of the font.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function getFontDisposal(ByVal fontdeclare As String) As CodeMethodInvokeExpression

        Dim disposeFont As New CodeMethodInvokeExpression

        With disposeFont
            .Method.TargetObject = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, fontdeclare)
            .Method.MethodName = "Dispose"
        End With

        Return disposeFont
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a consolidated font.
    ''' </summary>
    ''' <param name="shareFormat">The font shared by the various consolidated objects.</param>
    ''' <param name="Declarations">Declarations section of the code class being created.</param>
    ''' <param name="InitGraphics">InitGraphics section of the code class being created.</param>
    ''' <param name="RenderGDI">The RenderGDI method of the code class being created.</param>
    ''' <param name="DisposeGDI">The DisposeGDI section of the code class being created.</param>
    ''' <param name="ExportSettings">The current export settings.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub emitFont(ByVal shareFormat As ConsolidatedItem, ByVal Declarations As CodeDom.CodeTypeMemberCollection, _
  ByVal InitGraphics As CodeDom.CodeMemberMethod, _
  ByVal RenderGDI As CodeDom.CodeStatementCollection, _
  ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
  ByVal ExportSettings As ExportSettings)

        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.emitFont")

        'Create the font object
        Dim fontDeclaration As New CodeMemberField

        With fontDeclaration
            .Name = shareFormat.Name
            .Type = New CodeTypeReference(GetType(System.Drawing.Font))
            .Attributes = MemberAttributes.Private
            .InitExpression = getFontInitializer(DirectCast(shareFormat.item, Font), ExportSettings)
        End With

        Declarations.Add(fontDeclaration)

        With DisposeGDI
            .Statements.Add(getFontDisposal(fontDeclaration.Name))
        End With
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Begins the process of exporting all consolidated resources to declarations
    ''' </summary>
    ''' <param name="Declarations">Declarations section of the code class being created.</param>
    ''' <param name="InitGraphics">InitGraphics section of the code class being created.</param>
    ''' <param name="RenderGDI">The RenderGDI method of the code class being created.</param>
    ''' <param name="DisposeGDI">The DisposeGDI section of the code class being created.</param>
    ''' <param name="ExportSettings">The current export settings.</param>
    ''' <remarks>When this method is called, each arraylist of consolidated items is examined 
    ''' for objects.  As objects are found, they are emitted.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub emit(ByVal Declarations As CodeDom.CodeTypeMemberCollection, _
  ByVal InitGraphics As CodeDom.CodeMemberMethod, _
  ByVal RenderGDI As CodeDom.CodeStatementCollection, _
  ByVal DisposeGDI As CodeDom.CodeMemberMethod, _
  ByVal ExportSettings As ExportSettings)

        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.emit")

        'emit shared fills
        For Each sharedFill As ConsolidatedItem In _Fills
            If sharedFill.matches > 1 Then
                emitFill(sharedFill, Declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings)
            End If
        Next

        'emit shared strokes
        For Each sharedStroke As ConsolidatedItem In _Strokes
            If sharedStroke.matches > 1 Then
                emitStroke(sharedStroke, Declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings)
            End If
        Next

        'emit shared fonts
        For Each sharedFont As ConsolidatedItem In _Fonts
            If sharedFont.matches > 1 Then
                emitFont(sharedFont, Declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings)
            End If
        Next


        'emit shared string formats
        For Each sharedFormat As ConsolidatedItem In _Formats
            If sharedFormat.matches > 1 Then
                emitFormat(sharedFormat, Declarations, InitGraphics, RenderGDI, DisposeGDI, ExportSettings)
            End If
        Next

    End Sub

#End Region


#Region "Consolidation Test and Implementation"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if a fill is in the consolidation set and adds the fill if it's not.
    ''' </summary>
    ''' <param name="fillToCheck">The fill to check and add</param>
    ''' -----------------------------------------------------------------------------
    Public Sub consolidateFill(ByVal fillToCheck As GDIFill)
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.consolidateFill")

        For Each consolidatedFill As ConsolidatedItem In _Fills
            If GDIFill.ExportEquality(DirectCast(consolidatedFill.item, GDIFill), fillToCheck) Then
                'We've already got it in the set
                consolidatedFill.matches += 1
                Return
            End If
        Next

        addConsolidatedFill(fillToCheck)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a fill to the set of consolidated fills.
    ''' </summary>
    ''' <param name="fill">The fill to add to the consolidated set.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub addConsolidatedFill(ByVal fill As GDIFill)
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.addConsolidatedFill")

        Dim sName As String = "brush" & _NextFillID
        _Fills.Add(New ConsolidatedItem(fill, sName))
        _NextFillID += 1
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a name to use when rendering a fill to code.  This method checks if 
    ''' the incoming fill is equal for exporting to one in the consolidation set and returns 
    ''' the equivalent name if it is.
    ''' </summary>
    ''' <param name="fillToCheck">The fill to check to see if it is consolidated</param>
    ''' <returns>A string value containing the name of the fill to use instead of the original name,
    ''' or an empty string if no consolidation was found.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function getFillName(ByVal fillToCheck As GDIFill) As String
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.getFillName")

        For Each consolidatedFill As ConsolidatedItem In _Fills
            If GDIFill.ExportEquality(DirectCast(consolidatedFill.item, GDIFill), fillToCheck) Then
                Return consolidatedFill.Name
            End If
        Next

        Return String.Empty
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if a given fill has a match in the consolidated fill set.
    ''' </summary>
    ''' <param name="fillToCheck">The fill to check for matches</param>
    ''' <returns>A Boolean indicating whether an equivalent fill has already been added to the set.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function hasFillMatch(ByVal fillToCheck As GDIFill) As Boolean
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.hasFillMatch")

        For Each consolidatedFill As ConsolidatedItem In _Fills
            If GDIFill.ExportEquality(DirectCast(consolidatedFill.item, GDIFill), fillToCheck) Then
                If consolidatedFill.matches > 1 Then
                    Return True
                Else
                    Return False
                End If
            End If
        Next

        Return False
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if a stroke is in the consolidated set or not.  If it is not,
    '''  adds it to the consolidated set.
    ''' </summary>
    ''' <param name="strokeToCheck">The stroke to check for participation in the consolidation 
    ''' set.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub consolidateStroke(ByVal strokeToCheck As GDIStroke)
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.consolidateStroke")

        For Each consolidatedStroke As ConsolidatedItem In _Strokes
            If GDIStroke.op_Equality(DirectCast(consolidatedStroke.item, GDIStroke), strokeToCheck) Then
                consolidatedStroke.matches += 1
                Return
            End If
        Next

        addConsolidatedStroke(strokeToCheck)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds as stroke to the consolidated set.
    ''' </summary>
    ''' <param name="stroke">The stroke to add to the set</param>
    ''' -----------------------------------------------------------------------------
    Private Sub addConsolidatedStroke(ByVal stroke As GDIStroke)
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.addConsolidatedStroke")

        Dim sName As String = "stroke" & _NextStrokeID
        _Strokes.Add(New ConsolidatedItem(stroke, sName))
        _NextStrokeID += 1
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the equivalent name to use when emitting this stroke or an empty string if no
    ''' equivalent is available.
    ''' </summary>
    ''' <param name="strokeToCheck">The stroke to check for consolidation.</param>
    ''' <returns>A string value containing the name of the stroke to use instead of the original name,
    ''' or an empty string if no consolidation was found.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function getStrokeName(ByVal strokeToCheck As GDIStroke) As String
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.getStrokeName")

        For Each consolidatedStroke As ConsolidatedItem In _Strokes
            If GDIStroke.op_Equality(DirectCast(consolidatedStroke.item, GDIStroke), strokeToCheck) Then
                Return consolidatedStroke.Name
            End If
        Next

        Return String.Empty
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determine if a given stroke has an equivalent in the consolidation set.
    ''' </summary>
    ''' <param name="strokeToCheck">The stroke to check for equivalents for</param>
    ''' <returns>A Boolean indicating whether an equivalent stroke was found.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function hasStrokeMatch(ByVal strokeToCheck As GDIStroke) As Boolean
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.hasStrokeMatch")

        For Each consolidatedStroke As ConsolidatedItem In _Strokes
            If GDIStroke.op_Equality(DirectCast(consolidatedStroke.item, GDIStroke), strokeToCheck) Then
                If consolidatedStroke.matches > 1 Then
                    Return True
                Else
                    Return False
                End If
            End If
        Next

        Return False
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if a font is in the consolidation set and adds it if it is not.
    ''' </summary>
    ''' <param name="fontToCheck">The font to check for consolidation</param>
    ''' -----------------------------------------------------------------------------
    Public Sub consolidateFont(ByVal fontToCheck As Font)
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.consolidateFont")

        For Each consolidatedFont As ConsolidatedItem In _Fonts

            If fontEquality(DirectCast(consolidatedFont.item, Font), fontToCheck) Then
                consolidatedFont.matches += 1
                Return
            End If
        Next

        addConsolidatedFont(fontToCheck)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a font to the consolidated font set.
    ''' </summary>
    ''' <param name="font">The font to add to the set</param>
    ''' -----------------------------------------------------------------------------
    Private Sub addConsolidatedFont(ByVal font As Font)
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.addConsolidatedFont")

        Dim sName As String = "font" & _NextFontID
        _Fonts.Add(New ConsolidatedItem(font, sName))
        _NextFontID += 1
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if a font has an equivalent in the consolidation set and returns the name 
    ''' of the consolidated object if it does.  Otherwise returns an empty string.
    ''' </summary>
    ''' <param name="fontToCheck">The font to check for equivalency in the set.</param>
    ''' <returns>A string containing the name of an equivalent font or an 
    ''' empty string if there is no equivalent.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function getFontName(ByVal fontToCheck As Font) As String
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.getFontName")

        For Each consolidatedFont As ConsolidatedItem In _Fonts

            If fontEquality(DirectCast(consolidatedFont.item, Font), fontToCheck) Then
                Return consolidatedFont.Name
            End If
        Next

        Return String.Empty
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if a font has a match in the consolidated set.
    ''' </summary>
    ''' <param name="fontToCheck">the font to check for matches for</param>
    ''' <returns>A Boolean indicating if the font has a consolidated match.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function hasFontMatch(ByVal fontToCheck As Font) As Boolean
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.hasFontMatch")

        For Each consolidatedFont As ConsolidatedItem In _Fonts

            If fontEquality(DirectCast(consolidatedFont.item, Font), fontToCheck) Then
                If consolidatedFont.matches > 1 Then
                    Return True
                Else
                    Return False
                End If
            End If
        Next

        Return False
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Checks if a given string format exists in the consolidated set and adds it to the 
    ''' consolidated set if it does not have a match.
    ''' </summary>
    ''' <param name="stringFormatToCheck">The string format to check for participation in 
    ''' the consolidation set.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub consolidateStringFormats(ByVal stringFormatToCheck As StringFormat)
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.consolidateStringFormats")

        For Each consolidatedStringFormat As ConsolidatedItem In _Formats

            If stringFormatEquality(DirectCast(consolidatedStringFormat.item, StringFormat), stringFormatToCheck) Then
                consolidatedStringFormat.matches += 1
                Return
            End If
        Next

        addConsolidatedStringFormat(stringFormatToCheck)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Adds a string format to the consolidated set.
    ''' </summary>
    ''' <param name="format">The string format to consolidate</param>
    ''' -----------------------------------------------------------------------------
    Private Sub addConsolidatedStringFormat(ByVal format As StringFormat)
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.addConsolidatedStringFormat")

        Dim sName As String = "stringformat" & _NextFormatID
        _Formats.Add(New ConsolidatedItem(format, sName))
        _NextFormatID += 1
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if a string format has an equivalent in the consolidation set and returns the name 
    ''' of the consolidated object if it does.  Otherwise returns an empty string.
    ''' </summary>
    ''' <param name="stringFormatToCheck">The string format to compare for consolidation</param>
    ''' <returns>The name to use if a match is found in the set.  An empty string otherwise.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function getStringFormatName(ByVal stringFormatToCheck As StringFormat) As String
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.getStringFormatName")

        For Each consolidatedStringFormat As ConsolidatedItem In _Formats

            If stringFormatEquality(DirectCast(consolidatedStringFormat.item, StringFormat), stringFormatToCheck) Then
                Return consolidatedStringFormat.Name
            End If
        Next

        Return String.Empty
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Checks whether a specific string format has a match in the consolidation set.
    ''' </summary>
    ''' <param name="stringFormatToCheck">The format to check for existence in the set</param>
    ''' <returns>A Boolean indicating whether the string format was found or not.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function hasStringFormatMatch(ByVal stringFormatToCheck As StringFormat) As Boolean
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.hasStringFormatMatch")

        For Each consolidatedStringFormat As ConsolidatedItem In _Formats

            If stringFormatEquality(DirectCast(consolidatedStringFormat.item, StringFormat), stringFormatToCheck) Then
                If consolidatedStringFormat.matches > 1 Then
                    Return True
                Else
                    Return False
                End If
            End If
        Next

        Return False
    End Function

#End Region


#Region "Consolidation Equality Testing"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if two string formats are equal.
    ''' </summary>
    ''' <param name="strFormat1">The first string format to compare.</param>
    ''' <param name="strFormat2">The second string format to compare</param>
    ''' <returns>A Boolean indicating whether, for our purposes, string formats are equal.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function stringFormatEquality(ByVal strFormat1 As StringFormat, ByVal strFormat2 As StringFormat) As Boolean
        If strFormat1.Alignment = strFormat2.Alignment Then
            Return True
        Else
            Return False
        End If

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines if two fonts are equal (have the same properties)
    ''' </summary>
    ''' <param name="font1">The first font to test for equality</param>
    ''' <param name="font2">The second font to test for equality</param>
    ''' <returns>A Boolean indicating if the two fonts are equal.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function fontEquality(ByVal font1 As Font, ByVal font2 As Font) As Boolean

        If font1.Name = font2.Name AndAlso font1.Size = font2.Size AndAlso _
        font1.Unit = font2.Unit AndAlso font1.GdiCharSet = font2.GdiCharSet AndAlso _
        font1.Style = font2.Style Then
            Return True
        Else
            Return False
        End If

    End Function

#End Region

#Region "Cleanup and Disposal"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the ExportConsolidate class
    ''' </summary>
    ''' <param name="disposing">Whether unmanaged resources are being disposed or not.</param>
    ''' -----------------------------------------------------------------------------
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            _Fills.Clear()
            _Strokes.Clear()
            _Fonts.Clear()
            _Formats.Clear()

            _Fills = Nothing
            _Strokes = Nothing
            _Fonts = Nothing
            _Formats = Nothing

        End If
    End Sub

#End Region


End Class
