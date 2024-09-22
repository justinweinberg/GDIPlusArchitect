Imports System.CodeDom

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : ExportPrint
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Responsible for emitting a GDI+ Architect PrintDocument to code.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class ExportPrint
    Inherits BaseExport


#Region "Local fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Method where graphic resources are initialized.  This is similar to the 
    ''' windows form's InitializeComponent method.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _initGraphics As CodeMemberMethod
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposal section of the outgoing print document.  Allows for disposing of 
    ''' custom graphic resources created at the field level of the print document.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _disposeGraphics As CodeMemberMethod

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' class level member declaration section.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Members As CodeTypeMemberCollection

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Used to call the custom  DisposeGraphics method by overriding the default dispose
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _InheritedDispose As CodeMemberMethod
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Used to override the onPrintPage event so the custom printing can be called
    ''' exporting each page in turn.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _PrintOverride As CodeMemberMethod
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Used to produce the actual graphics code for print page events.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _PrintGraphics As CodeMemberMethod
#End Region


#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the class given a specific document to export
    ''' </summary>
    ''' <param name="doc">The GDIDocument to export</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal doc As GDIDocument)
        MyBase.New(doc)
    End Sub
#End Region

#Region "Code emit related functionality"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Builds the actual class that holds all of the print related code.
    ''' </summary>
    ''' <returns>A CodeTypeDeclaration to which the class will add methods and members.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function CreatePRINTDeclaration() As CodeDom.CodeTypeDeclaration
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.CreatePrintDeclaration")

        Dim PRINTDeclaration As New CodeTypeDeclaration

        Dim PRINTConstruct As New CodeConstructor
        Dim invokeInitGraphics As New CodeMethodInvokeExpression
        PRINTConstruct.Attributes = MemberAttributes.Public

        invokeInitGraphics.Method = New CodeMethodReferenceExpression(New CodeThisReferenceExpression, "InitializeGraphics")
        PRINTConstruct.Statements.Add(invokeInitGraphics)

        With PRINTDeclaration
            .BaseTypes.Add(New CodeTypeReference(GetType(System.Drawing.Printing.PrintDocument)))
            .IsClass = True
            .Attributes = MemberAttributes.Public
            .Members.Add(PRINTConstruct)
            .Comments.Add(New CodeCommentStatement("GDIPlus Architect PrintDocument Class Output"))
            .Name = _Document.ExportSettings.ClassName
        End With

        Return PRINTDeclaration

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a Boolean statement capable of comparing the currentPage variable 
    ''' for equality to a specific page (the page passed in to this method) 
    ''' </summary>
    ''' <param name="pg">The page to check for equality against.</param>
    ''' <returns>A statement that checks if the current page matches the page number.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function createPageConditional(ByVal pg As GDIPage) As CodeConditionStatement
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.createPageConditional")

        Dim bConditional As New CodeConditionStatement
        bConditional.Condition = New CodeBinaryOperatorExpression( _
        New CodeFieldReferenceExpression( _
        New CodeThisReferenceExpression, "currentPage"), _
        CodeBinaryOperatorType.ValueEquality, New CodePrimitiveExpression(pg.PageNum))

        Return bConditional
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits all of the code required to render a specific page to the print document.
    ''' </summary>
     ''' <param name="pg">The page to emit</param>
    ''' -----------------------------------------------------------------------------
    Private Sub createPage(ByVal pg As GDIPage)
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.createPage")

        Dim conditionalPrint As CodeConditionStatement = createPageConditional(pg)

        For Each obj As GDIObject In pg.GDIObjects
            obj.emit(_Members, _initGraphics, conditionalPrint.TrueStatements, _disposeGraphics, _ExportSettings, _Consolidated)
        Next

        _PrintGraphics.Statements.Add(conditionalPrint)
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits the OnPrintPage override needed for GDI+ Architect's method of handling pages.
    ''' </summary>
    ''' <returns>The OnPrintPage method </returns>
    ''' -----------------------------------------------------------------------------
    Private Function createPrintpageOverride() As CodeMemberMethod
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.createPrintpageOverride")

        Dim mbrPrintPageOverride As New CodeMemberMethod
        Dim graphicsdeclare As CodeVariableDeclarationStatement = New CodeVariableDeclarationStatement(GetType(System.Drawing.Graphics), "g")

        Dim argInvoke As New CodeFieldReferenceExpression(New CodeArgumentReferenceExpression("e"), "Graphics")

        graphicsdeclare.InitExpression = argInvoke

        Dim assignMorePagesTrue As New CodeAssignStatement(New CodeFieldReferenceExpression(New CodeArgumentReferenceExpression("e"), "HasMorePages"), New CodePrimitiveExpression(True))
        Dim assignMorePagesFalse As New CodeAssignStatement(New CodeFieldReferenceExpression(New CodeArgumentReferenceExpression("e"), "HasMorePages"), New CodePrimitiveExpression(False))
        Dim x As CodeBinaryOperatorExpression

        Dim incrementStatement As New CodeBinaryOperatorExpression(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, "currentPage"), CodeBinaryOperatorType.Add, New CodePrimitiveExpression(1))
        Dim assignPageIncrement As New CodeAssignStatement(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, "currentPage"), incrementStatement)

        Dim printConditional As New CodeConditionStatement

        Dim callPrintGraphics As New CodeMethodInvokeExpression

        With callPrintGraphics
            .Method = New CodeMethodReferenceExpression(New CodeThisReferenceExpression, "_PrintGraphics")
            .Parameters.Add(New CodeVariableReferenceExpression("g"))
        End With

        printConditional.Condition = New CodeBinaryOperatorExpression(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, "currentPage"), CodeBinaryOperatorType.LessThan, New CodeFieldReferenceExpression(New CodeThisReferenceExpression, "totalPages"))

        With printConditional.TrueStatements
            .Add(assignMorePagesTrue)
            .Add(assignPageIncrement)
        End With

        printConditional.FalseStatements.Add(assignMorePagesFalse)

        With mbrPrintPageOverride
            .Name = "OnPrintPage"
            .ReturnType = Nothing
            .Parameters.Add(New CodeParameterDeclarationExpression(GetType(System.Drawing.Printing.PrintPageEventArgs), "e"))
            .Attributes = MemberAttributes.Family Or MemberAttributes.Override
            .Statements.Add(graphicsdeclare)
            .Statements.Add(callPrintGraphics)
            .Statements.Add(printConditional)
        End With

        Return mbrPrintPageOverride
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a series of counters to track the current page and the total pages
    ''' </summary>
    ''' <param name="decl">Declarations section of the class to place these counters in.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub createPrintCounters(ByVal decl As CodeTypeMemberCollection)
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.createPrintCounters")

        Dim curpagestatement As New CodeMemberField(GetType(Int32), "currentPage")
        curpagestatement.Comments.Add(New CodeCommentStatement("Current page"))

        Dim totalpagestatement As New CodeMemberField(GetType(Int32), "totalPages")
        totalpagestatement.Comments.Add(New CodeCommentStatement("Total pages"))

        curpagestatement.InitExpression = New CodePrimitiveExpression(1)
        totalpagestatement.InitExpression = New CodePrimitiveExpression(_Document.Count)

        decl.Add(curpagestatement)
        decl.Add(totalpagestatement)

    End Sub

    'Public Sub _PrintGraphics(ByVal g As System.Drawing.Graphics, ByVal pagenum As Int32)
    '    g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
    '    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Generates the _PrintGraphics method to which the actual graphic code is appended.
    ''' </summary>
    ''' <returns>A method capable of rendering all of the graphics.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function CreatePrintGraphicsMember() As CodeMemberMethod
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.CreatePrintGraphicsMember")

        Dim mbrPrintGraphics As New CodeMemberMethod
        mbrPrintGraphics.Attributes = MemberAttributes.Private
        mbrPrintGraphics.Name = "_PrintGraphics"
        mbrPrintGraphics.Parameters.Add(New CodeParameterDeclarationExpression(GetType(System.Drawing.Graphics), "g"))

        mbrPrintGraphics.Statements.Add(MyBase.createSmoothingModeAssignment)
        mbrPrintGraphics.Statements.Add(MyBase.createTextHintAssignment)

        Return mbrPrintGraphics
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Main entry point for generating a print document.
    ''' </summary>
    ''' <returns>A string containing all of the code required to generate a print document.</returns>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function Generate() As String
        Trace.WriteLineIf(Session.Tracing, "ExportConslidated.GeneratePrintDocument")

        Dim sw As IO.StringWriter = New IO.StringWriter
        Dim codegen As System.CodeDom.Compiler.ICodeGenerator = getGenerator(sw)
        Dim cop As New System.CodeDom.Compiler.CodeGeneratorOptions

        sw.WriteLine()

        Dim classDeclaration As CodeTypeDeclaration = Me.CreatePRINTDeclaration


        _Members = classDeclaration.Members
        _initGraphics = MyBase.CreateInitGraphics()
        _disposeGraphics = MyBase.CreateDisposeGraphics()
        _InheritedDispose = MyBase.CreateInheritedDispose
        _PrintOverride = createPrintpageOverride()
        createPrintCounters(classDeclaration.Members)
        _PrintGraphics = CreatePrintGraphicsMember()

        Dim getAbsoluteImage As CodeMemberMethod = MyBase.CreateFromAbsolute
        Dim getRelativeImage As CodeMemberMethod = MyBase.CreateFromRelative
        Dim getEmbeddedImage As CodeMemberMethod = MyBase.CreateFromEmbedded

        With classDeclaration.Members
            .Add(_initGraphics)
            .Add(_PrintOverride)
            .Add(_disposeGraphics)
            .Add(_InheritedDispose)
            .Add(_PrintGraphics)
            If MyBase._AbsoluteMethodNeeded Then
                .Add(getAbsoluteImage)
            End If

            If MyBase._RelativeMethodNeeded Then
                .Add(getRelativeImage)
            End If

            If MyBase._EmbedMethodNeeded Then
                .Add(getEmbeddedImage)
            End If

            If MyBase._RotateMethodNeeded Then
                .Add(CreateRotateStart)
                .Add(CreateRotateEnd)
                CreateRotateLocals(_Members)
            End If
        End With

        _Consolidated.emit(_Members, _initGraphics, _PrintGraphics.Statements, _disposeGraphics, _ExportSettings)

        For Each pg As GDIPage In _Document
            createPage(pg)
        Next


        codegen.GenerateCodeFromType(classDeclaration, sw, cop)
        sw.Close()

        Return sw.ToString

    End Function

#End Region
End Class
