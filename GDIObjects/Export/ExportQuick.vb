Imports System.CodeDom
Imports System.CodeDom.Compiler


''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : ExportQuick
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Responsible for generating GDI+ Architect's Quick Code (A subset of code for selected objects
''' on the drawing surface).
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class ExportQuick
    Inherits BaseExport

#Region "Local Fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The collection of objects that code is being generated for.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _GDIObjects As GDIObjCol
#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the QuickCode object given a parent document objects belong to, and the 
    ''' objects that are to be exported.
    ''' </summary>
    ''' <param name="doc">The document that the objects code generation belongs to is being created for.</param>
    ''' <param name="objects">The objects to generate quick code for.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal doc As GDIDocument, ByVal objects As GDIObjCol)
        MyBase.New(doc, objects)
        Trace.WriteLineIf(Session.Tracing, "ExportQuick.new")

        _GDIObjects = objects
    End Sub

#End Region

#Region "Code generation related functionality"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Main entry point for generating the quick code class.
    ''' </summary>
    ''' <returns>A string containing the quick code </returns>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function Generate() As String
        Trace.WriteLineIf(Session.Tracing, "ExportQuick.GenerateClassDocument")

        Dim sw As IO.StringWriter = New IO.StringWriter
        Dim codegen As System.CodeDom.Compiler.ICodeGenerator = MyBase.getGenerator(sw)
        Dim cop As New System.CodeDom.Compiler.CodeGeneratorOptions

        sw.WriteLine()

        Dim classDeclaration As CodeTypeDeclaration = Me.CreateCLASSDeclaration


        'These 4 items are used to construct the class.  The rest doesn't change based on the items in question
        Dim classvariables As CodeTypeMemberCollection = classDeclaration.Members


        Dim renderGraphics As CodeMemberMethod = CreateCLASSRenderGraphics()

        'Method creation
        Dim initGraphics As CodeMemberMethod = MyBase.CreateInitGraphics()
        Dim disposeGraphics As CodeMemberMethod = MyBase.CreateDisposeGraphics()
        Dim inheritedDispose As CodeMemberMethod = MyBase.CreateInheritedDispose()

        Dim getAbsoluteImage As CodeMemberMethod = MyBase.CreateFromAbsolute
        Dim getRelativeImage As CodeMemberMethod = MyBase.CreateFromRelative
        Dim getEmbeddedImage As CodeMemberMethod = MyBase.CreateFromEmbedded

        With classDeclaration.Members
            .Add(initGraphics)
            .Add(renderGraphics)
            .Add(disposeGraphics)
            .Add(inheritedDispose)

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
                CreateRotateLocals(classvariables)
            End If
        End With


        _Consolidated.emit(classvariables, initGraphics, renderGraphics.Statements, disposeGraphics, _ExportSettings)

        For Each obj As GDIObject In _GDIObjects
            obj.emit(classvariables, initGraphics, renderGraphics.Statements, disposeGraphics, _ExportSettings, _Consolidated)
        Next

        codegen.GenerateCodeFromType(classDeclaration, sw, cop)
        sw.Close()

        Return sw.ToString

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates the quick code class declaration.
    ''' </summary>
    ''' <returns>A class for holding the quick code.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function CreateCLASSDeclaration() As CodeDom.CodeTypeDeclaration
        Trace.WriteLineIf(Session.Tracing, "ExportQuick.CreateCLASSDeclaration")

        Dim CLASSDeclaration As New CodeTypeDeclaration

        Dim CLASSConstruct As New CodeConstructor
        CLASSConstruct.Attributes = MemberAttributes.Public
        Dim invokeInitGraphics As New CodeMethodInvokeExpression

        invokeInitGraphics.Method = New CodeMethodReferenceExpression(New CodeThisReferenceExpression, "InitializeGraphics")
        CLASSConstruct.Statements.Add(invokeInitGraphics)

        With CLASSDeclaration
            .BaseTypes.Add(New CodeTypeReference(GetType(System.ComponentModel.Component)))
            .IsClass = True
            .Attributes = MemberAttributes.Public
            .Members.Add(CLASSConstruct)
            .Comments.Add(New CodeCommentStatement("GDIPlus Architect Component Class Output"))
            .Name = _Document.ExportSettings.ClassName
        End With

        Return CLASSDeclaration

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates the Quick Code's Render Graphics method to which objects can append their drawing code.
    ''' </summary>
    ''' <returns>A CodeMemberMethod holding the render graphics method</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function CreateCLASSRenderGraphics() As CodeMemberMethod
        Trace.WriteLineIf(Session.Tracing, "ExportQuick.CreateCLASSRenderGraphics")

        Dim mbrMetRenderGraphics As New CodeMemberMethod

        With mbrMetRenderGraphics
            .Name = "RenderGraphics"
            .Parameters.Add(New CodeParameterDeclarationExpression(GetType(System.Drawing.Graphics), "g"))
            .Attributes = MemberAttributes.Public
        End With

        Return mbrMetRenderGraphics
    End Function

#End Region

End Class
