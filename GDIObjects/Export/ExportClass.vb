Imports System.CodeDom
Imports System.CodeDom.Compiler


''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : ExportClass
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Responsible for emitting a GDI+ Architect Graphics class from a document.
''' </summary>
''' <remarks>
''' GDIDocuments can be of two distinct types: PrintDocuments or GraphicsClass style 
''' documents.  This class is responsible for generating the GraphicsClas style 
''' export.
''' </remarks>
''' -----------------------------------------------------------------------------
Friend Class ExportClass
    Inherits BaseExport


#Region "Constructors"




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new instance of the ExportClass.
    ''' </summary>
    ''' <param name="doc">The GDIDocument being exported to code.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal doc As GDIDocument)
        MyBase.New(doc)
        Trace.WriteLineIf(Session.Tracing, "ExportClass.New")
    End Sub
#End Region

#Region "Base Class Implementers"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Called to generate the graphics class code.
    ''' </summary>
    ''' <returns>A string containing the generated code.</returns>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function Generate() As String
        Trace.WriteLineIf(Session.Tracing, "ExportClass.GenerateClassDocument")

        'create a string writer, and a CodeGenerator for the selected export language
        Dim sw As IO.StringWriter = New IO.StringWriter
        Dim codegen As System.CodeDom.Compiler.ICodeGenerator = MyBase.getGenerator(sw)
        Dim cop As New System.CodeDom.Compiler.CodeGeneratorOptions

        sw.WriteLine()

        'Create a class declaration (the code class itself.)
        Dim classDeclaration As CodeTypeDeclaration = Me.CreateCLASSDeclaration


        'get a reference to the field level collection of the class
        Dim classvariables As CodeTypeMemberCollection = classDeclaration.Members

        'Create a rendergraphics method. This is a custom GDI+ Architect method where graphics rendering code goes.
        Dim renderGraphics As CodeMemberMethod = CreateCLASSRenderGraphics()

        'Create an initgraphics method.  This is a custom GDI+ Architect method where initialization code goes.
        Dim initGraphics As CodeMemberMethod = MyBase.CreateInitGraphics()
        'Create a disposeGraphics method.  This is where GDI+ Architect cleans up System.Drawing resoures with .dispose methods.
        Dim disposeGraphics As CodeMemberMethod = MyBase.CreateDisposeGraphics()
        'Create a method that inherits from dispose and calls disposegraphics as needed. 
        Dim inheritedDispose As CodeMemberMethod = MyBase.CreateInheritedDispose()


        'Create custom methods for retrieving absolute, relative, and embedded images.
        Dim getAbsoluteImage As CodeMemberMethod = MyBase.CreateFromAbsolute
        Dim getRelativeImage As CodeMemberMethod = MyBase.CreateFromRelative
        Dim getEmbeddedImage As CodeMemberMethod = MyBase.CreateFromEmbedded

        'Populate the class with our code sections
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

        'emit consolidated objects first.  These are objects that were identified as 
        'equivalent for the purposes of export.
        _Consolidated.emit(classvariables, initGraphics, renderGraphics.Statements, disposeGraphics, _ExportSettings)

        'emit each object in the export set in turn. Notice that each object is asked to export 
        'itself.
        For Each pg As GDIPage In _Document
            For Each obj As GDIObject In pg.GDIObjects
                obj.Emit(classvariables, initGraphics, renderGraphics.Statements, disposeGraphics, _ExportSettings, _Consolidated)
            Next
        Next

        'Write the code out
        codegen.GenerateCodeFromType(classDeclaration, sw, cop)
        sw.Close()

        'Return the code
        Return sw.ToString

    End Function


#End Region

#Region "Code emit helpers"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a code class to which all other code is added.  This is the actual class
    ''' that will be emitted (the Graphics class itself).
    ''' </summary>
    ''' <returns>A CodeTypeDeclaration containing the class</returns>
    ''' -----------------------------------------------------------------------------
    Private Function CreateCLASSDeclaration() As CodeDom.CodeTypeDeclaration
        Trace.WriteLineIf(Session.Tracing, "ExportClass.CreateClassDeclaration")

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
    ''' Creates the renderGraphics method.
    ''' </summary>
    ''' <returns>a CodeMemberMethod containing the render graphics method.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function CreateCLASSRenderGraphics() As CodeMemberMethod
        Trace.WriteLineIf(Session.Tracing, "ExportClass.CreateClassRenderGraphics")

        Dim mbrMetRenderGraphics As New CodeMemberMethod

        With mbrMetRenderGraphics
            .Name = "RenderGraphics"
            .Parameters.Add(New CodeParameterDeclarationExpression(GetType(System.Drawing.Graphics), "g"))
            .Attributes = MemberAttributes.Public
        End With

        mbrMetRenderGraphics.Statements.Add(MyBase.createSmoothingModeAssignment)
        mbrMetRenderGraphics.Statements.Add(MyBase.createTextHintAssignment)

        Return mbrMetRenderGraphics
    End Function

#End Region

End Class
