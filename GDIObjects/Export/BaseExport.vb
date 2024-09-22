Imports System.CodeDom
Imports System.CodeDom.Compiler

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : BaseExport
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' BaseExport is a set of common functionality used for exporting code for both 
''' ''' the GDI+ Architect PrintClass and the GDI+ Architect GraphicsClass.  
''' </summary>
''' <remarks>The goal of this base class is to try to extrapolate the common properties 
''' between these two export options into a single more manageable class.
''' </remarks>
''' -----------------------------------------------------------------------------
Friend MustInherit Class BaseExport

#Region "Local Fields"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The GDIDocument being exported to code.  
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _Document As GDIDocument


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether a routine to retrieve embedded resources is required or not.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _EmbedMethodNeeded As Boolean = False

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether a routine to retrieve resources from an absolute path is required or not.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _AbsoluteMethodNeeded As Boolean = False

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether a routine to retrieve relative resources is needed or not.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _RelativeMethodNeeded As Boolean = False
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether a routine to rotate resources is needed or not.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _RotateMethodNeeded As Boolean = False

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A set of all consolidated objects.  
    ''' Consolidated objects are used to avoid emitting unnecessary code.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _Consolidated As New ExportConsolidate

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The current export settings.  See ExportSettings for more information.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected _ExportSettings As ExportSettings

#End Region


#Region "Requires Implementation"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Required for inheritors to implement.  
    ''' Called to perform code generation.
    ''' </summary>
    ''' <returns>A string containing all of the generated code in the appropriate language.</returns>
    ''' -----------------------------------------------------------------------------
    Public MustOverride Function Generate() As String

#End Region

#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructor used when exporting an entire GDIDocument to code.
    ''' </summary>
    ''' <param name="doc">The document to export</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal doc As GDIDocument)
        Trace.WriteLineIf(Session.Tracing, "BaseExport.New")

        _Document = doc
        _ExportSettings = doc.ExportSettings

        checkImages()

        createConsolidatedObjects()
        checkRotation()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructor used to generate code for a set of selected objects.
    ''' </summary>
    ''' <param name="doc">The GDIDocument containing the selected objects.</param>
    ''' <param name="objects">The set of objects to export.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal doc As GDIDocument, ByVal objects As GDIObjCol)
        Trace.WriteLineIf(Session.Tracing, "BaseExport.New.Quick")

        _Document = doc
        _ExportSettings = doc.ExportSettings

        checkImages(objects)
        createConsolidatedObjects(objects)
        checkRotation(objects)
    End Sub


#End Region


#Region "Code Generation Helpers"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a CodeGenerator appropriate to the current language
    ''' </summary>
    ''' <param name="sw">A string writer to write the code to.</param>
    ''' <returns>An ICodeGenerator for either C# or VB.NET depending on ExportOption settings.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function getGenerator(ByVal sw As IO.StringWriter) As ICodeGenerator
        Trace.WriteLineIf(Session.Tracing, "BaseExport.getGenerator for: " & _Document.ExportSettings.Language.ToString)

        Dim codegen As ICodeGenerator

        Select Case _Document.ExportSettings.Language
            Case EnumCodeTypes.eCSharp
                Dim csProvider As New Microsoft.CSharp.CSharpCodeProvider
                codegen = csProvider.CreateGenerator(sw)
                csProvider.Dispose()

            Case EnumCodeTypes.eVB
                Dim vbProvider As New Microsoft.VisualBasic.VBCodeProvider
                codegen = vbProvider.CreateGenerator(sw)
                vbProvider.Dispose()
        End Select

        Return codegen

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a TextHint statement.
    ''' </summary>
    ''' <returns>A CodeAssignStatement used to assign a TextHint to the graphics object</returns>
    ''' <remarks>This creates the statement: ( g.TextRenderingHint = ... )
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Protected Function createTextHintAssignment() As CodeAssignStatement

        Dim hintassign As New CodeDom.CodeAssignStatement
        Dim hintTypeStatement As New CodeFieldReferenceExpression


        With hintTypeStatement
            .TargetObject = New CodeTypeReferenceExpression(GetType(System.Drawing.Text.TextRenderingHint))
            .FieldName = System.Enum.GetName(GetType(System.Drawing.Text.TextRenderingHint), _Document.TextRenderingHint)
        End With


        hintassign.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression("g"), "TextRenderingHint")
        hintassign.Right = hintTypeStatement

        Return hintassign
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a SmoothingMode assignment statement
    ''' </summary>
    ''' <returns>A CodeAssignStatement used to assign the current graphic object's smoothing mode.</returns>
    ''' <remarks>This returns the statement: ( g.SmoothingMode = ... )
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Protected Function createSmoothingModeAssignment() As CodeAssignStatement
        Dim smoothassign As New CodeDom.CodeAssignStatement

        Dim smoothTypeStatement As New CodeFieldReferenceExpression


        With smoothTypeStatement
            .TargetObject = New CodeTypeReferenceExpression(GetType(System.Drawing.drawing2d.SmoothingMode))
            .FieldName = System.Enum.GetName(GetType(System.Drawing.drawing2d.SmoothingMode), _Document.SmoothingMode)
        End With

        smoothassign.Left = New CodeFieldReferenceExpression(New CodeVariableReferenceExpression("g"), "SmoothingMode")
        smoothassign.Right = smoothTypeStatement

        Return smoothassign
    End Function


#End Region


#Region "Consolidation related methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Examine each GDIObject to be exported and checks which properties of the objects 
    ''' can be consolidated.
    ''' </summary>
    ''' <param name="objects">The set of objects being exported</param>
    ''' <remarks>
    ''' This overload is used when a specific set of objects are being exported rather 
    ''' than the entire document.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Protected Sub createConsolidatedObjects(ByVal objects As GDIObjCol)
        Trace.WriteLineIf(Session.Tracing, "BaseExport.createConsolidatedObjects.Quick")

        For Each pg As GDIPage In _Document
            For Each obj As GDIObject In objects

                If TypeOf obj Is GDIFilledShape Then
                    Dim fillobj As GDIFilledShape = DirectCast(obj, GDIFilledShape)
                    If fillobj.DrawFill AndAlso fillobj.Fill.Consolidate AndAlso _
                       _Document.ExportSettings.OverrideConsolidateFill = False Then
                        _Consolidated.consolidateFill(fillobj.Fill)
                    End If
                End If

                If TypeOf obj Is GDIShape Then
                    Dim strokedObj As GDIShape = DirectCast(obj, GDIShape)
                    If strokedObj.DrawStroke AndAlso strokedObj.Stroke.Consolidate AndAlso _
                    _Document.ExportSettings.OverrideConsolidateStroke = False Then
                        _Consolidated.consolidateStroke(strokedObj.Stroke)
                    End If
                End If

                If TypeOf obj Is GDIText Then
                    Dim gdiText As GDIText = DirectCast(obj, GDIText)
                    If gdiText.ConsolidateFont AndAlso _
                        _Document.ExportSettings.OverrideConsolidateFont = False Then
                        _Consolidated.consolidateFont(gdiText.Font)
                    End If
                    If gdiText.ConsolidateStringFormat AndAlso _
                        _Document.ExportSettings.OverrideConsolidateStringFormats = False Then
                        _Consolidated.consolidateStringFormats(gdiText.StringFormat)
                    End If
                End If

                If TypeOf obj Is GDIField Then
                    Dim gdiField As GDIField = DirectCast(obj, GDIField)
                    If gdiField.ConsolidateFont AndAlso _
                        _Document.ExportSettings.OverrideConsolidateFont = False Then
                        _Consolidated.consolidateFont(gdiField.Font)
                    End If
                    If gdiField.ConsolidateStringFormat AndAlso _
                        _Document.ExportSettings.OverrideConsolidateStringFormats = False Then
                        _Consolidated.consolidateStringFormats(gdiField.StringFormat)
                    End If
                End If
            Next
        Next

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Examine each GDIObject to be exported and checks which properties of the objects 
    ''' can be consolidated.
    ''' </summary>
    ''' <remarks>This overload is intended for when an entire document is being exported.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Protected Sub createConsolidatedObjects()
        Trace.WriteLineIf(Session.Tracing, "BaseExport.createConsolidatedObjects")

        For Each pg As GDIPage In _Document
            For Each obj As GDIObject In pg.GDIObjects

                If TypeOf obj Is GDIFilledShape Then
                    Dim fillobj As GDIFilledShape = DirectCast(obj, GDIFilledShape)
                    If fillobj.DrawFill AndAlso fillobj.Fill.Consolidate AndAlso _
                       _Document.ExportSettings.OverrideConsolidateFill = False Then
                        _Consolidated.consolidateFill(fillobj.Fill)
                    End If
                End If

                If TypeOf obj Is GDIShape Then
                    Dim strokedObj As GDIShape = DirectCast(obj, GDIShape)
                    If strokedObj.DrawStroke AndAlso strokedObj.Stroke.Consolidate AndAlso _
                    _Document.ExportSettings.OverrideConsolidateStroke = False Then
                        _Consolidated.consolidateStroke(strokedObj.Stroke)
                    End If
                End If

                If TypeOf obj Is GDIText Then
                    Dim gdiText As GDIText = DirectCast(obj, GDIText)
                    If gdiText.ConsolidateFont AndAlso _
                        _Document.ExportSettings.OverrideConsolidateFont = False Then
                        _Consolidated.consolidateFont(gdiText.Font)
                    End If
                    If gdiText.ConsolidateStringFormat AndAlso _
                        _Document.ExportSettings.OverrideConsolidateStringFormats = False Then
                        _Consolidated.consolidateStringFormats(gdiText.StringFormat)
                    End If
                End If

                If TypeOf obj Is GDIField Then
                    Dim gdiField As GDIField = DirectCast(obj, GDIField)
                    If gdiField.ConsolidateFont AndAlso _
                        _Document.ExportSettings.OverrideConsolidateFont = False Then
                        _Consolidated.consolidateFont(gdiField.Font)
                    End If
                    If gdiField.ConsolidateStringFormat AndAlso _
                        _Document.ExportSettings.OverrideConsolidateStringFormats = False Then
                        _Consolidated.consolidateStringFormats(gdiField.StringFormat)
                    End If
                End If
            Next
        Next
    End Sub

#End Region

#Region "Export Checkers"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Examines each object in the export set to see if one has been rotated.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected Sub checkRotation()
        Trace.WriteLineIf(Session.Tracing, "BaseExport.checkRotation")

        For Each pg As GDIPage In _Document
            For Each obj As GDIObject In pg.GDIObjects
                If obj.Rotation > 0 Then
                    _RotateMethodNeeded = True
                    Exit Sub
                End If
            Next
        Next
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Examines each GDIObject in a specific set of objects and checks whether 
    ''' any have been rotated.
    ''' </summary>
    ''' <param name="objCol">A collection of GDIObjects to check for rotation.</param>
    ''' -----------------------------------------------------------------------------
    Protected Sub checkRotation(ByVal objCol As GDIObjCol)
        Trace.WriteLineIf(Session.Tracing, "BaseExport.checkRotation.Quick")
        For Each obj As GDIObject In objCol
            If obj.Rotation > 0 Then
                _RotateMethodNeeded = True
                Exit Sub
            End If
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines which extra routines need to be embedded in the outgoing code to 
    ''' handle image loading and rendering.   This overload is intended for when a subset 
    ''' of objects are being exported.
    ''' </summary>
    ''' <param name="objCol">The set of objects being exported.</param>
    ''' -----------------------------------------------------------------------------
    Protected Sub checkImages(ByVal objCol As GDIObjCol)
        Trace.WriteLineIf(Session.Tracing, "BaseExport.checkImages.Quick")


        For Each obj As GDIObject In objCol
            If TypeOf obj Is GDIImage Then
                Dim objImg As GDILinkedImage = DirectCast(obj, GDILinkedImage)

                Select Case objImg.RuntimeSource

                    Case EnumLinkType.AbsolutePath
                        _AbsoluteMethodNeeded = True
                    Case EnumLinkType.Embedded
                        _EmbedMethodNeeded = True
                    Case EnumLinkType.RelativePath
                        _RelativeMethodNeeded = True
                End Select


            ElseIf TypeOf obj Is GDIFilledShape Then

                Dim fillobj As GDIFilledShape = DirectCast(obj, GDIFilledShape)

                If TypeOf fillobj.Fill Is GDITexturedFill AndAlso fillobj.DrawFill = True Then

                    Dim textureFill As GDITexturedFill = DirectCast(fillobj.Fill, GDITexturedFill)
                    Select Case textureFill.RuntimeSource
                        Case EnumLinkType.AbsolutePath
                            _AbsoluteMethodNeeded = True
                        Case EnumLinkType.Embedded
                            _EmbedMethodNeeded = True
                        Case EnumLinkType.RelativePath
                            _RelativeMethodNeeded = True
                    End Select
                End If

            End If
        Next

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determines which extra routines need to be embedded in the outgoing code to 
    ''' handle image loading and rendering.   This overload is intended for when an entire 
    ''' document is being exported.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Protected Sub checkImages()
        Trace.WriteLineIf(Session.Tracing, "BaseExport.checkImages")

        For Each pg As GDIPage In _Document
            For Each obj As GDIObject In pg.GDIObjects
                If TypeOf obj Is GDIImage Then
                    Dim objImg As GDILinkedImage = DirectCast(obj, GDILinkedImage)

                    Select Case objImg.RuntimeSource

                        Case EnumLinkType.AbsolutePath
                            _AbsoluteMethodNeeded = True
                        Case EnumLinkType.Embedded
                            _EmbedMethodNeeded = True
                        Case EnumLinkType.RelativePath
                            _RelativeMethodNeeded = True
                    End Select


                ElseIf TypeOf obj Is GDIFilledShape Then

                    Dim fillobj As GDIFilledShape = DirectCast(obj, GDIFilledShape)

                    If TypeOf fillobj.Fill Is GDITexturedFill AndAlso fillobj.DrawFill = True Then

                        Dim textureFill As GDITexturedFill = DirectCast(fillobj.Fill, GDITexturedFill)
                        Select Case textureFill.RuntimeSource
                            Case EnumLinkType.AbsolutePath
                                _AbsoluteMethodNeeded = True
                            Case EnumLinkType.Embedded
                                _EmbedMethodNeeded = True
                            Case EnumLinkType.RelativePath
                                _RelativeMethodNeeded = True
                        End Select
                    End If

                End If
            Next
        Next
    End Sub

#End Region


#Region "Rotation code emit related methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a method that begins the rotation of an object.
    ''' </summary>
    ''' <returns>A method capable of beginning rotation.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function CreateRotateStart() As CodeMemberMethod
        Trace.WriteLineIf(Session.Tracing, "BaseExport.CreateRotateStart")

        Dim MethCreateRotateStart As New CodeMemberMethod

        With MethCreateRotateStart
            .Name = "beginRotation"
            .ReturnType = Nothing
            .Attributes = MemberAttributes.Private
            .Parameters.Add(New CodeParameterDeclarationExpression(New CodeTypeReference(GetType(Drawing.Graphics)), "g"))
            .Parameters.Add(New CodeParameterDeclarationExpression(New CodeTypeReference(GetType(Single)), "rotation"))
            .Parameters.Add(New CodeParameterDeclarationExpression(New CodeTypeReference(GetType(PointF)), "rotationPoint"))
        End With

        Dim declareMatrix As New CodeVariableDeclarationStatement
        Dim initMatrix As New CodeObjectCreateExpression
        initMatrix.CreateType = New CodeTypeReference(GetType(Drawing2D.Matrix))

        With declareMatrix
            .InitExpression = initMatrix
            .Name = "mtx"
            .Type = New CodeTypeReference(GetType(Drawing2D.Matrix))
        End With

        Dim beginCon As New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeArgumentReferenceExpression("g"), "BeginContainer"))

        Dim setContainer As New CodeAssignStatement
        setContainer.Left = New CodeFieldReferenceExpression(New CodeThisReferenceExpression, "gContainer")
        setContainer.Right = beginCon

        Dim setTransform As New CodeAssignStatement
        setTransform.Left = New CodeFieldReferenceExpression(New CodeArgumentReferenceExpression("g"), "Transform")
        setTransform.Right = New CodeVariableReferenceExpression("mtx")

        Dim invokeRotate As New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeVariableReferenceExpression("mtx"), "RotateAt"))

        With invokeRotate
            .Parameters.Add(New CodeArgumentReferenceExpression("rotation"))
            .Parameters.Add(New CodeArgumentReferenceExpression("rotationPoint"))
        End With

        Dim mtxDispose As New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeVariableReferenceExpression("mtx"), "Dispose"))

        With MethCreateRotateStart
            .Statements.Add(declareMatrix)
            .Statements.Add(setContainer)
            .Statements.Add(invokeRotate)
            .Statements.Add(setTransform)
            .Statements.Add(mtxDispose)
        End With


        Return MethCreateRotateStart
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a method that ends the rotation of an object.
    ''' </summary>
    ''' <returns>A method capable of ending rotation.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function CreateRotateEnd() As CodeMemberMethod
        Trace.WriteLineIf(Session.Tracing, "BaseExport.CreateRotateStart")

        Dim MethCreateRotateEnd As New CodeMemberMethod

        With MethCreateRotateEnd
            .Name = "endRotation"
            .ReturnType = Nothing
            .Attributes = MemberAttributes.Private
            .Parameters.Add(New CodeParameterDeclarationExpression(GetType(Drawing.Graphics), "g"))
        End With


        Dim gRef As New CodeArgumentReferenceExpression("g")
        Dim endContainer As New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(gRef, "EndContainer"))

        endContainer.Parameters.Add(New CodeFieldReferenceExpression(New CodeThisReferenceExpression, "gContainer"))


        MethCreateRotateEnd.Statements.Add(endContainer)


        Return MethCreateRotateEnd
    End Function

 

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits Fields required to rotate an object on the surface.
    ''' </summary>
    ''' <param name="declarations">Declarations section of the graphics object</param>
    ''' <remarks>At this time, only the gContainer member is emitted.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Protected Sub CreateRotateLocals(ByVal declarations As CodeTypeMemberCollection)
        'Declare the rectangle

        Dim declareContainer As New CodeMemberField

        With declareContainer
            .InitExpression = Nothing
            .Name = "gContainer"
            .Attributes = MemberAttributes.Private
            .Type = New CodeTypeReference(GetType(Drawing2D.GraphicsContainer))
        End With

        declarations.Add(declareContainer)
    End Sub

#End Region

#Region "Image retrieval related emits"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a method capable of retrieving embedded resources.
    ''' </summary>
    ''' <returns>A CodeMemberMethod object with the code to retrieve embedded resources.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function CreateFromEmbedded() As CodeMemberMethod
        Trace.WriteLineIf(Session.Tracing, "BaseExport.CreateFromEmbedded")

        Dim MethCreateEmbedded As New CodeMemberMethod

        With MethCreateEmbedded
            .Name = "ImageFromEmbedded"
            .ReturnType = New CodeTypeReference(GetType(System.Drawing.Bitmap))
            .Attributes = MemberAttributes.Private
            .Parameters.Add(New CodeParameterDeclarationExpression(New CodeTypeReference(GetType(String)), "embeddedPath"))
        End With


        Dim getAssemblyCall As New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeTypeReferenceExpression(GetType(System.Reflection.Assembly)), "GetExecutingAssembly"))

        Dim streamDeclare As New CodeVariableDeclarationStatement(GetType(System.IO.Stream), "imageStream")

        Dim manifestCall As New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(getAssemblyCall, "GetManifestResourceStream"))

        manifestCall.Parameters.Add(New CodeArgumentReferenceExpression("embeddedPath"))

        streamDeclare.InitExpression = manifestCall

        Dim createBitmap As New CodeVariableDeclarationStatement(New CodeTypeReference(GetType(System.Drawing.Bitmap)), "img")
        createBitmap.InitExpression = New CodeObjectCreateExpression(New CodeTypeReference(GetType(System.Drawing.Bitmap)), New CodeVariableReferenceExpression("imageStream"))
        Dim returnBitmap As New CodeMethodReturnStatement(New CodeVariableReferenceExpression("img"))

        With MethCreateEmbedded
            .Statements.Add(streamDeclare)
            .Statements.Add(createBitmap)
            .Statements.Add(returnBitmap)
        End With


        Return MethCreateEmbedded
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a method capable of retrieving relative resources.
    ''' </summary>
    ''' <returns>A CodeMemberMethod object with the code to retrieve relative resources.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function CreateFromRelative() As CodeMemberMethod
        Trace.WriteLineIf(Session.Tracing, "BaseExport.CreateFromRelative")

        Dim MethCreateRelative As New CodeMemberMethod

        With MethCreateRelative
            .Name = "ImageFromRelativePath"

            .ReturnType = New CodeTypeReference(GetType(System.Drawing.Bitmap))
            .Attributes = MemberAttributes.Private
            .Parameters.Add(New CodeParameterDeclarationExpression(New CodeTypeReference(GetType(String)), "relativePath"))
        End With

        Dim getAssemblyCall As New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeTypeReferenceExpression(GetType(System.Reflection.Assembly)), "GetExecutingAssembly"))
        Dim locPathGet As New CodeFieldReferenceExpression(getAssemblyCall, "Location")

        Dim getDirectoryName As New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeTypeReferenceExpression(GetType(System.IO.Path)), "GetDirectoryName"))
        getDirectoryName.Parameters.Add(locPathGet)

        Dim executionpathDeclare As New CodeVariableDeclarationStatement(GetType(String), "executingPath")
        executionpathDeclare.InitExpression = getDirectoryName

        Dim createBitmap As New CodeVariableDeclarationStatement(New CodeTypeReference(GetType(System.Drawing.Bitmap)), "img")

        '  Dim x As String = String.Concat(s1,  "/", s2)

        Dim pathConcat As New CodeMethodInvokeExpression(New CodeMethodReferenceExpression(New CodeTypeReferenceExpression(GetType(String)), "Concat"))
        pathConcat.Parameters.Add(New CodeVariableReferenceExpression("executingPath"))
        pathConcat.Parameters.Add(New CodePrimitiveExpression("/"))
        pathConcat.Parameters.Add(New CodeVariableReferenceExpression("relativePath"))


        createBitmap.InitExpression = New CodeObjectCreateExpression(New CodeTypeReference(GetType(System.Drawing.Bitmap)), pathConcat)

        Dim returnBitmap As New CodeMethodReturnStatement(New CodeVariableReferenceExpression("img"))

        With MethCreateRelative
            .Statements.Add(executionpathDeclare)
            .Statements.Add(createBitmap)
            .Statements.Add(returnBitmap)
        End With

        Return MethCreateRelative
    End Function



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Emits a method capable of retrieving absolute resources.
    ''' </summary>
    ''' <returns>A CodeMemberMethod object with the code to retrieve absolute resources.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function CreateFromAbsolute() As CodeMemberMethod
        Trace.WriteLineIf(Session.Tracing, "BaseExport.CreateFromAbsolute")

        Dim MethCreateAbsolute As New CodeMemberMethod

        With MethCreateAbsolute
            .Name = "ImageFromAbsolutePath"
            .ReturnType = New CodeTypeReference(GetType(System.Drawing.Bitmap))
            .Attributes = MemberAttributes.Private
            .Parameters.Add(New CodeParameterDeclarationExpression(New CodeTypeReference(GetType(String)), "absolutePath"))
        End With

        Dim createBitmap As New CodeVariableDeclarationStatement(New CodeTypeReference(GetType(System.Drawing.Bitmap)), "img")
        createBitmap.InitExpression = New CodeObjectCreateExpression(New CodeTypeReference(GetType(System.Drawing.Bitmap)), New CodeArgumentReferenceExpression("absolutePath"))

        Dim returnBitmap As New CodeMethodReturnStatement(New CodeVariableReferenceExpression("img"))

        With MethCreateAbsolute
            .Statements.Add(createBitmap)
            .Statements.Add(returnBitmap)
        End With


        Return MethCreateAbsolute
    End Function

#End Region

#Region "GDI+ Architect Custom emits"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates the InitializeGraphics method for outgoing code.  The InitializeGraphics 
    ''' method is where GDI+ properties are assigned much in the same format as the 
    ''' InitializeComponent method.
    ''' </summary>
    ''' <returns>A CodeDOM.CodeMemberMethod to which other code pieces can be appended.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function CreateInitGraphics() As CodeMemberMethod
        Trace.WriteLineIf(Session.Tracing, "BaseExport.CreateInitGraphics")

        Dim MethCreateGraphics As New CodeMemberMethod

        With MethCreateGraphics
            .Name = "InitializeGraphics"
            .ReturnType = Nothing
            .Attributes = MemberAttributes.Private
        End With

        Return MethCreateGraphics
    End Function



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates the inherited override dispose.  
    ''' </summary>
    ''' <returns>A codeDOM.CodeMemberMethod that overrides the default dispose constructor 
    ''' to call the GDI+ Architect custom disposegraphics method in the flow of dispose.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function CreateInheritedDispose() As CodeDom.CodeMemberMethod
        Trace.WriteLineIf(Session.Tracing, "BaseExport.CreateInheritedDispose")

        Dim ProtectedDispose As New CodeMemberMethod

        With ProtectedDispose
            .Name = "Dispose"
            .Attributes = MemberAttributes.Overloaded Or MemberAttributes.Family Or MemberAttributes.Override
            .Parameters.Add(New CodeParameterDeclarationExpression(GetType(Boolean), "disposing"))
        End With
        Dim invokeDisposeGraphics As New CodeMethodInvokeExpression
        invokeDisposeGraphics.Method = New CodeMethodReferenceExpression(New CodeThisReferenceExpression, "DisposeGraphics")

        Dim booleanEvaluation As New CodeConditionStatement

        With booleanEvaluation
            .Condition = New CodeVariableReferenceExpression("disposing")
            .TrueStatements.Add(invokeDisposeGraphics)
        End With

        ProtectedDispose.Statements.Add(booleanEvaluation)

        Return ProtectedDispose
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates the DisposeGraphics section of the GDI+ Architect class file.  This allows 
    ''' the class to custom dispose its created resources.
    ''' </summary>
    ''' <returns>A CodeMemberMethod to which will be added
    ''' resources needing disposal.</returns>
    ''' -----------------------------------------------------------------------------
    Protected Function CreateDisposeGraphics() As CodeMemberMethod
        Trace.WriteLineIf(Session.Tracing, "BaseExport.CreateDisposeGraphics")

        Dim MethDisposeGraphics As New CodeMemberMethod
        With MethDisposeGraphics
            .Name = "DisposeGraphics"
            .Comments.Add(New CodeCommentStatement("Required to dispose of created resources"))
            .Attributes = MemberAttributes.Private
        End With

        Return MethDisposeGraphics
    End Function

#End Region

End Class
