
''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : GraphicsManager
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Holds methods that are required for graphic operations in multiple places within 
''' the project but don't conform to the object model.  Basically a utility class for 
''' graphic related operations.
''' </summary>
''' -----------------------------------------------------------------------------
Public Class GraphicsManager

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Temporary window used to retrieve graphics for objects that need a graphics context to 
    ''' perform operations but would not have one in the normal flow. For example, the text object 
    ''' needs to measure character ranges, which  is a method of the graphics objects. 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _tempWin As New System.Windows.Forms.Form

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A graphics container used to draw in scale mode as needed.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shared _GraphicsContainer As Drawing2D.GraphicsContainer



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a graphics context for operations that need one to do their work.
    ''' </summary>
    ''' <returns>A System.Drawing.Graphics context based on the current system settings.</returns>
    ''' <remarks>Ideally this would not be necessary, however, there are times when the 
    ''' structure of the System.Drawing Namespace requires a graphics context when it is
    ''' difficult to get at one, for example region operations and measuring character ranges.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Function getTempGraphics() As Drawing.Graphics
        Return _tempWin.CreateGraphics
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A matrix used to store the original graphics scale mode before scale mode 
    ''' changes.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _OriginalMatrix As Drawing2D.Matrix

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Begins a scaled graphics container.  Assumes a call will later be made to EndScaleMode.
    ''' </summary>
    ''' <param name="g">A System.Drawing.Graphics object</param>
    ''' <param name="fScale">The scale for the current view.</param>
    ''' <remarks>This is called when an object that references the document 
    ''' needs a custom scaling container.  It allows the object to maintain coordinates independent from any 
    ''' particular zoom setting.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Friend Sub BeginScaleMode(ByVal g As Graphics, ByVal fScale As Single)

        Dim priorSmooth As Drawing2D.SmoothingMode = g.SmoothingMode
        Dim priorHint As System.Drawing.Text.TextRenderingHint = g.TextRenderingHint


        'Notice that the original transformation matrix is cached.  This is so it can later 
        'be reverted to. 
        _OriginalMatrix = g.Transform()

        _GraphicsContainer = g.BeginContainer

        g.ScaleTransform(fScale, fScale)
        g.SmoothingMode = priorSmooth
        g.TextRenderingHint = priorHint

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Stops drawing under the current scale, closing the graphics container.  
    ''' Should be called after BeginScaleMode.
    ''' </summary>
    ''' <param name="g">The current graphics object</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub EndScaleMode(ByVal g As Graphics)
        g.EndContainer(_GraphicsContainer)
        g.Transform = _OriginalMatrix

        _OriginalMatrix.Dispose()

        _OriginalMatrix = Nothing
        _GraphicsContainer = Nothing

    End Sub





#Region "Image Related Methods"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Given an absolute path to an image, returns a bitmap of the image.
    ''' </summary>
    ''' <param name="sAbsPath">An absolute path to the resource to retrieve.</param>
    ''' <returns>A bitmap containing the resource specified in sAbsPath</returns>
    ''' -----------------------------------------------------------------------------
    Public Function ImageFromAbsolutePath(ByVal sAbsPath As String) As Bitmap
        Dim img As Bitmap = New Bitmap(sAbsPath)

        Return img

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Attempts to load an image from a file for situations where the file couldn't 
    ''' originally be found.  In other words, if there is a link to an image 
    ''' (texture resource or external bitmap resource) that has been moved or deleted, 
    ''' this method prompts the user to select an alternate image source to replace 
    ''' the missing resource with.
    ''' </summary>
    ''' <param name="sMSG">A string containing the reason the user is being prompted 
    ''' for an image file.</param>
    ''' <returns>A path to an alternate file chosen by the user, or an empty string 
    ''' if the user elected not to choose a file.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function AttemptLoadImage(ByVal sMSG As String) As String

        Dim fOpen As System.Windows.Forms.OpenFileDialog
        Dim tempBMP As Bitmap
        Dim iResp As System.Windows.Forms.DialogResult

        MsgBox(sMSG)
        Try
            fOpen = New System.Windows.Forms.OpenFileDialog
            With fOpen
                .Filter = GDIImage.CONST_IMG_FILTER
                .Title = "Select an alternate bitmap"
                .CheckFileExists = True
                .CheckPathExists = True
                iResp = .ShowDialog()
            End With


            If iResp = Windows.Forms.DialogResult.OK Then
                tempBMP = New Bitmap(fOpen.FileName)
                Return fOpen.FileName
            Else
                Return String.Empty
            End If



        Catch ex As Exception
            Return String.Empty
        Finally
            If Not fOpen Is Nothing Then
                fOpen.Dispose()
            End If

            If Not tempBMP Is Nothing Then
                tempBMP.Dispose()
                tempBMP = Nothing
            End If
        End Try

    End Function

#End Region
End Class
