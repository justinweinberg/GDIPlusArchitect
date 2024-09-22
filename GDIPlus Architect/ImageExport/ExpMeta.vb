Imports System.Drawing.Imaging
Imports GDIObjects
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : ExpMeta
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Responsible for Exporting the current page to a windows meta file.
''' </summary>
''' <remarks>
''' GDI+ can export very tiny metafiles.  It's a good export choice from GDI+ 
''' based applications.
''' </remarks>
''' -----------------------------------------------------------------------------
Friend Class ExpMeta
    Inherits Exp

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new intance of ExpMeta
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New("Windows Metafile", "image/x-wmf", "wmf |*.wmf| All files |*.*")
    End Sub

#End Region

#Region "Base Class Implementors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a windows metafile resource.  This is a bit different than creating a 
    ''' bitmap, as a device pointer is used to instantiate the metafile. 
    ''' </summary>
    ''' <param name="sFileName">The desired name of the metafile</param>
    ''' <returns>A metafile to which graphics can be rendered.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function createMetaFile(ByVal sFileName As String) As Metafile
        Dim g As Graphics
        Dim mf As System.Drawing.Imaging.Metafile = Nothing
        Dim hdc As IntPtr = IntPtr.Zero

        Try
            'Ask Drawing.Graphics to get its device context
            g = App.MDIMain.CreateGraphics()
            hdc = g.GetHdc()
            'From the device context, create the metafile.
            mf = New System.Drawing.Imaging.Metafile(sFileName, hdc)

        Catch ex As Exception
            Throw New Exception("Unable to allocate a windows metafile")
        Finally
            'Release the device context
            If Not g Is Nothing Then
                g.ReleaseHdc(hdc)
            End If
            If Not g Is Nothing Then
                g.Dispose()
            End If
        End Try


    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Generates a MetaFile of the current page.
    ''' </summary>
    ''' <param name="sFileName">The path to save the metafile to.</param>
    ''' <param name="doc">The parent document being exported.</param>
    ''' <param name="pg">The page being exported.</param>
    ''' <returns>A metafile containing the document.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function generateMetafile(ByVal sFileName As String, ByVal doc As GDIDocument, ByVal pg As GDIPage) As Metafile


        Dim g As Graphics


        Try
            'Get a metafile reference and a grapics context from the meta file
            Dim mf As Metafile = createMetaFile(sFileName)
            g = Graphics.FromImage(mf)

            'Fill the background with the document's backcolor
            Dim bBack As New SolidBrush(doc.BackColor)
            g.FillRectangle(bBack, doc.RectPageSize)
            bBack.Dispose()


            'Set the text hint and smoothing mode.
            g.TextRenderingHint = doc.TextRenderingHint
            g.SmoothingMode = doc.SmoothingMode
            'Draw the metafile to a bitmap
            pg.DrawToMetafile(g)


            Return mf


        Catch ex As Exception
            MsgBox("Unable to create metafile.")
        Finally
            If Not g Is Nothing Then
                g.Dispose()
            End If


        End Try



    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exports the metafile to an image.
    ''' </summary>
    ''' <param name="sFullPath">The path to save the jpeg to.</param>
    ''' <param name="doc">The parent document being exported.</param>
    ''' <param name="pg">The page being exported.</param>
    ''' <remarks>Unlike other image formats, you may notice that there is no "save" portion 
    ''' of the export.  This is because metafiles draw directly to the file.</remarks>
    ''' -----------------------------------------------------------------------------
    Public Overrides Sub Export(ByVal doc As GDIDocument, ByVal pg As GDIPage, ByVal sFullpath As String)

        Dim mf As System.Drawing.Imaging.Metafile = generateMetafile(sFullpath, doc, pg)

        mf.Dispose()
    End Sub

#End Region
End Class
