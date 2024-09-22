Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : Icons
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Responsible for managing all icons and cursors used in GDI+ Architect.  Exposes 
''' an ImageList that objects can attach to for their icons, and an enumeration that 
''' objects can use to get at their specific icons.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class IconManager
    Inherits System.ComponentModel.Component
 
#Region "Type Declarations"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An enumeration of the various icons used in GDI+ Architect
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Enum EnumIcons As Integer
        '''<summary>new class icon</summary>
        newclass = 0
        '''<summary>new print document icon</summary>
        newprintdoc
        '''<summary>open icon</summary>
        open
        '''<summary>save icon</summary>
        save
        '''<summary>copy icon</summary>
        copy
        '''<summary>cut icon</summary>
        cut
        '''<summary>paste icon</summary>
        paste
        '''<summary>delete icon</summary>
        delete
        '''<summary>properties icon</summary>
        properties
        '''<summary>print preview icon</summary>
        printpreview
        '''<summary>undo icon</summary>
        undo
        '''<summary>redo icon</summary>
        redo
        '''<summary>print icon</summary>
        print
        '''<summary>save all</summary>
        saveall
        '''<summary>zoom  tool icon</summary>
        zoom
        '''<summary>text tool icon</summary>
        alpha
        '''<summary>bucket tool icon</summary>
        bucket
        '''<summary>ellipse tool icon</summary>
        circle
        '''<summary>dropper tool icon</summary>
        drop
        '''<summary>GDIField tool icon</summary>
        field
        '''<summary>import image tool icon</summary>
        import
        '''<summary>line tool icon</summary>
        line
        '''<summary>magnify tool icon</summary>
        magnify
        '''<summary>pen tool icon</summary>
        pen
        '''<summary>pointer tool icon</summary>
        pointer
        '''<summary>rectangle tool icon</summary>
        rect
        '''<summary>hand tool icon</summary>
        hand
        '''<summary>toolbox panel icon</summary>
        toolbox
        '''<summary>propery panel icon</summary>
        propwin
        '''<summary>history panel icon</summary>
        historywin
        '''<summary>page panel icon</summary>
        pagewin
        '''<summary>info panel icon.  Not used.</summary>
        infowin
        '''<summary>help contentsicon</summary>
        helpcontents
        '''<summary>help index icon</summary>
        helpindex
        '''<summary>help search icon</summary>
        helpsearch
        '''<summary>application icon</summary>
        app
        '''<summary>down scroller icon</summary>
        DownButton
        '''<summary>up scroller icon</summary>
        UpButton
        '''<summary>align objects left icon</summary>
        AlignLeft
        '''<summary>align objects middle icon</summary>
        AlignMiddle
        '''<summary>align objects right icon</summary>
        AlignRight
        '''<summary>align objects top icon</summary>
        AlignTop
        '''<summary>align objects center icon</summary>
        AlignCenter
        '''<summary>align objects bottom icon</summary>
        AlignBottom
        '''<summary>same size height icon</summary>
        SameSizeHeight
        '''<summary>same size width icon</summary>
        SameSizeWidth
        '''<summary>same size both icon</summary>
        SameSizeBoth
        '''<summary>vertical space equal icon</summary>
        VertSpaceEqual
        '''<summary>horizontal space equal icon</summary>
        HorizSpaceEqual
        '''<summary>sand backward icon</summary>
        SendBackward
        '''<summary>send to back con</summary>
        SendToBack
        '''<summary>bring forward icon</summary>
        BringForward
        '''<summary>bring to front icon</summary>
        BringToFront
        '''<summary>Align panel icon</summary>
        AlignMain
    End Enum

#End Region

#Region "Local Fields"

    '''<summary>An image list of all the icons used in GDI+ Architect</summary>
    Private _IconImageList As New ImageList

    '''<summary>The magnify cursor</summary>
    Private _CursorMagnify As Cursor = LoadCursor("GDIPlusArchitect.Magnify.cur")

    '''<summary>The eye dropper cursor</summary>
    Private _CursorDropper As Cursor = LoadCursor("GDIPlusArchitect.Dropper.cur")
    '''<summary>The fill bucket cursor</summary>
    Private _CursorBucket As Cursor = LoadCursor("GDIPlusArchitect.Bucket.cur")
    '''<summary>The hand cursor</summary>
    Private _CursorHand As Cursor = LoadCursor("GDIPlusArchitect.Hand.cur")

#End Region


#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns an image list that other objects in the application use to assign their 
    ''' icons
    ''' </summary>
    ''' <value>A populated image list containing all the icons used in the app.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property IconImageList() As ImageList
        Get
            Return _IconImageList
        End Get

    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the magnify cursor.
    ''' </summary>
    ''' <value>The magnify cursor</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property CursorMagnify() As Cursor
        Get
            Return _CursorMagnify
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the dropper cursor. 
    ''' </summary>
    ''' <value>The magnify cursor</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property CursorDropper() As Cursor
        Get
            Return _CursorDropper
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the bucket cursor.
    ''' </summary>
    ''' <value>The magnify cursor</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property CursorBucket() As Cursor
        Get
            Return _CursorBucket
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the hand cursor.
    ''' </summary>
    ''' <value>The magnify cursor</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property CursorHand() As Cursor
        Get
            Return _CursorHand
        End Get
    End Property
#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of the Icons class. Responsible for loading the icons 
    ''' into the image list
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()

        _IconImageList.TransparentColor = Color.FromArgb(255, 0, 255)

        With _IconImageList.Images
            .Add(GetBitmap("GDIPlusArchitect.newclass.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.newprintdoc.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.open.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.save.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.copy.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.cut.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.paste.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.delete.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.properties.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.printpreview.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.undo.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.redo.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.print.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.saveall.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.zoom.bmp"))

            .Add(GetBitmap("GDIPlusArchitect.alpha.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.bucket.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.circle.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.drop.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.field.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.import.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.line.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.magnify.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.pen.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.pointer.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.rect.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.hand.bmp"))

            .Add(GetBitmap("GDIPlusArchitect.toolbox.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.propwin.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.historywin.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.pagewin.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.infowin.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.helpcontents.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.helpindex.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.helpsearch.bmp"))
            .Add(GetIconResource("GDIPlusArchitect.app16.ico"))

            .Add(GetBitmap("GDIPlusArchitect.ScrollDown.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.ScrollUp.bmp"))


            .Add(GetBitmap("GDIPlusArchitect.alignLeft.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.alignMiddle.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.alignRight.bmp"))


            .Add(GetBitmap("GDIPlusArchitect.alignTop.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.alignCenter.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.alignBottom.bmp"))


            .Add(GetBitmap("GDIPlusArchitect.sameSizeHeight.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.sameSizeWidth.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.sameSizeBoth.bmp"))


            .Add(GetBitmap("GDIPlusArchitect.vertSpaceEqual.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.horizSpaceEqual.bmp"))

            .Add(GetBitmap("GDIPlusArchitect.sendBackward.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.sendToBack.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.bringForward.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.bringToFront.bmp"))
            .Add(GetBitmap("GDIPlusArchitect.AlignMain.bmp"))
        End With
    End Sub


#End Region

#Region "Helper Methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Helper function to load cursors from an identifier.
    ''' </summary>
    ''' <param name="strIdentifier">A  path to the embedded cursor resource</param>
    ''' <returns>The cursor contained in the resource.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function LoadCursor(ByVal strIdentifier As String) As Cursor
        Dim s As System.IO.Stream = _
        System.Reflection.Assembly.GetEntryAssembly.GetManifestResourceStream(strIdentifier)

        Dim c As Cursor = New Cursor(s)

        s.Close()

        Return c
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Helper function to get an icon from an embedded resource.
    ''' </summary>
    ''' <param name="strIdentifier">A path to the embedded resource.</param>
    ''' <returns></returns>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetIconResource(ByVal strIdentifier As String) As System.Drawing.Icon
        Dim s As System.IO.Stream = System.Reflection.Assembly.GetEntryAssembly.GetManifestResourceStream(strIdentifier)
        ' read the resource from the returned stream
        Dim tempicon As Icon = New System.Drawing.Icon(s)

        ' close the stream
        s.Close()

        Return tempicon

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets a bitmap from an embedded  resource.
    ''' </summary>
    ''' <param name="strIdentifier">A path to the embedded bitmap resource.</param>
    ''' <returns>The Drawing.Bitmap file at the path</returns>
    ''' <remarks>The extra code here is because closing the stream causes 
    ''' the returned bitmap to become invalid.
    ''' Calling Clone() doesn't actually clone the bitmap, instead, it returns 
    ''' another bitmap which references the same stream.  Furthmore, GetManifestResourceStream
    ''' returns an unmanagedmemorystream, which isn't something I felt comfortable leaving 
    ''' open.  Instead, I copy the unmanaged stream into a managed memory stream via an 
    ''' array of bytes.  The memory stream still must remain open, but at least it's 
    ''' a managed memory stream.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private Function GetBitmap(ByVal strIdentifier As String) As System.Drawing.Bitmap
        Dim s As IO.Stream = (GetType(MDIMain).Module.Assembly.GetManifestResourceStream(strIdentifier))

        'Create an array of bytes
        Dim b(CInt(s.Length)) As Byte

        'Read the resource into the array of bytes
        s.Read(b, 0, b.Length)
        s.Close()

        'Create a memory stream based off of the array of bytes
        Dim mem As New IO.MemoryStream(b)

        'Assign the memory stream
        Dim tempbmp As New Bitmap(mem)

 
        Return tempbmp

    End Function

#End Region

#Region "Disposal and Cleanup"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the list of icons and cursors
    ''' </summary>
    ''' <param name="disposing"></param>
    ''' -----------------------------------------------------------------------------
    Public Overloads Sub dispose(ByVal disposing As Boolean)
        If Not _IconImageList Is Nothing Then
            _IconImageList.dispose()
        End If
        If Not CursorMagnify Is Nothing Then
            CursorMagnify.Dispose()
        End If
        If Not CursorDropper Is Nothing Then
            CursorDropper.Dispose()
        End If
        If Not CursorBucket Is Nothing Then
            CursorBucket.Dispose()
        End If
        If Not CursorHand Is Nothing Then
            CursorHand.Dispose()
        End If

        MyBase.dispose(disposing)

    End Sub

#End Region

End Class
