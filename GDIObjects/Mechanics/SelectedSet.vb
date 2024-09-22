Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : SelectedSet
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Represents the set of currently selected GDIObjects in a GDIDocument.
''' </summary>
''' <remarks>There is a bit of ambiguity on the real difference between the purpose 
''' of the SelectedSet and the GDIObjCol.  The only real difference is that selected set
''' is designed specifically to work with selected objects, whereas the GDIObjCol is intended
''' to work with any set of GDIObjects.
''' </remarks>
'''  -----------------------------------------------------------------------------
<Serializable()> _
Public Class SelectedSet

#Region "Non Serialized Fields"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A back reference to the GDIDocument the SelectSet belongs to.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _Document As GDIDocument

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A point used to note where a drag operation on the surface began.  This point is used 
    ''' to determine how much offset from the initial position the set of objects have been 
    ''' dragged
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _DragStartPoint As Point

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The current bounds (a rectangle) of the selected set of objects prior 
    ''' to a drag operation
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
       Private _DragBoundsCurrent As Rectangle
 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The initial bounds (a rectangle) of the selected set of objects prior 
    ''' to a drag operation
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    <NonSerialized()> _
    Private _DragBoundsStart As Rectangle

#End Region

#Region "Local Fields"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The set of selected objects.  This is stored as a GDIObjCol which is a class 
    ''' used by GDIPages as well as the SelectedSet object to manipulate groups of GDIObjects.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _SelectedObjects As GDIObjects.GDIObjCol



#End Region


#Region "Event Declarations"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Raised when the set of selected objects changes.  Notifies the parent document 
    ''' which in turn broadcasts the event to interested subscribers.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Event SelectionChanged(ByVal s As Object, ByVal e As EventArgs)
#End Region


#Region "Constructors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new SelectedSet.
    ''' </summary>
    ''' <param name="doc">A document to which the SelectedSet belongs.</param>
    ''' <remarks>There is one and only one SelectedSet object per GDIDocument
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal doc As GDIDocument)
        _Document = doc
        _SelectedObjects = New GDIObjCol
    End Sub


#End Region


#Region "Code emit related methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Generates code for the selected set of objects
    ''' </summary>
    ''' <returns>A string of code that will render the selected objects.</returns>
    ''' -----------------------------------------------------------------------------
    Friend Function QuickCode() As String
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.QuickCode")

        Dim exportCode As New ExportQuick(_Document, _SelectedObjects)

        Return exportCode.Generate()
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Generates SVG for the selected set of objects.
    ''' </summary>
    ''' <returns>A string of XML code that will render the selected objects.</returns>
    ''' -----------------------------------------------------------------------------
    Friend Function QuickSVG() As String
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.QuickCode")

        Dim xmlDoc As New System.Xml.XmlDocument

        xmlDoc.XmlResolver = Nothing

        xmlDoc.LoadXml(String.Format("<svg width='{0}px' height='{1}px' version='1.1'></svg>", _Document.RectPageSize.Width, _Document.RectPageSize.Height))
        Dim defNode As Xml.XmlNode = xmlDoc.CreateNode(Xml.XmlNodeType.Element, "defs", String.Empty)

        Dim group As Xml.XmlNode = xmlDoc.ChildNodes(0)
        xmlDoc.ChildNodes(0).AppendChild(defNode)
        For Each obj As GDIObject In _SelectedObjects
            obj.toXML(xmlDoc, defNode, group)
        Next

        Return prettyFormat(xmlDoc)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Pretty formats the outgoing SVG XML.
    ''' </summary>
    ''' <param name="voXML">a DOM of xml to pretty format.</param>
    ''' <returns>The XML content of the DOM with pretty formatting.</returns>
    ''' -----------------------------------------------------------------------------
    Private Function prettyFormat(ByVal voXML As Xml.XmlDocument) As String
        Dim oWriter As IO.TextWriter = New IO.StringWriter

        Dim oXMLWriter As Xml.XmlTextWriter = New Xml.XmlTextWriter(oWriter)

        oXMLWriter.WriteDocType("svg", "-//W3C//DTD SVG 1.1//EN", "http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd", Nothing)

        oXMLWriter.Formatting = Xml.Formatting.Indented
        voXML.Save(oXMLWriter)

        Return oWriter.ToString
    End Function

#End Region


#Region "Methods that set shared properties"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Transforms a set of selected objects by a width and height scale factor
    ''' </summary>
    ''' <param name="fPercWidth">The width percent to transform by</param>
    ''' <param name="fPercHeight">The height percent to transform by</param>
    ''' -----------------------------------------------------------------------------
    Public Sub Transform(ByVal fPercWidth As Single, ByVal fPercHeight As Single)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.scaletransform")
        Trace.WriteLineIf(Session.Tracing, "fPercWidth: " & fPercWidth)
        Trace.WriteLineIf(Session.Tracing, "fPercHeight: " & fPercHeight)


        Dim rectBounds As Rectangle = Me.Bounds()

        For Each obj As GDIObject In _SelectedObjects
            obj.ScaleTransform(rectBounds, fPercWidth, fPercHeight)
        Next

        _Document.recordHistory("Scale Transform")
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the font property of all selected objects, where appropriate 
    ''' (text or field objects) to a common value.
    ''' </summary>
    ''' <param name="Font">The new font to use</param>
    ''' -----------------------------------------------------------------------------
    Public Sub setFont(ByVal Font As Font)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.setFont")

        For Each obj As GDIObject In _SelectedObjects
            If TypeOf obj Is GDIText Then
                Dim GDIText As GDIText = DirectCast(obj, GDIText)
                GDIText.Font = Font
            ElseIf TypeOf obj Is GDIField Then
                Dim gdifield As GDIField = DirectCast(obj, GDIField)
                gdifield.Font = Font
            End If
        Next

        _Document.recordHistory("Font changed")

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the stroke on all relevant selected objects to a common stoke.
    ''' </summary>
    ''' <param name="stroke">The shared stroke to set for the objects.</param>
    ''' <remarks>Only records history if an object was actually affected.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub setStroke(ByVal stroke As GDIStroke)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.setStroke")

        Dim bAtLeastOneStroked As Boolean = False
        For Each obj As GDIObject In _SelectedObjects
            If TypeOf obj Is GDIShape Then

                DirectCast(obj, GDIShape).Stroke = stroke
                bAtLeastOneStroked = True
            End If
        Next

        If bAtLeastOneStroked Then
            _Document.recordHistory("Stroke Changed")
        End If
    End Sub

    ''' ----------------------------------------------------------------------
    ''' <summary>
    ''' Sets the fill on all relevant selected objects to a common fill.
    ''' </summary>
    ''' <param name="fill">The shared fill to set for all relevant objects (fillable objects) 
    ''' in the set.</param>
    ''' <remarks>Only records history if an object was actually affected.
    ''' </remarks>
    '''  -----------------------------------------------------------------------------
    Public Sub setFill(ByVal fill As GDIFill)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.setFill")

        Dim bAtLeastOneFilled As Boolean = False

        For Each obj As GDIObject In _SelectedObjects
            If TypeOf obj Is GDIFilledShape Then
                DirectCast(obj, GDIFilledShape).Fill = fill
                bAtLeastOneFilled = True
            End If
        Next

        If bAtLeastOneFilled Then
            _Document.recordHistory("Fill Changed")
        End If

    End Sub
#End Region

#Region "Cut,Copy, Paste methods"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Places the set of selected objects in "cut mode".  This removes them from the 
    ''' current canvas and populates the clipboard with the set.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub Cut()
      
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.Cut")

        Dim ido As New System.Windows.Forms.DataObject
        ido.SetData(_SelectedObjects.Format.Name, False, _SelectedObjects)
        System.Windows.Forms.Clipboard.SetDataObject(ido, True)

        If Not _SelectedObjects.Count = 0 Then
            _Document.CurrentPage.GDIObjects.Remove(_SelectedObjects)
            _SelectedObjects.Clear()
            _Document.recordHistory("Cut Objects")
        End If
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Copies the set of selected objects to the clipboard.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub Copy()
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.Copy")

        Dim ido As New System.Windows.Forms.DataObject
        ido.SetData(_SelectedObjects.Format.Name, False, _SelectedObjects)
        System.Windows.Forms.Clipboard.SetDataObject(ido, True)
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Implements a paste attributes feature.  Paste attributes takes the common set 
    ''' of attributes from a specific GDIObject and applies them to selected objects.
    ''' </summary>
    ''' <param name="d">DataObject that should contain a GDIObject if the command 
    ''' is relevant.</param>
    ''' <remarks>There's a lot of code in this procedure, but it's more complex than it 
    ''' looks.  The majority of the code checks whether attributes can be pasted.  The 
    ''' second bit of code assigns the matching attributes to the selected set.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub PasteAttributes(ByVal d As System.windows.forms.IDataObject)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.PasteAttributes")

        Dim gdiCol As GDIObjects.GDIObjCol = _
              DirectCast(d.GetData(GDIObjects.GDIObjCol.Format.Name, False), GDIObjects.GDIObjCol)

        If gdiCol.Count > 0 Then
            Dim gdiobj As GDIObject = gdiCol.Item(0)

            'Paste fill attributes, if applicable
            If TypeOf gdiobj Is GDIFilledShape Then
                Dim refFilledShape As GDIFilledShape = DirectCast(gdiobj, GDIFilledShape)

                For Each obj As GDIObject In _SelectedObjects
                    If TypeOf obj Is GDIFilledShape Then
                        Dim targetobj As GDIFilledShape = DirectCast(obj, GDIFilledShape)
                        targetobj.Fill = refFilledShape.Fill
                        targetobj.TrackFill = refFilledShape.TrackFill
                        targetobj.DrawFill = refFilledShape.DrawFill
                        targetobj = Nothing
                    End If
                Next

                refFilledShape = Nothing

            End If

            'Paste stroke attributes, if applicable
            If TypeOf gdiobj Is GDIShape Then
                Dim refShape As GDIShape = DirectCast(gdiobj, GDIShape)
                For Each obj As GDIObject In _SelectedObjects
                    If TypeOf obj Is GDIShape Then
                        Dim targetObj As GDIShape = DirectCast(obj, GDIShape)
                        targetObj.Stroke = refShape.Stroke
                        targetObj.DrawStroke = refShape.DrawStroke
                        targetObj = Nothing
                    End If
                Next

                refShape = Nothing
            End If

            'Paste text attributes, if applicable
            If TypeOf gdiobj Is GDIText Then
                Dim refText As GDIText = DirectCast(gdiobj, GDIText)

                For Each obj As GDIObject In _SelectedObjects
                    If TypeOf obj Is GDIText Then
                        Dim targetObj As GDIText = DirectCast(obj, GDIText)
                        targetObj.Font = refText.Font
                        targetObj.Alignment = refText.Alignment
                        targetObj = Nothing
                    ElseIf TypeOf obj Is GDIField Then
                        Dim targetObj As GDIField = DirectCast(obj, GDIField)
                        targetObj.Font = refText.Font
                        targetObj.Alignment = refText.Alignment
                        targetObj = Nothing
                    End If
                Next

                refText = Nothing

            End If
            'Paste field attributes, if applicable
            If TypeOf gdiobj Is GDIField Then
                Dim refField As GDIField = DirectCast(gdiobj, GDIField)

                For Each obj As GDIObject In _SelectedObjects
                    If TypeOf obj Is GDIText Then
                        Dim targetObj As GDIText = DirectCast(obj, GDIText)
                        targetObj.Font = refField.Font
                        targetObj.Alignment = refField.Alignment
                        targetObj = Nothing
                    ElseIf TypeOf obj Is GDIField Then
                        Dim targetObj As GDIField = DirectCast(obj, GDIField)
                        targetObj.Font = refField.Font
                        targetObj.Alignment = refField.Alignment
                        targetObj = Nothing
                    End If
                Next

                refField = Nothing

            End If

            _Document.recordHistory("Attributes Pasted")
        End If



    End Sub



#End Region


#Region "Alignment and Sizing"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the selected objects along a vertical center
    ''' </summary>
    ''' <param name="alignmode">The current align to settings (margin/canvas/normal) from 
    ''' the user interface project.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub AlignVertCenter(ByVal alignmode As EnumAlignMode)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.AlignVertCenter")

        Select Case alignmode
            Case EnumAlignMode.eCanvas
                _SelectedObjects.alignVertCenter(Me._Document.RectPageSize.X + (Me._Document.RectPageSize.Width \ 2))
            Case EnumAlignMode.eMargins
                _SelectedObjects.alignVertCenter(Me._Document.PrintableArea.X + (Me._Document.PrintableArea.Width \ 2))
            Case EnumAlignMode.eNormal
                _SelectedObjects.alignVertCenter()
        End Select


        _Document.recordHistory("Align Vertical")
    End Sub



       ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the selected objects along a horizontal center
    ''' </summary>
    ''' <param name="alignmode">The current align to settings (margin/canvas/normal) from 
    ''' the user interface project.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub AlignHorizCenter(ByVal alignmode As EnumAlignMode)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.AlignHorizCenter")

        Select Case alignmode
            Case EnumAlignMode.eCanvas
                _SelectedObjects.alignHorizCenter(Me._Document.RectPageSize.Y + (Me._Document.RectPageSize.Height \ 2))
            Case EnumAlignMode.eMargins
                _SelectedObjects.alignHorizCenter(Me._Document.PrintableArea.Y + (Me._Document.PrintableArea.Height \ 2))
            Case EnumAlignMode.eNormal
                _SelectedObjects.alignHorizCenter()
        End Select

        _Document.recordHistory("Align Horizontal")
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the selected objects rightmost depending on the alignmode setting.
    ''' </summary>
    ''' <param name="alignmode">The current align to settings (margin/canvas/normal) from 
    ''' the user interface project.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub AlignRight(ByVal alignmode As EnumAlignMode)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.AlignRight")

        Select Case alignmode
            Case EnumAlignMode.eCanvas
                _SelectedObjects.alignRight(Me._Document.RectPageSize.Right)
            Case EnumAlignMode.eMargins
                _SelectedObjects.alignRight(Me._Document.PrintableArea.Right)
            Case EnumAlignMode.eNormal
                _SelectedObjects.alignRight()
        End Select
        _Document.recordHistory("Align Right")
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns the selected objects leftmost depending on the alignmode setting.
    ''' </summary>
    ''' <param name="alignmode">The current align to settings (margin/canvas/normal) from 
    ''' the user interface project.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub alignLeft(ByVal alignmode As EnumAlignMode)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.alignLeft")
        Select Case alignmode
            Case EnumAlignMode.eCanvas
                _SelectedObjects.alignLeft(Me._Document.RectPageSize.Left)
            Case EnumAlignMode.eMargins
                _SelectedObjects.alignLeft(Me._Document.PrintableArea.Left)
            Case EnumAlignMode.eNormal
                _SelectedObjects.alignLeft()
        End Select

        _Document.recordHistory("Align Left")
    End Sub
 

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Attempts to evenly distribute the heights of selected objects.
    ''' </summary>
    ''' <param name="alignmode">The current align to settings (margin/canvas/normal) from 
    ''' the user interface project.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub DistributeHeights(ByVal alignmode As EnumAlignMode)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.DistributeHeights")

        Select Case alignmode
            Case EnumAlignMode.eCanvas
                _SelectedObjects.distributeHeights(Me._Document.RectPageSize)
            Case EnumAlignMode.eMargins
                _SelectedObjects.distributeHeights(Me._Document.PrintableArea)
            Case EnumAlignMode.eNormal
                _SelectedObjects.distributeHeights()
        End Select


        _Document.recordHistory("Distribute Heights")
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Attempts to evenly distribute the widths of the objects.
    ''' </summary>
    ''' <param name="alignmode">The current align to settings (margin/canvas/normal) from 
    ''' the user interface project.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub distributeWidths(ByVal alignmode As EnumAlignMode)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.distributeWidths")
        Select Case alignmode
            Case EnumAlignMode.eCanvas
                _SelectedObjects.distributeWidths(Me._Document.RectPageSize)
            Case EnumAlignMode.eMargins
                _SelectedObjects.distributeWidths(Me._Document.PrintableArea)
            Case EnumAlignMode.eNormal
                _SelectedObjects.distributeWidths()
        End Select

        _Document.recordHistory("Distribute Width")
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets objects to the same width depending on the settings of the alignmode argument.
    ''' </summary>
    ''' <param name="alignmode">The current align to settings (margin/canvas/normal) from 
    ''' the user interface project.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub MakeSameSizeWidth(ByVal alignmode As EnumAlignMode)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.MakeSameSizeWidth")
        Select Case alignmode
            Case EnumAlignMode.eCanvas
                _SelectedObjects.makeWidthEqual(Me._Document.RectPageSize.Width)
            Case EnumAlignMode.eMargins
                _SelectedObjects.makeWidthEqual(Me._Document.PrintableArea.Width)
            Case EnumAlignMode.eNormal
                _SelectedObjects.makeWidthEqual(Me.LastSelected)
        End Select

        _Document.recordHistory("Make Same Width")
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets objects to the same height depending on the settings of the alignmode argument.
    ''' </summary>
    ''' <param name="alignmode">The current align to settings (margin/canvas/normal) from 
    ''' the user interface project.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub MakeSameSizeHeight(ByVal alignmode As EnumAlignMode)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.MakeSameSizeHeight")
        Select Case alignmode
            Case EnumAlignMode.eCanvas
                _SelectedObjects.makeHeightEqual(Me._Document.RectPageSize.Height)
            Case EnumAlignMode.eMargins
                _SelectedObjects.makeHeightEqual(Me._Document.PrintableArea.Height)
            Case EnumAlignMode.eNormal
                _SelectedObjects.makeHeightEqual(Me.LastSelected)
        End Select

        _Document.recordHistory("Make Same Height")
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Makes objects have the same height and width depending on settings in the alignmode
    ''' argument.
    ''' </summary>
    ''' <param name="alignmode">The current align to settings (margin/canvas/normal) from 
    ''' the user interface project.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub MakeSameSizeBoth(ByVal alignmode As EnumAlignMode)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.MakeSameSizeBoth")
        Select Case alignmode
            Case EnumAlignMode.eCanvas
                _SelectedObjects.makeHeightWidthEqual(Me._Document.RectPageSize.Size)
            Case EnumAlignMode.eMargins
                _SelectedObjects.makeHeightWidthEqual(Me._Document.PrintableArea.Size)
            Case EnumAlignMode.eNormal
                _SelectedObjects.makeHeightWidthEqual(Me.LastSelected)
        End Select

        _Document.recordHistory("Make Same Width and Height")
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns a set of objects topmost depending on settings in the alignmode argument.
    ''' </summary>
    ''' <param name="alignmode">The current align to settings (margin/canvas/normal) from 
    ''' the user interface project.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub AlignTop(ByVal alignmode As EnumAlignMode)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.AlignTop")

        Select Case alignmode
            Case EnumAlignMode.eCanvas
                _SelectedObjects.alignTop(Me._Document.RectPageSize.Top)
            Case EnumAlignMode.eMargins
                _SelectedObjects.alignTop(Me._Document.PrintableArea.Top)
            Case EnumAlignMode.eNormal
                _SelectedObjects.alignTop()
        End Select

        _Document.recordHistory("Align Top")
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Aligns a set of objects bottommost depending on settings in the alignmode argument.
    ''' </summary>
    ''' <param name="alignmode">The current align to settings (margin/canvas/normal) from 
    ''' the user interface project.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub AlignBottom(ByVal alignmode As EnumAlignMode)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.AlignBottom")

        Select Case alignmode
            Case EnumAlignMode.eCanvas
                _SelectedObjects.alignBottom(Me._Document.RectPageSize.Bottom)
            Case EnumAlignMode.eMargins
                _SelectedObjects.alignBottom(Me._Document.PrintableArea.Bottom)
            Case EnumAlignMode.eNormal
                _SelectedObjects.alignBottom()
        End Select
        _Document.recordHistory("Align Bottom")
    End Sub


#End Region


#Region "Dragging, Positioning, and Selecting related members"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Selects a single specified item.
    ''' </summary>
    ''' <param name="obj">The object to select</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub SelectOneItem(ByVal obj As GDIObject)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.SelectOneItem")

        _SelectedObjects.Clear()
        _SelectedObjects.Add(obj)

        RaiseEvent SelectionChanged(Me, EventArgs.Empty)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Lets the selected set know that a drag operation has been completed.  This 
    ''' method stops drag operations for each of its items as well.
    ''' </summary>
    ''' <remarks>The history for this operation is recorded in the user interface layer.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub NotifyDragEnd()
        _DragBoundsCurrent = Rectangle.Empty
        _DragStartPoint = Point.Empty

        For Each gdiobject As GDIObjects.GDIObject In _SelectedObjects
            gdiobject.endDrag()
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Nudges the set of selected objects by an x and/or y factor.
    ''' </summary>
    ''' <param name="x">The X amount to nudge</param>
    ''' <param name="y">The Y amount to nudge</param>
    ''' -----------------------------------------------------------------------------
    Public Sub Nudge(ByVal x As Int32, ByVal y As Int32)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.Nudge: " & x & "," & y)

        _SelectedObjects.Nudge(x, y)

        _Document.recordHistory("Nudged objects")
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Explicitly sets the selection set to a new set of objects
    ''' </summary>
    ''' <param name="newSelection">The set of objects to select.</param>
    ''' -----------------------------------------------------------------------------
    Friend Sub SetSelection(ByVal newSelection As GDIObjCol)
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.SetSelection")

        _SelectedObjects.Clear()

        For Each obj As GDIObject In newSelection
            _SelectedObjects.Add(obj)
        Next

        RaiseEvent SelectionChanged(Me, EventArgs.Empty)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Starts a drag operation for the selected set of objects.  The point parameter 
    ''' is the initial point to begin dragging from.
    ''' </summary>
    ''' <param name="ptSnapped">The point of reference from which the drag operation 
    ''' began.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub initDragPoint(ByVal ptSnapped As Point)
        For Each gdiobject As GDIObjects.GDIObject In _SelectedObjects
            gdiobject.startDrag(ptSnapped)
        Next

        _DragBoundsStart = Me.Bounds()
        _DragStartPoint = ptSnapped
        _DragBoundsCurrent = _DragBoundsStart
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the drag rectangle based upon the most recent mouse cursor point.
    ''' </summary>
    ''' <param name="pt">The point at which the cursor was recorded</param>
    ''' <remarks>
    ''' What this math says is that the current drag rectangle is 
    ''' the difference between where the drag started and the current point minus the 
    ''' initial mouse down point.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub updateDragRect(ByVal pt As Point)
        _DragBoundsCurrent.X = _DragBoundsStart.X + (pt.X - _DragStartPoint.X)
        _DragBoundsCurrent.Y = _DragBoundsStart.Y + (pt.Y - _DragStartPoint.Y)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the drag positions of each object in the selected set. 
    ''' </summary>
    ''' <param name="pt">The point at which the last mouse position was recorded.</param>
    ''' <remarks>Each object is responsible for maintaining its own drag points as they 
    ''' are moved upon the designer surface.  This method allows those objects to update 
    ''' their internal drag points and set their new bounds accordingly.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub UpdateDraggedPositions(ByVal pt As Point)

        For Each gdiObj As GDIObjects.GDIObject In _SelectedObjects
            gdiObj.updateDrag(pt)
        Next
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the selected set given a drag rectangle and whether the shift key is currently down.
    ''' </summary>
    ''' <param name="rect">The rect of being dragged across to select objects</param>
    ''' <param name="bShiftDown">A Boolean indicating if the shift key is being held down during this 
    ''' operation.</param>
    ''' <remarks>
    '''  What this is designed to do is, given a selected set and a rectangle of new selections, 
    ''' set the selection set according to rules that most graphic applications use in these situations.  
    ''' Typically this means inverting the selection, unless the  shift key is down, 
    ''' which implies adding to the collection.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub handleSelection(ByVal rect As Rectangle, ByVal bShiftDown As Boolean)
        Dim tempcol As GDIObjects.GDIObjCol = _Document.CurrentPage.GDIObjects.getBoundObjects(rect)

        Dim tempSet As New GDIObjects.GDIObjCol


        If bShiftDown Then
            For Each gdiobject As GDIObjects.GDIObject In tempcol
                If _SelectedObjects.Contains(gdiobject) Then
                    _SelectedObjects.Remove(gdiobject)
                Else
                    _SelectedObjects.Add(gdiobject)
                End If
            Next

        Else
            _SelectedObjects.Clear()
            _SelectedObjects.AddRange(tempcol)
        End If

        RaiseEvent SelectionChanged(Me, EventArgs.Empty)

    End Sub


#End Region



#Region "Drawing Related Members"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws the selected objects.  Simply asks each object in the set to draw itself 
    ''' as "Selected" (with highlights and handle points as appropriate).
    ''' </summary>
    ''' <param name="g">Graphics to draw against </param>
    ''' <param name="fScale">The current zoom factor of the page.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub DrawSelected(ByVal g As Graphics, ByVal fScale As Single)
        _SelectedObjects.DrawSelected(g, fScale)
    End Sub

#End Region

#Region "ZOrder related members"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Brings the set of selected objects forward-most on the canvas.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub BringToFront()
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.BringToFront")

        If Not _SelectedObjects.Count = 0 Then
            _Document.CurrentPage.GDIObjects.bringToFront(_SelectedObjects)
            _Document.recordHistory("Objects To Front")
        End If
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sends the set of selected objects to the back of the canvas.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub SendToBack()
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.SendToBack")

        If Not _SelectedObjects.Count = 0 Then
            _Document.CurrentPage.GDIObjects.sendToback(_SelectedObjects)
            _Document.recordHistory("Objects To Back")
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sends the set of selected objects back a step.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub SendBackward()
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.SendBackward")

        If Not _SelectedObjects.Count = 0 Then
            _Document.CurrentPage.GDIObjects.sendBackward(_SelectedObjects)
            _Document.recordHistory("Objects Sent Backward")
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sends the set of selected objects forward a step
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub BringForward()
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.BringForward")

        If Not _SelectedObjects.Count = 0 Then
            _Document.CurrentPage.GDIObjects.bringForward(_SelectedObjects)
            _Document.recordHistory("Objects Brought Forward")
        End If
    End Sub

#End Region


#Region "Property Accessors"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the last item in the selected set of objects.  Sometimes the application 
    ''' needs to know this when the user performs an action that expects a single object
    ''' but multiple objects are selected.  This helps provide that ability.
    ''' </summary>
    ''' <value>The last selected object in the selected set.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property LastSelected() As GDIObjects.GDIObject
        Get
            If _SelectedObjects.Count > 0 Then
                Return _SelectedObjects.Item(_SelectedObjects.Count - 1)
            Else
                Return Nothing
            End If
        End Get
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the count of selected objects.
    ''' </summary>
    ''' <value>An integer containing the total count of the objects.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Count() As Int32
        Get
            Return _SelectedObjects.Count
        End Get
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a rectangle that bounds all of the selected objects.
    ''' </summary>
    ''' <value>A rectangle that bounds the selected objects.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property CurrentBounds() As Rectangle
        Get
            Return _DragBoundsCurrent
        End Get
    End Property

#End Region

#Region "Selection and Deselection"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deselects all of the selected objects on the canvas.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub DeselectAll()
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.DeselectAll")

        _SelectedObjects.Clear()
        RaiseEvent SelectionChanged(Me, EventArgs.Empty)

    End Sub

#End Region



#Region "Hit Test Related methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handles a selected object that has been hit on the surface.
    ''' </summary>
    ''' <param name="hitObject">The object to handle a hit for</param>
    ''' <param name="bShiftDown">Whether the Shift key is being held during this hit test</param>
    ''' <remarks>
    ''' The reason the Selected Set handles this action is because 
    ''' different selections happen depending on whether the selected object is in the set. 
    ''' For example, the object is removed from the set if it is already 
    ''' in the set and the user is holding shift down. 
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub handleNewSelection(ByVal hitObject As GDIObject, ByVal bShiftDown As Boolean)

        If Not bShiftDown Then

            If Not hitObject Is Nothing Then

                If Not _SelectedObjects.Contains(hitObject) Then
                    If _SelectedObjects.Count > 0 Then
                        _SelectedObjects.Clear()
                    End If

                    _SelectedObjects.Add(hitObject)

                End If
            Else
                _SelectedObjects.Clear()
            End If

        Else
            If Not hitObject Is Nothing Then
                If _SelectedObjects.Count > 0 Then
                    If _SelectedObjects.Contains(hitObject) Then
                        _SelectedObjects.Remove(hitObject)
                    Else
                        _SelectedObjects.Add(hitObject)
                    End If
                Else
                    _SelectedObjects.Add(hitObject)
                End If

            End If
        End If

        RaiseEvent SelectionChanged(Me, EventArgs.Empty)

    End Sub
#End Region

#Region "Misc Methods"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deletes all of the currently selected objects from the page / document.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub DeleteAll()
        Trace.WriteLineIf(Session.Tracing, "SelectedSet.DeleteAll")

        If Not _SelectedObjects.Count = 0 Then
            _Document.CurrentPage.GDIObjects.Remove(_SelectedObjects)
            _SelectedObjects.Clear()
            _Document.recordHistory("Delete Objects")
        End If

        RaiseEvent SelectionChanged(Me, EventArgs.Empty)

    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns the Bounds of all selected objects.
    ''' </summary>
    ''' <value>A rectangle that tightly encapsulates all of the selected objects.
    ''' </value>
    ''' <remarks>Note how these bounds are computed.  Each object is added 
    ''' to a graphics path and then the bounds of the path are computed and returned.
    ''' The only tricky part is rotating each object by its rotation value as they are 
    ''' added to the path.
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Private ReadOnly Property Bounds() As Rectangle
        Get

            Dim gPath As New Drawing2D.GraphicsPath

            For Each obj As GDIObject In _SelectedObjects
                If obj.Rotation > 0 Then
                    Dim mtx As New Drawing2D.Matrix
                    mtx.RotateAt(obj.Rotation, obj.RotationPoint)
                    Dim tempPath As New Drawing2D.GraphicsPath

                    tempPath.AddRectangle(obj.Bounds)
                    tempPath.Transform(mtx)
                    gPath.AddPath(tempPath, False)
                    tempPath.Dispose()
                    mtx.Dispose()
                Else
                    gPath.AddRectangle(obj.Bounds)
                End If

            Next

            Dim out As Rectangle = Rectangle.Round(gPath.GetBounds())
            gPath.Dispose()

            Return out
        End Get
    End Property



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Converts the set of selected objects to an array.  This is used to allow 
    ''' the visual studio property grid control to display the selected objects.
    ''' </summary>
    ''' <returns>An array of GDIObjects</returns>
    ''' -----------------------------------------------------------------------------
    Public Function ToArray() As GDIObject()
        Dim arrObjects(_SelectedObjects.Count - 1) As GDIObject

        For i As Int32 = 0 To _SelectedObjects.Count - 1
            arrObjects(i) = _SelectedObjects.Item(i)
        Next

        Return arrObjects
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Checks whether a specific GDIObject is in the set of selected objects.
    ''' </summary>
    ''' <param name="value">The GDIObject to test for participation in the selected set.</param>
    ''' <returns>A Boolean indicating whether the item is contained within the selected set.</returns>
    ''' -----------------------------------------------------------------------------
    Public Function Contains(ByVal value As GDIObject) As Boolean
        Return _SelectedObjects.Contains(value)
    End Function

#End Region

End Class
