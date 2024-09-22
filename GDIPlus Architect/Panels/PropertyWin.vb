Imports GDIObjects
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : PropertyWin
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Holds an instance of a property grid which displays the properties of the selected 
''' object(s) on the design surface.
''' </summary>
''' <remarks>The cboItems combobox is owner drawn.  Much of the code in this module 
''' is to support this.
''' 
''' If you have never used the property grid control, you will probably wonder a couple things.
''' 
''' Q) Where is the code that defines the look and feel of items in the property grid?
''' A) This code is specified in "Attributes" on the object's public properties.  With attributes
''' you can assign a default value, a description, the default editor (for example an image file 
''' choose editor called an ImageFileChooser), and much more.  If you are curious about how a 
''' property was able to do what it does in the grid, go to the object file where it is declared 
''' and look at the attributes declared on the property
''' 
''' Q) Where is the code that assigns the selected value to the property?
''' A) There is none.  If you need to control this, you control it with the SET code 
''' in the property and with attributes as discussed above.
''' 
''' Q) Fine, but what if I want to report back to a user that a complex property is bad instead of 
''' just quietly making it valid?
''' A) Throw a format exception.  The grid itself will handle it since the grid is assigning 
''' the property.
''' 
''' Q) What if I want to do general things every time properties change?
''' A) Handle the propertyvaluechanged event.
''' 
''' </remarks>
''' -----------------------------------------------------------------------------
Friend Class PropertyWin
    Inherits UserControl



#Region "Local fields"


    '''<summary>Used to denote the last item for for the owner drawn dropdown list</summary>
    Private _LastIndex As Int32 = -1
    '''<summary>Holds a reference to the last selection rectangle for the owner
    '''  drawn dropdown list</summary>
    Private _LastSelectedBounds As Rectangle

    '''<summary>The font used to draw the item's name in the owner drawn
    '''  combo box.  Items in this box are composed of two strings in 
    ''' similar fashion to Visual Studio's property panel drop down box.
    '''  The first component is the name of the object on the surface. 
    '''  This is drawn bold.  The second portion is the type of item. 
    '''  This is drawn with the normal combo box font.</summary>
    Private _ItemNameFont As Font

    '''<summary>Whether to stop looking at updates.  This is to avoid infinite loops in callbacks when a new item is selected in the drop down list.  Typically the cboItems combo box responds to changes in history, selected item, and document to update itself to the appropriate selected item.  But when an item is selected in the drop down, we don't want it to update.  This boolean implements that mechanism.</summary>
    Private _StopUpdates As Boolean = False

#End Region

#Region "Constructors"


    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        _ItemNameFont = New Font(cboItems.Font, FontStyle.Bold)

        'Add any initialization after the InitializeComponent() call

    End Sub
#End Region

#Region " Windows Form Designer generated code "


    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not _ItemNameFont Is Nothing Then
                _ItemNameFont.Dispose()
            End If
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Private WithEvents propGrid As System.Windows.Forms.PropertyGrid
    Private WithEvents cboItems As System.Windows.Forms.ComboBox

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.propGrid = New System.Windows.Forms.PropertyGrid
        Me.cboItems = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'propGrid
        '
        Me.propGrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.propGrid.CommandsVisibleIfAvailable = True
        Me.propGrid.LargeButtons = False
        Me.propGrid.LineColor = System.Drawing.SystemColors.ScrollBar
        Me.propGrid.Location = New System.Drawing.Point(0, 24)
        Me.propGrid.Name = "propGrid"
        Me.propGrid.PropertySort = System.Windows.Forms.PropertySort.Alphabetical
        Me.propGrid.Size = New System.Drawing.Size(384, 384)
        Me.propGrid.TabIndex = 0
        Me.propGrid.Text = "PropertyGrid"
        Me.propGrid.ViewBackColor = System.Drawing.SystemColors.Window
        Me.propGrid.ViewForeColor = System.Drawing.SystemColors.WindowText
        '
        'cboItems
        '
        Me.cboItems.Dock = System.Windows.Forms.DockStyle.Top
        Me.cboItems.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboItems.Location = New System.Drawing.Point(0, 0)
        Me.cboItems.Name = "cboItems"
        Me.cboItems.Size = New System.Drawing.Size(384, 21)
        Me.cboItems.TabIndex = 1
        '
        'PropertyWin
        '
        Me.Controls.Add(Me.propGrid)
        Me.Controls.Add(Me.cboItems)
        Me.Name = "PropertyWin"
        Me.Size = New System.Drawing.Size(384, 408)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "Refresh"



    Public Sub refreshPanel()
        updateDropDownItems()
        refreshView()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Updates the list of drop down items with the list of objects contained in the 
    ''' the current GDIDocument.
    ''' </summary>
    ''' <param name="doc">The current GDIDocument </param>
    ''' -----------------------------------------------------------------------------
    Private Sub updateDropDownItems()
        Trace.WriteLineIf(App.TraceOn, "PropertyWin.update")

        If Not MDIMain.ActiveDocument Is Nothing AndAlso Not _StopUpdates Then
            cboItems.BeginUpdate()
            cboItems.Items.Clear()

            For Each page As GDIPage In MDIMain.ActiveDocument
                For Each obj As GDIObject In page.GDIObjects
                    cboItems.Items.Add(obj)
                Next
            Next

            cboItems.Sorted = True
            cboItems.DrawMode = DrawMode.OwnerDrawFixed

        End If

    End Sub




    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Refreshes the property grid
    ''' </summary>
    ''' <param name="doc"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub refreshView()
        Trace.WriteLineIf(App.TraceOn, "PropertyWin.refreshView")
        If Not MDIMain.ActiveDocument Is Nothing Then
            Dim objs() As GDIObject = MDIMain.ActiveDocument.Selected.ToArray

            propGrid.SelectedObjects = objs

            If Not _StopUpdates Then
                If MDIMain.ActiveDocument.Selected.Count = 1 Then
                    If Not cboItems.SelectedItem Is objs(0) Then
                        cboItems.SelectedItem = objs(0)
                        cboItems.Invalidate()
                    End If
                Else
                    cboItems.SelectedIndex = -1
                End If
            End If
        Else
            propGrid.SelectedObject = Nothing
        End If

    End Sub

#End Region

#Region "Event Handlers"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in properties.  Asks the current GDIDocument to record a 
    ''' history marker for the changed property.
    ''' </summary>
    ''' <param name="s"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub pg_PropertyValueChanged(ByVal s As System.Object, ByVal e As System.Windows.Forms.PropertyValueChangedEventArgs) Handles propGrid.PropertyValueChanged
        If Not MDIMain.ActiveDocument Is Nothing Then

            Dim lbl As String = e.ChangedItem.Label & " set to " & e.ChangedItem.Value.ToString
            MDIMain.ActiveDocument.recordHistory(lbl)
            MDIMain.RefreshActiveDocument()
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the item list drop down box.  This box holds a list of the items
    ''' contained in the current document.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>Notice the use of _StopUpdates here.  This is used to short circuit updates to the 
    ''' control.  A clearer way would have been to remove listeners, but 
    ''' </remarks>
        ''' -----------------------------------------------------------------------------
    Private Sub cboItems_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboItems.SelectedValueChanged
        If cboItems.SelectedIndex > -1 Then

            Dim selItem As GDIObject = DirectCast(cboItems.SelectedItem, GDIObject)


            _StopUpdates = True
            MDIMain.ActiveDocument.Selected.DeselectAll()
            MDIMain.ActiveDocument.SelectOneItem(selItem)
            _StopUpdates = False

        End If


    End Sub

#End Region


#Region "Owner drawn list"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws selected items in cboItems.  cboItems is owner drawn, and this method 
    ''' handles rendering the selected item in the dropdown.
    ''' </summary>
    ''' <param name="iIndex"></param>
    ''' <param name="rect"></param>
    ''' <param name="g"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub drawSelectedItem(ByVal iIndex As Int32, ByVal rect As Rectangle, ByVal g As Graphics)

        Dim bText As SolidBrush
        Dim bBackColor As SolidBrush

        If cboItems.DroppedDown = True Then
            bBackColor = New SolidBrush(Color.Blue)
            bText = New SolidBrush(Color.White)
        Else
            bBackColor = New SolidBrush(Color.White)
            bText = New SolidBrush(Color.Black)
        End If



        g.FillRectangle(bBackColor, rect)

        Dim curItem As GDIObject = DirectCast(cboItems.Items(iIndex), GDIObject)
        Dim sizefBold As SizeF = g.MeasureString(curItem.Name, _ItemNameFont, New PointF(rect.X, (rect.Height - cboItems.Font.Height) \ 2 + rect.Top), StringFormat.GenericTypographic)

        g.DrawString(curItem.Name, _ItemNameFont, bText, rect.X, _
                    ((rect.Height - cboItems.Font.Height) \ 2) + rect.Top)
        g.DrawString(curItem.ClassName, cboItems.Font, bText, rect.X + sizefBold.Width + 15, _
                    ((rect.Height - cboItems.Font.Height) \ 2) + rect.Top)

        bText.Dispose()
        bBackColor.Dispose()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws unselected items in cboItems.  cboItems is owner drawn, and this method 
    ''' handles rendering the unselected items in the dropdown.
    ''' </summary>
    ''' <param name="iIndex"></param>
    ''' <param name="rect"></param>
    ''' <param name="g"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub drawUnselectedItem(ByVal iIndex As Int32, ByVal rect As Rectangle, ByVal g As Graphics)
        Dim bText As SolidBrush

        bText = New SolidBrush(Color.Black)

        g.FillRectangle(Brushes.White, rect)

        Dim curItem As GDIObject = DirectCast(cboItems.Items(iIndex), GDIObject)
        Dim sizefBold As SizeF = g.MeasureString(curItem.Name, _ItemNameFont, New PointF(rect.X, (rect.Height - cboItems.Font.Height) \ 2 + rect.Top), StringFormat.GenericTypographic)

        g.DrawString(curItem.Name, _ItemNameFont, bText, rect.X, _
                    ((rect.Height - cboItems.Font.Height) \ 2) + rect.Top)
        g.DrawString(curItem.ClassName, cboItems.Font, bText, rect.X + sizefBold.Width + 15, _
                    ((rect.Height - cboItems.Font.Height) \ 2) + rect.Top)

        bText.Dispose()
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws a single item in the cboItems box (handles the cboItems.DrawItem event).  
    ''' In owner drawn controls, the control is given a DrawItem event for each item 
    ''' it needs to render to the surface.  This method handles drawing each item in turn.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>For more information on owner drawn listcontrols, check MSDN.
    ''' </remarks>
        ''' -----------------------------------------------------------------------------
    Private Sub cboItems_DrawItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles cboItems.DrawItem




        Dim g As Graphics = e.Graphics

        If e.Index > -1 AndAlso e.Index < cboItems.Items.Count AndAlso cboItems.Items.Count > 0 Then

            'Draw current sel item

            If e.Index = cboItems.SelectedIndex Then
                drawSelectedItem(e.Index, e.Bounds, g)
                If _LastIndex <> -1 AndAlso _LastIndex <> e.Index AndAlso _LastIndex < cboItems.Items.Count Then
                    drawUnselectedItem(_LastIndex, _LastSelectedBounds, g)
                End If

                _LastIndex = e.Index
                _LastSelectedBounds = e.Bounds
            Else
                drawUnselectedItem(e.Index, e.Bounds, g)
            End If


        End If


    End Sub

#End Region
End Class
