Imports GDIObjects
''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : HistoryWindow
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Holds a panel of history information.  This panel displays the history of changes
''' for the selected GDIDocument up to the limit specified under Options for undo steps.
''' Additionally, it allows the user to click and select history points that they wish 
''' to navigate to.
''' </summary>
''' <remarks>The history list is owner drawn.  This allows for item in the past to be drawn 
''' with the normal look and feel and items in the future (items past the current history 
''' position) to be drawn grayed.  For more information on owner drawn ListControls,
''' see MSDN.
''' </remarks>
''' -----------------------------------------------------------------------------
Friend Class HistoryWin
    Inherits System.Windows.Forms.UserControl

#Region "Local fields"
    '''<summary>Indicates if history is being updated and subsequent 
    ''' updates should be ignored</summary>
    Private _Updating As Boolean = False


    '''<summary>Current index in the history list.  This always is equal 
    ''' to the current document's History object's history 
    ''' position unless there isn't a current document, 
    ''' in which case this is set to -1</summary>
    Private _Index As Int32 = -1


    '''<summary>The last selected index.  Used for owner drawing the list</summary>
    Private _LastSelectedIndex As Int32 = -1


    '''<summary>The last selected item's bounds.  Used for owner drawing the list.</summary>
    Private _LastSelectedBounds As Rectangle

#End Region

#Region "Constructors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of a HistoryWin
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()


        'Add any initialization after the InitializeComponent() call

    End Sub

#End Region

#Region " Windows Form Designer generated code "



    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
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
    Private WithEvents lstHistory As System.Windows.Forms.ListBox

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lstHistory = New System.Windows.Forms.ListBox
        Me.SuspendLayout()
        '
        'lstHistory
        '
        Me.lstHistory.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lstHistory.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstHistory.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
        Me.lstHistory.Location = New System.Drawing.Point(0, 0)
        Me.lstHistory.Name = "lstHistory"
        Me.lstHistory.Size = New System.Drawing.Size(180, 156)
        Me.lstHistory.TabIndex = 0
        '
        'HistoryWin
        '
        Me.Controls.Add(Me.lstHistory)
        Me.Name = "HistoryWin"
        Me.Size = New System.Drawing.Size(180, 180)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Refresh"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Refreshes the list of history items
    ''' </summary>
    ''' <param name="doc"></param>
    ''' -----------------------------------------------------------------------------
    Public Sub refreshPanel()

        _Updating = True

        If MDIMain.ActiveDocument Is Nothing Then
            _Index = -1
            lstHistory.Items.Clear()
        Else
            _Index = MDIMain.ActiveDocument.CurHistoryPos

            With lstHistory
                 .Items.Clear()

                MDIMain.ActiveDocument.PopulateHistoryList(lstHistory)
         


                .SelectedIndex = _Index

             End With
        End If

 
        _Updating = False
    End Sub
#End Region

#Region "Owner drawn list"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws the selected history item to the list.
    ''' </summary>
    ''' <param name="iIndex">Index of the selected items</param>
    ''' <param name="rect">Bounds of the selected item</param>
    ''' <param name="g">Graphics context to draw against</param>
    ''' -----------------------------------------------------------------------------
    Private Sub drawSelectedItem(ByVal iIndex As Int32, ByVal rect As Rectangle, ByVal g As Graphics)
        g.FillRectangle(Brushes.Blue, rect)
        g.DrawString(CStr(iIndex + 1) & ". " & lstHistory.Items(iIndex).ToString, lstHistory.Font, Brushes.White, _
        rect.X, ((rect.Height - lstHistory.Font.Height) \ 2) + rect.Top)
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draws unselected history items to the list
    ''' </summary>
    ''' <param name="iIndex">Index of the unselected items</param>
    ''' <param name="rect">Bounds of the unselected item</param>
    ''' <param name="g">Graphics context to draw against</param>
    ''' -----------------------------------------------------------------------------
    Private Sub drawUnselectedItem(ByVal iIndex As Int32, ByVal rect As Rectangle, ByVal g As Graphics)
        Dim bText As SolidBrush

        If iIndex <= _Index Then
            bText = New SolidBrush(Color.Black)
        Else
            bText = New SolidBrush(Color.LightGray)

        End If

        g.FillRectangle(Brushes.White, rect)
        g.DrawString(CStr(iIndex + 1) & ". " & lstHistory.Items(iIndex).ToString, lstHistory.Font, bText, rect.X, _
                    ((rect.Height - lstHistory.Font.Height) \ 2) + rect.Top)

        bText.Dispose()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Owner draws the history list with its items.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub lstHistory_DrawItem(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles lstHistory.DrawItem

        Dim g As Graphics = e.Graphics

        If e.Index > -1 AndAlso e.Index < lstHistory.Items.Count AndAlso lstHistory.Items.Count > 0 Then

            'Draw current selected item
            If e.Index = lstHistory.SelectedIndex Then
                drawSelectedItem(e.Index, e.Bounds, g)
                If _LastSelectedIndex <> -1 AndAlso _LastSelectedIndex <> e.Index AndAlso _LastSelectedIndex < lstHistory.Items.Count Then
                    drawUnselectedItem(_LastSelectedIndex, _LastSelectedBounds, g)
                End If

                _LastSelectedIndex = e.Index
                _LastSelectedBounds = e.Bounds
            Else
                drawUnselectedItem(e.Index, e.Bounds, g)
            End If
        Else

        End If


    End Sub

#End Region

#Region "event handlers"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Enables and disables the underlying history list when the HistoryWin's enabled 
    ''' state changes.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub HistoryWindow_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.EnabledChanged
        If Me.Enabled Then
            lstHistory.Enabled = True
        Else
            lstHistory.Items.Clear()
            lstHistory.Enabled = False
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Responds to changes in the selected item within the history list and sets the 
    ''' selected history position in the current document based on the value.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' -----------------------------------------------------------------------------
    Private Sub lstHistory_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstHistory.SelectedValueChanged
        If Not _Updating Then
            If Not lstHistory.SelectedIndex = -1 Then
                MDIMain.ActiveDocument.setHistoryPos(lstHistory.SelectedIndex)
            End If
        End If
    End Sub

#End Region

End Class
