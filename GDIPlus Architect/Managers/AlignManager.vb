Imports GDIObjects

''' -----------------------------------------------------------------------------
''' Project	 : GDIPlus Architect
''' Class	 : AlignManager
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Handles alignment in the application.  Manages the current align to mode (shared 
''' by menus and panels) and provides wrapper methods for aligning objects on the 
''' surface.
''' </summary>
''' -----------------------------------------------------------------------------
Friend Class AlignManager

    '''<summary>The current align mode for objects on the surface.</summary>
    Private _AlignMode As EnumAlignMode

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The current align mode.
    ''' </summary>
    ''' <value></value>
    ''' -----------------------------------------------------------------------------
    Public Property AlignMode() As EnumAlignMode
        Get
            Return _AlignMode
        End Get
        Set(ByVal Value As EnumAlignMode)
            _AlignMode = Value
        End Set
    End Property


    '''<summary>Invokes a same size both</summary>
    Public Sub SameSizeBoth()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.MakeSameSizeBoth(_AlignMode)
        End If
    End Sub

    '''<summary>Invokes a same size height</summary>
    Public Sub SameSizeHeight()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.MakeSameSizeHeight(_AlignMode)
        End If
    End Sub

    '''<summary>Invokes a same size width</summary>
    Public Sub SameSizeWidth()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.MakeSameSizeWidth(_AlignMode)
        End If
    End Sub

    '''<summary>Invokes equal horizontal spacing </summary>
    Public Sub HorizSpaceEqual()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.distributeWidths(_AlignMode)
        End If
    End Sub


    '''<summary>Invokes equal vertical spacing </summary>
    Public Sub VertSpaceEqual()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.DistributeHeights(_AlignMode)
        End If
    End Sub

    '''<summary>Invokes align right</summary>
    Public Sub AlignRight()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.AlignRight(_AlignMode)
        End If
    End Sub


    '''<summary>Invokes align middle</summary>
    Public Sub AlignMiddle()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.AlignVertCenter(_AlignMode)
        End If
    End Sub



    '''<summary>Invokes align left</summary>
    Public Sub AlignLeft()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.alignLeft(_AlignMode)
        End If
    End Sub


    '''<summary>Invokes align top</summary>
    Public Sub AlignTop()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.AlignTop(_AlignMode)
        End If
    End Sub


    '''<summary>Invokes align center</summary>
    Public Sub AlignCenter()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.AlignHorizCenter(_AlignMode)
        End If
    End Sub

    '''<summary>Invokes align bottom</summary>
    Public Sub AlignBottom()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.AlignBottom(_AlignMode)
        End If
    End Sub

    '''<summary>Invokes bring forward</summary>
    Public Sub BringForward()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.BringForward()
        End If
    End Sub

    '''<summary>Invokes even height distribution</summary>
    Public Sub DistributeHeights()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.DistributeHeights(_AlignMode)
        End If
    End Sub

    '''<summary>Invokes align horizontal center</summary>
    Public Sub AlignHorizCenter()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.AlignHorizCenter(_AlignMode)
        End If
    End Sub

    '''<summary>Invokes align vertical center</summary>
    Public Sub AlignVertCenter()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.AlignVertCenter(_AlignMode)
        End If
    End Sub



    Public Sub DistributeWidths()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.distributeWidths(_AlignMode)
        End If
    End Sub

    'Invokes send backward 
    Public Sub SendBackward()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.SendBackward()
        End If
    End Sub

    'Invokes a bring to front 
    Public Sub BringToFront()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.BringToFront()
        End If
    End Sub


    'Invokes a send to back
    Public Sub SendToBack()
        If App.ToolManager.ToolInUse = False Then
            MDIMain.ActiveDocument.Selected.SendToBack()
        End If
    End Sub
End Class
