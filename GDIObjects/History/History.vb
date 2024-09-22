

''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : History
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a container for the history of a specific document (a set of history items 
''' recorded each time the document goes through a change that records undo or 
''' redo information).
''' </summary>
''' -----------------------------------------------------------------------------
Public Class History
    Inherits CollectionBase
    Implements IDisposable

#Region "Local Fields"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether History has been disposed or not.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Disposed As Boolean = False



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Current history position  within the history
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _HistoryIndex As Int32 = -1

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The number of allowed undo steps for this document.  This is set under preferences 
    ''' in the user interface, and probably doesn't belong inside this object.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _MaxUndoSteps As Int32
#End Region


#Region "Constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Creates a new instance of history.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Friend Sub New()
        _MaxUndoSteps = Session.Settings.UndoSteps - 1
    End Sub
#End Region


#Region "Property Accessors"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets the history item at the current position within the document.
    ''' </summary>
    ''' <value>The history item at the current position</value>
    ''' -----------------------------------------------------------------------------
    Friend ReadOnly Property CurHistItem() As HistoryItem
        Get
            Return DirectCast(innerlist(_HistoryIndex), HistoryItem)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gets or sets the current history position.
    ''' </summary>
    ''' <value>An integer within the bounds of the history collection</value>
    ''' <remarks>This is not intended to be called directly by other classes in the project 
    ''' with the exception of the GDIDocument class. 
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Friend Property CurPos() As Int32
        Get
            Return _HistoryIndex
        End Get
        Set(ByVal Value As Int32)
            If Value >= 0 AndAlso Value < Me.Count Then
                _HistoryIndex = Value
            End If
        End Set
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Indexes the history by a specific value
    ''' </summary>
    ''' <param name="Index">The index value to return a history position at.</param>
    ''' <value>Returns the history item at the specified index</value>
    ''' -----------------------------------------------------------------------------
    Default Friend ReadOnly Property HistoryItems(ByVal Index As Int32) As HistoryItem
        Get
            Return DirectCast(InnerList(Index), HistoryItem)
        End Get
    End Property

#End Region

#Region "Shadowed methods"

    'The following are shadowed to disallow certain functions on history to consumers by 
    'forcing a change in scope.

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Shadows the add member of the CollectionBase to make it do nothing and hide it.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shadows Sub add()

    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Shadows the remove member of the CollectionBase to make it do nothing and hide it.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shadows Sub remove()

    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Shadows the removeat member of the CollectionBase to make it do nothing and hide it.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shadows Sub removeat()

    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Shadows the insert member of the CollectionBase to make it do nothing and hide it.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private Shadows Sub insert()

    End Sub


#End Region



#Region "Module level members"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' "Undoes" a step of history.
    ''' </summary>
    ''' <remarks>
    ''' This method actually just moves the history position.
    ''' The caller, which is a GDIDocument, is actually responsible for raising the history
    ''' changed event which lets consumers know that history has been modified. 
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub undo()
        If _HistoryIndex > 0 Then
            _HistoryIndex -= 1
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' "Redoes" a step of history. 
    ''' </summary>
    ''' <remarks>
    ''' This method actually just moves history forward a step.
    ''' The caller, which is a GDIDocument, is actually responsible for raising the history
    ''' changed event which lets consumers know that history has been modified.  
    ''' </remarks>
    ''' -----------------------------------------------------------------------------
    Public Sub redo()
        If _HistoryIndex < list.Count - 1 Then
            _HistoryIndex += 1
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Persists a moment of history information.  This means creating a history item 
    ''' instance for the document and with the specific label and adding it to the history 
    ''' object, and also a
    ''' </summary>
    ''' <param name="doc">The document to persist into history</param>
    ''' <param name="sLabel">The label to give the document.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub recordHistory(ByVal doc As GDIObjects.GDIDocument, ByVal sLabel As String)
        _HistoryIndex += 1

        Dim histItem As New HistoryItem(sLabel, doc)

        If innerlist.Count > _HistoryIndex Then
            For i As Int32 = innerlist.Count - 1 To _HistoryIndex Step -1
                innerlist.RemoveAt(i)
            Next
        ElseIf innerlist.Count - 1 = _MaxUndoSteps Then
            innerlist.RemoveAt(0)
            _HistoryIndex -= 1
        End If

        innerlist.Add(histItem)

    End Sub

#End Region

#Region "Disposal and Cleanup"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of allocated history resources
    ''' </summary>
    ''' <param name="disposing">Whether unmanaged resources are being disposed or not</param>
    ''' -----------------------------------------------------------------------------
    Public Overloads Sub Dispose(ByVal disposing As Boolean)
        If Not _Disposed Then
            If disposing Then

                For Each histitem As HistoryItem In Me.InnerList
                    histitem.Dispose()
                Next
            End If
        End If

        _Disposed = True
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the history 
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Overloads Sub Dispose() Implements System.IDisposable.Dispose
        Me.Dispose(True)
    End Sub
#End Region


End Class
