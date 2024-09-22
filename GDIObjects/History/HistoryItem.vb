''' -----------------------------------------------------------------------------
''' Project	 : GDIObjects
''' Class	 : HistoryItem
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Provides a container for a snapshot of history (for undo and redo).  Each time 
''' history is recorded, a new history item is created and added to the History thus far.
''' This allows the user to undo and redo as well as select a specific history spot to 
''' continue from.
''' </summary>
''' -----------------------------------------------------------------------------

Friend Class HistoryItem
    Implements IDisposable




#Region "local declarations"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Whether this history item has been disposed or not.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Disposed As Boolean = False

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A byte array that holds the GDIDocument recoded in this history item.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Snapshot As Byte()
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Label that this history item uses to display itself in lists.
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Private _Name As String
#End Region

#Region "Constructors"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Constructs a new history item instance.
    ''' </summary>
    ''' <param name="sLabel">The name to give this history position</param>
    ''' <param name="doc">The GDIDocument to be serialized into history.</param>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal sLabel As String, ByVal doc As GDIObjects.GDIDocument)
        _Name = sLabel
        recordbitStream(doc)
    End Sub
#End Region


#Region "Base Class Overrides"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns a sting representation of the history item object.
    ''' </summary>
    ''' <returns>The label of the history item</returns>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function ToString() As String
        Return _Name
    End Function
#End Region


#Region "Persistence Related methods"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Serializes a GDIDocument to a bit stream for more compact storage within the 
    ''' history item object.
    ''' </summary>
    ''' <param name="doc">The document to serialize to history.</param>
    ''' -----------------------------------------------------------------------------
    Private Sub recordbitStream(ByVal doc As GDIObjects.GDIDocument)

        Dim ioStream As System.IO.MemoryStream

        Try
            ioStream = New IO.MemoryStream

            Dim binSerial As New System.Runtime.Serialization.formatters.Binary.BinaryFormatter

            binSerial.Serialize(ioStream, doc)

            _Snapshot = ioStream.ToArray()

        Catch e As Exception
            Throw e
        Finally
            If Not ioStream Is Nothing Then
                ioStream.Close()
                ioStream = Nothing
            End If

        End Try
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retrieves the previously serialized document from this history item's bitstream.
    ''' </summary>
    ''' <value>The GDIdocument recorded when this history item was created.</value>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property HistoricDocument() As GDIObjects.GDIDocument
        Get
            'load objects from binary
            Dim ioStream As System.IO.MemoryStream
            Try
                Dim binSerial As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter

                ioStream = New IO.MemoryStream(_Snapshot)

                Return DirectCast(binSerial.Deserialize(ioStream), GDIObjects.GDIDocument)

            Catch e As Exception
                Throw New ApplicationException("File to retrieve history.", e)

            Finally
                If Not ioStream Is Nothing Then
                    ioStream.Close()
                    ioStream = Nothing
                End If

            End Try
        End Get

    End Property

#End Region

#Region "Dispoal and Cleanup"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Disposes of the history item.
    ''' </summary>
    ''' <param name="disposing">Whether unmanaged resources are being disposed or not.</param>
    ''' -----------------------------------------------------------------------------
    Public Overloads Sub Dispose(ByVal disposing As Boolean)
        If Not _Disposed Then
            If disposing Then
                _Snapshot = Nothing
            End If
        End If

        _Disposed = True
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Default dispose method
    ''' </summary>
    ''' -----------------------------------------------------------------------------
    Public Overloads Sub Dispose() Implements System.IDisposable.Dispose
        Me.Dispose(True)
    End Sub
#End Region
End Class
