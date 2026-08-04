''' <summary>
''' Provides progress information for file and directory copy or deletion operations.
''' </summary>
''' <remarks>
''' Instances are immutable after construction. Negative numeric values supplied to the constructor are normalized to zero.
''' </remarks>
Public NotInheritable Class ProgressEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ProgressEventArgs"/> class.
    ''' </summary>
    ''' <param name="TotalSize">
    ''' The total number of bytes represented by the operation.
    ''' </param>
    ''' <param name="HandledSize">
    ''' The number of bytes already processed.
    ''' </param>
    ''' <param name="CurrentPath">
    ''' The path of the file currently being processed, or <see langword="Nothing"/> when no specific path applies.
    ''' </param>
    ''' <param name="ProcessedItems">
    ''' The number of files already completed.
    ''' </param>
    ''' <param name="TotalItems">
    ''' The total number of files represented by the operation.
    ''' </param>
    Public Sub New(TotalSize As Long, HandledSize As Long, CurrentPath As String, ProcessedItems As Long, TotalItems As Long)
        Me.TotalSize = Math.Max(0, TotalSize)
        Me.HandledSize = Math.Max(0, HandledSize)
        Me.CurrentPath = CurrentPath
        Me.ProcessedItems = Math.Max(0, ProcessedItems)
        Me.TotalItems = Math.Max(0, TotalItems)
    End Sub
    ''' <summary>
    ''' Gets the total number of bytes represented by the operation.
    ''' </summary>
    ''' <value>
    ''' The total operation size, in bytes.
    ''' </value>
    Public ReadOnly Property TotalSize As Long
    ''' <summary>
    ''' Gets the number of bytes that have already been processed.
    ''' </summary>
    ''' <value>
    ''' The accumulated processed size, in bytes.
    ''' </value>
    Public ReadOnly Property HandledSize As Long
    ''' <summary>
    ''' Gets the path of the file currently being processed.
    ''' </summary>
    ''' <value>
    ''' The current file path, or <see langword="Nothing"/> when no specific file is associated with the progress notification.
    ''' </value>
    Public ReadOnly Property CurrentPath As String
    ''' <summary>
    ''' Gets the number of files that have been completely processed.
    ''' </summary>
    ''' <value>
    ''' The completed file count.
    ''' </value>
    Public ReadOnly Property ProcessedItems As Long
    ''' <summary>
    ''' Gets the total number of files represented by the operation.
    ''' </summary>
    ''' <value>
    ''' The total file count.
    ''' </value>
    Public ReadOnly Property TotalItems As Long
    ''' <summary>
    ''' Gets the calculated completion percentage of the operation.
    ''' </summary>
    ''' <value>
    ''' An integer from 0 through 100 representing the operation completion percentage.
    ''' </value>
    ''' <remarks>
    ''' Byte-based progress is preferred when a positive total size is available. Item-based progress is used when the total byte size is zero. An operation with no bytes and no items is considered complete.
    ''' </remarks>
    Public ReadOnly Property PercentCompleted As Integer
        Get
            If TotalItems > 0 AndAlso ProcessedItems >= TotalItems Then Return 100
            If TotalSize > 0 Then Return CInt(Math.Clamp(Math.Floor(HandledSize / CDbl(TotalSize) * 100.0), 0.0, 100.0))
            If TotalItems > 0 Then Return CInt(Math.Clamp(Math.Floor(ProcessedItems / CDbl(TotalItems) * 100.0), 0.0, 100.0))
            Return 100
        End Get
    End Property
End Class