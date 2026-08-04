Imports System.ComponentModel
''' <summary>
''' Provides data for the <see cref="BusyOverlay.CancellationRequested"/> event.
''' </summary>
Public NotInheritable Class BusyOverlayCancellationEventArgs
    Inherits CancelEventArgs
    Private ReadOnly _CancellableOperationCount As Integer
    ''' <summary>
    ''' Initializes a new instance of the <see cref="BusyOverlayCancellationEventArgs"/> class.
    ''' </summary>
    ''' <param name="CancellableOperationCount">The number of active operations that can receive cancellation.</param>
    Public Sub New(CancellableOperationCount As Integer)
        _CancellableOperationCount = CancellableOperationCount
    End Sub
    ''' <summary>
    ''' Gets the number of active operations whose cancellation tokens will be canceled.
    ''' </summary>
    ''' <value>The number of cancellable operations.</value>
    Public ReadOnly Property CancellableOperationCount As Integer
        Get
            Return _CancellableOperationCount
        End Get
    End Property
End Class
