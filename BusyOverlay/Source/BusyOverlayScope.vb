Imports System.Threading
''' <summary>
''' Represents one reference-counted busy operation started through <see cref="BusyOverlay.BeginOperation"/>.
''' </summary>
''' <remarks>Dispose the scope exactly once when the protected work finishes. Disposing an already disposed scope has no effect.</remarks>
Public NotInheritable Class BusyOverlayScope
    Implements IDisposable
    Private _Owner As BusyOverlay
    Private _Disposed As Integer
    ''' <summary>
    ''' Initializes a scope owned by the specified overlay.
    ''' </summary>
    ''' <param name="Owner">The overlay that owns the operation.</param>
    Friend Sub New(Owner As BusyOverlay)
        _Owner = Owner
    End Sub
    ''' <summary>
    ''' Completes this operation and allows the overlay to hide when no other operation is active.
    ''' </summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        If Interlocked.Exchange(_Disposed, 1) <> 0 Then Return
        Dim Owner As BusyOverlay = Interlocked.Exchange(_Owner, Nothing)
        If Owner IsNot Nothing Then Owner.EndScopedOperation()
    End Sub
End Class
