''' <summary>Resumes a tracker once when a suspension scope is disposed.</summary>
Friend NotInheritable Class TrackingSuspension
    Implements IDisposable
    Private _Tracker As ChangeTracker
    ''' <summary>Initializes a scope for an already suspended tracker.</summary>
    ''' <param name="Tracker">The tracker that owns the suspension count.</param>
    Public Sub New(Tracker As ChangeTracker)
        _Tracker = Tracker
    End Sub
    ''' <summary>Releases this suspension once, refreshing when the outermost scope ends.</summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        If _Tracker Is Nothing Then Return
        _Tracker.VerifySuspensionAccess()
        Dim Tracker As ChangeTracker = _Tracker
        _Tracker = Nothing
        Tracker.EndSuspension()
    End Sub
End Class
