''' <summary>Provides the centralized change state after a Boolean transition.</summary>
Public NotInheritable Class HasChangesChangedEventArgs
    Inherits EventArgs
    ''' <summary>Initializes the event data.</summary>
    ''' <param name="HasChanges">Whether at least one tracked property differs from its baseline.</param>
    ''' <param name="ChangedPropertyCount">The number of properties differing from their baselines.</param>
    Public Sub New(HasChanges As Boolean, ChangedPropertyCount As Integer)
        Me.HasChanges = HasChanges
        Me.ChangedPropertyCount = ChangedPropertyCount
    End Sub
    ''' <summary>Gets the centralized Boolean change state.</summary>
    Public ReadOnly Property HasChanges As Boolean
    ''' <summary>Gets the number of changed properties at the time of the event.</summary>
    Public ReadOnly Property ChangedPropertyCount As Integer
End Class
