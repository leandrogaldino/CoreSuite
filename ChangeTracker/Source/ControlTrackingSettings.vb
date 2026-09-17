Imports System.ComponentModel
''' <summary>Stores the designer configuration and runtime subscriptions for one control.</summary>
Friend NotInheritable Class ControlTrackingSettings
    ''' <summary>Gets or sets whether the control participates in tracking.</summary>
    Public Property TrackChanges As Boolean
    ''' <summary>Gets or sets the normalized comma-separated property names.</summary>
    Public Property TrackedProperties As String = String.Empty
    ''' <summary>Gets or sets the normalized comma-separated alternative event names.</summary>
    Public Property TrackedEvents As String = String.Empty
    ''' <summary>Gets or sets whether runtime notification handlers are attached.</summary>
    Public Property AutomaticTracking As Boolean = True
    ''' <summary>Gets the property snapshots created when tracking starts.</summary>
    Public ReadOnly Property Properties As New List(Of TrackedPropertyState)
    ''' <summary>Gets the subscriptions owned by the tracker for this control.</summary>
    Public ReadOnly Property Subscriptions As New List(Of TrackingSubscription)
End Class
''' <summary>Stores a scalar property's baseline and last observed value.</summary>
Friend NotInheritable Class TrackedPropertyState
    ''' <summary>Gets or sets the control owning the tracked property.</summary>
    Public Property TargetControl As Control
    ''' <summary>Gets or sets the descriptor used to read and observe the property.</summary>
    Public Property Descriptor As PropertyDescriptor
    ''' <summary>Gets or sets the accepted baseline value.</summary>
    Public Property OriginalValue As Object
    ''' <summary>Gets or sets the most recently observed value.</summary>
    Public Property CurrentValue As Object
End Class
''' <summary>Owns an event subscription and removes the exact handler on disposal.</summary>
Friend NotInheritable Class TrackingSubscription
    Implements IDisposable
    Private _Detach As Action
    ''' <summary>Initializes a subscription with its matching unsubscription action.</summary>
    ''' <param name="Detach">The action that removes the registered handler.</param>
    Public Sub New(Detach As Action)
        _Detach = Detach
    End Sub
    ''' <summary>Removes the registered handler at most once.</summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        Dim Detach As Action = _Detach
        _Detach = Nothing
        If Detach IsNot Nothing Then Detach()
    End Sub
End Class
