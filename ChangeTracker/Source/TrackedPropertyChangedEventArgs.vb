''' <summary>Describes an observed value change and the centralized state of the complete tracker.</summary>
Public NotInheritable Class TrackedPropertyChangedEventArgs
    Inherits EventArgs
    ''' <summary>Initializes the event data.</summary>
    ''' <param name="TargetControl">The control whose value changed.</param>
    ''' <param name="PropertyName">The tracked property name.</param>
    ''' <param name="OriginalValue">The accepted baseline value.</param>
    ''' <param name="PreviousValue">The value from the previous observation.</param>
    ''' <param name="CurrentValue">The newly observed value.</param>
    ''' <param name="HasChanges">The centralized state after the complete refresh.</param>
    ''' <param name="ChangedPropertyCount">The total number of properties differing from their baselines.</param>
    Public Sub New(TargetControl As Control, PropertyName As String, OriginalValue As Object, PreviousValue As Object, CurrentValue As Object, HasChanges As Boolean, ChangedPropertyCount As Integer)
        Me.TargetControl = TargetControl
        Me.PropertyName = PropertyName
        Me.OriginalValue = OriginalValue
        Me.PreviousValue = PreviousValue
        Me.CurrentValue = CurrentValue
        Me.HasChanges = HasChanges
        Me.ChangedPropertyCount = ChangedPropertyCount
    End Sub
    ''' <summary>Gets the control whose value changed.</summary>
    Public ReadOnly Property TargetControl As Control
    ''' <summary>Gets the tracked property name.</summary>
    Public ReadOnly Property PropertyName As String
    ''' <summary>Gets the accepted baseline value.</summary>
    Public ReadOnly Property OriginalValue As Object
    ''' <summary>Gets the previously observed value, which can differ from the baseline.</summary>
    Public ReadOnly Property PreviousValue As Object
    ''' <summary>Gets the newly observed value.</summary>
    Public ReadOnly Property CurrentValue As Object
    ''' <summary>Gets whether this property differs from its accepted baseline.</summary>
    Public ReadOnly Property IsChanged As Boolean
        Get
            Return Not Object.Equals(OriginalValue, CurrentValue)
        End Get
    End Property
    ''' <summary>Gets whether any tracked property differs from its baseline.</summary>
    Public ReadOnly Property HasChanges As Boolean
    ''' <summary>Gets the total number of changed properties after the complete refresh.</summary>
    Public ReadOnly Property ChangedPropertyCount As Integer
End Class
