''' <summary>Describes one property that differs from the accepted baseline.</summary>
''' <remarks>Values are scalar snapshots. The control reference remains live and can subsequently be disposed.</remarks>
Public NotInheritable Class ChangeTrackingResult
    ''' <summary>Initializes an immutable description of a changed property.</summary>
    ''' <param name="TargetControl">The control owning the property.</param>
    ''' <param name="PropertyName">The tracked property name.</param>
    ''' <param name="OriginalValue">The accepted baseline value.</param>
    ''' <param name="CurrentValue">The observed current value.</param>
    Public Sub New(TargetControl As Control, PropertyName As String, OriginalValue As Object, CurrentValue As Object)
        ArgumentNullException.ThrowIfNull(TargetControl)
        ArgumentException.ThrowIfNullOrWhiteSpace(PropertyName)
        Me.TargetControl = TargetControl
        Me.PropertyName = PropertyName
        Me.OriginalValue = OriginalValue
        Me.CurrentValue = CurrentValue
    End Sub
    ''' <summary>Gets the control owning the property.</summary>
    Public ReadOnly Property TargetControl As Control
    ''' <summary>Gets the tracked property name.</summary>
    Public ReadOnly Property PropertyName As String
    ''' <summary>Gets the value captured by the last successful baseline operation.</summary>
    Public ReadOnly Property OriginalValue As Object
    ''' <summary>Gets the value observed when this result was created.</summary>
    Public ReadOnly Property CurrentValue As Object
End Class
