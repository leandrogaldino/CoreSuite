''' <summary>
''' Provides data for the <see cref="ValidationProvider.ValidatingControl"/> event and allows application-defined rules to change the final result.
''' </summary>
Public NotInheritable Class ValidatingControlEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ValidatingControlEventArgs"/> class.
    ''' </summary>
    ''' <param name="TargetControl">The control being evaluated.</param>
    ''' <param name="Value">The value resolved for the control.</param>
    ''' <param name="GroupName">The validation group assigned to the control.</param>
    ''' <param name="IsValid">The result produced by the built-in rules.</param>
    ''' <param name="ErrorMessage">The message produced by the built-in rules.</param>
    ''' <param name="FailureReason">The built-in rule that failed.</param>
    Public Sub New(TargetControl As Control, Value As Object, GroupName As String, IsValid As Boolean, ErrorMessage As String, FailureReason As ValidationFailureReason)
        ArgumentNullException.ThrowIfNull(TargetControl)
        Me.TargetControl = TargetControl
        Me.Value = Value
        Me.GroupName = GroupName
        Me.IsValid = IsValid
        Me.ErrorMessage = ErrorMessage
        Me.FailureReason = FailureReason
    End Sub
    ''' <summary>
    ''' Gets the control being evaluated.
    ''' </summary>
    ''' <value>The current validation target.</value>
    Public ReadOnly Property TargetControl As Control
    ''' <summary>
    ''' Gets the value resolved for the control.
    ''' </summary>
    ''' <value>The value used by the built-in validation rules.</value>
    Public ReadOnly Property Value As Object
    ''' <summary>
    ''' Gets the validation group assigned to the control.
    ''' </summary>
    ''' <value>The configured group name.</value>
    Public ReadOnly Property GroupName As String
    ''' <summary>
    ''' Gets or sets a value indicating whether the control is valid.
    ''' </summary>
    ''' <value><see langword="True"/> to accept the value; otherwise, <see langword="False"/>.</value>
    Public Property IsValid As Boolean
    ''' <summary>
    ''' Gets or sets the message displayed when the control is invalid.
    ''' </summary>
    ''' <value>The final validation error message.</value>
    Public Property ErrorMessage As String
    ''' <summary>
    ''' Gets or sets the reason associated with the final validation result.
    ''' </summary>
    ''' <value>A <see cref="ValidationFailureReason"/> value.</value>
    Public Property FailureReason As ValidationFailureReason
End Class
