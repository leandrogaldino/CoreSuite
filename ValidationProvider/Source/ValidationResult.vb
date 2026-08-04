''' <summary>
''' Represents the result of validating a single Windows Forms control.
''' </summary>
Public NotInheritable Class ValidationResult
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ValidationResult"/> class.
    ''' </summary>
    ''' <param name="TargetControl">The control that was evaluated.</param>
    ''' <param name="GroupName">The validation group assigned to the control.</param>
    ''' <param name="IsValid">A value indicating whether the control passed validation.</param>
    ''' <param name="ErrorMessage">The error message assigned to the control.</param>
    ''' <param name="FailureReason">The rule that caused validation to fail.</param>
    Public Sub New(TargetControl As Control, GroupName As String, IsValid As Boolean, ErrorMessage As String, FailureReason As ValidationFailureReason)
        Me.TargetControl = TargetControl
        Me.GroupName = GroupName
        Me.IsValid = IsValid
        Me.ErrorMessage = ErrorMessage
        Me.FailureReason = FailureReason
    End Sub
    ''' <summary>
    ''' Gets the control that was evaluated.
    ''' </summary>
    ''' <value>The validated Windows Forms control.</value>
    Public ReadOnly Property TargetControl As Control
    ''' <summary>
    ''' Gets the validation group assigned to the control.
    ''' </summary>
    ''' <value>The group name, or an empty string when the control is not grouped.</value>
    Public ReadOnly Property GroupName As String
    ''' <summary>
    ''' Gets a value indicating whether the control passed validation.
    ''' </summary>
    ''' <value><see langword="True"/> when the value is valid; otherwise, <see langword="False"/>.</value>
    Public ReadOnly Property IsValid As Boolean
    ''' <summary>
    ''' Gets the error message assigned to the control.
    ''' </summary>
    ''' <value>An empty string for a valid control; otherwise, the displayed validation message.</value>
    Public ReadOnly Property ErrorMessage As String
    ''' <summary>
    ''' Gets the rule that caused validation to fail.
    ''' </summary>
    ''' <value>A <see cref="ValidationFailureReason"/> value describing the result.</value>
    Public ReadOnly Property FailureReason As ValidationFailureReason
End Class
