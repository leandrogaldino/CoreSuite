''' <summary>
''' Provides data for the <see cref="ValidationProvider.ControlValidated"/> event.
''' </summary>
Public NotInheritable Class ControlValidatedEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ControlValidatedEventArgs"/> class.
    ''' </summary>
    ''' <param name="Result">The completed control validation result.</param>
    Public Sub New(Result As ValidationResult)
        ArgumentNullException.ThrowIfNull(Result)
        Me.Result = Result
    End Sub
    ''' <summary>
    ''' Gets the completed control validation result.
    ''' </summary>
    ''' <value>The result produced for the validated control.</value>
    Public ReadOnly Property Result As ValidationResult
End Class
