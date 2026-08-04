''' <summary>
''' Provides data for the <see cref="ValidationProvider.ValidationValueRequested"/> event.
''' </summary>
Public NotInheritable Class ValidationValueRequestedEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ValidationValueRequestedEventArgs"/> class.
    ''' </summary>
    ''' <param name="TargetControl">The control whose value is being resolved.</param>
    ''' <param name="Value">The value resolved automatically by the provider.</param>
    Public Sub New(TargetControl As Control, Value As Object)
        If TargetControl Is Nothing Then Throw New ArgumentNullException(NameOf(TargetControl))
        Me.TargetControl = TargetControl
        Me.Value = Value
    End Sub
    ''' <summary>
    ''' Gets the control whose value is being resolved.
    ''' </summary>
    ''' <value>The control about to be validated.</value>
    Public ReadOnly Property TargetControl As Control
    ''' <summary>
    ''' Gets or sets the value that will be evaluated by the validation rules.
    ''' </summary>
    ''' <value>The automatically resolved value or an application-defined replacement.</value>
    Public Property Value As Object
    ''' <summary>
    ''' Gets or sets a value indicating whether the application supplied the final validation value.
    ''' </summary>
    ''' <value><see langword="True"/> when <see cref="Value"/> should replace the automatically resolved value; otherwise, <see langword="False"/>.</value>
    Public Property Handled As Boolean
End Class
