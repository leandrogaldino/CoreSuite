''' <summary>
''' Provides data for the <see cref="NavigationView.NavigationFailed"/> event.
''' </summary>
Public NotInheritable Class NavigationFailedEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="NavigationFailedEventArgs"/> class.
    ''' </summary>
    ''' <param name="TargetPage">The page that could not be displayed.</param>
    ''' <param name="Exception">The exception raised while creating or displaying the page.</param>
    Public Sub New(TargetPage As NavigationPage, Exception As Exception)
        ArgumentNullException.ThrowIfNull(TargetPage)
        ArgumentNullException.ThrowIfNull(Exception)
        Me.TargetPage = TargetPage
        Me.Exception = Exception
    End Sub
    ''' <summary>
    ''' Gets the page that could not be displayed.
    ''' </summary>
    ''' <value>The requested target page.</value>
    Public ReadOnly Property TargetPage As NavigationPage
    ''' <summary>
    ''' Gets the exception raised by the navigation operation.
    ''' </summary>
    ''' <value>The navigation exception.</value>
    Public ReadOnly Property Exception As Exception
    ''' <summary>
    ''' Gets or sets a value indicating whether the exception has been handled by the application.
    ''' </summary>
    ''' <value><see langword="True"/> to suppress rethrowing the exception; otherwise, <see langword="False"/>.</value>
    Public Property Handled As Boolean
End Class
