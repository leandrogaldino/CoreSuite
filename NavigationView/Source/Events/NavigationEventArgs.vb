''' <summary>
''' Provides data for the <see cref="NavigationView.Navigated"/> event.
''' </summary>
Public NotInheritable Class NavigationEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="NavigationEventArgs"/> class.
    ''' </summary>
    ''' <param name="PreviousPage">The page displayed before navigation, or <see langword="Nothing"/>.</param>
    ''' <param name="CurrentPage">The page displayed after navigation.</param>
    Public Sub New(PreviousPage As NavigationPage, CurrentPage As NavigationPage)
        Me.PreviousPage = PreviousPage
        Me.CurrentPage = CurrentPage
    End Sub
    ''' <summary>
    ''' Gets the page displayed before navigation.
    ''' </summary>
    ''' <value>The previous page, or <see langword="Nothing"/>.</value>
    Public ReadOnly Property PreviousPage As NavigationPage
    ''' <summary>
    ''' Gets the page displayed after navigation.
    ''' </summary>
    ''' <value>The current page.</value>
    Public ReadOnly Property CurrentPage As NavigationPage
End Class
