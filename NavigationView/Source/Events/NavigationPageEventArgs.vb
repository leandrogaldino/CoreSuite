''' <summary>
''' Provides data for events associated with a <see cref="NavigationPage"/> instance.
''' </summary>
Public Class NavigationPageEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="NavigationPageEventArgs"/> class.
    ''' </summary>
    ''' <param name="Page">The page associated with the event.</param>
    Public Sub New(Page As NavigationPage)
        ArgumentNullException.ThrowIfNull(Page)
        Me.Page = Page
    End Sub
    ''' <summary>
    ''' Gets the page associated with the event.
    ''' </summary>
    ''' <value>The affected page.</value>
    Public ReadOnly Property Page As NavigationPage
End Class
