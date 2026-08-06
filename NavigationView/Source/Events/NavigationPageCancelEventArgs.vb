Imports System.ComponentModel
''' <summary>
''' Provides data for the cancelable <see cref="NavigationView.PageClosing"/> event.
''' </summary>
Public NotInheritable Class NavigationPageCancelEventArgs
    Inherits CancelEventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="NavigationPageCancelEventArgs"/> class.
    ''' </summary>
    ''' <param name="Page">The page whose created control is about to close.</param>
    Public Sub New(Page As NavigationPage)
        ArgumentNullException.ThrowIfNull(Page)
        Me.Page = Page
    End Sub
    ''' <summary>
    ''' Gets the page whose created control is about to close.
    ''' </summary>
    ''' <value>The affected page.</value>
    Public ReadOnly Property Page As NavigationPage
End Class
