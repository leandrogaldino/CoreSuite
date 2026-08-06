Imports System.ComponentModel
''' <summary>
''' Provides data for the cancelable <see cref="NavigationView.Navigating"/> event.
''' </summary>
Public NotInheritable Class NavigationCancelEventArgs
    Inherits CancelEventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="NavigationCancelEventArgs"/> class.
    ''' </summary>
    ''' <param name="CurrentPage">The page currently displayed, or <see langword="Nothing"/>.</param>
    ''' <param name="TargetPage">The page requested by the navigation operation.</param>
    Public Sub New(CurrentPage As NavigationPage, TargetPage As NavigationPage)
        Me.CurrentPage = CurrentPage
        Me.TargetPage = TargetPage
    End Sub
    ''' <summary>
    ''' Gets the page currently displayed.
    ''' </summary>
    ''' <value>The current page, or <see langword="Nothing"/> when no page is selected.</value>
    Public ReadOnly Property CurrentPage As NavigationPage
    ''' <summary>
    ''' Gets the page requested by the navigation operation.
    ''' </summary>
    ''' <value>The target page.</value>
    Public ReadOnly Property TargetPage As NavigationPage
End Class
