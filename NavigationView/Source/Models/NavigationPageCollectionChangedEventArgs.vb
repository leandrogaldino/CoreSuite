Friend Enum NavigationPageCollectionChangeAction
    Add
    Remove
    Replace
    Reset
    ItemChanged
End Enum
Friend NotInheritable Class NavigationPageCollectionChangedEventArgs
    Inherits EventArgs
    Public Sub New(Action As NavigationPageCollectionChangeAction, Page As NavigationPage, OldPages As IReadOnlyList(Of NavigationPage))
        Me.Action = Action
        Me.Page = Page
        Me.OldPages = OldPages
    End Sub
    Public ReadOnly Property Action As NavigationPageCollectionChangeAction
    Public ReadOnly Property Page As NavigationPage
    Public ReadOnly Property OldPages As IReadOnlyList(Of NavigationPage)
End Class
