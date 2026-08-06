Partial Public Class NavigationView
    Private Sub PagesChanged(Sender As Object, E As NavigationPageCollectionChangedEventArgs)
        If E.OldPages IsNot Nothing Then
            For Each OldPage As NavigationPage In E.OldPages
                Dim WasSelected As Boolean = ReferenceEquals(OldPage, _SelectedPage)
                If WasSelected Then SetSelectedPage(Nothing)
                If OldPage.ReleaseControl() Then RaiseEvent PageClosed(Me, New NavigationPageEventArgs(OldPage))
            Next
        End If
        RebuildNavigationButtons()
        If _AutoNavigateFirstPage AndAlso _IsLoaded AndAlso Not IsInDesignMode AndAlso _SelectedPage Is Nothing Then ScheduleFirstPageNavigation()
    End Sub
    Private Sub ScheduleFirstPageNavigation()
        If _IsDisposing OrElse IsDisposed Then Return
        If IsHandleCreated Then
            BeginInvoke(New MethodInvoker(AddressOf NavigateFirstAvailablePage))
        Else
            NavigateFirstAvailablePage()
        End If
    End Sub
    Private Function NavigateFirstAvailablePage() As Boolean
        If _SelectedPage IsNot Nothing Then Return True
        For Each Page As NavigationPage In _Pages
            If Page.Visible AndAlso Page.Enabled Then Return Navigate(Page)
        Next
        Return False
    End Function
End Class
