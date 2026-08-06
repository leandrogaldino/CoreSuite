Partial Public Class NavigationView
    ''' <summary>
    ''' Displays the page identified by the specified key.
    ''' </summary>
    ''' <param name="Key">The case-insensitive key of the page to display.</param>
    ''' <returns><see langword="True"/> when the page is displayed or already selected; otherwise, <see langword="False"/>.</returns>
    Public Overloads Function Navigate(Key As String) As Boolean
        If String.IsNullOrWhiteSpace(Key) Then Throw New ArgumentException("The page key cannot be empty.", NameOf(Key))
        Dim Page As NavigationPage = _Pages.FindByKey(Key)
        If Page Is Nothing Then Throw New KeyNotFoundException($"No navigation page with the key '{Key}' was found.")
        Return Navigate(Page)
    End Function
    ''' <summary>
    ''' Displays the specified page.
    ''' </summary>
    ''' <param name="Page">The page to display. It must belong to this control.</param>
    ''' <returns><see langword="True"/> when the page is displayed or already selected; otherwise, <see langword="False"/>.</returns>
    Public Overloads Function Navigate(Page As NavigationPage) As Boolean
        ArgumentNullException.ThrowIfNull(Page)
        EnsurePageBelongsToControl(Page)
        If Not Enabled OrElse Not Page.Enabled Then Return False
        If ReferenceEquals(_SelectedPage, Page) Then
            If Page.IsCreated Then Return True
            Return ReloadCurrentPage()
        End If
        Dim PreviousPage As NavigationPage = _SelectedPage
        Dim NavigatingArgs As New NavigationCancelEventArgs(PreviousPage, Page)
        RaiseEvent Navigating(Me, NavigatingArgs)
        If NavigatingArgs.Cancel Then Return False
        Dim WasCreated As Boolean = Page.IsCreated
        Dim PageControl As UserControl = Nothing
        Try
            PageControl = If(WasCreated, Page.CachedControl, Page.CreateNewControl())
            PreparePageControl(Page, PageControl)
            If Not WasCreated Then
                Page.AssignControl(PageControl)
                RaiseEvent PageCreated(Me, New NavigationPageEventArgs(Page))
            End If
            PageControl.BringToFront()
            PageControl.Show()
            If PreviousPage IsNot Nothing AndAlso PreviousPage.CachedControl IsNot Nothing Then PreviousPage.CachedControl.Hide()
            SetSelectedPage(Page)
            If PreviousPage IsNot Nothing AndAlso PreviousPage.CacheMode = NavigationPageCacheMode.Recreate AndAlso PreviousPage.ReleaseControl() Then RaiseEvent PageClosed(Me, New NavigationPageEventArgs(PreviousPage))
            RaiseEvent Navigated(Me, New NavigationEventArgs(PreviousPage, Page))
            Return True
        Catch Ex As Exception
            If Not WasCreated AndAlso PageControl IsNot Nothing AndAlso Not ReferenceEquals(Page, _SelectedPage) Then
                If ReferenceEquals(Page.CachedControl, PageControl) Then Page.AssignControl(Nothing)
                DisposeFailedControl(Page, PageControl)
            End If
            Return HandleNavigationFailure(Page, Ex)
        End Try
    End Function
    ''' <summary>
    ''' Disposes and recreates the control owned by the selected page.
    ''' </summary>
    ''' <returns><see langword="True"/> when the selected page is recreated; otherwise, <see langword="False"/>.</returns>
    Public Function ReloadCurrentPage() As Boolean
        Dim Page As NavigationPage = _SelectedPage
        If Page Is Nothing OrElse Not Enabled OrElse Not Page.Enabled Then Return False
        Dim NavigatingArgs As New NavigationCancelEventArgs(Page, Page)
        RaiseEvent Navigating(Me, NavigatingArgs)
        If NavigatingArgs.Cancel Then Return False
        Dim NewControl As UserControl = Nothing
        Dim OldControl As UserControl = Page.CachedControl
        Try
            NewControl = Page.CreateNewControl()
            If ReferenceEquals(NewControl, OldControl) Then Throw New InvalidOperationException("A page factory must return a new UserControl when the page is reloaded.")
            PreparePageControl(Page, NewControl)
            NewControl.BringToFront()
            NewControl.Show()
        Catch Ex As Exception
            If NewControl IsNot Nothing AndAlso Not ReferenceEquals(NewControl, OldControl) Then
                DisposeFailedControl(Page, NewControl)
            End If
            Return HandleNavigationFailure(Page, Ex)
        End Try
        If OldControl IsNot Nothing Then
            If OldControl.Parent IsNot Nothing Then OldControl.Parent.Controls.Remove(OldControl)
            OldControl.Dispose()
            RaiseEvent PageClosed(Me, New NavigationPageEventArgs(Page))
        End If
        Page.AssignControl(NewControl)
        RaiseEvent PageCreated(Me, New NavigationPageEventArgs(Page))
        UpdateSelectedButton()
        RaiseEvent Navigated(Me, New NavigationEventArgs(Page, Page))
        Return True
    End Function
    ''' <summary>
    ''' Closes and disposes the created control owned by the page with the specified key without removing its definition.
    ''' </summary>
    ''' <param name="Key">The case-insensitive key of the page to close.</param>
    ''' <returns><see langword="True"/> when a created or selected page was closed; otherwise, <see langword="False"/>.</returns>
    Public Overloads Function ClosePage(Key As String) As Boolean
        If String.IsNullOrWhiteSpace(Key) Then Throw New ArgumentException("The page key cannot be empty.", NameOf(Key))
        Dim Page As NavigationPage = _Pages.FindByKey(Key)
        If Page Is Nothing Then Throw New KeyNotFoundException($"No navigation page with the key '{Key}' was found.")
        Return ClosePage(Page)
    End Function
    ''' <summary>
    ''' Closes and disposes the created control owned by the specified page without removing its definition.
    ''' </summary>
    ''' <param name="Page">The page to close. It must belong to this control.</param>
    ''' <returns><see langword="True"/> when a created or selected page was closed; otherwise, <see langword="False"/>.</returns>
    Public Overloads Function ClosePage(Page As NavigationPage) As Boolean
        ArgumentNullException.ThrowIfNull(Page)
        EnsurePageBelongsToControl(Page)
        Dim WasSelected As Boolean = ReferenceEquals(Page, _SelectedPage)
        If Not WasSelected AndAlso Not Page.IsCreated Then Return False
        Dim ClosingArgs As New NavigationPageCancelEventArgs(Page)
        RaiseEvent PageClosing(Me, ClosingArgs)
        If ClosingArgs.Cancel Then Return False
        If WasSelected Then SetSelectedPage(Nothing)
        Dim WasReleased As Boolean = Page.ReleaseControl()
        If WasReleased Then RaiseEvent PageClosed(Me, New NavigationPageEventArgs(Page))
        Return WasSelected OrElse WasReleased
    End Function
    ''' <summary>
    ''' Disposes every cached page control except the currently displayed control.
    ''' </summary>
    ''' <returns>The number of page controls that were closed.</returns>
    Public Overloads Function ClearCache() As Integer
        Return ClearCache(False)
    End Function
    ''' <summary>
    ''' Disposes cached page controls and optionally includes the currently displayed control.
    ''' </summary>
    ''' <param name="IncludeCurrentPage"><see langword="True"/> to close the selected page as well; otherwise, <see langword="False"/>.</param>
    ''' <returns>The number of page controls that were closed.</returns>
    Public Overloads Function ClearCache(IncludeCurrentPage As Boolean) As Integer
        Dim ClosedCount As Integer
        If _Pages.Count = 0 Then Return ClosedCount
        Dim Snapshot(_Pages.Count - 1) As NavigationPage
        For Index As Integer = 0 To _Pages.Count - 1
            Snapshot(Index) = _Pages(Index)
        Next
        For Each Page As NavigationPage In Snapshot
            If (IncludeCurrentPage OrElse Not ReferenceEquals(Page, _SelectedPage)) AndAlso Page.IsCreated AndAlso ClosePage(Page) Then ClosedCount += 1
        Next
        Return ClosedCount
    End Function
    ''' <summary>
    ''' Determines whether the page identified by the specified key currently owns a created control.
    ''' </summary>
    ''' <param name="Key">The case-insensitive page key.</param>
    ''' <returns><see langword="True"/> when the page owns a non-disposed control; otherwise, <see langword="False"/>.</returns>
    Public Function IsPageCreated(Key As String) As Boolean
        Dim Page As NavigationPage = _Pages.FindByKey(Key)
        Return Page IsNot Nothing AndAlso Page.IsCreated
    End Function
    ''' <summary>
    ''' Gets a page's currently cached control without creating it.
    ''' </summary>
    ''' <param name="Key">The case-insensitive page key.</param>
    ''' <returns>The cached control, or <see langword="Nothing"/> when the page is missing or has not been created.</returns>
    Public Function GetCachedControl(Key As String) As UserControl
        Dim Page As NavigationPage = _Pages.FindByKey(Key)
        Return Page?.CachedControl
    End Function
    Private Sub PreparePageControl(Page As NavigationPage, PageControl As UserControl)
        If IsControlOwnedByAnotherPage(Page, PageControl) Then Throw New InvalidOperationException("A page factory must return a UserControl that is not owned by another navigation page.")
        If PageControl.Parent IsNot Nothing AndAlso Not ReferenceEquals(PageControl.Parent, _ContentPanel) Then Throw New InvalidOperationException("A navigation page factory must return a UserControl that is not parented by another container.")
        PageControl.Dock = DockStyle.Fill
        PageControl.Margin = New Padding(0)
        PageControl.Visible = False
        If PageControl.Parent Is Nothing Then _ContentPanel.Controls.Add(PageControl)
    End Sub
    Private Function IsControlOwnedByAnotherPage(Page As NavigationPage, PageControl As UserControl) As Boolean
        For Each ExistingPage As NavigationPage In _Pages
            If Not ReferenceEquals(ExistingPage, Page) AndAlso ReferenceEquals(ExistingPage.CachedControl, PageControl) Then Return True
        Next
        Return False
    End Function
    Private Sub DisposeFailedControl(Page As NavigationPage, PageControl As UserControl)
        If IsControlOwnedByAnotherPage(Page, PageControl) Then Return
        If PageControl.Parent IsNot Nothing AndAlso Not ReferenceEquals(PageControl.Parent, _ContentPanel) Then Return
        If PageControl.Parent IsNot Nothing Then PageControl.Parent.Controls.Remove(PageControl)
        PageControl.Dispose()
    End Sub
    Private Sub SetSelectedPage(Page As NavigationPage)
        If ReferenceEquals(_SelectedPage, Page) Then Return
        _SelectedPage = Page
        UpdateSelectedButton()
        RaiseEvent SelectedPageChanged(Me, EventArgs.Empty)
    End Sub
    Private Function HandleNavigationFailure(Page As NavigationPage, Ex As Exception) As Boolean
        Dim FailureArgs As New NavigationFailedEventArgs(Page, Ex)
        RaiseEvent NavigationFailed(Me, FailureArgs)
        If FailureArgs.Handled Then Return False
        System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(Ex).Throw()
        Return False
    End Function
    Private Sub EnsurePageBelongsToControl(Page As NavigationPage)
        If Not _Pages.Contains(Page) Then Throw New ArgumentException("The page does not belong to this NavigationView.", NameOf(Page))
    End Sub
End Class
