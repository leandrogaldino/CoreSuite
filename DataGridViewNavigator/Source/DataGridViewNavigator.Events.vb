Imports System.Windows.Forms

Public Partial Class DataGridViewNavigator

    ''' <summary>
    ''' Reports exceptions raised by a ToolStrip navigation click.
    ''' Explicit calls to MoveToXxxAsync continue to propagate exceptions via Task.
    ''' </summary>
    Public Event NavigationFailed As EventHandler(Of NavigationErrorEventArgs)

    Private Sub AttachGrid(Grid As DataGridView)
        AddHandler Grid.DataSourceChanged, AddressOf DataGridView_DataSourceChanged
        AddHandler Grid.DataBindingComplete, AddressOf DataGridView_DataBindingComplete
        AddHandler Grid.SelectionChanged, AddressOf DataGridView_SelectionChanged
        AddHandler Grid.RowsAdded, AddressOf DataGridView_RowsAdded
        AddHandler Grid.RowsRemoved, AddressOf DataGridView_RowsRemoved
        AddHandler Grid.RowStateChanged, AddressOf DataGridView_RowStateChanged
        AddHandler Grid.Sorted, AddressOf DataGridView_Sorted
        AddHandler Grid.Disposed, AddressOf DataGridView_Disposed
        AddHandler Grid.HandleCreated, AddressOf DataGridView_HandleCreated
    End Sub

    Private Sub DetachGrid(Grid As DataGridView)
        If Grid Is Nothing Then Return

        RemoveHandler Grid.DataSourceChanged, AddressOf DataGridView_DataSourceChanged
        RemoveHandler Grid.DataBindingComplete, AddressOf DataGridView_DataBindingComplete
        RemoveHandler Grid.SelectionChanged, AddressOf DataGridView_SelectionChanged
        RemoveHandler Grid.RowsAdded, AddressOf DataGridView_RowsAdded
        RemoveHandler Grid.RowsRemoved, AddressOf DataGridView_RowsRemoved
        RemoveHandler Grid.RowStateChanged, AddressOf DataGridView_RowStateChanged
        RemoveHandler Grid.Sorted, AddressOf DataGridView_Sorted
        RemoveHandler Grid.Disposed, AddressOf DataGridView_Disposed
        RemoveHandler Grid.HandleCreated, AddressOf DataGridView_HandleCreated
    End Sub

    Private Sub GridStructureChanged(sender As Object)
        If Not ReferenceEquals(sender, _DataGridView) OrElse _IsDisposed Then Return
        RefreshButtons()
    End Sub

    Private Sub DataGridView_DataSourceChanged(sender As Object, e As EventArgs)
        If ReferenceEquals(sender, _DataGridView) Then _PendingSelectionRestoreIndex = -1
        GridStructureChanged(sender)
    End Sub

    Private Sub DataGridView_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
        GridStructureChanged(sender)
        If ReferenceEquals(sender, _DataGridView) Then QueueSelectionRestore()
    End Sub

    Private Sub DataGridView_HandleCreated(sender As Object, e As EventArgs)
        If ReferenceEquals(sender, _DataGridView) Then QueueSelectionRestore()
    End Sub

    Private Sub DataGridView_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs)
        GridStructureChanged(sender)
    End Sub

    Private Sub DataGridView_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs)
        If ReferenceEquals(sender, _DataGridView) Then RequestSelectionRestore(e.RowIndex)
        GridStructureChanged(sender)
    End Sub

    Private Sub DataGridView_RowStateChanged(sender As Object, e As DataGridViewRowStateChangedEventArgs)
        If (e.StateChanged And DataGridViewElementStates.Visible) <> DataGridViewElementStates.None Then
            GridStructureChanged(sender)
        End If
    End Sub

    Private Sub DataGridView_Sorted(sender As Object, e As EventArgs)
        GridStructureChanged(sender)
    End Sub

    Private Sub DataGridView_SelectionChanged(sender As Object, e As EventArgs)
        If Not ReferenceEquals(sender, _DataGridView) OrElse _IsDisposed Then Return
        RefreshButtons()
    End Sub

    Private Sub DataGridView_Disposed(sender As Object, e As EventArgs)
        If Not ReferenceEquals(sender, _DataGridView) Then Return
        DetachGrid(_DataGridView)
        _GridGeneration += 1
        _PendingSelectionRestoreIndex = -1
        _SelectionRestoreQueued = False
        _DataGridView = Nothing
        RefreshButtons()
    End Sub

    Private Async Sub FirstButton_Click(sender As Object, e As EventArgs)
        Try
            Await MoveToFirstAsync()
        Catch ex As Exception
            ReportNavigationFailure(NavigationDirection.First, ex)
        End Try
    End Sub

    Private Async Sub PreviousButton_Click(sender As Object, e As EventArgs)
        Try
            Await MoveToPreviousAsync()
        Catch ex As Exception
            ReportNavigationFailure(NavigationDirection.Previous, ex)
        End Try
    End Sub

    Private Async Sub NextButton_Click(sender As Object, e As EventArgs)
        Try
            Await MoveToNextAsync()
        Catch ex As Exception
            ReportNavigationFailure(NavigationDirection.[Next], ex)
        End Try
    End Sub

    Private Async Sub LastButton_Click(sender As Object, e As EventArgs)
        Try
            Await MoveToLastAsync()
        Catch ex As Exception
            ReportNavigationFailure(NavigationDirection.Last, ex)
        End Try
    End Sub

    Private Sub ReportNavigationFailure(Direction As NavigationDirection, ex As Exception)
        ' A click must not surface an unhandled exception through Async Sub.
        Try
            Trace.TraceError("DataGridViewNavigator: navigation {0} failed: {1}", Direction, ex)
        Catch
            ' Trace listeners are external; reporting failures must not escape Async Sub.
        End Try
        Try
            RaiseEvent NavigationFailed(Me, New NavigationErrorEventArgs(Direction, ex))
        Catch notificationError As Exception
            ' Client event handlers must not crash the WinForms synchronization context.
            Try
                Trace.TraceError("DataGridViewNavigator: NavigationFailed handler failed: {0}", notificationError)
            Catch
                ' Diagnostics are best-effort; the original callback error was handled.
            End Try
        End Try
    End Sub
End Class
