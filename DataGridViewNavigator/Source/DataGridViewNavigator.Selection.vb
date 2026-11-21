Imports System.Windows.Forms

Public Partial Class DataGridViewNavigator

    ''' <summary>
    ''' Row removal can temporarily leave a bound grid without an active selection.
    ''' Restore the nearest surviving row AFTER the DataGridView finishes updating
    ''' its currency/rows; never attempt to move CurrentCell inside RowsRemoved.
    ''' </summary>
    Private Sub RequestSelectionRestore(RemovedRowIndex As Integer)
        If _IsDisposed OrElse Not HasUsableGrid() Then Return
        If _PendingSelectionRestoreIndex < 0 Then
            _PendingSelectionRestoreIndex = RemovedRowIndex
        Else
            _PendingSelectionRestoreIndex = Math.Min(_PendingSelectionRestoreIndex, RemovedRowIndex)
        End If
        QueueSelectionRestore()
    End Sub

    Private Sub QueueSelectionRestore()
        If _IsDisposed OrElse _SelectionRestoreQueued OrElse
           _PendingSelectionRestoreIndex < 0 OrElse Not HasUsableGrid() Then Return

        Dim Grid As DataGridView = _DataGridView
        If Not Grid.IsHandleCreated Then Return

        Dim Generation As Integer = _GridGeneration
        _SelectionRestoreQueued = True
        Try
            Grid.BeginInvoke(New MethodInvoker(Sub()
                                                   If Generation <> _GridGeneration OrElse
                                                      Not ReferenceEquals(Grid, _DataGridView) Then Return
                                                   _SelectionRestoreQueued = False
                                                   RestoreSelectionAfterRemoval()
                                               End Sub))
        Catch ex As InvalidOperationException
            _SelectionRestoreQueued = False
            Trace.TraceInformation("DataGridViewNavigator: selection recovery deferred: {0}", ex.Message)
        End Try
    End Sub

    Private Sub RestoreSelectionAfterRemoval()
        If _IsDisposed OrElse Not HasUsableGrid() OrElse
           _PendingSelectionRestoreIndex < 0 Then Return
        If _IsNavigating Then Return

        Dim Grid As DataGridView = _DataGridView
        If Grid.SelectedRows.Count > 0 Then
            Dim selectedIndex As Integer = Grid.SelectedRows(0).Index
            If CanNavigateToRow(selectedIndex) Then
                _PendingSelectionRestoreIndex = -1
                RefreshButtons()
                Return
            End If
        End If

        Dim Closest As Integer = GetNearestNavigableIndex(_PendingSelectionRestoreIndex)
        If Closest < 0 Then
            _PendingSelectionRestoreIndex = -1
            RefreshButtons()
            Return
        End If

        Try
            SelectRow(Grid, Closest)
            EnsureVisibleRow(Grid, Closest)
            _PendingSelectionRestoreIndex = -1
            RefreshButtons()
        Catch ex As InvalidOperationException
            Trace.TraceWarning("DataGridViewNavigator: unable to restore selection yet: {0}", ex)
        End Try
    End Sub

    Private Function GetNearestNavigableIndex(PreferredIndex As Integer) As Integer
        If Not HasUsableGrid() Then Return -1
        Dim LastIndex As Integer = _DataGridView.Rows.Count - 1
        Dim StartIndex As Integer = Math.Min(Math.Max(0, PreferredIndex), LastIndex)
        For Index As Integer = StartIndex To LastIndex
            If CanNavigateToRow(Index) Then Return Index
        Next
        For Index As Integer = StartIndex - 1 To 0 Step -1
            If CanNavigateToRow(Index) Then Return Index
        Next
        Return -1
    End Function

End Class
