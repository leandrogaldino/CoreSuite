Imports System.ComponentModel
Imports System.Windows.Forms

Public Partial Class DataGridViewNavigator

    ''' <summary>Moves synchronously to the first visible, non-new row.</summary>
    Public Sub MoveToFirst()
        EnsureSynchronousNavigation()
        MoveToFirstAsync().GetAwaiter().GetResult()
    End Sub

    ''' <summary>Moves synchronously to the previous visible, non-new row.</summary>
    Public Sub MoveToPrevious()
        EnsureSynchronousNavigation()
        MoveToPreviousAsync().GetAwaiter().GetResult()
    End Sub

    ''' <summary>Moves synchronously to the next visible, non-new row.</summary>
    Public Sub MoveToNext()
        EnsureSynchronousNavigation()
        MoveToNextAsync().GetAwaiter().GetResult()
    End Sub

    ''' <summary>Moves synchronously to the last visible, non-new row.</summary>
    Public Sub MoveToLast()
        EnsureSynchronousNavigation()
        MoveToLastAsync().GetAwaiter().GetResult()
    End Sub

    ''' <summary>Moves asynchronously to the first visible, non-new row.</summary>
    Public Function MoveToFirstAsync() As Task
        Return MoveToAsync(GetFirstRowIndex())
    End Function

    ''' <summary>Moves asynchronously to the previous visible, non-new row.</summary>
    Public Function MoveToPreviousAsync() As Task
        Dim CurrentIndex As Integer = GetSelectedRowIndex()
        ' With no current row, Previous starts at the end, instead of doing nothing.
        Return MoveToAsync(If(CurrentIndex < 0, GetLastRowIndex(), GetPreviousRowIndex(CurrentIndex)))
    End Function

    ''' <summary>Moves asynchronously to the next visible, non-new row.</summary>
    Public Function MoveToNextAsync() As Task
        Dim CurrentIndex As Integer = GetSelectedRowIndex()
        ' With no current row, Next starts at the beginning, instead of doing nothing.
        Return MoveToAsync(If(CurrentIndex < 0, GetFirstRowIndex(), GetNextRowIndex(CurrentIndex)))
    End Function

    ''' <summary>Moves asynchronously to the last visible, non-new row.</summary>
    Public Function MoveToLastAsync() As Task
        Return MoveToAsync(GetLastRowIndex())
    End Function

    Private Async Function MoveToAsync(RowIndex As Integer) As Task
        If _IsDisposed OrElse _IsNavigating OrElse Not HasUsableGrid() Then Return
        If Not CanNavigateToRow(RowIndex) OrElse GetSelectedRowIndex() = RowIndex Then Return

        Dim Grid As DataGridView = _DataGridView
        Dim Source As Object = Grid.DataSource
        Dim Binding As BindingSource = TryCast(Source, BindingSource)
        Dim BindingDataSource As Object = Binding?.DataSource
        Dim BindingDataMember As String = Binding?.DataMember
        Dim OriginalTarget As DataGridViewRow = Grid.Rows(RowIndex)
        Dim TargetItem As Object = OriginalTarget.DataBoundItem

        ' Preserve the source record so a change during an Await is not overwritten.
        ' SelectionChanged event counts can change during ordinary rebinding too.
        Dim OriginIndex As Integer = GetSelectedRowIndex()
        Dim OriginRow As DataGridViewRow = If(OriginIndex >= 0, Grid.Rows(OriginIndex), Nothing)
        Dim OriginItem As Object = OriginRow?.DataBoundItem

        _IsNavigating = True
        Try
            RefreshButtons()
            If Not Await ExecuteBeforeMoveAsync() Then Return
            If _IsDisposed OrElse Not HasUsableGrid() Then Return
            If Not ReferenceEquals(_DataGridView, Grid) OrElse
               Not ReferenceEquals(Grid.DataSource, Source) Then Return
            ' A BindingSource stays the same object when its underlying table changes.
            If Binding IsNot Nothing AndAlso
               (Not ReferenceEquals(Binding.DataSource, BindingDataSource) OrElse
                Binding.DataMember <> BindingDataMember) Then Return

            If Not OriginStillCurrent(Grid, OriginRow, OriginItem) Then Return

            ' Binding may replace DataGridViewRow/DataRowView instances.
            ' Resolve the original record rather than comparing UI row references.
            Dim DestinationIndex As Integer = FindOriginalTargetIndex(Grid, OriginalTarget, TargetItem)
            If DestinationIndex < 0 OrElse GetSelectedRowIndex() = DestinationIndex Then Return

            SelectRow(Grid, DestinationIndex)
            EnsureVisibleRow(Grid, DestinationIndex)
            Await ExecuteAfterMoveAsync()
        Finally
            _IsNavigating = False
            If Not _IsDisposed Then
                RefreshButtons()
                QueueSelectionRestore()
            End If
        End Try
    End Function

    Private Function OriginStillCurrent(Grid As DataGridView, OriginalRow As DataGridViewRow, OriginalItem As Object) As Boolean
        Dim CurrentIndex As Integer = GetSelectedRowIndex()
        If OriginalRow Is Nothing Then Return CurrentIndex < 0
        If CurrentIndex < 0 Then Return False
        Dim CurrentRow As DataGridViewRow = Grid.Rows(CurrentIndex)
        Return SameRecord(OriginalRow, OriginalItem, CurrentRow)
    End Function

    Private Function FindOriginalTargetIndex(Grid As DataGridView, OriginalRow As DataGridViewRow, OriginalItem As Object) As Integer
        For Index As Integer = 0 To Grid.Rows.Count - 1
            If Not CanNavigateToRow(Index) Then Continue For
            If SameRecord(OriginalRow, OriginalItem, Grid.Rows(Index)) Then Return Index
        Next
        Return -1
    End Function

    Private Shared Function SameRecord(OriginalRow As DataGridViewRow, OriginalItem As Object, CandidateRow As DataGridViewRow) As Boolean
        If OriginalItem Is Nothing Then
            ' The row instance is the identity of an unbound row.
            Return ReferenceEquals(OriginalRow, CandidateRow)
        End If

        Dim CandidateItem As Object = CandidateRow.DataBoundItem
        If CandidateItem Is Nothing Then Return False
        If TypeOf OriginalItem Is DataRowView Then
            Dim CandidateView As DataRowView = TryCast(CandidateItem, DataRowView)
            Return CandidateView IsNot Nothing AndAlso ReferenceEquals(DirectCast(OriginalItem, DataRowView).Row, CandidateView.Row)
        End If
        If OriginalItem.GetType().IsValueType OrElse TypeOf OriginalItem Is String Then
            Return Object.Equals(OriginalItem, CandidateItem)
        End If
        Return ReferenceEquals(OriginalItem, CandidateItem)
    End Function

    Private Async Function ExecuteBeforeMoveAsync() As Task(Of Boolean)
        Dim Args As New CancelEventArgs()

        ActionBeforeMove?.Invoke(Args)
        If Args.Cancel Then Return False

        If ActionBeforeMoveAsync IsNot Nothing Then Await ActionBeforeMoveAsync.Invoke(Args)
        Return Not Args.Cancel
    End Function

    Private Async Function ExecuteAfterMoveAsync() As Task
        ActionAfterMove?.Invoke()
        If ActionAfterMoveAsync IsNot Nothing Then Await ActionAfterMoveAsync.Invoke()
    End Function

    Private Shared Sub SelectRow(grid As DataGridView, rowIndex As Integer)
        Dim Row As DataGridViewRow = grid.Rows(rowIndex)
        Dim CellToSelect As DataGridViewCell = Nothing

        For Each Cell As DataGridViewCell In Row.Cells
            If Cell.Visible Then
                CellToSelect = Cell
                Exit For
            End If
        Next

        ' First update CurrentCell, which also synchronizes the CurrencyManager.
        ' Clearing the selection before this can leave CurrentRow stale.
        If CellToSelect IsNot Nothing AndAlso Not ReferenceEquals(grid.CurrentCell, CellToSelect) Then
            grid.CurrentCell = CellToSelect
        End If
        Row.Selected = True
    End Sub

    Private Sub EnsureSynchronousNavigation()
        ThrowIfDisposed()
        If ActionBeforeMoveAsync IsNot Nothing OrElse ActionAfterMoveAsync IsNot Nothing Then
            Throw New InvalidOperationException("Synchronous navigation cannot be used while asynchronous callbacks are configured. Use the corresponding asynchronous method.")
        End If
    End Sub
End Class
