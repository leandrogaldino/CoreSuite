Imports System.ComponentModel

''' <summary>
''' Provides navigation functionality for a <see cref="DataGridView"/> using four
''' <see cref="ToolStripButton"/> controls (First, Previous, Next and Last).
''' </summary>
''' <remarks>
''' <para>
''' When associated with a DataGridView, this component automatically configures
''' <see cref="DataGridView.SelectionMode"/> to <see cref="DataGridViewSelectionMode.FullRowSelect"/>
''' and <see cref="DataGridView.MultiSelect"/> to <c>False</c>.
''' </para>
''' <para>
''' Synchronous and asynchronous callbacks can be assigned to execute custom logic
''' before and after navigation.
''' </para>
''' </remarks>
Public Class DataGridViewNavigator
    Inherits Component

    Private _DataGridView As DataGridView
    Private _FirstButton As ToolStripButton
    Private _PreviousButton As ToolStripButton
    Private _NextButton As ToolStripButton
    Private _LastButton As ToolStripButton
    Private _IsNavigating As Boolean

    ''' <summary>
    ''' Gets or sets the synchronous action executed before a navigation operation.
    ''' </summary>
    ''' <remarks>
    ''' Set <see cref="CancelEventArgs.Cancel"/> to <c>True</c> to cancel the navigation.
    ''' </remarks>
    Public Property ActionBeforeMove As Action(Of CancelEventArgs)

    ''' <summary>
    ''' Gets or sets the asynchronous action executed before a navigation operation.
    ''' </summary>
    ''' <remarks>
    ''' Set <see cref="CancelEventArgs.Cancel"/> to <c>True</c> to cancel the navigation.
    ''' </remarks>
    Public Property ActionBeforeMoveAsync As Func(Of CancelEventArgs, Task)

    ''' <summary>
    ''' Gets or sets the synchronous action executed after a successful navigation operation.
    ''' </summary>
    Public Property ActionAfterMove As Action

    ''' <summary>
    ''' Gets or sets the asynchronous action executed after a successful navigation operation.
    ''' </summary>
    Public Property ActionAfterMoveAsync As Func(Of Task)

    ''' <summary>
    ''' Gets a value indicating whether a navigation operation is currently being executed.
    ''' </summary>
    Public ReadOnly Property IsNavigating As Boolean
        Get
            Return _IsNavigating
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets the DataGridView that will be assigned the navigation functionality.
    ''' </summary>
    Public Property DataGridView As DataGridView
        Get
            Return _DataGridView
        End Get
        Set(value As DataGridView)
            If ReferenceEquals(_DataGridView, value) Then Return

            If _DataGridView IsNot Nothing Then
                RemoveHandler _DataGridView.DataSourceChanged, AddressOf DataGridView_DataSourceChanged
                RemoveHandler _DataGridView.SelectionChanged, AddressOf DataGridView_SelectionChanged
            End If

            _DataGridView = value

            If _DataGridView IsNot Nothing Then
                _DataGridView.MultiSelect = False
                _DataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect
                AddHandler _DataGridView.DataSourceChanged, AddressOf DataGridView_DataSourceChanged
                AddHandler _DataGridView.SelectionChanged, AddressOf DataGridView_SelectionChanged
            End If

            RefreshButtons()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the button used to navigate to the first row.
    ''' </summary>
    Public Property FirstButton As ToolStripButton
        Get
            Return _FirstButton
        End Get
        Set(value As ToolStripButton)
            If ReferenceEquals(_FirstButton, value) Then Return
            If _FirstButton IsNot Nothing Then RemoveHandler _FirstButton.Click, AddressOf FirstButton_Click

            _FirstButton = value

            If _FirstButton IsNot Nothing Then AddHandler _FirstButton.Click, AddressOf FirstButton_Click
            RefreshButtons()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the button used to navigate to the previous row.
    ''' </summary>
    Public Property PreviousButton As ToolStripButton
        Get
            Return _PreviousButton
        End Get
        Set(value As ToolStripButton)
            If ReferenceEquals(_PreviousButton, value) Then Return
            If _PreviousButton IsNot Nothing Then RemoveHandler _PreviousButton.Click, AddressOf PreviousButton_Click

            _PreviousButton = value

            If _PreviousButton IsNot Nothing Then AddHandler _PreviousButton.Click, AddressOf PreviousButton_Click
            RefreshButtons()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the button used to navigate to the next row.
    ''' </summary>
    Public Property NextButton As ToolStripButton
        Get
            Return _NextButton
        End Get
        Set(value As ToolStripButton)
            If ReferenceEquals(_NextButton, value) Then Return
            If _NextButton IsNot Nothing Then RemoveHandler _NextButton.Click, AddressOf NextButton_Click

            _NextButton = value

            If _NextButton IsNot Nothing Then AddHandler _NextButton.Click, AddressOf NextButton_Click
            RefreshButtons()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the button used to navigate to the last row.
    ''' </summary>
    Public Property LastButton As ToolStripButton
        Get
            Return _LastButton
        End Get
        Set(value As ToolStripButton)
            If ReferenceEquals(_LastButton, value) Then Return
            If _LastButton IsNot Nothing Then RemoveHandler _LastButton.Click, AddressOf LastButton_Click

            _LastButton = value

            If _LastButton IsNot Nothing Then AddHandler _LastButton.Click, AddressOf LastButton_Click
            RefreshButtons()
        End Set
    End Property

    Private Sub DataGridView_DataSourceChanged(sender As Object, e As EventArgs)
        RefreshButtons()
    End Sub

    Private Sub DataGridView_SelectionChanged(sender As Object, e As EventArgs)
        RefreshButtons()
    End Sub

    Private Async Sub FirstButton_Click(sender As Object, e As EventArgs)
        Await MoveToFirstAsync()
    End Sub

    Private Async Sub PreviousButton_Click(sender As Object, e As EventArgs)
        Await MoveToPreviousAsync()
    End Sub

    Private Async Sub NextButton_Click(sender As Object, e As EventArgs)
        Await MoveToNextAsync()
    End Sub

    Private Async Sub LastButton_Click(sender As Object, e As EventArgs)
        Await MoveToLastAsync()
    End Sub

    ''' <summary>
    ''' Moves synchronously to the first row.
    ''' </summary>
    ''' <remarks>
    ''' This method cannot be used when asynchronous navigation callbacks are configured.
    ''' </remarks>
    Public Sub MoveToFirst()
        EnsureSynchronousNavigation()
        MoveToFirstAsync().GetAwaiter().GetResult()
    End Sub

    ''' <summary>
    ''' Moves synchronously to the previous row.
    ''' </summary>
    ''' <remarks>
    ''' This method cannot be used when asynchronous navigation callbacks are configured.
    ''' </remarks>
    Public Sub MoveToPrevious()
        EnsureSynchronousNavigation()
        MoveToPreviousAsync().GetAwaiter().GetResult()
    End Sub

    ''' <summary>
    ''' Moves synchronously to the next row.
    ''' </summary>
    ''' <remarks>
    ''' This method cannot be used when asynchronous navigation callbacks are configured.
    ''' </remarks>
    Public Sub MoveToNext()
        EnsureSynchronousNavigation()
        MoveToNextAsync().GetAwaiter().GetResult()
    End Sub

    ''' <summary>
    ''' Moves synchronously to the last row.
    ''' </summary>
    ''' <remarks>
    ''' This method cannot be used when asynchronous navigation callbacks are configured.
    ''' </remarks>
    Public Sub MoveToLast()
        EnsureSynchronousNavigation()
        MoveToLastAsync().GetAwaiter().GetResult()
    End Sub

    ''' <summary>
    ''' Moves asynchronously to the first row.
    ''' </summary>
    Public Async Function MoveToFirstAsync() As Task
        Await MoveToAsync(GetFirstRowIndex())
    End Function

    ''' <summary>
    ''' Moves asynchronously to the previous row.
    ''' </summary>
    Public Async Function MoveToPreviousAsync() As Task
        Dim CurrentIndex = GetSelectedRowIndex()
        If CurrentIndex < 0 Then Return
        Await MoveToAsync(GetPreviousRowIndex(CurrentIndex))
    End Function

    ''' <summary>
    ''' Moves asynchronously to the next row.
    ''' </summary>
    Public Async Function MoveToNextAsync() As Task
        Dim CurrentIndex = GetSelectedRowIndex()
        If CurrentIndex < 0 Then Return
        Await MoveToAsync(GetNextRowIndex(CurrentIndex))
    End Function

    ''' <summary>
    ''' Moves asynchronously to the last row.
    ''' </summary>
    Public Async Function MoveToLastAsync() As Task
        Await MoveToAsync(GetLastRowIndex())
    End Function

    Private Async Function MoveToAsync(RowIndex As Integer) As Task
        If _IsNavigating OrElse _DataGridView Is Nothing Then Return
        If Not CanNavigateToRow(RowIndex) Then Return
        If GetSelectedRowIndex() = RowIndex Then Return

        _IsNavigating = True
        RefreshButtons()

        Try
            If Not Await ExecuteBeforeMoveAsync() Then Return

            SelectRow(RowIndex)
            EnsureVisibleRow(RowIndex)

            Await ExecuteAfterMoveAsync()
        Finally
            _IsNavigating = False
            RefreshButtons()
        End Try
    End Function

    Private Async Function ExecuteBeforeMoveAsync() As Task(Of Boolean)
        Dim E As New CancelEventArgs()

        ActionBeforeMove?.Invoke(E)
        If E.Cancel Then Return False

        If ActionBeforeMoveAsync IsNot Nothing Then Await ActionBeforeMoveAsync.Invoke(E)

        Return Not E.Cancel
    End Function

    Private Async Function ExecuteAfterMoveAsync() As Task
        ActionAfterMove?.Invoke()
        If ActionAfterMoveAsync IsNot Nothing Then Await ActionAfterMoveAsync.Invoke()
    End Function

    Private Sub SelectRow(RowIndex As Integer)
        Dim Row = _DataGridView.Rows(RowIndex)
        Dim CellToSelect As DataGridViewCell = Nothing

        For Each Cell As DataGridViewCell In Row.Cells
            If Cell.Visible Then
                CellToSelect = Cell
                Exit For
            End If
        Next

        _DataGridView.ClearSelection()

        If CellToSelect IsNot Nothing Then _DataGridView.CurrentCell = CellToSelect
        Row.Selected = True
    End Sub

    Private Function GetSelectedRowIndex() As Integer
        If _DataGridView Is Nothing Then Return -1

        If _DataGridView.SelectedRows.Count > 0 Then
            Dim Index = _DataGridView.SelectedRows(0).Index
            If CanNavigateToRow(Index) Then Return Index
        End If

        If _DataGridView.CurrentRow IsNot Nothing AndAlso CanNavigateToRow(_DataGridView.CurrentRow.Index) Then Return _DataGridView.CurrentRow.Index

        Return -1
    End Function

    Private Function GetFirstRowIndex() As Integer
        If _DataGridView Is Nothing Then Return -1

        For Index = 0 To _DataGridView.Rows.Count - 1
            If CanNavigateToRow(Index) Then Return Index
        Next

        Return -1
    End Function

    Private Function GetPreviousRowIndex(CurrentIndex As Integer) As Integer
        If _DataGridView Is Nothing Then Return -1

        For Index = CurrentIndex - 1 To 0 Step -1
            If CanNavigateToRow(Index) Then Return Index
        Next

        Return -1
    End Function

    Private Function GetNextRowIndex(CurrentIndex As Integer) As Integer
        If _DataGridView Is Nothing Then Return -1

        For Index = CurrentIndex + 1 To _DataGridView.Rows.Count - 1
            If CanNavigateToRow(Index) Then Return Index
        Next

        Return -1
    End Function

    Private Function GetLastRowIndex() As Integer
        If _DataGridView Is Nothing Then Return -1

        For Index = _DataGridView.Rows.Count - 1 To 0 Step -1
            If CanNavigateToRow(Index) Then Return Index
        Next

        Return -1
    End Function

    Private Function CanNavigateToRow(RowIndex As Integer) As Boolean
        If _DataGridView Is Nothing Then Return False
        If RowIndex < 0 OrElse RowIndex >= _DataGridView.Rows.Count Then Return False

        Dim Row = _DataGridView.Rows(RowIndex)
        Return Not Row.IsNewRow AndAlso Row.Visible
    End Function

    Private Function IsDefinedButtons() As Boolean
        Return _FirstButton IsNot Nothing AndAlso _PreviousButton IsNot Nothing AndAlso _NextButton IsNot Nothing AndAlso _LastButton IsNot Nothing
    End Function

    Private Sub EnsureSynchronousNavigation()
        If ActionBeforeMoveAsync IsNot Nothing OrElse ActionAfterMoveAsync IsNot Nothing Then Throw New InvalidOperationException("Synchronous navigation cannot be used when asynchronous callbacks are configured. Use the corresponding asynchronous navigation method.")
    End Sub

    ''' <summary>
    ''' Ensures that the specified row is visible within the DataGridView.
    ''' </summary>
    Public Sub EnsureVisibleRow(RowToShow As Integer)
        If _DataGridView Is Nothing Then Return
        EnsureVisibleRow(_DataGridView, RowToShow)
    End Sub

    ''' <summary>
    ''' Ensures that the specified row is visible in the provided DataGridView.
    ''' </summary>
    Public Shared Sub EnsureVisibleRow(Dgv As DataGridView, RowToShow As Integer)
        ArgumentNullException.ThrowIfNull(Dgv)
        If RowToShow < 0 OrElse RowToShow >= Dgv.Rows.Count Then Return

        Dim Row = Dgv.Rows(RowToShow)
        If Row.IsNewRow OrElse Not Row.Visible OrElse Row.Displayed Then Return

        Dgv.FirstDisplayedScrollingRowIndex = RowToShow
    End Sub

    ''' <summary>
    ''' Updates the enabled state of the navigation buttons according to the currently selected row.
    ''' </summary>
    Public Sub RefreshButtons()
        If Not IsDefinedButtons() Then Return

        If _DataGridView Is Nothing OrElse _IsNavigating Then
            SetButtonsEnabled(False)
            Return
        End If

        Dim CurrentIndex = GetSelectedRowIndex()

        If CurrentIndex < 0 Then
            SetButtonsEnabled(False)
            Return
        End If

        Dim FirstIndex = GetFirstRowIndex()
        Dim LastIndex = GetLastRowIndex()

        _FirstButton.Enabled = FirstIndex >= 0 AndAlso CurrentIndex <> FirstIndex
        _PreviousButton.Enabled = GetPreviousRowIndex(CurrentIndex) >= 0
        _NextButton.Enabled = GetNextRowIndex(CurrentIndex) >= 0
        _LastButton.Enabled = LastIndex >= 0 AndAlso CurrentIndex <> LastIndex
    End Sub

    Private Sub SetButtonsEnabled(Enabled As Boolean)
        If _FirstButton IsNot Nothing Then _FirstButton.Enabled = Enabled
        If _PreviousButton IsNot Nothing Then _PreviousButton.Enabled = Enabled
        If _NextButton IsNot Nothing Then _NextButton.Enabled = Enabled
        If _LastButton IsNot Nothing Then _LastButton.Enabled = Enabled
    End Sub

    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            If _DataGridView IsNot Nothing Then
                RemoveHandler _DataGridView.DataSourceChanged, AddressOf DataGridView_DataSourceChanged
                RemoveHandler _DataGridView.SelectionChanged, AddressOf DataGridView_SelectionChanged
            End If

            If _FirstButton IsNot Nothing Then RemoveHandler _FirstButton.Click, AddressOf FirstButton_Click
            If _PreviousButton IsNot Nothing Then RemoveHandler _PreviousButton.Click, AddressOf PreviousButton_Click
            If _NextButton IsNot Nothing Then RemoveHandler _NextButton.Click, AddressOf NextButton_Click
            If _LastButton IsNot Nothing Then RemoveHandler _LastButton.Click, AddressOf LastButton_Click
        End If

        MyBase.Dispose(Disposing)
    End Sub
End Class