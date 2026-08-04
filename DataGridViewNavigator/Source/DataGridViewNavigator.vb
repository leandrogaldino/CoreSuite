Imports System.ComponentModel
''' <summary>
''' Provides navigation functionality for a <see cref="DataGridView"/> using four
''' <see cref="ToolStripButton"/> controls (First, Previous, Next and Last).
''' </summary>
''' <remarks>
''' <para>
''' When associated with a DataGridView, this component automatically configures
''' <see cref="DataGridView.SelectionMode"/> to
''' <see cref="DataGridViewSelectionMode.FullRowSelect"/> and
''' <see cref="DataGridView.MultiSelect"/> to <c>False</c>.
''' </para>
''' <para>
''' Optional callbacks can be assigned to <see cref="ActionBeforeMove"/> and
''' <see cref="ActionAfterMove"/> to execute custom logic before and after navigation.
''' </para>
''' </remarks>
Public Class DataGridViewNavigator
    Inherits Component
    Private _DataGridView As New DataGridView
    Private _FirstButton As ToolStripButton
    Private _PreviousButton As ToolStripButton
    Private _NextButton As ToolStripButton
    Private _LastButton As ToolStripButton
    ''' <summary>
    ''' Gets or sets the action executed immediately before a navigation operation.
    ''' </summary>
    ''' <remarks>
    ''' This action can be used to validate pending changes before changing the selected row.
    ''' Set <see cref="CancelNextMove"/> to <c>True</c> to cancel the navigation.
    ''' </remarks>
    Public Property ActionBeforeMove As Action
    ''' <summary>
    ''' Gets or sets the action executed immediately after a successful navigation operation.
    ''' </summary>
    ''' <remarks>
    ''' This action is invoked after the selected row has changed.
    ''' </remarks>
    Public Property ActionAfterMove As Action
    ''' <summary>
    ''' The DataGridView that will be assigned the navigation functionalities.
    ''' </summary>
    Public Property DataGridView As DataGridView
        Get
            Return _DataGridView
        End Get
        Set(value As DataGridView)
            _DataGridView = value
            If _DataGridView IsNot Nothing Then
                AddHandler DataGridView.DataSourceChanged, AddressOf DataGridView_DataSourceChanged
                AddHandler DataGridView.RowEnter, AddressOf DataGridView_RowEnter
                If IsDefinedButtons() Then RefreshButtons()
            End If
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the next navigation operation should be canceled.
    ''' </summary>
    ''' <remarks>
    ''' This property is typically set inside <see cref="ActionBeforeMove"/> to prevent the navigation
    ''' from continuing. After each navigation attempt, its value is automatically reset to <c>False</c>.
    ''' </remarks>
    Public Property CancelNextMove As Boolean
    Private Sub DataGridView_DataSourceChanged(sender As Object, e As EventArgs)
        If IsDefinedButtons() Then RefreshButtons()
    End Sub
    Private Sub DataGridView_RowEnter(sender As Object, e As DataGridViewCellEventArgs)
        If DataGridView.MultiSelect Then DataGridView.MultiSelect = False
        If DataGridView.SelectionMode <> DataGridViewSelectionMode.FullRowSelect Then DataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        If IsDefinedButtons() And DataGridView.SelectedRows.Count > 0 Then
            RefreshButtons()
        End If
    End Sub
    Private Function IsDefinedButtons() As Boolean
        If FirstButton IsNot Nothing And
                PreviousButton IsNot Nothing And
                NextButton IsNot Nothing And
                LastButton IsNot Nothing Then
            Return True
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' Gets or sets the button used to navigate to the first row.
    ''' </summary>
    Public Property FirstButton As ToolStripButton
        Get
            Return _FirstButton
        End Get
        Set(value As ToolStripButton)
            _FirstButton = value
            If _FirstButton IsNot Nothing Then
                _FirstButton = value
                If IsDefinedButtons() Then RefreshButtons()
                AddHandler FirstButton.Click, AddressOf FirstButton_Click
            End If
        End Set
    End Property
    Private Sub FirstButton_Click(sender As Object, e As EventArgs)
        MoveToFirst()
    End Sub
    ''' <summary>
    ''' Gets or sets the button used to navigate to the previous row.
    ''' </summary>
    Public Property PreviousButton As ToolStripButton
        Get
            Return _PreviousButton
        End Get
        Set(value As ToolStripButton)
            _PreviousButton = value
            If _PreviousButton IsNot Nothing Then
                _PreviousButton = value
                If IsDefinedButtons() Then RefreshButtons()
                AddHandler PreviousButton.Click, AddressOf PreviousButton_Click
            End If
        End Set
    End Property
    Private Sub PreviousButton_Click(sender As Object, e As EventArgs)
        MoveToPrevious()
    End Sub
    ''' <summary>
    ''' Gets or sets the button used to navigate to the next row.
    ''' </summary>
    Public Property NextButton As ToolStripButton
        Get
            Return _NextButton
        End Get
        Set(value As ToolStripButton)
            _NextButton = value
            If NextButton IsNot Nothing Then
                _NextButton = value
                If IsDefinedButtons() Then RefreshButtons()
                AddHandler NextButton.Click, AddressOf NextButton_Click
            End If
        End Set
    End Property
    Private Sub NextButton_Click(sender As Object, e As EventArgs)
        MoveToNext()
    End Sub
    ''' <summary>
    ''' Gets or sets the button used to navigate to the last row.
    ''' </summary>
    Public Property LastButton As ToolStripButton
        Get
            Return _LastButton
        End Get
        Set(value As ToolStripButton)
            _LastButton = value
            If _LastButton IsNot Nothing Then
                _LastButton = value
                If IsDefinedButtons() Then RefreshButtons()
                AddHandler LastButton.Click, AddressOf LastButton_Click
            End If
        End Set
    End Property
    Private Sub LastButton_Click(sender As Object, e As EventArgs)
        MoveToLast()
    End Sub
    ''' <summary>
    ''' Ensures that the specified row is visible within the DataGridView.
    ''' </summary>
    ''' <param name="RowToShow">
    ''' The zero-based index of the row to display.
    ''' </param>
    Public Sub EnsureVisibleRow(RowToShow As Integer)
        If RowToShow >= 0 AndAlso RowToShow < _DataGridView.RowCount Then
            _DataGridView.Rows(RowToShow).Selected = True
            Dim CountVisible = _DataGridView.DisplayedRowCount(False)
            Dim FirstVisible = _DataGridView.FirstDisplayedScrollingRowIndex
            If RowToShow < FirstVisible Then
                _DataGridView.FirstDisplayedScrollingRowIndex = RowToShow
            ElseIf RowToShow >= FirstVisible + CountVisible Then
                _DataGridView.FirstDisplayedScrollingRowIndex = RowToShow - CountVisible + If(CountVisible > 0, 1, 0)
            End If
        End If
    End Sub
    ''' <summary>
    ''' Ensures that the specified row is visible in the provided <see cref="DataGridView"/>.
    ''' </summary>
    ''' <param name="Dgv">
    ''' The <see cref="DataGridView"/> whose row should be brought into view.
    ''' </param>
    ''' <param name="RowToShow">
    ''' The zero-based index of the row to display.
    ''' </param>
    Public Shared Sub EnsureVisibleRow(Dgv As DataGridView, RowToShow As Integer)
        If RowToShow >= 0 AndAlso RowToShow < Dgv.RowCount Then
            Dgv.Rows(RowToShow).Selected = True
            Dim CountVisible = Dgv.DisplayedRowCount(False)
            Dim FirstVisible = Dgv.FirstDisplayedScrollingRowIndex
            If RowToShow < FirstVisible Then
                Dgv.FirstDisplayedScrollingRowIndex = RowToShow
            ElseIf RowToShow >= FirstVisible + CountVisible Then
                Dgv.FirstDisplayedScrollingRowIndex = RowToShow - CountVisible + If(CountVisible > 0, 1, 0)
            End If
        End If
    End Sub
    ''' <summary>
    ''' Updates the enabled state of the navigation buttons according to the currently selected row.
    ''' </summary>
    Public Sub RefreshButtons()
        If _FirstButton IsNot Nothing And _PreviousButton IsNot Nothing And _NextButton IsNot Nothing And _LastButton IsNot Nothing Then
            If _DataGridView.SelectedRows.Count > 0 Then
                _FirstButton.Enabled = _DataGridView.SelectedRows(0).Index > 0
                _PreviousButton.Enabled = _DataGridView.SelectedRows(0).Index > 0
                _NextButton.Enabled = _DataGridView.SelectedRows(0).Index < _DataGridView.Rows.Count - 1
                _LastButton.Enabled = _DataGridView.SelectedRows(0).Index < _DataGridView.Rows.Count - 1
            Else
                _FirstButton.Enabled = False
                _PreviousButton.Enabled = False
                _NextButton.Enabled = False
                _LastButton.Enabled = False
            End If
        End If
    End Sub
    ''' <summary>
    ''' Moves to the first row of the DataGridView.
    ''' </summary>
    Public Sub MoveToFirst()
        If ActionBeforeMove IsNot Nothing Then ActionBeforeMove.Invoke
        If Not CancelNextMove Then
            If _DataGridView.SelectedRows.Count = 1 Then
                If _DataGridView.SelectedRows(0).Index > 0 Then
                    _DataGridView.Rows(0).Selected = True
                    RefreshButtons()
                    If _DataGridView.SelectedRows.Count > 0 Then EnsureVisibleRow(_DataGridView.SelectedRows(0).Index)
                End If
            End If
            If ActionAfterMove IsNot Nothing Then ActionAfterMove.Invoke
            CancelNextMove = False
        End If
    End Sub
    ''' <summary>
    ''' Moves to the previous row of the DataGridView.
    ''' </summary>
    Public Sub MoveToPrevious()
        If ActionBeforeMove IsNot Nothing Then ActionBeforeMove.Invoke
        If Not CancelNextMove Then
            If _DataGridView.SelectedRows.Count = 1 Then
                If _DataGridView.SelectedRows(0).Index > 0 Then
                    _DataGridView.Rows(_DataGridView.SelectedRows(0).Index - 1).Selected = True
                    RefreshButtons()
                    If _DataGridView.SelectedRows.Count > 0 Then EnsureVisibleRow(_DataGridView.SelectedRows(0).Index)
                End If
            End If
            If ActionAfterMove IsNot Nothing Then ActionAfterMove.Invoke
            CancelNextMove = False
        End If
    End Sub
    ''' <summary>
    ''' Moves to the next row of the DataGridView.
    ''' </summary>
    Public Sub MoveToNext()
        If ActionBeforeMove IsNot Nothing Then ActionBeforeMove.Invoke
        If Not CancelNextMove Then
            If _DataGridView.SelectedRows.Count = 1 Then
                If _DataGridView.SelectedRows(0).Index < _DataGridView.Rows.Count - 1 Then
                    _DataGridView.Rows(_DataGridView.SelectedRows(0).Index + 1).Selected = True
                    RefreshButtons()
                    If _DataGridView.SelectedRows.Count > 0 Then EnsureVisibleRow(_DataGridView.SelectedRows(0).Index)
                End If
            End If
            If ActionAfterMove IsNot Nothing Then ActionAfterMove.Invoke
            CancelNextMove = False
        End If
    End Sub
    ''' <summary>
    ''' Moves to the last row of the DataGridView.
    ''' </summary>
    Public Sub MoveToLast()
        If ActionBeforeMove IsNot Nothing Then ActionBeforeMove.Invoke
        If Not CancelNextMove Then
            If _DataGridView.SelectedRows.Count = 1 Then
                If _DataGridView.SelectedRows(0).Index < _DataGridView.Rows.Count - 1 Then
                    _DataGridView.Rows(_DataGridView.Rows.Count - 1).Selected = True
                    RefreshButtons()
                    If _DataGridView.SelectedRows.Count > 0 Then EnsureVisibleRow(_DataGridView.SelectedRows(0).Index)
                End If
            End If
            If ActionAfterMove IsNot Nothing Then ActionAfterMove.Invoke
            CancelNextMove = False
        End If
    End Sub
End Class
