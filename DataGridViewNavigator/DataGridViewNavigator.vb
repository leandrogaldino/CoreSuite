Imports System.ComponentModel
Imports System.Windows.Forms

''' <summary>
''' <para>Associates the DataGridView control with ToolStripButtons to navigate between available rows.</para>
''' <para>The value of the DataGridView's SelectionMode property is changed to FullRowSelect at runtime.<br/>The value of the DataGridView's MultiSelect property is changed to False at runtime.</para>
''' <para>To use this control, you must first define the buttons and the DataGridView. <br/> If you need to handle actions before and after moving between rows, you must assign an action to the ActionBeforeMove and ActionAfterMove fields like: ActionBeforeMove = New Action(AddressOf SomeProcedure).</para>
''' </summary>
Public Class DataGridViewNavigator
    Inherits Component
    Private _DataGridView As New DataGridView
    Private _FirstButton As ToolStripButton
    Private _PreviousButton As ToolStripButton
    Private _NextButton As ToolStripButton
    Private _LastButton As ToolStripButton
    Public ActionBeforeMove As Action
    Public ActionAfterMove As Action

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
    ''' The Button that will be assigned the functionality to navigate to the first row.
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
    ''' The Button that will be assigned the functionality to navigate to the previous row.
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
    ''' The Button that will be assigned the functionality to navigate to the next row.
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
    ''' The Button that will be assigned the functionality to navigate to the last row.
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
    ''' Make sure the specified line will be visible in the DataGridView.
    ''' </summary>
    ''' <param name="RowToShow">The row that will be shown.</param>
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
    ''' Update the navigation buttons, enabling and disabling them if necessary.
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

    Public Property CancelNextMove As Boolean

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