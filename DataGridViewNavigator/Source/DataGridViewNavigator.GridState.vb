Imports System.Windows.Forms

Public Partial Class DataGridViewNavigator

    Private Function HasUsableGrid() As Boolean
        Return _DataGridView IsNot Nothing AndAlso Not _DataGridView.IsDisposed AndAlso Not _DataGridView.Disposing
    End Function

    Private Function GetSelectedRowIndex() As Integer
        If Not HasUsableGrid() Then Return -1

        If _DataGridView.SelectedRows.Count > 0 Then
            Dim Index As Integer = _DataGridView.SelectedRows(0).Index
            If CanNavigateToRow(Index) Then Return Index
        End If

        Dim CurrentRow As DataGridViewRow = _DataGridView.CurrentRow
        If CurrentRow IsNot Nothing AndAlso CanNavigateToRow(CurrentRow.Index) Then Return CurrentRow.Index
        Return -1
    End Function

    Private Function GetFirstRowIndex() As Integer
        If Not HasUsableGrid() Then Return -1
        For Index As Integer = 0 To _DataGridView.Rows.Count - 1
            If CanNavigateToRow(Index) Then Return Index
        Next
        Return -1
    End Function

    Private Function GetPreviousRowIndex(CurrentIndex As Integer) As Integer
        If Not HasUsableGrid() Then Return -1
        For Index As Integer = CurrentIndex - 1 To 0 Step -1
            If CanNavigateToRow(Index) Then Return Index
        Next
        Return -1
    End Function

    Private Function GetNextRowIndex(CurrentIndex As Integer) As Integer
        If Not HasUsableGrid() Then Return -1
        For Index As Integer = CurrentIndex + 1 To _DataGridView.Rows.Count - 1
            If CanNavigateToRow(Index) Then Return Index
        Next
        Return -1
    End Function

    Private Function GetLastRowIndex() As Integer
        If Not HasUsableGrid() Then Return -1
        For Index As Integer = _DataGridView.Rows.Count - 1 To 0 Step -1
            If CanNavigateToRow(Index) Then Return Index
        Next
        Return -1
    End Function

    Private Function CanNavigateToRow(RowIndex As Integer) As Boolean
        If Not HasUsableGrid() Then Return False
        If RowIndex < 0 OrElse RowIndex >= _DataGridView.Rows.Count Then Return False

        Dim Row As DataGridViewRow = _DataGridView.Rows(RowIndex)
        Return Not Row.IsNewRow AndAlso Row.Visible
    End Function

    ''' <summary>Scrolls the attached grid to make the specified row visible.</summary>
    Public Sub EnsureVisibleRow(RowToShow As Integer)
        If Not HasUsableGrid() Then Return
        EnsureVisibleRow(_DataGridView, RowToShow)
    End Sub

    ''' <summary>Scrolls a grid to make the specified row visible, when possible.</summary>
    Public Shared Sub EnsureVisibleRow(Grid As DataGridView, RowToShow As Integer)
        If Grid Is Nothing Then Throw New ArgumentNullException(NameOf(Grid))
        If Grid.IsDisposed OrElse Grid.Disposing Then Return
        If RowToShow < 0 OrElse RowToShow >= Grid.Rows.Count Then Return

        Dim Row As DataGridViewRow = Grid.Rows(RowToShow)
        If Row.IsNewRow OrElse Not Row.Visible OrElse Row.Displayed OrElse Row.Frozen Then Return
        If Not Grid.IsHandleCreated Then Return

        Grid.FirstDisplayedScrollingRowIndex = RowToShow
    End Sub

    ''' <summary>
    ''' Updates navigation button states and the position label. Called automatically
    ''' for selection, binding, row additions/removals, visibility and sorting changes.
    ''' </summary>
    Public Sub RefreshButtons()
        If _IsDisposed Then Return

        If Not HasUsableGrid() Then
            SetButtonsEnabled(False)
            UpdatePositionLabel(0, 0)
            Return
        End If

        Dim CurrentIndex As Integer = GetSelectedRowIndex()
        Dim Total As Integer = 0
        Dim Position As Integer = 0
        Dim FirstIndex As Integer = -1
        Dim LastIndex As Integer = -1
        Dim PreviousExists As Boolean = False
        Dim NextExists As Boolean = False

        For Index As Integer = 0 To _DataGridView.Rows.Count - 1
            If Not CanNavigateToRow(Index) Then Continue For

            Total += 1
            If FirstIndex < 0 Then FirstIndex = Index
            LastIndex = Index
            If Index = CurrentIndex Then Position = Total
            If CurrentIndex >= 0 Then
                If Index < CurrentIndex Then PreviousExists = True
                If Index > CurrentIndex Then NextExists = True
            End If
        Next

        UpdatePositionLabel(Position, Total)
        If _IsNavigating Then
            SetButtonsEnabled(False)
            Return
        End If

        ' When there are rows but no selection, navigation is still possible.
        SetButtonEnabled(_FirstButton, FirstIndex >= 0 AndAlso CurrentIndex <> FirstIndex)
        SetButtonEnabled(_PreviousButton, If(CurrentIndex < 0, LastIndex >= 0, PreviousExists))
        SetButtonEnabled(_NextButton, If(CurrentIndex < 0, FirstIndex >= 0, NextExists))
        SetButtonEnabled(_LastButton, LastIndex >= 0 AndAlso CurrentIndex <> LastIndex)
    End Sub

    Private Sub UpdatePositionLabel(Position As Integer, Total As Integer)
        If _PositionLabel IsNot Nothing AndAlso Not _PositionLabel.IsDisposed Then
            _PositionLabel.Text = String.Format(_PositionFormat, Position, Total)
        End If
    End Sub

    Private Shared Sub SetButtonEnabled(Button As ToolStripButton, Enabled As Boolean)
        If Button IsNot Nothing AndAlso Not Button.IsDisposed Then Button.Enabled = Enabled
    End Sub

    Private Sub SetButtonsEnabled(Enabled As Boolean)
        SetButtonEnabled(_FirstButton, Enabled)
        SetButtonEnabled(_PreviousButton, Enabled)
        SetButtonEnabled(_NextButton, Enabled)
        SetButtonEnabled(_LastButton, Enabled)
    End Sub
End Class
