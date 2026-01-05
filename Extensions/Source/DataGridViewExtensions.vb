Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Imports System.Windows.Forms

Namespace Extensions
    Public Module DataGridViewExtensions
        ''' <summary>
        ''' Extension method to populate a <see cref="DataGridView"/> from a collection,
        ''' preserving selection and scroll position.
        ''' Optionally adds an extra column with item order as the first column.
        ''' </summary>
        ''' <param name="Dgv">The DataGridView to populate.</param>
        ''' <param name="Collection">The collection to bind to the DataGridView.</param>
        ''' <param name="KeepSelection">If true, preserves the currently selected row.</param>
        ''' <param name="KeepSort">If true, preserves the current sort column and order.</param>
        ''' <param name="ShowOrder">If true, adds a column displaying the order of the items as the first column.</param>
        <Extension()>
        Public Sub Fill(Dgv As DataGridView, Collection As IEnumerable, Optional KeepSelection As Boolean = True, Optional KeepSort As Boolean = True, Optional ShowOrder As Boolean = True)
            Dim SelectedIndex As Integer = -1
            Dim FirstDisplayed As Integer = -1
            Dim SortColumn As DataGridViewColumn = Nothing
            Dim SortOrder As SortOrder = SortOrder.None
            If KeepSelection AndAlso Dgv.SelectedRows.Count = 1 Then
                SelectedIndex = Dgv.SelectedRows(0).Index
                FirstDisplayed = Dgv.FirstDisplayedScrollingRowIndex
            End If
            If KeepSort AndAlso Dgv.SortedColumn IsNot Nothing Then
                SortColumn = Dgv.SortedColumn
                SortOrder = If(Dgv.SortOrder = SortOrder.Descending, SortOrder.Descending, SortOrder.Ascending)
            End If
            Dim Table As DataTable = Collection.ToTable()
            If ShowOrder Then
                Table.Columns.Add("Order", GetType(Integer))
                For i As Integer = 0 To Table.Rows.Count - 1
                    Table.Rows(i)("Order") = i + 1
                Next
                Table.Columns("Order").SetOrdinal(0)
            End If
            Dgv.DataSource = Table
            If KeepSort AndAlso SortColumn IsNot Nothing AndAlso SortOrder <> SortOrder.None Then
                Dim Direction As ListSortDirection = If(SortOrder = SortOrder.Descending, ListSortDirection.Descending, ListSortDirection.Ascending)
                Dgv.Sort(Dgv.Columns(SortColumn.Index), Direction)
            End If
            If KeepSelection AndAlso Dgv.Rows.Count > 0 Then
                If SelectedIndex >= 0 AndAlso SelectedIndex < Dgv.Rows.Count Then
                    Dgv.Rows(SelectedIndex).Selected = True
                ElseIf Dgv.Rows.Count > 0 Then
                    Dgv.Rows(Dgv.Rows.Count - 1).Selected = True
                End If
                If FirstDisplayed >= 0 AndAlso FirstDisplayed < Dgv.Rows.Count Then
                    Dgv.FirstDisplayedScrollingRowIndex = FirstDisplayed
                End If
            End If
        End Sub
    End Module
End Namespace
