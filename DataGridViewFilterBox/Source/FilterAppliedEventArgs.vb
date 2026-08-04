Imports System.Collections.ObjectModel
''' <summary>
''' Provides data for the <see cref="DataGridViewFilterBox.FilterApplied"/> event.
''' </summary>
Public Class FilterAppliedEventArgs
    Inherits EventArgs
    Private ReadOnly _ColumnNames As ReadOnlyCollection(Of String)
    ''' <summary>
    ''' Initializes a new instance of the <see cref="FilterAppliedEventArgs"/> class.
    ''' </summary>
    ''' <param name="FilterText">The text used to build the filter.</param>
    ''' <param name="FilterExpression">The complete expression assigned to the target <see cref="DataView"/>.</param>
    ''' <param name="MatchedRowCount">The number of rows visible after filtering.</param>
    ''' <param name="ColumnNames">The data columns included in the generated expression.</param>
    Public Sub New(FilterText As String, FilterExpression As String, MatchedRowCount As Integer, ColumnNames As IEnumerable(Of String))
        Me.FilterText = FilterText
        Me.FilterExpression = FilterExpression
        Me.MatchedRowCount = MatchedRowCount
        _ColumnNames = New List(Of String)(ColumnNames).AsReadOnly()
    End Sub
    ''' <summary>
    ''' Gets the text used to build the filter.
    ''' </summary>
    ''' <value>The text entered in the filter box.</value>
    Public ReadOnly Property FilterText As String
    ''' <summary>
    ''' Gets the complete expression assigned to the target <see cref="DataView.RowFilter"/> property.
    ''' </summary>
    ''' <value>The generated expression combined with any previously existing filter.</value>
    Public ReadOnly Property FilterExpression As String
    ''' <summary>
    ''' Gets the number of rows visible after the filter was applied.
    ''' </summary>
    ''' <value>The filtered <see cref="DataView.Count"/> value.</value>
    Public ReadOnly Property MatchedRowCount As Integer
    ''' <summary>
    ''' Gets the data columns included in the generated filter expression.
    ''' </summary>
    ''' <value>A read-only list of data column names.</value>
    Public ReadOnly Property ColumnNames As IReadOnlyList(Of String)
        Get
            Return _ColumnNames
        End Get
    End Property
End Class
