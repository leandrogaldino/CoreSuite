Imports System.ComponentModel
''' <summary>
''' Represents a column sorting definition used when building an ORDER BY clause in a query.
''' </summary>
Public Class QueryOrderBy

    ''' <summary>
    ''' Gets or sets the column reference used for sorting.
    ''' </summary>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Category("QueriedBox")>
    <Description("Gets or sets the column reference used for sorting.")>
    Public Property Column As New QueryColumnReference
    ''' <summary>
    ''' Gets or sets the sorting direction applied to the column.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the sorting direction applied to the column.")>
    Public Property Direction As QueryOrderByDirection = QueryOrderByDirection.Ascending
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryOrderBy"/> class.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryOrderBy"/> class
    ''' with the specified column and sorting direction.
    ''' </summary>
    ''' <param name="Column">
    ''' The column reference used for sorting.
    ''' </param>
    ''' <param name="Direction">
    ''' The sorting direction applied to the column.
    ''' </param>
    Public Sub New(Column As QueryColumnReference, Optional Direction As QueryOrderByDirection = QueryOrderByDirection.Ascending)
        Me.Column = Column
        Me.Direction = Direction
    End Sub

    ''' <summary>
    ''' Returns a string representation of the ORDER BY definition.
    ''' </summary>
    ''' <returns>
    ''' The column name followed by the SQL sorting direction.
    ''' </returns>
    Overrides Function ToString() As String
        Dim Col As String = If(String.IsNullOrEmpty(Column.ColumnName), "?", Column.ColumnName)
        Return $"{Col} {Direction.GetSqlValue().ToUpper()}"
    End Function
End Class
