
Imports System.ComponentModel

Public Class QueriedBoxOrderBy
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property Column As New QueriedBoxColumnReference
    Public Property Direction As QueryOrderByDirection = QueryOrderByDirection.Ascending
    Public Sub New()
    End Sub
    Public Sub New(Column As QueriedBoxColumnReference, Optional Direction As QueryOrderByDirection = QueryOrderByDirection.Ascending)
        Me.Column = Column
        Me.Direction = Direction
    End Sub
    Overrides Function ToString() As String
        Dim Col As String = If(String.IsNullOrEmpty(Column.ColumnName), "?", Column.ColumnName)
        Return $"{Col} {Direction}"
    End Function
End Class
