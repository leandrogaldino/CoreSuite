Imports System.ComponentModel

Public Class QueriedBoxCondition2
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property Column As New QueriedBoxColumnReference
    Public Property [Operator] As SearchOperator
    Public Property Value As String
    Public Property Relation As SearchRelation
    Public Sub New()
    End Sub
    Public Sub New(Column As QueriedBoxColumnReference, [Operator] As SearchOperator, Value As String, Relation As SearchRelation)
        Me.Column = Column
        Me.Operator = [Operator]
        Me.Value = Value
        Me.Relation = Relation
    End Sub
    Overrides Function ToString() As String
        Dim Col As String = If(String.IsNullOrEmpty(Column.ColumnName), "?", Column.ColumnName)
        Dim Val As String = If(String.IsNullOrEmpty(Value), "?", Value)
        Return $"{Col} {[Operator]} {Val}"
    End Function
End Class
