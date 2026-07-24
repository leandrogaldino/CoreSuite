Imports System.ComponentModel

Public Class QueriedBoxJoinCondition
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property LeftColumn As New QueriedBoxColumnReference
    Public Property [Operator] As SearchOperator
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property RightColumn As New QueriedBoxColumnReference
    Public Property Relation As SearchRelation
    Public Sub New()
    End Sub

    Public Sub New(LeftColumn As QueriedBoxColumnReference, [Operator] As SearchOperator, RightColumn As QueriedBoxColumnReference, Relation As SearchRelation)
        Me.LeftColumn = LeftColumn
        Me.Operator = [Operator]
        Me.RightColumn = RightColumn
        Me.Relation = Relation
    End Sub
    Public Overrides Function ToString() As String
        Dim LeftCol As String = If(String.IsNullOrEmpty(LeftColumn.ColumnName), "?", LeftColumn.ColumnName)
        Dim RightCol As String = If(String.IsNullOrEmpty(RightColumn.ColumnName), "?", RightColumn.ColumnName)
        Return $"{LeftCol} {[Operator]} {RightCol}"
    End Function
End Class