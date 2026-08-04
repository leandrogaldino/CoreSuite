Imports System.ComponentModel
''' <summary>
''' Represents a condition used to filter records in a query.
''' </summary>
Public Class QueryCondition
    ''' <summary>
    ''' Gets or sets the column reference used in the condition.
    ''' </summary>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Category("QueriedBox")>
    <Description("Gets or sets the column reference used in the condition.")>
    Public Property Column As New QueryColumnReference
    ''' <summary>
    ''' Gets or sets the comparison operator applied to the column value.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the comparison operator applied to the column value.")>
    Public Property [Operator] As QueryConditionOperator
    ''' <summary>
    ''' Gets or sets the values used by the condition.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the values used by the condition.")>
    Public Property Values As String()
    ''' <summary>
    ''' Gets or sets the logical relation applied between this condition and other conditions.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the logical relation applied between this condition and other conditions.")>
    Public Property Relation As QueryRelation
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryCondition"/> class.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryCondition"/> class
    ''' with the specified column, operator, values, and relation.
    ''' </summary>
    ''' <param name="Column">
    ''' The column reference used in the condition.
    ''' </param>
    ''' <param name="Operator">
    ''' The comparison operator applied to the column value.
    ''' </param>
    ''' <param name="Values">
    ''' The values used by the condition.
    ''' </param>
    ''' <param name="Relation">
    ''' The logical relation applied between conditions.
    ''' </param>
    Public Sub New(Column As QueryColumnReference, [Operator] As QueryConditionOperator, Values As String(), Relation As QueryRelation)
        Me.Column = Column
        Me.Operator = [Operator]
        Me.Values = Values
        Me.Relation = Relation
    End Sub
    ''' <summary>
    ''' Returns a string representation of the query condition.
    ''' </summary>
    ''' <returns>
    ''' The condition formatted using the column, SQL operator, and assigned values.
    ''' </returns>
    Public Overrides Function ToString() As String
        Dim Values As New List(Of String)
        If Me.Values IsNot Nothing Then
            For Each v In Me.Values
                Values.Add(If(String.IsNullOrWhiteSpace(v), "?", v))
            Next v
        End If
        Dim ValueText As String
        Select Case [Operator]
            Case QueryConditionOperator.Between
                ValueText = If(Values.Count >= 2, $"{Values(0)} AND {Values(1)}", "? AND ?")
            Case QueryConditionOperator.In, QueryConditionOperator.NotIn
                ValueText = $"({String.Join(", ", Values)})"
            Case Else
                ValueText = If(Values.Count > 0, Values(0), "?")
        End Select
        Return $"{Column.ColumnName} {[Operator].GetSqlValue()} {ValueText}"
    End Function
End Class