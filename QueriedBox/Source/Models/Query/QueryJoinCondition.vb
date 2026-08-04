Imports System.ComponentModel
''' <summary>
''' Represents a join condition used to define the relationship between two columns in a query.
''' </summary>
Public Class QueryJoinCondition
    ''' <summary>
    ''' Gets or sets the left column reference used in the join condition.
    ''' </summary>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Category("QueriedBox")>
    <Description("Gets or sets the left column reference used in the join condition.")>
    Public Property LeftColumn As New QueryColumnReference
    ''' <summary>
    ''' Gets or sets the comparison operator used between the columns.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the comparison operator used between the columns.")>
    Public Property [Operator] As QueryJoinConditionOperator
    ''' <summary>
    ''' Gets or sets the right column reference used in the join condition.
    ''' </summary>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Category("QueriedBox")>
    <Description("Gets or sets the right column reference used in the join condition.")>
    Public Property RightColumn As New QueryColumnReference
    ''' <summary>
    ''' Gets or sets the logical relation applied to the join condition.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the logical relation applied to the join condition.")>
    Public Property Relation As QueryRelation
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryJoinCondition"/> class.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryJoinCondition"/> class
    ''' with the specified columns, operator, and relation.
    ''' </summary>
    ''' <param name="LeftColumn">
    ''' The left column reference used in the join condition.
    ''' </param>
    ''' <param name="Operator">
    ''' The comparison operator applied between the columns.
    ''' </param>
    ''' <param name="RightColumn">
    ''' The right column reference used in the join condition.
    ''' </param>
    ''' <param name="Relation">
    ''' The logical relation applied to the condition.
    ''' </param>
    Public Sub New(LeftColumn As QueryColumnReference, [Operator] As QueryJoinConditionOperator, RightColumn As QueryColumnReference, Relation As QueryRelation)
        Me.LeftColumn = LeftColumn
        Me.Operator = [Operator]
        Me.RightColumn = RightColumn
        Me.Relation = Relation
    End Sub
    ''' <summary>
    ''' Returns a string representation of the join condition.
    ''' </summary>
    ''' <returns>
    ''' The join condition formatted using the selected columns and SQL operator.
    ''' </returns>
    Public Overrides Function ToString() As String
        Dim LeftCol As String = If(String.IsNullOrEmpty(LeftColumn.ColumnName), "?", LeftColumn.ColumnName)
        Dim RightCol As String = If(String.IsNullOrEmpty(RightColumn.ColumnName), "?", RightColumn.ColumnName)
        Return $"{LeftCol} {[Operator].GetSqlValue()} {RightCol}"
    End Function
End Class