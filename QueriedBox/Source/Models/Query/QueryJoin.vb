Imports System.Collections.ObjectModel
Imports System.ComponentModel
''' <summary>
''' Represents a JOIN clause definition used when building a query.
''' </summary>
Public Class QueryJoin
    ''' <summary>
    ''' Gets or sets the type of JOIN operation to apply.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the type of JOIN operation to apply.")>
    Public Property Type As QueryJoinType = QueryJoinType.Inner
    ''' <summary>
    ''' Gets or sets the table involved in the JOIN operation.
    ''' </summary>
    <TypeConverter(GetType(ExpandableObjectConverter))>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Category("QueriedBox")>
    <Description("Gets or sets the table involved in the JOIN operation.")>
    Public Property Table As New QueryTable
    ''' <summary>
    ''' Gets or sets the collection of conditions used to define the JOIN relationship.
    ''' </summary>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Category("QueriedBox")>
    <Description("Gets or sets the collection of conditions used to define the JOIN relationship.")>
    Public Property Conditions As New Collection(Of QueryJoinCondition)
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryJoin"/> class.
    ''' </summary>
    Public Sub New()
    End Sub
    Public Sub New(Type As QueryJoinType, Table As QueryTable, Conditions As Collection(Of QueryJoinCondition))
        Me.Type = Type
        Me.Table = Table
        Me.Conditions = Conditions
    End Sub
    ''' <summary>
    ''' Returns the SQL representation of the configured JOIN clause.
    ''' </summary>
    ''' <returns>
    ''' A SQL JOIN expression containing the join type, table, alias, and configured conditions.
    ''' </returns>
    Public Overrides Function ToString() As String
        Dim JoinTable As String = If(String.IsNullOrWhiteSpace(Table.TableName), "?", Table.TableName)
        If Not String.IsNullOrWhiteSpace(Table.TableAlias) Then
            JoinTable &= $" AS {Table.TableAlias}"
        End If
        Dim JoinConditions As New List(Of String)
        For i As Integer = 0 To Conditions.Count - 1
            Dim ConditionText As String = Conditions(i).ToString()
            If i > 0 Then
                ConditionText = $"{Conditions(i - 1).Relation.GetSqlValue()} {ConditionText}"
            End If
            JoinConditions.Add(ConditionText)
        Next
        Return $"{Type.GetSqlValue()} JOIN {JoinTable} ON {String.Join(" ", JoinConditions)}"
    End Function
End Class


