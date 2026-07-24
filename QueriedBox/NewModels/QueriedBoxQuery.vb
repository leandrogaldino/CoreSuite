Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Drawing.Design
Imports Microsoft.DotNet.DesignTools.Editors

<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class QueriedBoxQuery
    <TypeConverter(GetType(ExpandableObjectConverter))>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property Table As New QueriedBoxTable
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(CollectionEditor), GetType(UITypeEditor))>
    Public Property Columns As New Collection(Of QueriedBoxColumn)
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(CollectionEditor), GetType(UITypeEditor))>
    Public Property Joins As New Collection(Of QueriedBoxJoin)
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(CollectionEditor), GetType(UITypeEditor))>
    Public Property Conditions As New Collection(Of QueriedBoxCondition2)
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(CollectionEditor), GetType(UITypeEditor))>
    Public Property Parameters As New Collection(Of QueriedBoxParameter)
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(CollectionEditor), GetType(UITypeEditor))>
    Public Property OrderBy As New Collection(Of QueriedBoxOrderBy)
    Public Property Limit As Integer?
    Public Property Offset As Integer?
    Public Sub New()
    End Sub
    Public Sub New(Table As QueriedBoxTable, Columns As Collection(Of QueriedBoxColumn), Joins As Collection(Of QueriedBoxJoin), Conditions As Collection(Of QueriedBoxCondition2), Parameters As Collection(Of QueriedBoxParameter), OrderBy As Collection(Of QueriedBoxOrderBy), Limit As Integer?, Offset As Integer?)
        Me.Table = Table
        Me.Columns = Columns
        Me.Joins = Joins
        Me.Conditions = Conditions
        Me.Parameters = Parameters
        Me.OrderBy = OrderBy
        Me.Limit = Limit
        Me.Offset = Offset
    End Sub
    Public Overrides Function ToString() As String
        Dim TableSql As String = If(String.IsNullOrWhiteSpace(Table.TableName), "?", Table.TableName)
        If Not String.IsNullOrWhiteSpace(Table.Alias) Then
            TableSql &= $" AS {Table.Alias}"
        End If
        Dim ColumnsSql As New List(Of String)
        For Each Column In Columns
            Dim Text = If(String.IsNullOrWhiteSpace(Column.ColumnName), "?", Column.ColumnName)
            If Not String.IsNullOrWhiteSpace(Column.ColumnAlias) Then
                Text &= $" AS {Column.ColumnAlias}"
            End If
            ColumnsSql.Add(Text)
        Next Column
        If ColumnsSql.Count = 0 Then
            ColumnsSql.Add("*")
        End If
        Dim JoinsSql As New List(Of String)
        For Each j In Joins
            Dim JoinTable = If(String.IsNullOrWhiteSpace(j.Table.TableName), "?", j.Table.TableName)
            If Not String.IsNullOrWhiteSpace(j.Table.Alias) Then
                JoinTable &= $" AS {j.Table.Alias}"
            End If
            Dim JoinConditions As New List(Of String)
            For Each c In j.Conditions
                Dim LeftCol = If(String.IsNullOrWhiteSpace(c.LeftColumn.ColumnName), "?", c.LeftColumn.ColumnName)
                Dim RightCol = If(String.IsNullOrWhiteSpace(c.RightColumn.ColumnName), "?", c.RightColumn.ColumnName)
                JoinConditions.Add($"{LeftCol} {c.Operator.ToString().ToUpper()} {RightCol}")
            Next c
            JoinsSql.Add($"{j.Type} JOIN {JoinTable} ON {String.Join(" AND ", JoinConditions)}")
        Next j
        Dim WhereSql As New List(Of String)
        For i = 0 To Conditions.Count - 1
            Dim c = Conditions(i)
            Dim Parameter = Parameters.FirstOrDefault(Function(p) p.ParameterName = c.Value)
            Dim Value = If(Parameter Is Nothing, c.Value, Parameter.ParameterValue)
            If String.IsNullOrWhiteSpace(Value) Then
                Value = "?"
            End If
            Dim Text = $"{c.Column.ColumnName} {c.Operator.ToString().ToUpper()} {Value}"
            If i > 0 Then
                Text = $"{Conditions(i - 1).Relation.ToString().ToUpper()} {Text}"
            End If
            WhereSql.Add(Text)
        Next i
        Dim OrderSql As New List(Of String)
        For Each o In OrderBy
            Dim Col = If(String.IsNullOrWhiteSpace(o.Column.ColumnName), "?", o.Column.ColumnName)
            OrderSql.Add($"{Col} {o.Direction.ToString().ToUpper()}")
        Next o
        Dim Sql As String =
        $"SELECT {String.Join(", ", ColumnsSql)} FROM {TableSql}"

        If JoinsSql.Count > 0 Then
            Sql &= " " & String.Join(" ", JoinsSql)
        End If

        If WhereSql.Count > 0 Then
            Sql &= " WHERE " & String.Join(" ", WhereSql)
        End If

        If OrderSql.Count > 0 Then
            Sql &= " ORDER BY " & String.Join(", ", OrderSql)
        End If

        If Limit.HasValue Then
            Sql &= $" LIMIT {Limit.Value}"
        End If

        If Offset.HasValue Then
            Sql &= $" OFFSET {Offset.Value}"
        End If

        Return Sql

    End Function
End Class
