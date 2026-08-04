Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Drawing.Design
Imports Microsoft.DotNet.DesignTools.Editors
''' <summary>
''' Represents a query definition containing tables, columns, joins, conditions, parameters,
''' ordering, and pagination settings used to build SQL statements.
''' </summary>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class Query
    Private _Dialect As SqlDialect?
    ''' <summary>
    ''' Gets or sets the main table used in the query.
    ''' </summary>
    <TypeConverter(GetType(ExpandableObjectConverter))>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Category("QueriedBox")>
    <Description("Gets or sets the main table used in the query.")>
    Public Property Table As New QueryTable
    ''' <summary>
    ''' Gets or sets the name of the primary key column used by the query.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the name of the primary key column used by the query.")>
    Public Property PrimaryKeyColumnName As String
    ''' <summary>
    ''' Gets or sets the collection of columns included in the SELECT statement.
    ''' </summary>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(CollectionEditor), GetType(UITypeEditor))>
    <Category("QueriedBox")>
    <Description("Gets or sets the collection of columns included in the SELECT statement.")>
    Public Property Columns As New Collection(Of QueryColumn)
    ''' <summary>
    ''' Gets or sets the collection of JOIN definitions applied to the query.
    ''' </summary>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(CollectionEditor), GetType(UITypeEditor))>
    <Category("QueriedBox")>
    <Description("Gets or sets the collection of JOIN definitions applied to the query.")>
    Public Property Joins As New Collection(Of QueryJoin)
    ''' <summary>
    ''' Gets or sets the collection of filtering conditions applied to the query.
    ''' </summary>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(CollectionEditor), GetType(UITypeEditor))>
    <Category("QueriedBox")>
    <Description("Gets or sets the collection of filtering conditions applied to the query.")>
    Public Property Conditions As New Collection(Of QueryCondition)
    ''' <summary>
    ''' Gets or sets the collection of parameters used by the query.
    ''' </summary>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(CollectionEditor), GetType(UITypeEditor))>
    <Category("QueriedBox")>
    <Description("Gets or sets the collection of parameters used by the query.")>
    Public Property Parameters As New Collection(Of QueryParameter)
    ''' <summary>
    ''' Gets or sets the collection of sorting definitions applied to the query.
    ''' </summary>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(CollectionEditor), GetType(UITypeEditor))>
    <Category("QueriedBox")>
    <Description("Gets or sets the collection of sorting definitions applied to the query.")>
    Public Property OrderBy As New Collection(Of QueryOrderBy)
    ''' <summary>
    ''' Gets or sets the maximum number of records returned by the query.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the maximum number of records returned by the query.")>
    Public Property Limit As Integer? = 500
    ''' <summary>
    ''' Gets or sets the number of records skipped before returning query results.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the number of records skipped before returning query results.")>
    Public Property Offset As Integer?
    ''' <summary>
    ''' Gets or sets a value indicating whether the query should return only distinct records.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets whether the query should return only distinct records.")>
    Public Property Distinct As Boolean = False

    ''' <summary>
    ''' Gets or sets the SQL dialect used to generate the query.
    ''' </summary>
    ''' <remarks>
    ''' When set to <see langword="Nothing"/>, the value defined by
    ''' <see cref="QueriedBox.SqlDialect"/> is used.
    ''' </remarks>
    <Category("QueriedBox")>
    <Description("Gets or sets the SQL dialect used to generate the query.")>
    Public Property Dialect As SqlDialect?
        Get
            Return _Dialect
        End Get
        Set(value As SqlDialect?)
            _Dialect = value
        End Set
    End Property
    ''' <summary>
    ''' Gets the SQL dialect effectively used to generate the query.
    ''' </summary>
    ''' <remarks>
    ''' Returns the value of <see cref="Dialect"/> when specified; otherwise,
    ''' returns the global default defined by <see cref="QueriedBox.SqlDialect"/>.
    ''' </remarks>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property EffectiveDialect As SqlDialect
        Get
            Return If(Dialect, QueriedBox.SqlDialect)
        End Get
    End Property
    ''' <summary>
    ''' Initializes a new instance of the <see cref="Query"/> class.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="Query"/> class
    ''' with the specified query components.
    ''' </summary>
    ''' <param name="Table">The main table used in the query.</param>
    ''' <param name="Columns">The columns included in the SELECT statement.</param>
    ''' <param name="Joins">The JOIN definitions applied to the query.</param>
    ''' <param name="Conditions">The filtering conditions applied to the query.</param>
    ''' <param name="Parameters">The parameters used by the query.</param>
    ''' <param name="OrderBy">The sorting definitions applied to the query.</param>
    ''' <param name="Limit">The maximum number of records returned.</param>
    ''' <param name="Offset">The number of records skipped before returning results.</param>
    Public Sub New(Table As QueryTable, Columns As Collection(Of QueryColumn), Joins As Collection(Of QueryJoin), Conditions As Collection(Of QueryCondition), Parameters As Collection(Of QueryParameter), OrderBy As Collection(Of QueryOrderBy), Limit As Integer?, Offset As Integer?)
        Me.Table = Table
        Me.Columns = Columns
        Me.Joins = Joins
        Me.Conditions = Conditions
        Me.Parameters = Parameters
        Me.OrderBy = OrderBy
        Me.Limit = Limit
        Me.Offset = Offset
    End Sub
    ''' <summary>
    ''' Returns a string representation of the complete SQL query.
    ''' </summary>
    ''' <returns>
    ''' The generated SQL SELECT statement.
    ''' </returns>
    Public Overrides Function ToString() As String
        Dim CurrentDialect = EffectiveDialect
        Dim Sql As String = $"SELECT {GetSelect()}"
        Dim Joins = GetJoins()
        If Joins <> "" Then Sql &= $" {Joins}"
        Dim WhereSql = GetWhere()
        If WhereSql <> "" Then Sql &= $" WHERE {WhereSql}"
        Dim Order = GetOrder()
        If Order <> "" Then Sql &= $" ORDER BY {Order}"
        Dim Pagination = CurrentDialect.GetLimitOffset(Limit, Offset)
        If Pagination <> "" Then
            Sql &= $" {Pagination}"
        End If
        Return Sql
    End Function
    ''' <summary>
    ''' Generates the query with an additional column added to the SELECT statement.
    ''' </summary>
    ''' <param name="AdditionalSelectColumn">The additional column expression to include.</param>
    ''' <returns>The generated SQL query.</returns>
    Public Function ToStringWithAdditionalSelectColumn(AdditionalSelectColumn As String) As String
        Return ToStringWithAdditions(AdditionalSelectColumn, Nothing)
    End Function
    ''' <summary>
    ''' Generates the query with an additional filtering condition added to the WHERE clause.
    ''' </summary>
    ''' <param name="AdditionalWhereCondition">The additional WHERE condition expression.</param>
    ''' <returns>The generated SQL query.</returns>
    Public Function ToStringWithAdditionalWhereCondition(AdditionalWhereCondition As String) As String
        Return ToStringWithAdditions(Nothing, AdditionalWhereCondition)
    End Function
    ''' <summary>
    ''' Generates the query with optional additional SELECT columns and WHERE conditions.
    ''' </summary>
    ''' <param name="AdditionalSelectColumn">An optional additional column expression.</param>
    ''' <param name="AdditionalWhereCondition">An optional additional filtering condition.</param>
    ''' <returns>The generated SQL query.</returns>
    Public Function ToStringWithAdditions(Optional AdditionalSelectColumn As String = Nothing, Optional AdditionalWhereCondition As String = Nothing) As String
        Dim CurrentDialect = EffectiveDialect
        Dim Sql As String = "SELECT "
        If Not String.IsNullOrWhiteSpace(AdditionalSelectColumn) Then
            Sql &= $"{AdditionalSelectColumn}, "
        End If
        Sql &= GetSelect()
        Dim Joins = GetJoins()
        If Joins <> "" Then Sql &= $" {Joins}"
        Dim WhereSql = GetWhere()
        If Not String.IsNullOrWhiteSpace(AdditionalWhereCondition) OrElse Not String.IsNullOrWhiteSpace(WhereSql) Then
            Sql &= " WHERE "
            If Not String.IsNullOrWhiteSpace(AdditionalWhereCondition) Then
                Sql &= AdditionalWhereCondition
                If Not String.IsNullOrWhiteSpace(WhereSql) Then Sql &= " AND "
            End If
            If Not String.IsNullOrWhiteSpace(WhereSql) Then Sql &= WhereSql
        End If
        Dim Order = GetOrder()
        If Order <> "" Then Sql &= $" ORDER BY {Order}"
        Dim Pagination = CurrentDialect.GetLimitOffset(Limit, Offset)
        If Pagination <> "" Then
            Sql &= $" {Pagination}"
        End If
        Sql &= ";"
        Return Sql
    End Function
    ''' <summary>
    ''' Generates the SELECT clause of the query, including selected columns and the main table.
    ''' </summary>
    ''' <returns>
    ''' The SQL SELECT expression containing columns and the source table.
    ''' </returns>
    Public Function GetSelect() As String
        Dim CurrentDialect = EffectiveDialect
        Dim TableSql As String = Table.ToString()
        TableSql = If(String.IsNullOrWhiteSpace(TableSql), "?", TableSql)
        Dim ColumnsSql As New List(Of String)
        For Each Column In Columns
            ColumnsSql.Add(Column.ToString(CurrentDialect))
        Next
        If ColumnsSql.Count = 0 Then
            ColumnsSql.Add("*")
        End If
        Dim DistinctSql As String = If(Distinct, "DISTINCT ", "")
        Dim FullSelect As String = $"{DistinctSql}{String.Join(", ", ColumnsSql)} FROM {TableSql}"
        Return FullSelect
    End Function
    ''' <summary>
    ''' Generates the JOIN clauses defined in the query.
    ''' </summary>
    ''' <returns>
    ''' The SQL JOIN expressions or an empty string when no joins are configured.
    ''' </returns>
    Public Function GetJoins() As String
        If Joins.Count = 0 Then Return ""
        Return String.Join(" ", Joins)
    End Function
    ''' <summary>
    ''' Generates the SQL WHERE clause based on the configured query conditions.
    ''' </summary>
    ''' <returns>
    ''' The generated SQL WHERE expression, or an empty string when no conditions are configured.
    ''' </returns>
    Public Function GetWhere() As String
        If Conditions.Count = 0 Then Return ""
        Dim WhereSql As New List(Of String)
        For i = 0 To Conditions.Count - 1
            Dim Text = Conditions(i).ToString()
            If i > 0 Then
                Text = $"{Conditions(i - 1).Relation.ToString().ToUpper()} {Text}"
            End If
            WhereSql.Add(Text)
        Next i
        Return String.Join(" ", WhereSql)
    End Function
    ''' <summary>
    ''' Generates the ORDER BY clause based on the configured sorting definitions.
    ''' </summary>
    ''' <returns>
    ''' The SQL ORDER BY expression or an empty string when no sorting is configured.
    ''' </returns>
    Public Function GetOrder() As String
        If OrderBy.Count = 0 Then Return ""
        Dim OrderSql As New List(Of String)
        For Each o In OrderBy
            OrderSql.Add(o.ToString())
        Next o
        Return String.Join(", ", OrderSql)
    End Function
End Class
