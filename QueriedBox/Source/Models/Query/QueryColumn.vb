Imports System.ComponentModel

''' <summary>
''' Represents a database column definition used when building a query and displaying query results.
''' </summary>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class QueryColumn
    ''' <summary>
    ''' Gets or sets the name of the database column.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the name of the database column.")>
    Public Property ColumnName As String
    ''' <summary>
    ''' Gets or sets the alias assigned to the column in the query result.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the alias assigned to the column in the query result.")>
    Public Property ColumnAlias As String
    ''' <summary>
    ''' Gets or sets the value expression used when the column value is null.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the value expression used when the column value is null.")>
    Public Property IfNull As String
    ''' <summary>
    ''' Gets or sets the QueriedBox-specific options for the column.
    ''' </summary>
    <TypeConverter(GetType(ExpandableObjectConverter))>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Category("QueriedBox")>
    <Description("Gets or sets the QueriedBox-specific options for the column.")>
    Public Property Options As New QueryColumnOptions
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryColumn"/> class.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryColumn"/> class with the specified column configuration.
    ''' </summary>
    ''' <param name="Name">
    ''' The name of the database column.
    ''' </param>
    ''' <param name="Alias">
    ''' The alias assigned to the column in the query result.
    ''' </param>
    ''' <param name="IfNull">
    ''' The fallback value or expression used when the column value is <see langword="Nothing"/>.
    ''' </param>
    ''' <param name="Options">
    ''' The options that control how the column is displayed and used by the query.
    ''' </param>
    Public Sub New(Name As String, [Alias] As String, IfNull As String, Options As QueryColumnOptions)
        Me.ColumnName = Name
        Me.ColumnAlias = [Alias]
        Me.IfNull = IfNull
        Me.Options = Options
    End Sub
    ''' <summary>
    ''' Returns a string representation of the column definition.
    ''' </summary>
    ''' <param name="Dialect">
    ''' The SQL dialect used to generate dialect-specific expressions.
    ''' </param>
    ''' <returns>
    ''' The SQL expression representing the column, including null handling and an alias when specified.
    ''' </returns>
    Public Overloads Function ToString(Dialect As SqlDialect) As String
        Dim Text = If(String.IsNullOrWhiteSpace(ColumnName), "?", ColumnName)
        If Not String.IsNullOrWhiteSpace(IfNull) Then
            Text = Dialect.GetIfNull(Text, IfNull)
        End If
        If Not String.IsNullOrWhiteSpace(ColumnAlias) Then
            Text &= $" AS {ColumnAlias}"
        End If
        Return Text
    End Function
End Class
