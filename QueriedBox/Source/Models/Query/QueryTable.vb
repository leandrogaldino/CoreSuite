Imports System.ComponentModel

''' <summary>
''' Represents a database table definition used when building a query.
''' </summary>
Public Class QueryTable
    ''' <summary>
    ''' Gets or sets the name of the database table.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the name of the database table.")>
    Public Property TableName As String
    ''' <summary>
    ''' Gets or sets the alias assigned to the database table.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the alias assigned to the database table.")>
    Public Property TableAlias As String
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryTable"/> class.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryTable"/> class
    ''' with the specified table name and alias.
    ''' </summary>
    ''' <param name="Name">
    ''' The name of the database table.
    ''' </param>
    ''' <param name="[Alias]">
    ''' The alias assigned to the database table.
    ''' </param>
    Public Sub New(Name As String, [Alias] As String)
        Me.TableName = Name
        Me.TableAlias = [Alias]
    End Sub

    ''' <summary>
    ''' Returns a string representation of the table definition, including its alias when specified.
    ''' </summary>
    ''' <returns>
    ''' The table name formatted with the alias using the SQL AS syntax when applicable.
    ''' </returns>
    Overrides Function ToString() As String
        Dim HasAlias As Boolean = Not String.IsNullOrEmpty(TableAlias)
        If HasAlias Then
            Return $"{TableName} {TableAlias}"
        Else
            Return $"{TableName}"
        End If
    End Function
End Class
