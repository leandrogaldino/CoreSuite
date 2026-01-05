Imports System.Data.Common

''' <summary>
''' Represents selection (SELECT) options for a MySQL table.
''' Used by the <see cref="MySqlRequest"/> class to execute queries
''' with filters, specific columns, ordering, and result limits.
''' </summary>
Public Class MySqlSelectOptions

    ''' <summary>
    ''' List of columns to be returned.
    ''' If <c>Nothing</c> or empty, all columns (*) will be returned.
    ''' </summary>
    Public Property Columns As List(Of String)

    ''' <summary>
    ''' Optional WHERE clause used to filter records.
    ''' </summary>
    Public Property Where As String

    ''' <summary>
    ''' Dictionary of query parameters, used when placeholders are defined.
    ''' </summary>
    Public Property QueryArgs As Dictionary(Of String, Object)

    ''' <summary>
    ''' Optional ORDER BY clause used to sort the results.
    ''' </summary>
    Public Property OrderBy As String

    ''' <summary>
    ''' Optional limit on the number of records to be returned.
    ''' </summary>
    Public Property Limit As Integer?

    ''' <summary>
    ''' Optional connection to be used.
    ''' If <c>Nothing</c>, the <see cref="MySqlRequest"/> class will
    ''' automatically create and manage the connection.
    ''' </summary>
    Public Property Connection As DbConnection

End Class