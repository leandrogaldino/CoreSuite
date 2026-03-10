Imports System.Data.Common

''' <summary>
''' Provides helper methods for database-related operations,
''' such as debugging SQL commands and inspecting generated queries.
''' </summary>
Public Class DatabaseHelper

    ''' <summary>
    ''' Outputs a SQL command for debugging purposes by replacing
    ''' its parameters with their corresponding values in the query text.
    ''' </summary>
    ''' <param name="Command">
    ''' The <see cref="DbCommand"/> containing the SQL statement and its parameters.
    ''' </param>
    Public Shared Sub DebugQuery(Command As DbCommand)
        Dim Query As String = Command.CommandText
        For Each Parameter As DbParameter In Command.Parameters
            Query = Query.Replace(Parameter.ParameterName, "'" & Parameter.Value & "'")
        Next
        Debug.Print(Query)
    End Sub

End Class