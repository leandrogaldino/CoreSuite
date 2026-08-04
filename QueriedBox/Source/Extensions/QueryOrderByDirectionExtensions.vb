Imports System.Runtime.CompilerServices
''' <summary>
''' Provides extension methods for working with <see cref="QueryOrderByDirection"/> values.
''' </summary>
Module QueryOrderByDirectionExtensions
    ''' <summary>
    ''' Converts a query ordering direction into its SQL representation.
    ''' </summary>
    ''' <param name="Value">The ordering direction to convert.</param>
    ''' <returns>The SQL sorting direction representation.</returns>
    <Extension>
    Public Function GetSqlValue(Value As QueryOrderByDirection) As String
        Select Case Value
            Case QueryOrderByDirection.Ascending
                Return "ASC"
            Case QueryOrderByDirection.Descending
                Return "DESC"
            Case Else
                Return ""
        End Select
    End Function
End Module
