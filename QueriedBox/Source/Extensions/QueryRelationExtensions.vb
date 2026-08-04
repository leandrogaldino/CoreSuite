Imports System.Runtime.CompilerServices
''' <summary>
''' Provides extension methods for working with <see cref="QueryRelation"/> values.
''' </summary>
Module QueryRelationExtensions
    ''' <summary>
    ''' Converts a query logical relation into its SQL representation.
    ''' </summary>
    ''' <param name="Value">The query relation to convert.</param>
    ''' <returns>The SQL logical operator representation.</returns>
    <Extension>
    Public Function GetSqlValue(Value As QueryRelation) As String
        Select Case Value
            Case QueryRelation.And
                Return "AND"
            Case QueryRelation.Or
                Return "OR"
            Case Else
                Return ""
        End Select
    End Function
End Module
