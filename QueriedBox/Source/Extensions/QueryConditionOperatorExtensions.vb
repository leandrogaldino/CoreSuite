Imports System.Runtime.CompilerServices
''' <summary>
''' Provides extension methods for working with <see cref="QueryConditionOperator"/> values.
''' </summary>
Module QueryConditionOperatorExtensions
    ''' <summary>
    ''' Converts a query comparison operator into its SQL representation.
    ''' </summary>
    ''' <param name="Value">The query operator to convert.</param>
    ''' <returns>The SQL operator representation.</returns>
    <Extension>
    Public Function GetSqlValue(Value As QueryConditionOperator) As String
        Select Case Value
            Case QueryConditionOperator.Between
                Return "BETWEEN"
            Case QueryConditionOperator.Equal
                Return "="
            Case QueryConditionOperator.GreaterThan
                Return ">"
            Case QueryConditionOperator.GreaterThanOrEqual
                Return ">="
            Case QueryConditionOperator.LessThan
                Return "<"
            Case QueryConditionOperator.LessThanOrEqual
                Return "<="
            Case QueryConditionOperator.Like
                Return "LIKE"
            Case QueryConditionOperator.NotEqual
                Return "<>"
            Case QueryConditionOperator.In
                Return "IN"
            Case QueryConditionOperator.NotIn
                Return "NOT IN"
            Case Else
                Return ""
        End Select
    End Function
End Module
