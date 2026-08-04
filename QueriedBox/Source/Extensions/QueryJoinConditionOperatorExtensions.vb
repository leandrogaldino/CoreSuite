Imports System.Runtime.CompilerServices
''' <summary>
''' Provides extension methods for working with <see cref="QueryJoinConditionOperator"/> values.
''' </summary>
Module QueryJoinConditionOperatorExtensions
    ''' <summary>
    ''' Converts a query comparison operator into its SQL representation.
    ''' </summary>
    ''' <param name="Value">The query operator to convert.</param>
    ''' <returns>The SQL operator representation.</returns>
    <Extension>
    Public Function GetSqlValue(Value As QueryJoinConditionOperator) As String
        Select Case Value
            Case QueryJoinConditionOperator.Equal
                Return "="
            Case QueryJoinConditionOperator.GreaterThan
                Return ">"
            Case QueryJoinConditionOperator.GreaterThanOrEqual
                Return ">="
            Case QueryJoinConditionOperator.LessThan
                Return "<"
            Case QueryJoinConditionOperator.LessThanOrEqual
                Return "<="
            Case QueryJoinConditionOperator.Like
                Return "LIKE"
            Case QueryJoinConditionOperator.NotEqual
                Return "<>"
            Case Else
                Return ""
        End Select
    End Function
End Module
