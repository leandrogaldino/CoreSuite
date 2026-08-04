Imports System.Runtime.CompilerServices
''' <summary>
''' Provides extension methods for working with <see cref="QueryJoinType"/> values.
''' </summary>
Module QueryJoinTypeExtensions
    ''' <summary>
    ''' Converts a query JOIN type into its SQL representation.
    ''' </summary>
    ''' <param name="Value">The JOIN type to convert.</param>
    ''' <returns>The SQL JOIN operator representation.</returns>
    <Extension>
    Public Function GetSqlValue(Value As QueryJoinType) As String
        Select Case Value
            Case QueryJoinType.Inner
                Return "INNER"
            Case QueryJoinType.Left
                Return "LEFT"
            Case QueryJoinType.Right
                Return "RIGHT"
            Case QueryJoinType.Full
                Return "FULL"
            Case Else
                Return ""
        End Select
    End Function
End Module
