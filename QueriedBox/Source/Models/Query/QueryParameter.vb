Imports System.ComponentModel
''' <summary>
''' Represents a parameter used to provide values for query conditions.
''' </summary>
Public Class QueryParameter
    ''' <summary>
    ''' Gets or sets the name of the parameter used in query conditions.
    ''' </summary>
    <Description("Gets or sets the name of the parameter used in query conditions.")>
    <Category("QueriedBox")>
    Public Property ParameterName As String
    ''' <summary>
    ''' Gets or sets the value assigned to the parameter used in query conditions.
    ''' </summary>
    <Description("Gets or sets the value assigned to the parameter used in query conditions.")>
    <Category("QueriedBox")>
    Public Property ParameterValue As String
    ''' <summary>
    ''' Returns a string representation of the query parameter.
    ''' </summary>
    ''' <returns>
    ''' A formatted string containing the parameter name and value.
    ''' </returns>
    Public Overrides Function ToString() As String
        If ParameterName <> Nothing And ParameterValue <> Nothing Then
            Return ParameterName & " = " & ParameterValue
        ElseIf ParameterName <> Nothing And ParameterValue = Nothing Then
            Return ParameterName & " = Nothing"
        Else
            Return "New Undefined" & MyBase.GetType.Name
        End If
    End Function
End Class