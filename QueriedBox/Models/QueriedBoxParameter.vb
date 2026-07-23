Imports System.ComponentModel

Public Class QueriedBoxParameter
    ''' <summary>
    ''' Define o nome do parâmetro utilizado na Query.
    ''' </summary>
    <Description("Define o nome do parâmetro utilizado nas condições da Query.")>
    <Category("Query")>
    Public Property ParameterName As String
    ''' <summary>
    ''' Define o valor do parâmetro utilizado na Query.
    ''' </summary>
    <Description("Define o valor do parâmetro utilizado nas condições da Query.")>
    <Category("Query")>
    Public Property ParameterValue As String
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