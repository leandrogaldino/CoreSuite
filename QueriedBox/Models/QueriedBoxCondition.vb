Imports System.ComponentModel
Imports CoreSuite.QueriedBox

Public Class QueriedBoxCondition
    ''' <summary>
    ''' Define o nome ou apelido da tabela do banco de dados onde será aplicada a condição.
    ''' </summary>
    <Description("Define o nome ou apelido da tabela do banco de dados onde será aplicada a condição.")>
    <Category("Query")>
    Public Property TableNameOrAlias As String
    ''' <summary>
    ''' Define o nome do campo do banco de dados onde será aplicada a condição.
    ''' </summary>
    <Description("Define o nome do campo do banco de dados onde será aplicada a condição.")>
    <Category("Query")>
    Public Property FieldName As String
    ''' <summary>
    ''' Define o operador da condição. Para o operador BETWEEN, separar os dois valores por ponto e vírgula (;).
    ''' </summary>
    <Description("Define o operador da condição. Para o operador BETWEEN, separar os dois valores por ponto e vírgula (;).")>
    <Category("Query")>
    <TypeConverter(GetType(OperatorFilterCollection))>
    Public Property [Operator] As String
    ''' <summary>
    ''' Define o valor a ser testado na condição.
    ''' </summary>
    <Description("Define o valor a ser testado na condição.")>
    <Category("Query")>
    Public Property Value As String
    Public Overrides Function ToString() As String
        If TableNameOrAlias <> Nothing And FieldName <> Nothing And [Operator] <> Nothing And Value <> Nothing Then
            If [Operator] = "BETWEEN" Then
                If Value.Split(";").Length = 2 Then
                    Return String.Format("{0}.{1} {2} {3} AND {4}", TableNameOrAlias, FieldName, [Operator], Value.Split(";").ElementAt(0), Value.Split(";").ElementAt(1))
                Else
                    Return "New Undefined" & MyBase.GetType.Name
                End If
            Else
                Return String.Format("{0}.{1} {2} {3}", TableNameOrAlias, FieldName, [Operator], Value)
            End If
        ElseIf TableNameOrAlias <> Nothing And FieldName <> Nothing And [Operator] <> Nothing And Value = Nothing Then
            If [Operator] = "BETWEEN" Then
                Return String.Format("{0}.{1} {2} {3} AND {4}", TableNameOrAlias, FieldName, [Operator], "Nothing", "Nothing")
            Else
                Return String.Format("{0}.{1} {2} {3}", TableNameOrAlias, FieldName, [Operator], "Nothing")
            End If
        Else
            Return "New Undefined" & MyBase.GetType.Name
        End If
    End Function
End Class