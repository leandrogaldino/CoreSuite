Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public Class QueriedBoxRelation
    ''' <summary>
    ''' Define condições adicionais para o relacionamento.
    ''' </summary>
    <Description("Define condições adicionais para o relacionamento.")>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Category("Query")>
    Public Property Conditions As New Collection(Of QueriedBoxCondition)

    ''' <summary>
    ''' Define o tipo de relacionamento.
    ''' </summary>
    <Description("Define o tipo de relacionamento.")>
    <TypeConverter(GetType(RelationFilterCollection))>
    <Category("Query")>
    Public Property RelationType As String
    ''' <summary>
    ''' Define o operador a ser utilizado na relação.
    ''' </summary>
    <Description("Define o operador a ser utilizado na relação.")>
    <Category("Query")>
    <TypeConverter(GetType(OperatorFilterCollection))>
    Public Property [Operator] As String
    ''' <summary>
    ''' Define o nome da tabela que será relacionada.
    ''' </summary>
    <Description("Define o nome da tabela que será relacionada.")>
    <Category("Query")>
    Public Property RelateTableName As String
    ''' <summary>
    ''' Define o apelido da tabela que será relacionada.
    ''' </summary>
    <Description("Define o apelido da tabela que será relacionada.")>
    <Category("Query")>
    Public Property RelateTableAlias As String
    ''' <summary>
    ''' Define o nome do campo que se será relacionado.
    ''' </summary>
    <Description("Define o nome do campo que se será relacionado.")>
    <Category("Query")>
    Public Property RelateFieldName As String
    ''' <summary>
    ''' Define o nome da tabela que será relacionada com a tabela principal.
    ''' </summary>
    <Description("Define o nome da tabela que será relacionada com a tabela principal.")>
    <Category("Query")>
    Public Property WithTableName As String
    ''' <summary>
    ''' Define o apelido da tabela que se será relacionada com a tabela principal.
    ''' </summary>
    <Description("Define o apelido da tabela que se será relacionada com a tabela principal.")>
    <Category("Query")>
    Public Property WithTableAlias As String
    ''' <summary>
    ''' Define o nome do campo que será relacionado com o campo da tabela principal.
    ''' </summary>
    <Description("Define o nome do campo que será relacionado com o campo da tabela principal.")>
    <Category("Query")>
    Public Property WithFieldName As String
    Public Sub New()
    End Sub
    Public Overrides Function ToString() As String
        If RelateTableName = Nothing Or WithTableName = Nothing Or RelateFieldName = Nothing Or WithFieldName = Nothing Or [Operator] = Nothing Or RelationType = Nothing Then
            Return "New Undefined " & MyBase.GetType.Name
        Else
            Return String.Format("{0} JOIN {1} AS {2} ON {2}.{3} {4} {5}.{6}",
                                     RelationType, RelateTableName, If(RelateTableAlias <> Nothing, RelateTableAlias, RelateTableName),
                                     RelateFieldName, [Operator], If(WithTableAlias <> Nothing, WithTableAlias, WithTableName), WithFieldName)
        End If
    End Function
End Class
