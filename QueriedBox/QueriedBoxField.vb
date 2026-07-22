Imports System.ComponentModel

Public Class QueriedBoxField
    ''' <summary>
    ''' Define o nome da tabela a qual está relacionada aos resultados que serão exibidos.
    ''' </summary>
    <Description("Define o nome da tabela a qual está relacionada aos resultados que serão exibidos.")>
    <Category("Query")>
    Public Property DisplayTableName As String
    ''' <summary>
    ''' Define o apelido da tabela a qual está relacionada aos resultados que serão exibidos.
    ''' </summary>
    <Description("Define o apelido da tabela a qual está relacionada aos resultados que serão exibidos.")>
    <Category("Query")>
    Public Property DisplayTableAlias As String
    ''' <summary>
    ''' Define o nome do campo que será que será exibido em conjunto nos resultados da pesquisa.
    ''' </summary>
    <Description("Define o nome do campo que será que será exibido em conjunto nos resultados da pesquisa.")>
    <Category("Query")>
    Public Property DisplayFieldName As String
    ''' <summary>
    ''' Define o apelido do campo que será que será exibido em conjunto nos resultados da pesquisa.
    ''' </summary>
    <Description("Define o apelido do campo que será que será exibido em conjunto nos resultados da pesquisa.")>
    <Category("Query")>
    Public Property DisplayFieldAlias As String
    ''' <summary>
    ''' Define o nome do campo da tabela que está atribuído como Primary Key.
    ''' </summary>
    <Description("Define o nome do campo da tabela que está atribuído como Primary Key.")>
    <Category("Query")>
    Public Property DisplayMainFieldName As String
    ''' <summary>
    ''' Define se o valor desse campo será congelado no controle.
    ''' </summary>
    <Description("Define se o valor desse campo será congelado no controle.")>
    <Category("Comportamento")>
    Public Property Freeze As Boolean
    ''' <summary>
    ''' Define um prefixo que será congelado com o resultado.
    ''' </summary>
    <Category("Comportamento")>
    <Description("Define um prefixo que será congelado com o resultado.")>
    Public Property Prefix As String
    ''' <summary>
    ''' Define um sufixo que será congelado com o resultado.
    ''' </summary>
    <Category("Comportamento")>
    <Description("Define um sufixo que será congelado com o resultado.")>
    Public Property Suffix As String
    ''' <summary>
    ''' Define se o campo será exibido no resultado.
    ''' </summary>
    <Category("Comportamento")>
    <Description("Define se o campo será exibido no resultado.")>
    Public Property Display As Boolean = True
    ''' <summary>
    ''' Define um valor a ser retornado caso a consulta retorne um valor nulo para esse campo.
    ''' </summary>
    <Description("Define um valor a ser retornado caso a consulta retorne um valor nulo para esse campo.")>
    <Category("Query")>
    Public Property IfNull As String
    ''' <summary>
    ''' Define como a largura da coluna será ajustada nos resultados.
    ''' </summary>
    <Category("Aparência")>
    <DefaultValue(GetType(Color), "Blue")>
    <Description("Define como a largura da coluna será ajustada nos resultados.")>
    Public Property DisplayFieldAutoSizeColumnMode As DataGridViewAutoSizeColumnMode
    Public Overrides Function ToString() As String
        If DisplayTableName <> Nothing Or DisplayFieldName <> Nothing Then
            Return If(DisplayTableAlias = Nothing, DisplayTableName, DisplayTableAlias) & "." & If(DisplayFieldAlias = Nothing, DisplayFieldName, DisplayFieldAlias)
        Else
            Return "Undefined " & MyBase.GetType.Name
        End If
    End Function
End Class
