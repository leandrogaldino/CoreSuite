Imports System.ComponentModel
Imports Microsoft.DotNet.DesignTools.Designers.Actions
Public Class QueriedTextBoxControlDesignerActionList
    Inherits DesignerActionList
    Private ReadOnly _Control As QueriedBox
    Public Sub New(Designer As QueriedTextBoxControlDesigner)
        MyBase.New(Designer.Component)
        _Control = CType(Designer.Component, QueriedBox)
    End Sub
    Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
        Dim Items As New DesignerActionItemCollection From {
            New DesignerActionPropertyItem(NameOf(QueryEnabled), "Query Enabled", "Query", "Define se as pesquisas estão habilitadas."),
            New DesignerActionPropertyItem(NameOf(CharactersToQuery), "Characters To Query", "Query", "Define a quantidade de caracteres necessários para iniciar a pesquisa."),
            New DesignerActionPropertyItem(NameOf(MainTableName), "Main Table Name", "Main Table", "Define o nome da tabela principal."),
            New DesignerActionPropertyItem(NameOf(MainTableAlias), "Main Table Alias", "Main Table", "Define o apelido da tabela principal."),
            New DesignerActionPropertyItem(NameOf(MainReturnFieldName), "Main Return Field Name", "Main Table", "Define o campo Primary Key da tabela principal."),
            New DesignerActionPropertyItem(NameOf(DisplayTableName), "Display Table Name", "Display", "Define a tabela relacionada aos resultados exibidos."),
            New DesignerActionPropertyItem(NameOf(DisplayTableAlias), "Display Table Alias", "Display", "Define o apelido da tabela exibida."),
            New DesignerActionPropertyItem(NameOf(DisplayFieldName), "Display Field Name", "Display", "Define o campo exibido no controle."),
            New DesignerActionPropertyItem(NameOf(DisplayFieldAlias), "Display Field Alias", "Display", "Define o apelido do campo exibido."),
            New DesignerActionPropertyItem(NameOf(DisplayMainFieldName), "Display Main Field Name", "Display", "Define o campo Primary Key relacionado."),
            New DesignerActionPropertyItem(NameOf(Limit), "Limit", "Query", "Define o limite máximo de resultados retornados.")
        }
        Return Items
    End Function
    Public Property QueryEnabled As Boolean
        Get
            Return _Control.QueryEnabled
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(QueryEnabled), value)
        End Set
    End Property
    Public Property QueryInterval As Integer
        Get
            Return _Control.QueryInterval
        End Get
        Set(value As Integer)
            SetProperty(NameOf(QueryInterval), value)
        End Set
    End Property
    Public Property CharactersToQuery As Integer
        Get
            Return _Control.CharactersToQuery
        End Get
        Set(value As Integer)
            SetProperty(NameOf(CharactersToQuery), value)
        End Set
    End Property
    Public Property MainTableName As String
        Get
            Return _Control.MainTableName
        End Get
        Set(value As String)
            SetProperty(NameOf(MainTableName), value)
        End Set
    End Property
    Public Property MainTableAlias As String
        Get
            Return _Control.MainTableAlias
        End Get
        Set(value As String)
            SetProperty(NameOf(MainTableAlias), value)
        End Set
    End Property
    Public Property MainReturnFieldName As String
        Get
            Return _Control.MainReturnFieldName
        End Get
        Set(value As String)
            SetProperty(NameOf(MainReturnFieldName), value)
        End Set
    End Property
    Public Property DisplayTableName As String
        Get
            Return _Control.DisplayTableName
        End Get
        Set(value As String)
            SetProperty(NameOf(DisplayTableName), value)
        End Set
    End Property
    Public Property DisplayTableAlias As String
        Get
            Return _Control.DisplayTableAlias
        End Get
        Set(value As String)
            SetProperty(NameOf(DisplayTableAlias), value)
        End Set
    End Property
    Public Property DisplayFieldName As String
        Get
            Return _Control.DisplayFieldName
        End Get
        Set(value As String)
            SetProperty(NameOf(DisplayFieldName), value)
        End Set
    End Property
    Public Property DisplayFieldAlias As String
        Get
            Return _Control.DisplayFieldAlias
        End Get
        Set(value As String)
            SetProperty(NameOf(DisplayFieldAlias), value)
        End Set
    End Property
    Public Property DisplayMainFieldName As String
        Get
            Return _Control.DisplayMainFieldName
        End Get
        Set(value As String)
            SetProperty(NameOf(DisplayMainFieldName), value)
        End Set
    End Property
    Public Property Limit As Integer
        Get
            Return _Control.Limit
        End Get
        Set(value As Integer)
            SetProperty(NameOf(Limit), value)
        End Set
    End Property
    Private Sub SetProperty(PropertyName As String, Value As Object)
        TypeDescriptor.GetProperties(_Control)(PropertyName).SetValue(_Control, Value)
    End Sub
End Class