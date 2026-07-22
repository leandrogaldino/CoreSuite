Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Windows.Forms.Design

Friend Class QueriedBoxDesignerActionList
    Inherits DesignerActionList
    Private Control As QueriedBox
    Private Designer As ControlDesigner
    Private ActionList As DesignerActionList
    Public Sub New(ByVal Designer As ControlDesigner, ByVal ActionList As DesignerActionList)
        MyBase.New(Designer.Component)
        Me.Designer = Designer
        Me.ActionList = ActionList
        Control = CType(Designer.Control, QueriedBox)
    End Sub
    Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
        Dim Items = New DesignerActionItemCollection From {
            New DesignerActionPropertyItem("QueryEnabled", "QueryEnabled", "Category1", "Define se as pesquisas estão habilitadas."),
            New DesignerActionPropertyItem("CharactersToQuery", "Characters To Query", "Category1", "Define a quantidade de caracteres necessários para iniciar a pesquisa."),
            New DesignerActionPropertyItem("MainTableName", "MainTableName", "Category2", "Define o nome da tabela principal."),
            New DesignerActionPropertyItem("MainTableAlias", "MainTableAlias", "Category2", "Define o apelido da tabela principal."),
            New DesignerActionPropertyItem("MainReturnFieldName", "MainReturnFieldName", "Category2", "Define o nome do campo atribuido como Primary Key na tabela principal."),
            New DesignerActionPropertyItem("DisplayTableName", "DisplayTableName", "Category3", "Define o nome da tabela a qual está relacionada aos resultados que serão exibidos."),
            New DesignerActionPropertyItem("DisplayTableAlias", "DisplayTableAlias", "Category3", "Define o apelido da tabela a qual está relacionada aos resultados que serão exibidos."),
            New DesignerActionPropertyItem("DisplayFieldName", "DisplayFieldName", "Category3", "Define o nome do campo que será congelado no controle quando escolhido pelo usuário."),
            New DesignerActionPropertyItem("DisplayFieldAlias", "DisplayFieldAlias", "Category3", "Define o apelido do campo que será congelado no controle quando escolhido pelo usuário."),
            New DesignerActionPropertyItem("DisplayMainFieldName", "DisplayMainFieldName", "Category3", "Define o nome do campo da tabela que está atribuído como Primary Key."),
            New DesignerActionPropertyItem("Limit", "Limit", "Category4", "Define o máximo de resultados que podem ser retornados pela pesquisa.")
        }
        Return Items
    End Function

    Public Property QueryEnabled As Boolean
        Get
            Return Control.QueryEnabled
        End Get
        Set(ByVal value As Boolean)
            TypeDescriptor.GetProperties(Component)("QueryEnabled").SetValue(Component, value)
        End Set
    End Property
    Public Property QueryInterval As Integer
        Get
            Return Control.QueryInterval
        End Get
        Set(ByVal value As Integer)
            TypeDescriptor.GetProperties(Component)("QueryInterval").SetValue(Component, value)
        End Set
    End Property
    Public Property CharactersToQuery As Integer
        Get
            Return Control.CharactersToQuery
        End Get
        Set(ByVal value As Integer)
            TypeDescriptor.GetProperties(Component)("CharactersToQuery").SetValue(Component, value)
        End Set
    End Property
    Public Property MainTableName As String
        Get
            Return Control.MainTableName
        End Get
        Set(ByVal value As String)
            TypeDescriptor.GetProperties(Component)("MainTableName").SetValue(Component, value)
        End Set
    End Property
    Public Property MainTableAlias As String
        Get
            Return Control.MainTableAlias
        End Get
        Set(ByVal value As String)
            TypeDescriptor.GetProperties(Component)("MainTableAlias").SetValue(Component, value)
        End Set
    End Property
    Public Property MainReturnFieldName As String
        Get
            Return Control.MainReturnFieldName
        End Get
        Set(ByVal value As String)
            TypeDescriptor.GetProperties(Component)("MainReturnFieldName").SetValue(Component, value)
        End Set
    End Property
    Public Property DisplayTableName As String
        Get
            Return Control.DisplayTableName
        End Get
        Set(ByVal value As String)
            TypeDescriptor.GetProperties(Component)("DisplayTableName").SetValue(Component, value)
        End Set
    End Property
    Public Property DisplayTableAlias As String
        Get
            Return Control.DisplayTableAlias
        End Get
        Set(ByVal value As String)
            TypeDescriptor.GetProperties(Component)("DisplayTableAlias").SetValue(Component, value)
        End Set
    End Property
    Public Property DisplayFieldName As String
        Get
            Return Control.DisplayFieldName
        End Get
        Set(ByVal value As String)
            TypeDescriptor.GetProperties(Component)("DisplayFieldName").SetValue(Component, value)
        End Set
    End Property
    Public Property DisplayFieldAlias As String
        Get
            Return Control.DisplayFieldAlias
        End Get
        Set(ByVal value As String)
            TypeDescriptor.GetProperties(Component)("DisplayFieldAlias").SetValue(Component, value)
        End Set
    End Property
    Public Property DisplayMainFieldName As String
        Get
            Return Control.DisplayMainFieldName
        End Get
        Set(ByVal value As String)
            TypeDescriptor.GetProperties(Component)("DisplayMainFieldName").SetValue(Component, value)
        End Set
    End Property
    Public Property Limit As Integer
        Get
            Return Control.Limit
        End Get
        Set(ByVal value As Integer)
            TypeDescriptor.GetProperties(Component)("Limit").SetValue(Component, value)
        End Set
    End Property
End Class
