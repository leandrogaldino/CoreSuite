Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Drawing.Design
Imports Microsoft.DotNet.DesignTools.Editors

Public Class QueriedBoxJoin
    Public Property Type As JoinType
    <TypeConverter(GetType(ExpandableObjectConverter))>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property Table As New QueriedBoxTable
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    <Editor(GetType(CollectionEditor), GetType(UITypeEditor))>
    Public Property Conditions As New Collection(Of QueriedBoxJoinCondition)
    Public Sub New()
    End Sub
    Public Sub New(Type As JoinType, Table As QueriedBoxTable, Conditions As Collection(Of QueriedBoxJoinCondition))
        Me.Type = Type
        Me.Table = Table
        Me.Conditions = Conditions
    End Sub
    Public Overrides Function ToString() As String
        Dim TableName As String = If(String.IsNullOrEmpty(Table.TableName), "?", Table.TableName)
        Return $"{Type} JOIN {TableName}"
    End Function
End Class


