Imports System.Collections.ObjectModel
Imports System.ComponentModel
''' <summary>
''' Represents a collection of column references used by a <see cref="DataGridViewFilterBox"/>.
''' </summary>
Public Class DataGridViewFilterColumnCollection
    Inherits Collection(Of DataGridViewFilterColumn)
    Friend Event Changed As EventHandler
    ''' <summary>
    ''' Adds a column reference with the specified name to the collection.
    ''' </summary>
    ''' <param name="ColumnName">The data column name, grid column name, or data property name to add.</param>
    ''' <returns>The newly created column reference.</returns>
    Public Overloads Function Add(ColumnName As String) As DataGridViewFilterColumn
        If String.IsNullOrWhiteSpace(ColumnName) Then Throw New ArgumentException("A column name is required.", NameOf(ColumnName))
        Dim Column As New DataGridViewFilterColumn(ColumnName)
        Add(Column)
        Return Column
    End Function
    ''' <summary>
    ''' Determines whether the collection contains the specified column reference.
    ''' </summary>
    ''' <param name="ColumnName">The column name to locate.</param>
    ''' <returns><see langword="True"/> when a case-insensitive match exists; otherwise, <see langword="False"/>.</returns>
    Public Overloads Function Contains(ColumnName As String) As Boolean
        Return IndexOf(ColumnName) >= 0
    End Function
    ''' <summary>
    ''' Returns the zero-based index of the specified column reference.
    ''' </summary>
    ''' <param name="ColumnName">The column name to locate.</param>
    ''' <returns>The zero-based index of the item, or <c>-1</c> when the item is not present.</returns>
    Public Overloads Function IndexOf(ColumnName As String) As Integer
        For Index As Integer = 0 To Count - 1
            If String.Equals(Me(Index).ColumnName, ColumnName, StringComparison.OrdinalIgnoreCase) Then Return Index
        Next
        Return -1
    End Function
    ''' <summary>
    ''' Inserts a column reference and notifies the owning control that its configuration changed.
    ''' </summary>
    ''' <param name="Index">The zero-based index at which to insert the item.</param>
    ''' <param name="Item">The column reference to insert.</param>
    Protected Overrides Sub InsertItem(Index As Integer, Item As DataGridViewFilterColumn)
        ValidateItem(Item, -1)
        MyBase.InsertItem(Index, Item)
        AddHandler Item.PropertyChanged, AddressOf ItemPropertyChanged
        OnChanged()
    End Sub
    ''' <summary>
    ''' Replaces a column reference and notifies the owning control that its configuration changed.
    ''' </summary>
    ''' <param name="Index">The zero-based index of the item to replace.</param>
    ''' <param name="Item">The replacement column reference.</param>
    Protected Overrides Sub SetItem(Index As Integer, Item As DataGridViewFilterColumn)
        ValidateItem(Item, Index)
        RemoveHandler Me(Index).PropertyChanged, AddressOf ItemPropertyChanged
        MyBase.SetItem(Index, Item)
        AddHandler Item.PropertyChanged, AddressOf ItemPropertyChanged
        OnChanged()
    End Sub
    ''' <summary>
    ''' Removes a column reference and notifies the owning control that its configuration changed.
    ''' </summary>
    ''' <param name="Index">The zero-based index of the item to remove.</param>
    Protected Overrides Sub RemoveItem(Index As Integer)
        RemoveHandler Me(Index).PropertyChanged, AddressOf ItemPropertyChanged
        MyBase.RemoveItem(Index)
        OnChanged()
    End Sub
    ''' <summary>
    ''' Removes all column references and notifies the owning control that its configuration changed.
    ''' </summary>
    Protected Overrides Sub ClearItems()
        If Count = 0 Then Return
        For Each Item As DataGridViewFilterColumn In Me
            RemoveHandler Item.PropertyChanged, AddressOf ItemPropertyChanged
        Next
        MyBase.ClearItems()
        OnChanged()
    End Sub
    Private Sub ValidateItem(Item As DataGridViewFilterColumn, IgnoredIndex As Integer)
        ArgumentNullException.ThrowIfNull(Item)
        If String.IsNullOrWhiteSpace(Item.ColumnName) Then Return
        For Index As Integer = 0 To Count - 1
            If Index <> IgnoredIndex AndAlso String.Equals(Me(Index).ColumnName, Item.ColumnName, StringComparison.OrdinalIgnoreCase) Then Throw New ArgumentException($"The column '{Item.ColumnName}' is already present in the collection.", NameOf(Item))
        Next
    End Sub
    Private Sub ItemPropertyChanged(sender As Object, e As PropertyChangedEventArgs)
        OnChanged()
    End Sub
    Private Sub OnChanged()
        RaiseEvent Changed(Me, EventArgs.Empty)
    End Sub
End Class
