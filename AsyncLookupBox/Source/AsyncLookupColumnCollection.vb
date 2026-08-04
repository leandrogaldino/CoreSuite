Imports System.Collections.ObjectModel
Imports System.ComponentModel
''' <summary>
''' Represents the ordered collection of columns displayed by an <see cref="AsyncLookupBox"/>.
''' </summary>
<Description("Contains the columns displayed in an AsyncLookupBox result list.")>
Public Class AsyncLookupColumnCollection
    Inherits Collection(Of AsyncLookupColumn)
    Friend Event Changed As EventHandler
    ''' <summary>
    ''' Adds a result column for the specified property path.
    ''' </summary>
    ''' <param name="PropertyName">The property path whose value is displayed.</param>
    ''' <returns>The column added to the collection.</returns>
    Public Overloads Function Add(PropertyName As String) As AsyncLookupColumn
        Dim Column As New AsyncLookupColumn(PropertyName)
        Add(Column)
        Return Column
    End Function
    ''' <summary>
    ''' Inserts a column into the collection and observes its configuration changes.
    ''' </summary>
    ''' <param name="Index">The zero-based insertion index.</param>
    ''' <param name="Item">The column to insert.</param>
    Protected Overrides Sub InsertItem(Index As Integer, Item As AsyncLookupColumn)
        If Item Is Nothing Then Throw New ArgumentNullException(NameOf(Item))
        MyBase.InsertItem(Index, Item)
        AddHandler Item.Changed, AddressOf ItemChanged
        OnChanged()
    End Sub
    ''' <summary>
    ''' Replaces a column and updates change tracking.
    ''' </summary>
    ''' <param name="Index">The zero-based index of the column to replace.</param>
    ''' <param name="Item">The replacement column.</param>
    Protected Overrides Sub SetItem(Index As Integer, Item As AsyncLookupColumn)
        If Item Is Nothing Then Throw New ArgumentNullException(NameOf(Item))
        RemoveHandler Me(Index).Changed, AddressOf ItemChanged
        MyBase.SetItem(Index, Item)
        AddHandler Item.Changed, AddressOf ItemChanged
        OnChanged()
    End Sub
    ''' <summary>
    ''' Removes a column and releases its change subscription.
    ''' </summary>
    ''' <param name="Index">The zero-based index of the column to remove.</param>
    Protected Overrides Sub RemoveItem(Index As Integer)
        RemoveHandler Me(Index).Changed, AddressOf ItemChanged
        MyBase.RemoveItem(Index)
        OnChanged()
    End Sub
    ''' <summary>
    ''' Removes all columns and releases their change subscriptions.
    ''' </summary>
    Protected Overrides Sub ClearItems()
        For Each Column As AsyncLookupColumn In Me
            RemoveHandler Column.Changed, AddressOf ItemChanged
        Next
        MyBase.ClearItems()
        OnChanged()
    End Sub
    Private Sub ItemChanged(Sender As Object, E As EventArgs)
        OnChanged()
    End Sub
    Private Sub OnChanged()
        RaiseEvent Changed(Me, EventArgs.Empty)
    End Sub
End Class
