Imports System.Collections.ObjectModel

''' <summary>
''' Represents a collection of <see cref="DataGridViewLayoutColumn"/> instances
''' that automatically assigns each column display index according to its position.
''' </summary>
''' <remarks>
''' When a column is inserted into the collection, its
''' <see cref="DataGridViewLayoutColumn.DisplayIndex"/> is automatically set
''' to the insertion index.
''' </remarks>
Public Class DataGridViewLayoutColumnCollection
    Inherits Collection(Of DataGridViewLayoutColumn)

    ''' <summary>
    ''' Inserts a column into the collection and assigns its display index.
    ''' </summary>
    ''' <param name="Index">The zero-based index at which the column is inserted.</param>
    ''' <param name="Item">The column configuration to insert.</param>
    ''' <exception cref="ArgumentNullException">
    ''' Thrown when <paramref name="Item"/> is <see langword="Nothing"/>.
    ''' </exception>
    Protected Overrides Sub InsertItem(Index As Integer, Item As DataGridViewLayoutColumn)
        ArgumentNullException.ThrowIfNull(Item)

        Item.DisplayIndex = Index
        MyBase.InsertItem(Index, Item)
    End Sub
End Class