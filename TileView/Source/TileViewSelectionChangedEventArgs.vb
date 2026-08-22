''' <summary>
''' Provides data for a <see cref="TileView"/> selection change.
''' </summary>
Public Class TileViewSelectionChangedEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TileViewSelectionChangedEventArgs"/> class.
    ''' </summary>
    ''' <param name="previousItem">The previously selected item, or <see langword="Nothing"/> when no item was selected.</param>
    ''' <param name="selectedItem">The newly selected item, or <see langword="Nothing"/> when the selection was cleared.</param>
    Public Sub New(previousItem As TileViewItem, selectedItem As TileViewItem)
        Me.PreviousItem = previousItem
        Me.SelectedItem = selectedItem
    End Sub
    ''' <summary>
    ''' Gets the item that was selected before the change.
    ''' </summary>
    Public ReadOnly Property PreviousItem As TileViewItem
    ''' <summary>
    ''' Gets the item selected after the change.
    ''' </summary>
    Public ReadOnly Property SelectedItem As TileViewItem
End Class
