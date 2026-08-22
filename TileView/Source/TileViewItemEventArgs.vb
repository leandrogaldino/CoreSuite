''' <summary>
''' Provides data for events associated with a <see cref="TileViewItem"/>.
''' </summary>
Public Class TileViewItemEventArgs
    Inherits EventArgs

    ''' <summary>
    ''' Initializes a new instance of the <see cref="TileViewItemEventArgs"/> class.
    ''' </summary>
    ''' <param name="Item">The item associated with the event.</param>
    Public Sub New(Item As TileViewItem)
        ArgumentNullException.ThrowIfNull(Item)
        Me.Item = Item
    End Sub

    ''' <summary>
    ''' Gets the item associated with the event.
    ''' </summary>
    Public ReadOnly Property Item As TileViewItem
End Class