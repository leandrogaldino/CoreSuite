''' <summary>
''' Provides data for events associated with a <see cref="StatusSelectorItem"/>.
''' </summary>
Public Class StatusSelectorItemEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="StatusSelectorItemEventArgs"/> class.
    ''' </summary>
    ''' <param name="Item">The item associated with the event.</param>
    Public Sub New(Item As StatusSelectorItem)
        Me.Item = Item
    End Sub
    ''' <summary>
    ''' Gets the item associated with the event.
    ''' </summary>
    Public ReadOnly Property Item As StatusSelectorItem
End Class
