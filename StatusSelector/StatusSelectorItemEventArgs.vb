''' <summary>
''' Provides data for events associated with a <see cref="StatusSelectorItem"/>.
''' </summary>
Public Class StatusSelectorItemEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="StatusSelectorItemEventArgs"/> class.
    ''' </summary>
    ''' <param name="item">The item associated with the event.</param>
    Public Sub New(item As StatusSelectorItem)
        Me.Item = item
    End Sub
    ''' <summary>
    ''' Gets the item associated with the event.
    ''' </summary>
    Public ReadOnly Property Item As StatusSelectorItem
End Class
