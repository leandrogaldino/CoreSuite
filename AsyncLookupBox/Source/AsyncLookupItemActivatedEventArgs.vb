''' <summary>
''' Provides internal data when a result is activated in the lookup drop-down.
''' </summary>
Friend NotInheritable Class AsyncLookupItemActivatedEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="AsyncLookupItemActivatedEventArgs"/> class.
    ''' </summary>
    ''' <param name="Item">The activated result object.</param>
    Public Sub New(Item As Object)
        Me.Item = Item
    End Sub
    ''' <summary>
    ''' Gets the activated result object.
    ''' </summary>
    Public ReadOnly Property Item As Object
End Class
