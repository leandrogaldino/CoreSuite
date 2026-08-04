''' <summary>
''' Provides data for the <see cref="AsyncLookupBox.SelectionChanged"/> event.
''' </summary>
Public NotInheritable Class AsyncLookupSelectionChangedEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Initializes a new instance of the <see cref="AsyncLookupSelectionChangedEventArgs"/> class.
    ''' </summary>
    ''' <param name="OldItem">The previously selected result object.</param>
    ''' <param name="NewItem">The newly selected result object.</param>
    ''' <param name="OldValue">The previously resolved value.</param>
    ''' <param name="NewValue">The newly resolved value.</param>
    Public Sub New(OldItem As Object, NewItem As Object, OldValue As Object, NewValue As Object)
        Me.OldItem = OldItem
        Me.NewItem = NewItem
        Me.OldValue = OldValue
        Me.NewValue = NewValue
    End Sub
    ''' <summary>
    ''' Gets the previously selected result object.
    ''' </summary>
    Public ReadOnly Property OldItem As Object
    ''' <summary>
    ''' Gets the newly selected result object, or <see langword="Nothing"/> when the selection was cleared.
    ''' </summary>
    Public ReadOnly Property NewItem As Object
    ''' <summary>
    ''' Gets the previously resolved selected value.
    ''' </summary>
    Public ReadOnly Property OldValue As Object
    ''' <summary>
    ''' Gets the newly resolved selected value, or <see langword="Nothing"/> when the selection was cleared.
    ''' </summary>
    Public ReadOnly Property NewValue As Object
End Class
