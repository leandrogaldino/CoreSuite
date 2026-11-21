''' <summary>Provides the previous and new selected values when ValueMember or the selected record changes.</summary>
Public NotInheritable Class LookupSelectedValueChangedEventArgs
    Inherits EventArgs

    Public Sub New(OldValue As Object, NewValue As Object)
        Me.OldValue = OldValue
        Me.NewValue = NewValue
    End Sub

    ''' <summary>Gets the identifier or value before the change.</summary>
    Public ReadOnly Property OldValue As Object

    ''' <summary>Gets the identifier or value after the change.</summary>
    Public ReadOnly Property NewValue As Object
End Class
