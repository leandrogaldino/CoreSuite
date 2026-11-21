''' <summary>Provides the record selected when the user activates the optional Ctrl+click selection link.</summary>
Public NotInheritable Class LookupSelectionLinkClickedEventArgs
    Inherits EventArgs

    Public Sub New(SelectedItem As Object, SelectedValue As Object)
        Me.SelectedItem = SelectedItem
        Me.SelectedValue = SelectedValue
    End Sub

    ''' <summary>Gets the selected business object.</summary>
    Public ReadOnly Property SelectedItem As Object

    ''' <summary>Gets the identifier defined by ValueMember.</summary>
    Public ReadOnly Property SelectedValue As Object
End Class
