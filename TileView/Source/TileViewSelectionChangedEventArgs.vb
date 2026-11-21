
''' <summary>Logical record selection event; data references remain valid after tile recycling.</summary>
Public NotInheritable Class TileViewSelectionChangedEventArgs
    Inherits EventArgs

    Public Sub New(PreviousSourceIndex As Integer, SourceIndex As Integer, PreviousDataItem As Object, DataItem As Object, PreviousItem As TileViewItem, SelectedItem As TileViewItem)
        Me.PreviousSourceIndex = PreviousSourceIndex
        Me.SourceIndex = SourceIndex
        Me.PreviousDataItem = PreviousDataItem
        Me.DataItem = DataItem
        Me.PreviousItem = PreviousItem
        Me.SelectedItem = SelectedItem
    End Sub

    Public ReadOnly Property PreviousSourceIndex As Integer
    Public ReadOnly Property SourceIndex As Integer
    Public ReadOnly Property PreviousDataItem As Object
    Public ReadOnly Property DataItem As Object
    Public ReadOnly Property PreviousItem As TileViewItem
    Public ReadOnly Property SelectedItem As TileViewItem
End Class
