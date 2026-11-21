
''' <summary>Visual tile and its current logical data indices at the time of an event.</summary>
Public NotInheritable Class TileViewItemEventArgs
    Inherits EventArgs

    Public Sub New(Item As TileViewItem)
        If Item Is Nothing Then Throw New ArgumentNullException(NameOf(Item))
        Me.Item = Item
        Me.DisplayIndex = Item.DisplayIndex
        Me.SourceIndex = Item.DataIndex
        Me.DataItem = Item.DataItem
    End Sub

    Public ReadOnly Property Item As TileViewItem
    Public ReadOnly Property DisplayIndex As Integer
    Public ReadOnly Property SourceIndex As Integer
    Public ReadOnly Property DataItem As Object
End Class
