Imports System.Collections
Imports System.Drawing

''' <summary>
''' Represents the collection of status items owned by a <see cref="StatusSelector"/>.
''' </summary>
Public NotInheritable Class StatusSelectorItemCollection
    Implements IList(Of StatusSelectorItem)
    Private ReadOnly _Owner As StatusSelector
    Private ReadOnly _Items As New List(Of StatusSelectorItem)
    Friend Sub New(Owner As StatusSelector)
        _Owner = Owner
    End Sub
    ''' <summary>
    ''' Gets or sets the item at the specified index.
    ''' </summary>
    Default Public Property Item(Index As Integer) As StatusSelectorItem Implements IList(Of StatusSelectorItem).Item
        Get
            Return _Items(Index)
        End Get
        Set(value As StatusSelectorItem)
            If value Is Nothing Then Throw New ArgumentNullException("value")
            Dim OldItem = _Items(Index)
            If ReferenceEquals(OldItem, value) Then Return
            _Owner.DetachItem(OldItem)
            _Items(Index) = value
            _Owner.AttachItem(value)
            _Owner.ItemsChanged()
        End Set
    End Property
    ''' <summary>
    ''' Gets the number of status items in the collection.
    ''' </summary>
    Public ReadOnly Property Count As Integer Implements ICollection(Of StatusSelectorItem).Count
        Get
            Return _Items.Count
        End Get
    End Property
    Private ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of StatusSelectorItem).IsReadOnly
        Get
            Return False
        End Get
    End Property
    ''' <summary>
    ''' Adds an existing item to the collection.
    ''' </summary>
    Public Sub Add(Item As StatusSelectorItem) Implements ICollection(Of StatusSelectorItem).Add
        If Item Is Nothing Then Throw New ArgumentNullException("item")
        _Items.Add(Item)
        _Owner.AttachItem(Item)
        _Owner.ItemsChanged()
    End Sub
    ''' <summary>
    ''' Creates and adds a status item.
    ''' </summary>
    Public Function Add(Text As String, Value As Object) As StatusSelectorItem
        Return Add(Text, Value, SystemColors.ControlText)
    End Function
    ''' <summary>
    ''' Creates and adds a status item with a custom foreground color.
    ''' </summary>
    Public Function Add(Text As String, Value As Object, foreColor As Color) As StatusSelectorItem
        Dim NewItem As New StatusSelectorItem(Text, Value, foreColor)
        Add(NewItem)
        Return NewItem
    End Function
    ''' <summary>
    ''' Creates and adds a status item with custom foreground and background colors.
    ''' </summary>
    Public Function Add(Text As String, Value As Object, ForeColor As Color, BackColor As Color) As StatusSelectorItem
        Dim NewItem As New StatusSelectorItem(Text, Value, ForeColor) With {.BackColor = BackColor}
        Add(NewItem)
        Return NewItem
    End Function
    ''' <summary>
    ''' Adds all supplied items to the collection.
    ''' </summary>
    Public Sub AddRange(Items As IEnumerable(Of StatusSelectorItem))
        If Items Is Nothing Then Throw New ArgumentNullException("items")
        _Owner.BeginItemsUpdate()
        Try
            For Each StatusItem As StatusSelectorItem In Items
                Add(StatusItem)
            Next
        Finally
            _Owner.EndItemsUpdate()
        End Try
    End Sub
    ''' <summary>
    ''' Removes all items from the collection.
    ''' </summary>
    Public Sub Clear() Implements ICollection(Of StatusSelectorItem).Clear
        For Each StatusItem As StatusSelectorItem In _Items
            _Owner.DetachItem(StatusItem)
        Next
        _Items.Clear()
        _Owner.ItemsChanged()
    End Sub
    ''' <summary>
    ''' Determines whether the collection contains the specified item.
    ''' </summary>
    Public Function Contains(Item As StatusSelectorItem) As Boolean Implements ICollection(Of StatusSelectorItem).Contains
        Return _Items.Contains(Item)
    End Function
    Public Sub CopyTo(Array As StatusSelectorItem(), ArrayIndex As Integer) Implements ICollection(Of StatusSelectorItem).CopyTo
        _Items.CopyTo(Array, ArrayIndex)
    End Sub
    Public Function GetEnumerator() As IEnumerator(Of StatusSelectorItem) Implements IEnumerable(Of StatusSelectorItem).GetEnumerator
        Return _Items.GetEnumerator()
    End Function
    Private Function GetNonGenericEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return _Items.GetEnumerator()
    End Function
    ''' <summary>
    ''' Gets the index of the specified item.
    ''' </summary>
    Public Function IndexOf(Item As StatusSelectorItem) As Integer Implements IList(Of StatusSelectorItem).IndexOf
        Return _Items.IndexOf(Item)
    End Function
    ''' <summary>
    ''' Inserts an item at the specified index.
    ''' </summary>
    Public Sub Insert(Index As Integer, Item As StatusSelectorItem) Implements IList(Of StatusSelectorItem).Insert
        If Item Is Nothing Then Throw New ArgumentNullException("item")
        _Items.Insert(Index, Item)
        _Owner.AttachItem(Item)
        _Owner.ItemsChanged()
    End Sub
    ''' <summary>
    ''' Removes the specified item from the collection.
    ''' </summary>
    Public Function Remove(Item As StatusSelectorItem) As Boolean Implements ICollection(Of StatusSelectorItem).Remove
        If Not _Items.Remove(Item) Then Return False
        _Owner.DetachItem(Item)
        _Owner.ItemsChanged()
        Return True
    End Function
    ''' <summary>
    ''' Removes the item at the specified index.
    ''' </summary>
    Public Sub RemoveAt(Index As Integer) Implements IList(Of StatusSelectorItem).RemoveAt
        Dim Item = _Items(Index)
        _Items.RemoveAt(Index)
        _Owner.DetachItem(Item)
        _Owner.ItemsChanged()
    End Sub
End Class
