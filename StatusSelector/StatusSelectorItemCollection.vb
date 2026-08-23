Imports System.Collections
Imports System.Drawing

''' <summary>
''' Represents the collection of status items owned by a <see cref="StatusSelector"/>.
''' </summary>
Public NotInheritable Class StatusSelectorItemCollection
    Implements IList(Of StatusSelectorItem)
    Private ReadOnly _Owner As StatusSelector
    Private ReadOnly _Items As New List(Of StatusSelectorItem)
    Friend Sub New(owner As StatusSelector)
        _Owner = owner
    End Sub
    ''' <summary>
    ''' Gets or sets the item at the specified index.
    ''' </summary>
    Default Public Property Item(index As Integer) As StatusSelectorItem Implements IList(Of StatusSelectorItem).Item
        Get
            Return _Items(index)
        End Get
        Set(value As StatusSelectorItem)
            ArgumentNullException.ThrowIfNull(value)
            Dim OldItem = _Items(index)
            If ReferenceEquals(OldItem, value) Then Return
            _Owner.DetachItem(OldItem)
            _Items(index) = value
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
    Public Sub Add(item As StatusSelectorItem) Implements ICollection(Of StatusSelectorItem).Add
        ArgumentNullException.ThrowIfNull(item)
        _Items.Add(item)
        _Owner.AttachItem(item)
        _Owner.ItemsChanged()
    End Sub
    ''' <summary>
    ''' Creates and adds a status item.
    ''' </summary>
    Public Function Add(text As String, value As Object) As StatusSelectorItem
        Return Add(text, value, SystemColors.ControlText)
    End Function
    ''' <summary>
    ''' Creates and adds a status item with a custom foreground color.
    ''' </summary>
    Public Function Add(text As String, value As Object, foreColor As Color) As StatusSelectorItem
        Dim NewItem As New StatusSelectorItem(text, value, foreColor)
        Add(NewItem)
        Return NewItem
    End Function
    ''' <summary>
    ''' Creates and adds a status item with custom foreground and background colors.
    ''' </summary>
    Public Function Add(text As String, value As Object, foreColor As Color, backColor As Color) As StatusSelectorItem
        Dim NewItem As New StatusSelectorItem(text, value, foreColor) With {.BackColor = backColor}
        Add(NewItem)
        Return NewItem
    End Function
    ''' <summary>
    ''' Adds all supplied items to the collection.
    ''' </summary>
    Public Sub AddRange(items As IEnumerable(Of StatusSelectorItem))
        ArgumentNullException.ThrowIfNull(items)
        _Owner.BeginItemsUpdate()
        Try
            For Each StatusItem As StatusSelectorItem In items
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
    Public Function Contains(item As StatusSelectorItem) As Boolean Implements ICollection(Of StatusSelectorItem).Contains
        Return _Items.Contains(item)
    End Function
    Public Sub CopyTo(array As StatusSelectorItem(), arrayIndex As Integer) Implements ICollection(Of StatusSelectorItem).CopyTo
        _Items.CopyTo(array, arrayIndex)
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
    Public Function IndexOf(item As StatusSelectorItem) As Integer Implements IList(Of StatusSelectorItem).IndexOf
        Return _Items.IndexOf(item)
    End Function
    ''' <summary>
    ''' Inserts an item at the specified index.
    ''' </summary>
    Public Sub Insert(index As Integer, item As StatusSelectorItem) Implements IList(Of StatusSelectorItem).Insert
        ArgumentNullException.ThrowIfNull(item)
        _Items.Insert(index, item)
        _Owner.AttachItem(item)
        _Owner.ItemsChanged()
    End Sub
    ''' <summary>
    ''' Removes the specified item from the collection.
    ''' </summary>
    Public Function Remove(item As StatusSelectorItem) As Boolean Implements ICollection(Of StatusSelectorItem).Remove
        If Not _Items.Remove(item) Then Return False
        _Owner.DetachItem(item)
        _Owner.ItemsChanged()
        Return True
    End Function
    ''' <summary>
    ''' Removes the item at the specified index.
    ''' </summary>
    Public Sub RemoveAt(index As Integer) Implements IList(Of StatusSelectorItem).RemoveAt
        Dim Item = _Items(index)
        _Items.RemoveAt(index)
        _Owner.DetachItem(Item)
        _Owner.ItemsChanged()
    End Sub
End Class
