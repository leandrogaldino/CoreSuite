Imports System.Collections.ObjectModel
''' <summary>
''' Represents the ordered collection of actions displayed by a <see cref="TextBoxActionPanel"/>.
''' </summary>
''' <remarks>The first visible item is placed at the right edge of the panel, and each following item is placed to its left.</remarks>
Public NotInheritable Class TextBoxActionCollection
    Inherits Collection(Of TextBoxAction)
    Friend Event Changed As EventHandler
    ''' <summary>
    ''' Adds and returns an action initialized with the specified values.
    ''' </summary>
    ''' <param name="Key">The identifier used to distinguish the action.</param>
    ''' <param name="Image">The image displayed by the action button.</param>
    ''' <param name="ToolTipText">The tooltip displayed for the action.</param>
    ''' <returns>The action added to the collection.</returns>
    Public Overloads Function Add(Key As String, Image As Image, ToolTipText As String) As TextBoxAction
        Dim Item As New TextBoxAction(Key, Image, ToolTipText)
        Add(Item)
        Return Item
    End Function
    ''' <summary>
    ''' Finds the first action whose key matches the specified value without regard to case.
    ''' </summary>
    ''' <param name="Key">The key to locate.</param>
    ''' <returns>The matching action, or <see langword="Nothing"/> when no action is found.</returns>
    Public Function FindByKey(Key As String) As TextBoxAction
        For Each Item As TextBoxAction In Items
            If String.Equals(Item.Key, Key, StringComparison.OrdinalIgnoreCase) Then Return Item
        Next
        Return Nothing
    End Function
    ''' <summary>
    ''' Inserts an action and begins observing changes to its presentation properties.
    ''' </summary>
    ''' <param name="Index">The zero-based insertion index.</param>
    ''' <param name="Item">The action to insert.</param>
    Protected Overrides Sub InsertItem(Index As Integer, Item As TextBoxAction)
        ValidateItem(Item, Nothing)
        AssignDefaultKey(Item)
        MyBase.InsertItem(Index, Item)
        AddHandler Item.Changed, AddressOf Item_Changed
        OnChanged()
    End Sub
    ''' <summary>
    ''' Replaces an action and updates property-change observation.
    ''' </summary>
    ''' <param name="Index">The zero-based index of the action to replace.</param>
    ''' <param name="Item">The replacement action.</param>
    Protected Overrides Sub SetItem(Index As Integer, Item As TextBoxAction)
        Dim CurrentItem As TextBoxAction = Items(Index)
        If ReferenceEquals(CurrentItem, Item) Then Return
        ValidateItem(Item, CurrentItem)
        AssignDefaultKey(Item)
        RemoveHandler CurrentItem.Changed, AddressOf Item_Changed
        MyBase.SetItem(Index, Item)
        AddHandler Item.Changed, AddressOf Item_Changed
        OnChanged()
    End Sub
    ''' <summary>
    ''' Removes an action and stops observing its presentation properties.
    ''' </summary>
    ''' <param name="Index">The zero-based index of the action to remove.</param>
    Protected Overrides Sub RemoveItem(Index As Integer)
        Dim Item As TextBoxAction = Items(Index)
        RemoveHandler Item.Changed, AddressOf Item_Changed
        MyBase.RemoveItem(Index)
        OnChanged()
    End Sub
    ''' <summary>
    ''' Removes all actions and stops observing their presentation properties.
    ''' </summary>
    Protected Overrides Sub ClearItems()
        For Each Item As TextBoxAction In Items
            RemoveHandler Item.Changed, AddressOf Item_Changed
        Next
        MyBase.ClearItems()
        OnChanged()
    End Sub
    Private Sub ValidateItem(Item As TextBoxAction, ReplacedItem As TextBoxAction)
        ArgumentNullException.ThrowIfNull(Item)
        For Each ExistingItem As TextBoxAction In Items
            If Not ReferenceEquals(ExistingItem, ReplacedItem) AndAlso ReferenceEquals(ExistingItem, Item) Then Throw New InvalidOperationException("The same TextBoxAction instance cannot be added to the collection more than once.")
        Next
    End Sub
    Private Sub AssignDefaultKey(Item As TextBoxAction)
        If Not String.IsNullOrWhiteSpace(Item.Key) Then Return
        Dim Number As Integer = Count + 1
        Dim Candidate As String = $"Action{Number}"
        While FindByKey(Candidate) IsNot Nothing
            Number += 1
            Candidate = $"Action{Number}"
        End While
        Item.Key = Candidate
    End Sub
    Private Sub Item_Changed(Sender As Object, E As EventArgs)
        OnChanged()
    End Sub
    Private Sub OnChanged()
        RaiseEvent Changed(Me, EventArgs.Empty)
    End Sub
    ''' <summary>
    ''' Gets the action identified by the specified key.
    ''' </summary>
    ''' <param name="Key">The case-insensitive key of the action to retrieve.</param>
    ''' <value>The action associated with the specified key.</value>
    ''' <exception cref="ArgumentException">
    ''' Thrown when <paramref name="Key"/> is empty or contains only white-space characters.
    ''' </exception>
    ''' <exception cref="KeyNotFoundException">
    ''' Thrown when no action has the specified key.
    ''' </exception>
    Default Public Overloads ReadOnly Property Item(Key As String) As TextBoxAction
        Get
            If String.IsNullOrWhiteSpace(Key) Then Throw New ArgumentException("The action key cannot be empty.", NameOf(Key))
            Dim Action As TextBoxAction = FindByKey(Key)
            If Action Is Nothing Then Throw New KeyNotFoundException($"No action with the key '{Key}' was found.")
            Return Action
        End Get
    End Property
End Class
