Imports System.Collections.ObjectModel
''' <summary>
''' Represents the ordered collection of pages managed by a <see cref="NavigationView"/>.
''' </summary>
Public NotInheritable Class NavigationPageCollection
    Inherits Collection(Of NavigationPage)
    Friend Event Changed As EventHandler(Of NavigationPageCollectionChangedEventArgs)
    ''' <summary>
    ''' Adds and returns a page that creates the specified <see cref="UserControl"/> type.
    ''' </summary>
    ''' <typeparam name="TControl">The page control type. It must provide a public parameterless constructor.</typeparam>
    ''' <param name="Key">The unique key used to locate the page.</param>
    ''' <param name="Text">The text displayed by the navigation button.</param>
    ''' <param name="Image">The optional image displayed by the navigation button.</param>
    ''' <returns>The page added to the collection.</returns>
    Public Overloads Function Add(Of TControl As {UserControl, New})(Key As String, Text As String, Optional Image As Image = Nothing) As NavigationPage
        Return Add(Key, Text, GetType(TControl), Image)
    End Function
    ''' <summary>
    ''' Adds and returns a page that creates a control of the specified type.
    ''' </summary>
    ''' <param name="Key">The unique key used to locate the page.</param>
    ''' <param name="Text">The text displayed by the navigation button.</param>
    ''' <param name="ControlType">A non-abstract <see cref="UserControl"/> type with a public parameterless constructor.</param>
    ''' <param name="Image">The optional image displayed by the navigation button.</param>
    ''' <returns>The page added to the collection.</returns>
    Public Overloads Function Add(Key As String, Text As String, ControlType As Type, Optional Image As Image = Nothing) As NavigationPage
        Dim Page As New NavigationPage(Key, Text, ControlType, Image)
        Add(Page)
        Return Page
    End Function
    ''' <summary>
    ''' Adds and returns a page that creates its control through a run-time factory.
    ''' </summary>
    ''' <param name="Key">The unique key used to locate the page.</param>
    ''' <param name="Text">The text displayed by the navigation button.</param>
    ''' <param name="Image">The optional image displayed by the navigation button.</param>
    ''' <param name="Factory">The factory that creates a new page control.</param>
    ''' <returns>The page added to the collection.</returns>
    Public Overloads Function Add(Key As String, Text As String, Image As Image, Factory As Func(Of UserControl)) As NavigationPage
        ArgumentNullException.ThrowIfNull(Factory)
        Dim Page As New NavigationPage(Key, Text, Image, Factory)
        Add(Page)
        Return Page
    End Function
    ''' <summary>
    ''' Finds the page whose key matches the specified value without regard to case.
    ''' </summary>
    ''' <param name="Key">The key to locate.</param>
    ''' <returns>The matching page, or <see langword="Nothing"/> when no page is found.</returns>
    Public Function FindByKey(Key As String) As NavigationPage
        If String.IsNullOrWhiteSpace(Key) Then Return Nothing
        For Each Page As NavigationPage In Items
            If String.Equals(Page.Key, Key, StringComparison.OrdinalIgnoreCase) Then Return Page
        Next
        Return Nothing
    End Function
    ''' <summary>
    ''' Determines whether the collection contains a page with the specified key.
    ''' </summary>
    ''' <param name="Key">The case-insensitive key to locate.</param>
    ''' <returns><see langword="True"/> when a matching page exists; otherwise, <see langword="False"/>.</returns>
    Public Overloads Function Contains(Key As String) As Boolean
        Return FindByKey(Key) IsNot Nothing
    End Function
    ''' <summary>
    ''' Returns the zero-based index of the page with the specified key.
    ''' </summary>
    ''' <param name="Key">The case-insensitive key to locate.</param>
    ''' <returns>The page index, or <c>-1</c> when no page is found.</returns>
    Public Overloads Function IndexOf(Key As String) As Integer
        For Index As Integer = 0 To Count - 1
            If String.Equals(Me(Index).Key, Key, StringComparison.OrdinalIgnoreCase) Then Return Index
        Next
        Return -1
    End Function
    ''' <summary>
    ''' Removes the page with the specified key.
    ''' </summary>
    ''' <param name="Key">The case-insensitive key of the page to remove.</param>
    ''' <returns><see langword="True"/> when a page was removed; otherwise, <see langword="False"/>.</returns>
    Public Overloads Function Remove(Key As String) As Boolean
        Dim Page As NavigationPage = FindByKey(Key)
        If Page Is Nothing Then Return False
        Return Remove(Page)
    End Function
    ''' <summary>
    ''' Gets the page identified by the specified key.
    ''' </summary>
    ''' <param name="Key">The case-insensitive key of the page to retrieve.</param>
    ''' <value>The page associated with the specified key.</value>
    ''' <exception cref="ArgumentException">Thrown when <paramref name="Key"/> is empty.</exception>
    ''' <exception cref="KeyNotFoundException">Thrown when no page has the specified key.</exception>
    Default Public Overloads ReadOnly Property Item(Key As String) As NavigationPage
        Get
            If String.IsNullOrWhiteSpace(Key) Then Throw New ArgumentException("The page key cannot be empty.", NameOf(Key))
            Dim Page As NavigationPage = FindByKey(Key)
            If Page Is Nothing Then Throw New KeyNotFoundException($"No navigation page with the key '{Key}' was found.")
            Return Page
        End Get
    End Property
    ''' <summary>
    ''' Inserts a page and begins observing changes to its properties.
    ''' </summary>
    ''' <param name="Index">The zero-based insertion index.</param>
    ''' <param name="Item">The page to insert.</param>
    Protected Overrides Sub InsertItem(Index As Integer, Item As NavigationPage)
        ValidateItem(Item, Nothing)
        AssignDefaultKey(Item)
        ValidateProposedKey(Item.Key, Nothing)
        MyBase.InsertItem(Index, Item)
        Item.Owner = Me
        AddHandler Item.Changed, AddressOf PageChanged
        OnChanged(New NavigationPageCollectionChangedEventArgs(NavigationPageCollectionChangeAction.Add, Item, Nothing))
    End Sub
    ''' <summary>
    ''' Replaces a page and updates property-change observation.
    ''' </summary>
    ''' <param name="Index">The zero-based index of the page to replace.</param>
    ''' <param name="Item">The replacement page.</param>
    Protected Overrides Sub SetItem(Index As Integer, Item As NavigationPage)
        Dim CurrentPage As NavigationPage = Items(Index)
        If ReferenceEquals(CurrentPage, Item) Then Return
        ValidateItem(Item, CurrentPage)
        AssignDefaultKey(Item)
        ValidateProposedKey(Item.Key, CurrentPage)
        RemoveHandler CurrentPage.Changed, AddressOf PageChanged
        CurrentPage.Owner = Nothing
        MyBase.SetItem(Index, Item)
        Item.Owner = Me
        AddHandler Item.Changed, AddressOf PageChanged
        OnChanged(New NavigationPageCollectionChangedEventArgs(NavigationPageCollectionChangeAction.Replace, Item, {CurrentPage}))
    End Sub
    ''' <summary>
    ''' Removes a page and stops observing its properties.
    ''' </summary>
    ''' <param name="Index">The zero-based index of the page to remove.</param>
    Protected Overrides Sub RemoveItem(Index As Integer)
        Dim Page As NavigationPage = Items(Index)
        RemoveHandler Page.Changed, AddressOf PageChanged
        Page.Owner = Nothing
        MyBase.RemoveItem(Index)
        OnChanged(New NavigationPageCollectionChangedEventArgs(NavigationPageCollectionChangeAction.Remove, Nothing, {Page}))
    End Sub
    ''' <summary>
    ''' Removes all pages and stops observing their properties.
    ''' </summary>
    Protected Overrides Sub ClearItems()
        If Count = 0 Then Return
        Dim OldPages(Count - 1) As NavigationPage
        For Index As Integer = 0 To Count - 1
            OldPages(Index) = Items(Index)
        Next
        For Each Page As NavigationPage In OldPages
            RemoveHandler Page.Changed, AddressOf PageChanged
            Page.Owner = Nothing
        Next
        MyBase.ClearItems()
        OnChanged(New NavigationPageCollectionChangedEventArgs(NavigationPageCollectionChangeAction.Reset, Nothing, OldPages))
    End Sub
    Friend Sub ValidateProposedKey(Key As String, IgnoredPage As NavigationPage)
        If String.IsNullOrWhiteSpace(Key) Then Throw New ArgumentException("The page key cannot be empty.", NameOf(Key))
        For Each Page As NavigationPage In Items
            If Not ReferenceEquals(Page, IgnoredPage) AndAlso String.Equals(Page.Key, Key, StringComparison.OrdinalIgnoreCase) Then Throw New ArgumentException($"A navigation page with the key '{Key}' already exists.", NameOf(Key))
        Next
    End Sub
    Private Sub ValidateItem(Item As NavigationPage, ReplacedPage As NavigationPage)
        ArgumentNullException.ThrowIfNull(Item)
        If Item.Owner IsNot Nothing AndAlso Not ReferenceEquals(Item, ReplacedPage) Then Throw New InvalidOperationException("The NavigationPage already belongs to a collection.")
        For Each ExistingPage As NavigationPage In Items
            If Not ReferenceEquals(ExistingPage, ReplacedPage) AndAlso ReferenceEquals(ExistingPage, Item) Then Throw New InvalidOperationException("The same NavigationPage instance cannot be added more than once.")
        Next
    End Sub
    Private Sub AssignDefaultKey(Page As NavigationPage)
        If Not String.IsNullOrWhiteSpace(Page.Key) Then Return
        Dim Number As Integer = Count + 1
        Dim Candidate As String = $"Page{Number}"
        While FindByKey(Candidate) IsNot Nothing
            Number += 1
            Candidate = $"Page{Number}"
        End While
        Page.Key = Candidate
    End Sub
    Private Sub PageChanged(Sender As Object, E As EventArgs)
        OnChanged(New NavigationPageCollectionChangedEventArgs(NavigationPageCollectionChangeAction.ItemChanged, DirectCast(Sender, NavigationPage), Nothing))
    End Sub
    Private Sub OnChanged(E As NavigationPageCollectionChangedEventArgs)
        RaiseEvent Changed(Me, E)
    End Sub
End Class
