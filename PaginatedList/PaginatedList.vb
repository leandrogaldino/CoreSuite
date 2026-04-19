Imports System.Collections.ObjectModel

''' <summary>
''' Generic class representing a paginated list. Allows splitting a collection of items into multiple pages,
''' where each page contains a specific number of items (defined by <see cref="PageSize"/>).
''' </summary>
''' <typeparam name="T">The type of items in the list.</typeparam>
Public Class PaginatedList(Of T)
    Inherits ObservableCollection(Of T)
    Private _PageSize As Integer

    ''' <summary>
    ''' Gets the list of pages, where each page is a sublist of paginated items.
    ''' </summary>
    ''' <remarks>
    ''' Pages are generated based on the value of <see cref="PageSize"/> and the original list of items.
    ''' Each page is a sublist containing up to <see cref="PageSize"/> items.
    ''' </remarks>
    ''' <returns>A list of lists containing the paginated items.</returns>
    Public ReadOnly Property Pages As List(Of List(Of T))
        Get
            Return GetPaginatedList(MyBase.Items, _PageSize)
        End Get
    End Property

    ''' <summary>
    ''' The page size, i.e., the number of items each page should contain.
    ''' </summary>
    ''' <remarks>
    ''' If the value of <see cref="PageSize"/> is changed, the pages will be recalculated based on the new size.
    ''' </remarks>
    ''' <value>The number of items per page.</value>
    Public Property PageSize As Integer
        Get
            Return _PageSize
        End Get
        Set(value As Integer)
            _PageSize = value
        End Set
    End Property

    ''' <summary>
    ''' Creates a new paginated list with the items in reverse order.
    ''' </summary>
    ''' <returns>A new paginated list with the items reversed.</returns>
    Public Function Reverse() As PaginatedList(Of T)
        Dim ReversedList As List(Of T)
        Dim ReversedPaginatedList As PaginatedList(Of T)
        ReversedList = New List(Of T)(MyBase.Items)
        ReversedList.Reverse()
        ReversedPaginatedList = New PaginatedList(Of T)(_PageSize)
        For Each i In ReversedList
            ReversedPaginatedList.Add(i)
        Next
        Return ReversedPaginatedList
    End Function

    ''' <summary>
    ''' Initializes a new instance of the <see cref="PaginatedList(Of T)"/> class.
    ''' </summary>
    ''' <param name="PageSize">The number of items per page.</param>
    ''' <exception cref="ArgumentOutOfRangeException">Thrown if <paramref name="PageSize"/> is less than or equal to zero.</exception>
    Public Sub New(PageSize As Integer)
        If PageSize <= 0 Then
            Throw New ArgumentOutOfRangeException("PageSize", "The number of elements per page must be greater than zero.")
        End If
        _PageSize = PageSize
    End Sub
    Protected Overrides Sub InsertItem(index As Integer, item As T)
        MyBase.InsertItem(index, item)
    End Sub
    Protected Overrides Sub ClearItems()
        MyBase.ClearItems()
    End Sub
    ''' <summary>
    ''' Gets the index of the page where the specified item is located.
    ''' </summary>
    ''' <param name="Item">The item to locate.</param>
    ''' <returns>The index of the page containing the item. Returns -1 if the item is not found.</returns>
    Public Function PageIndexOf(Item As T) As Integer
        For pageIndex As Integer = 0 To Pages.Count - 1
            If Pages(pageIndex).Contains(Item) Then
                Return pageIndex
            End If
        Next
        Return -1
    End Function
    Protected Overrides Sub RemoveItem(index As Integer)
        MyBase.RemoveItem(index)
    End Sub
    Protected Overrides Sub SetItem(index As Integer, item As T)
        MyBase.SetItem(index, item)
    End Sub
    Private Function GetPaginatedList(OriginalList As List(Of T), PageSize As Integer) As List(Of List(Of T))
        Dim paginatedList As New List(Of List(Of T))()
        Dim pageCount As Integer = Math.Ceiling(OriginalList.Count / PageSize)
        For i As Integer = 0 To pageCount - 1
            Dim currentPage As New List(Of T)()
            currentPage.AddRange(OriginalList.Skip(i * PageSize).Take(PageSize))
            paginatedList.Add(currentPage)
        Next
        Return paginatedList
    End Function
End Class