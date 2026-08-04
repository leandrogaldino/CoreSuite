Imports System.Threading
''' <summary>
''' Provides data for the <see cref="AsyncLookupBox.SearchRequested"/> event and receives the asynchronous search operation supplied by the application.
''' </summary>
Public NotInheritable Class AsyncLookupSearchRequestedEventArgs
    Inherits EventArgs
    Private _SearchTask As Task(Of IReadOnlyList(Of Object))
    ''' <summary>
    ''' Initializes a new instance of the <see cref="AsyncLookupSearchRequestedEventArgs"/> class.
    ''' </summary>
    ''' <param name="SearchText">The text entered by the user.</param>
    ''' <param name="CancellationToken">The token canceled when the request is replaced or explicitly canceled.</param>
    Public Sub New(SearchText As String, CancellationToken As CancellationToken)
        Me.SearchText = If(SearchText, String.Empty)
        Me.CancellationToken = CancellationToken
    End Sub
    ''' <summary>
    ''' Gets the text that must be searched.
    ''' </summary>
    ''' <value>The current lookup text.</value>
    Public ReadOnly Property SearchText As String
    ''' <summary>
    ''' Gets the token canceled when a newer search replaces this request, the text is cleared, or the control is disposed.
    ''' </summary>
    ''' <value>The cancellation token associated with this request.</value>
    Public ReadOnly Property CancellationToken As CancellationToken
    ''' <summary>
    ''' Gets or sets a value indicating whether this request should end without changing the current results.
    ''' </summary>
    ''' <value><see langword="True"/> to cancel request processing; otherwise, <see langword="False"/>.</value>
    Public Property Cancel As Boolean
    ''' <summary>
    ''' Gets a value indicating whether a task or immediate result collection was supplied.
    ''' </summary>
    ''' <value><see langword="True"/> when <see cref="SetSearchTask(Of TResult)"/> or <see cref="SetResults"/> was called.</value>
    Public ReadOnly Property HasSearchOperation As Boolean
        Get
            Return _SearchTask IsNot Nothing
        End Get
    End Property
    ''' <summary>
    ''' Supplies the task that retrieves lookup results.
    ''' </summary>
    ''' <typeparam name="TResult">A result type that implements <see cref="IEnumerable"/>, such as a list or array of business objects.</typeparam>
    ''' <param name="SearchTask">The task returned by the application search service.</param>
    ''' <remarks>This generic method accepts tasks returning <c>List(Of T)</c>, arrays, and other enumerable result types without requiring conversion to <c>IEnumerable(Of Object)</c>.</remarks>
    Public Sub SetSearchTask(Of TResult)(SearchTask As Task(Of TResult))
        ArgumentNullException.ThrowIfNull(SearchTask)
        If _SearchTask IsNot Nothing Then Throw New InvalidOperationException("A search operation has already been supplied for this request.")
        _SearchTask = ConvertSearchTaskAsync(SearchTask)
    End Sub
    ''' <summary>
    ''' Supplies an already available lookup-result collection.
    ''' </summary>
    ''' <param name="Results">The objects that must be displayed in the result list.</param>
    Public Sub SetResults(Results As IEnumerable)
        ArgumentNullException.ThrowIfNull(Results)
        If TypeOf Results Is String Then Throw New ArgumentException("A string is not a valid lookup-result collection.", NameOf(Results))
        If _SearchTask IsNot Nothing Then Throw New InvalidOperationException("A search operation has already been supplied for this request.")
        _SearchTask = Task.FromResult(Of IReadOnlyList(Of Object))(Results.Cast(Of Object)().ToList())
    End Sub
    Friend Function GetResultsAsync() As Task(Of IReadOnlyList(Of Object))
        Return _SearchTask
    End Function
    Private Shared Async Function ConvertSearchTaskAsync(Of TResult)(SearchTask As Task(Of TResult)) As Task(Of IReadOnlyList(Of Object))
        Dim Result As TResult = Await SearchTask
        Dim EnumerableResult As IEnumerable = TryCast(CType(Result, Object), IEnumerable)
        If EnumerableResult Is Nothing OrElse TypeOf EnumerableResult Is String Then Throw New InvalidOperationException("The supplied search task must return an enumerable collection of result objects.")
        Return EnumerableResult.Cast(Of Object)().ToList()
    End Function
End Class
