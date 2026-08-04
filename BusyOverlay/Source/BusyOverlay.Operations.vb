Imports System.Runtime.ExceptionServices
Imports System.Threading
Partial Public Class BusyOverlay
    ''' <summary>
    ''' Shows the overlay immediately and keeps it visible until <see cref="HideOverlay"/> is called.
    ''' </summary>
    ''' <remarks>This method is idempotent. Active scoped or asynchronous operations continue to keep the overlay visible after <see cref="HideOverlay"/> clears the manual state.</remarks>
    Public Sub ShowOverlay()
        EnsureCanOperate()
        Dim WasBusy As Boolean = IsBusy
        _ManualBusy = True
        _OperationDisplayRequested = True
        If Not WasBusy Then OnBusyStarted()
        SynchronizeOverlayVisibility()
    End Sub
    ''' <summary>
    ''' Clears the manual busy state created by <see cref="ShowOverlay"/>.
    ''' </summary>
    ''' <remarks>The overlay remains visible when a scoped or asynchronous operation is still active.</remarks>
    Public Sub HideOverlay()
        EnsureUiThread()
        If Not _ManualBusy Then Return
        Dim WasBusy As Boolean = IsBusy
        _ManualBusy = False
        If _ActiveOperationCount = 0 Then _OperationDisplayRequested = False
        SynchronizeOverlayVisibility()
        If WasBusy AndAlso Not IsBusy Then OnBusyEnded()
    End Sub
    ''' <summary>
    ''' Begins a reference-counted operation and shows the overlay immediately.
    ''' </summary>
    ''' <returns>A disposable scope that completes this operation.</returns>
    ''' <example>
    ''' <code>
    ''' Using BusyScope As BusyOverlayScope = BusyOverlay1.BeginOperation()
    '''     Await SaveChangesAsync()
    ''' End Using
    ''' </code>
    ''' </example>
    Public Function BeginOperation() As BusyOverlayScope
        EnsureCanOperate()
        Dim WasBusy As Boolean = IsBusy
        _ActiveOperationCount += 1
        _OperationDisplayRequested = True
        If Not WasBusy Then OnBusyStarted()
        SynchronizeOverlayVisibility()
        Return New BusyOverlayScope(Me)
    End Function
    ''' <summary>
    ''' Runs an asynchronous operation, applying delayed display and minimum display time automatically.
    ''' </summary>
    ''' <param name="Operation">The asynchronous operation to execute.</param>
    ''' <returns>A task that completes after the operation and overlay cleanup complete.</returns>
    Public Function RunAsync(Operation As Func(Of Task)) As Task
        ArgumentNullException.ThrowIfNull(Operation)
        Return RunCoreAsync(Function(UnusedToken As CancellationToken) InvokeWithoutResultAsync(Operation), False, CancellationToken.None)
    End Function
    ''' <summary>
    ''' Runs an asynchronous operation that receives a cancellation token controlled by the overlay.
    ''' </summary>
    ''' <param name="Operation">The asynchronous operation to execute.</param>
    ''' <param name="CancellationToken">An optional external cancellation token linked to the overlay request.</param>
    ''' <returns>A task that completes after the operation and overlay cleanup complete.</returns>
    Public Function RunAsync(Operation As Func(Of CancellationToken, Task), Optional CancellationToken As CancellationToken = Nothing) As Task
        ArgumentNullException.ThrowIfNull(Operation)
        Return RunCoreAsync(Function(operationToken As CancellationToken) InvokeWithoutResultAsync(Operation, operationToken), True, CancellationToken)
    End Function
    ''' <summary>
    ''' Runs an asynchronous operation that returns a result, applying delayed display and minimum display time automatically.
    ''' </summary>
    ''' <typeparam name="TResult">The type returned by the operation.</typeparam>
    ''' <param name="Operation">The asynchronous operation to execute.</param>
    ''' <returns>A task containing the operation result.</returns>
    Public Function RunAsync(Of TResult)(Operation As Func(Of Task(Of TResult))) As Task(Of TResult)
        ArgumentNullException.ThrowIfNull(Operation)
        Return RunCoreAsync(Function(UnusedToken As CancellationToken) Operation(), False, CancellationToken.None)
    End Function
    ''' <summary>
    ''' Runs an asynchronous operation that receives cancellation and returns a result.
    ''' </summary>
    ''' <typeparam name="TResult">The type returned by the operation.</typeparam>
    ''' <param name="Operation">The asynchronous operation to execute.</param>
    ''' <param name="cancellationToken">An optional external cancellation token linked to the overlay request.</param>
    ''' <returns>A task containing the operation result.</returns>
    Public Function RunAsync(Of TResult)(Operation As Func(Of CancellationToken, Task(Of TResult)), Optional cancellationToken As CancellationToken = Nothing) As Task(Of TResult)
        ArgumentNullException.ThrowIfNull(Operation)
        Return RunCoreAsync(Operation, True, cancellationToken)
    End Function
    ''' <summary>
    ''' Updates determinate progress and optionally replaces the detail text.
    ''' </summary>
    ''' <param name="Value">A value within <see cref="ProgressMinimum"/> and <see cref="ProgressMaximum"/>.</param>
    ''' <param name="DetailText">Optional detail text. Pass <see langword="Nothing"/> to keep the current text.</param>
    Public Sub ReportProgress(Value As Integer, Optional DetailText As String = Nothing)
        EnsureUiThread()
        SetProgress(Value, DetailText, True)
    End Sub
    ''' <summary>
    ''' Raises <see cref="CancellationRequested"/> and cancels every active cancellable operation when the request is accepted.
    ''' </summary>
    ''' <returns><see langword="True"/> when at least one cancellation source was canceled; otherwise, <see langword="False"/>.</returns>
    Public Function RequestCancellation() As Boolean
        EnsureUiThread()
        If _CancellationSources.Count = 0 Then Return False
        Dim EventData As New BusyOverlayCancellationEventArgs(_CancellationSources.Count)
        RaiseEvent CancellationRequested(Me, EventData)
        If EventData.Cancel Then Return False
        Dim CancellationRequested As Boolean
        For Each CancellationSource As CancellationTokenSource In _CancellationSources.ToArray()
            Try
                CancellationSource.Cancel()
                CancellationRequested = True
            Catch ex As ObjectDisposedException
            End Try
        Next
        If _View IsNot Nothing AndAlso Not _View.IsDisposed Then _View.ApplySettings()
        Return CancellationRequested
    End Function
    ''' <summary>
    ''' Completes one operation created by <see cref="BeginOperation"/> and hides the surface when no other owner remains.
    ''' </summary>
    Friend Sub EndScopedOperation()
        If _IsDisposed Then Return
        If _TargetControl IsNot Nothing AndAlso _TargetControl.IsHandleCreated AndAlso _TargetControl.InvokeRequired Then
            _TargetControl.BeginInvoke(New MethodInvoker(AddressOf EndScopedOperation))
            Return
        End If
        If _ActiveOperationCount <= 0 Then Return
        Dim WasBusy As Boolean = IsBusy
        _ActiveOperationCount -= 1
        If _ActiveOperationCount = 0 AndAlso Not _ManualBusy Then _OperationDisplayRequested = False
        RefreshOverlayAppearance()
        SynchronizeOverlayVisibility()
        If WasBusy AndAlso Not IsBusy Then OnBusyEnded()
    End Sub
    Private Async Function RunCoreAsync(Of TResult)(Operation As Func(Of CancellationToken, Task(Of TResult)), Cancellable As Boolean, CancellationToken As CancellationToken) As Task(Of TResult)
        EnsureCanOperate()
        Dim LinkedSource As CancellationTokenSource =
        CancellationTokenSource.CreateLinkedTokenSource(CancellationToken)
        RegisterAsyncOperation(LinkedSource, Cancellable)
        Dim Result As TResult = Nothing
        Dim CapturedException As Exception = Nothing
        Try
            LinkedSource.Token.ThrowIfCancellationRequested()
            Dim OperationTask As Task(Of TResult) = Operation(LinkedSource.Token)
            If OperationTask Is Nothing Then
                Throw New InvalidOperationException("The asynchronous operation returned Nothing instead of a Task.")
            End If
            If Not IsOverlayVisible AndAlso _OperationDisplayDelay > 0 Then
                Dim DelayTask As Task = Task.Delay(_OperationDisplayDelay, CancellationToken)
                Dim CompletedTask As Task = Await Task.WhenAny(OperationTask, DelayTask)
                If ReferenceEquals(CompletedTask, DelayTask) Then
                    RequestOperationDisplay()
                    Await Task.Yield()
                End If
            Else
                RequestOperationDisplay()
                Await Task.Yield()
            End If
            Result = Await OperationTask
        Catch ex As Exception
            CapturedException = ex
        End Try
        Await CompleteAsyncOperation(LinkedSource)
        If CapturedException IsNot Nothing Then
            ExceptionDispatchInfo.Capture(CapturedException).Throw()
        End If
        Return Result
    End Function
    Private Shared Async Function InvokeWithoutResultAsync(Operation As Func(Of Task)) As Task(Of Boolean)
        Await Operation()
        Return True
    End Function
    Private Shared Async Function InvokeWithoutResultAsync(Operation As Func(Of CancellationToken, Task), CancellationToken As CancellationToken) As Task(Of Boolean)
        Await Operation(CancellationToken)
        Return True
    End Function
    Private Sub RegisterAsyncOperation(CancellationSource As CancellationTokenSource, Cancellable As Boolean)
        Dim WasBusy As Boolean = IsBusy
        _ActiveOperationCount += 1
        If Cancellable Then _CancellationSources.Add(CancellationSource)
        If Not WasBusy Then OnBusyStarted()
        RefreshOverlayAppearance()
    End Sub
    Private Sub RequestOperationDisplay()
        If _IsDisposed OrElse _ActiveOperationCount <= 0 Then Return
        _OperationDisplayRequested = True
        SynchronizeOverlayVisibility()
    End Sub
    Private Async Function CompleteAsyncOperation(CancellationSource As CancellationTokenSource) As Task
        If Not _IsDisposed AndAlso _ActiveOperationCount = 1 AndAlso Not _ManualBusy AndAlso _OperationDisplayRequested AndAlso IsOverlayVisible AndAlso _MinimumOperationDisplayTime > 0 AndAlso _OverlayShownAt.HasValue Then
            Dim Elapsed As TimeSpan = DateTimeOffset.UtcNow - _OverlayShownAt.Value
            Dim RemainingMilliseconds As Integer = _MinimumOperationDisplayTime - CInt(Math.Floor(Elapsed.TotalMilliseconds))
            If RemainingMilliseconds > 0 Then Await Task.Delay(RemainingMilliseconds)
        End If
        Dim WasBusy As Boolean = IsBusy
        _CancellationSources.Remove(CancellationSource)
        CancellationSource.Dispose()
        If _ActiveOperationCount > 0 Then _ActiveOperationCount -= 1
        If _ActiveOperationCount = 0 AndAlso Not _ManualBusy Then _OperationDisplayRequested = False
        RefreshOverlayAppearance()
        SynchronizeOverlayVisibility()
        If WasBusy AndAlso Not IsBusy Then OnBusyEnded()
    End Function
    Private Sub SetProgress(Value As Integer, DetailText As String, RaiseChangedEvent As Boolean)
        ValidateRange(Value, _ProgressMinimum, _ProgressMaximum, NameOf(Value))
        Dim ValueChanged As Boolean = _ProgressValue <> Value
        Dim DetailChanged As Boolean = DetailText IsNot Nothing AndAlso Not String.Equals(_DetailText, DetailText, StringComparison.Ordinal)
        If Not ValueChanged AndAlso Not DetailChanged Then Return
        _ProgressValue = Value
        If DetailText IsNot Nothing Then _DetailText = DetailText
        RefreshOverlayAppearance()
        If RaiseChangedEvent Then OnProgressChanged()
    End Sub
    Private Sub OverlayView_CancellationClick(sender As Object, e As EventArgs)
        RequestCancellation()
    End Sub
    Private Sub EnsureCanOperate()
        ObjectDisposedException.ThrowIf(_IsDisposed, Me)
        If IsInDesignMode Then Throw New InvalidOperationException("BusyOverlay run-time methods cannot be used in the Windows Forms Designer.")
        If _TargetControl Is Nothing Then Throw New InvalidOperationException("TargetControl must be assigned before the overlay can be used.")
        If _TargetControl.IsDisposed Then Throw New ObjectDisposedException(_TargetControl.Name, "The busy overlay target has been disposed.")
        If TypeOf _TargetControl IsNot Form AndAlso _TargetControl.Parent Is Nothing Then Throw New InvalidOperationException("TargetControl must be a Form or belong to a parent control.")
        EnsureUiThread()
    End Sub
    Private Sub EnsureUiThread()
        If _TargetControl IsNot Nothing AndAlso _TargetControl.IsHandleCreated AndAlso _TargetControl.InvokeRequired Then Throw New InvalidOperationException("BusyOverlay must be accessed from the UI thread that owns TargetControl.")
    End Sub
End Class
