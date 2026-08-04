Imports System.Threading
Partial Public Class AsyncLookupBox
    Private Sub ScheduleSearch()
        _SearchTimer.Stop()
        If _IsDisposing OrElse IsInDesignMode() OrElse _SuppressTextChanged OrElse Not SearchEnabled OrElse Not Enabled OrElse Me.ReadOnly Then Return
        If TextLength = 0 Then
            CancelActiveSearch()
            _Results = Array.Empty(Of Object)()
            CloseDropDown()
            Return
        End If
        If Not Focused Then Return
        If TextLength < MinimumCharacters Then
            CancelActiveSearch()
            _Results = Array.Empty(Of Object)()
            Dim RemainingCharacters As Integer = MinimumCharacters - TextLength
            ShowDropDownStatus(FormatCharactersRemainingMessage(RemainingCharacters))
            Return
        End If
        _SearchTimer.Start()
    End Sub
    Private Function CanSearchCurrentText() As Boolean
        Return SearchEnabled AndAlso Enabled AndAlso Not Me.ReadOnly AndAlso TextLength >= MinimumCharacters
    End Function
    Private Async Sub SearchTimerTick(Sender As Object, E As EventArgs)
        _SearchTimer.Stop()
        Await PerformSearchAsync()
    End Sub
    Private Async Sub StartSearch()
        _SearchTimer.Stop()
        Await PerformSearchAsync()
    End Sub
    Private Async Function PerformSearchAsync() As Task(Of IReadOnlyList(Of Object))
        If Not CanSearchCurrentText() OrElse _IsDisposing OrElse IsInDesignMode() Then Return Array.Empty(Of Object)()
        CancelActiveSearch()
        _SearchVersion += 1
        Dim RequestVersion As Long = _SearchVersion
        Dim RequestCancellation As New CancellationTokenSource()
        _SearchCancellation = RequestCancellation
        Dim SearchText As String = Text
        Dim Stopwatch As Stopwatch = Stopwatch.StartNew()
        SetIsSearching(True)
        ShowDropDownStatus(LoadingText)
        Try
            Dim Args As New AsyncLookupSearchRequestedEventArgs(SearchText, RequestCancellation.Token)
            RaiseEvent SearchRequested(Me, Args)
            If Args.Cancel Then
                If _Results.Count > 0 Then
                    ShowDropDownResults(_Results)
                Else
                    CloseDropDown()
                End If
                Return _Results
            End If
            If Not Args.HasSearchOperation Then Throw New InvalidOperationException(SearchNotConfiguredText)
            Dim SearchResults As IReadOnlyList(Of Object) = Await Args.GetResultsAsync()
            If RequestCancellation.IsCancellationRequested OrElse RequestVersion <> _SearchVersion Then Return Array.Empty(Of Object)()
            Dim NonNullResults As List(Of Object) = If(SearchResults, Array.Empty(Of Object)()).Where(Function(Item) Item IsNot Nothing).ToList()
            Dim WasTruncated As Boolean = MaximumResults > 0 AndAlso NonNullResults.Count > MaximumResults
            If WasTruncated Then NonNullResults = NonNullResults.Take(MaximumResults).ToList()
            _Results = NonNullResults
            Stopwatch.Stop()
            If AutoSelectSingleResult AndAlso _Results.Count = 1 Then
                SelectItemCore(_Results(0))
            ElseIf _Results.Count = 0 Then
                ShowDropDownStatus(NoResultsText)
            Else
                ShowDropDownResults(_Results)
            End If
            RaiseEvent SearchCompleted(Me, New AsyncLookupSearchCompletedEventArgs(SearchText, _Results, Stopwatch.Elapsed, WasTruncated))
            Return _Results
        Catch Canceled As OperationCanceledException When RequestCancellation.IsCancellationRequested OrElse RequestVersion <> _SearchVersion
            Return Array.Empty(Of Object)()
        Catch Failure As Exception
            If RequestCancellation.IsCancellationRequested OrElse RequestVersion <> _SearchVersion Then Return Array.Empty(Of Object)()
            Stopwatch.Stop()
            _Results = Array.Empty(Of Object)()
            ShowDropDownStatus(SearchErrorText)
            RaiseEvent SearchFailed(Me, New AsyncLookupSearchFailedEventArgs(SearchText, Failure))
            Return _Results
        Finally
            If RequestVersion = _SearchVersion Then
                If ReferenceEquals(_SearchCancellation, RequestCancellation) Then _SearchCancellation = Nothing
                SetIsSearching(False)
            End If
            RequestCancellation.Dispose()
        End Try
    End Function
    Private Sub CancelActiveSearch()
        _SearchTimer.Stop()
        _SearchVersion += 1
        Dim Cancellation As CancellationTokenSource = _SearchCancellation
        _SearchCancellation = Nothing
        If Cancellation IsNot Nothing AndAlso Not Cancellation.IsCancellationRequested Then Cancellation.Cancel()
        SetIsSearching(False)
    End Sub
    Private Sub SetIsSearching(Value As Boolean)
        If _IsSearching = Value Then Return
        _IsSearching = Value
        If _ActionButton IsNot Nothing Then _ActionButton.IsSearching = Value
        UpdateActionButtonLayout()
        RaiseEvent IsSearchingChanged(Me, EventArgs.Empty)
    End Sub
    Private Function FormatCharactersRemainingMessage(RemainingCharacters As Integer) As String
        Try
            Return String.Format(CharactersRemainingText, RemainingCharacters)
        Catch InvalidFormat As FormatException
            Return CharactersRemainingText
        End Try
    End Function
End Class
