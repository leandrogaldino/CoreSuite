''' <summary>
''' Limits the frequency at which progress notifications are emitted.
''' </summary>
''' <remarks>
''' The throttle reduces synchronization-context traffic and UI updates during high-throughput file operations.
''' </remarks>
Friend NotInheritable Class ProgressThrottle
    ''' <summary>
    ''' Defines the minimum elapsed time required between non-forced progress reports.
    ''' </summary>
    Private ReadOnly ReportInterval As TimeSpan = TimeSpan.FromMilliseconds(75)
    ''' <summary>
    ''' Measures the elapsed time since the throttle instance was created.
    ''' </summary>
    Private ReadOnly _Stopwatch As Stopwatch = Stopwatch.StartNew()
    ''' <summary>
    ''' Stores the elapsed time at which the previous progress report was approved.
    ''' </summary>
    Private _LastReportTime As TimeSpan
    ''' <summary>
    ''' Determines whether a progress notification should be emitted at the current time.
    ''' </summary>
    ''' <param name="Force">
    ''' <see langword="True"/> to approve the notification regardless of the elapsed interval; otherwise, <see langword="False"/>.
    ''' </param>
    ''' <returns>
    ''' <see langword="True"/> when the notification should be emitted; otherwise, <see langword="False"/>.
    ''' </returns>
    Public Function ShouldReport(Optional Force As Boolean = False) As Boolean
        Dim CurrentTime As TimeSpan = _Stopwatch.Elapsed
        If Not Force AndAlso CurrentTime - _LastReportTime < ReportInterval Then Return False
        _LastReportTime = CurrentTime
        Return True
    End Function
End Class