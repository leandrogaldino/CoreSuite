Imports System.Net
Imports System.Net.Http
Imports System.Threading
''' <summary>
''' Provides internet connectivity monitoring, periodically checking access to known
''' endpoints and raising events when the connectivity status changes.
''' </summary>
Public Class Connectivity
    Implements IDisposable
    Private _LastStatus As Boolean
    Private _IsMonitoring As Boolean
    Private _Cancellation As CancellationTokenSource
    Private _HasStatus As Boolean
    Private Shared ReadOnly _Client As New HttpClient() With {.Timeout = TimeSpan.FromSeconds(3)}
    Private Shared ReadOnly _Addresses As String() = {
        "https://www.google.com/",
        "https://www.microsoft.com/",
        "https://www.facebook.com/",
        "https://aws.amazon.com/",
        "https://www.cloudflare.com/",
        "https://github.com/",
        "https://www.youtube.com/"
    }
    ''' <summary>
    ''' Occurs when the connectivity status changes (connected or disconnected).
    ''' </summary>
    Public Event ConnectivityChanged(sender As Object, e As ConnectivityEventArgs)
    ''' <summary>
    ''' Occurs when an internet connection is established.
    ''' </summary>
    Public Event Connected(sender As Object, e As EventArgs)
    ''' <summary>
    ''' Occurs when the internet connection is lost.
    ''' </summary>
    Public Event Disconnected(sender As Object, e As EventArgs)
    ''' <summary>
    ''' Gets or sets the time interval between connectivity checks while monitoring is active.
    ''' </summary>
    Public Property MonitoringInterval As TimeSpan = TimeSpan.FromSeconds(1)
    ''' <summary>
    ''' Gets a value indicating whether connectivity monitoring is currently running.
    ''' </summary>
    Public ReadOnly Property IsMonitoring As Boolean
        Get
            Return _IsMonitoring
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating whether an internet connection is available, based on the last check performed.
    ''' </summary>
    Public ReadOnly Property IsConnected As Boolean
        Get
            Return _HasStatus AndAlso _LastStatus
        End Get
    End Property
    ''' <summary>
    ''' Starts continuous background monitoring of internet connectivity.
    ''' </summary>
    Public Sub StartMonitoring()
        If _IsMonitoring Then Return
        _Cancellation?.Dispose()
        _Cancellation = New CancellationTokenSource()
        _IsMonitoring = True
        Dim Task As Task = MonitorConnectivity()
    End Sub
    ''' <summary>
    ''' Runs the asynchronous loop that periodically checks connectivity,
    ''' raising the corresponding events whenever the status changes.
    ''' </summary>
    ''' <returns>A task representing the asynchronous monitoring operation.</returns>
    Private Async Function MonitorConnectivity() As Task
        Try
            _LastStatus = Await IsAvailableAsync().ConfigureAwait(False)
            While _IsMonitoring
                Dim currentStatus = Await IsAvailableAsync().ConfigureAwait(False)
                If currentStatus <> _LastStatus Then
                    CheckStatusChange(currentStatus)
                End If
                Await Task.Delay(MonitoringInterval, _Cancellation.Token).ConfigureAwait(False)
            End While
        Catch ex As OperationCanceledException
        Catch ex As Exception
        End Try
    End Function
    ''' <summary>
    ''' Stops the currently running connectivity monitoring.
    ''' </summary>
    Public Sub StopMonitoring()
        If Not _IsMonitoring Then Return
        _IsMonitoring = False
        Dim cancellation = _Cancellation
        _Cancellation = Nothing
        cancellation?.Cancel()
        cancellation?.Dispose()
    End Sub
    ''' <summary>
    ''' Updates the stored connectivity status and raises the appropriate events
    ''' based on the new state.
    ''' </summary>
    ''' <param name="CurrentStatus">The currently detected connectivity status.</param>
    Private Sub CheckStatusChange(CurrentStatus As Boolean)
        _LastStatus = CurrentStatus
        _HasStatus = True
        Dim Args As New ConnectivityEventArgs(CurrentStatus)
        RaiseEvent ConnectivityChanged(Me, Args)
        If CurrentStatus Then
            RaiseEvent Connected(Me, EventArgs.Empty)
        Else
            RaiseEvent Disconnected(Me, EventArgs.Empty)
        End If
    End Sub
    ''' <summary>
    ''' Asynchronously checks whether an internet connection is available by
    ''' testing access to a list of known endpoints.
    ''' </summary>
    ''' <returns>A task that returns <c>True</c> if any endpoint responds successfully; otherwise, <c>False</c>.</returns>
    Public Async Function IsAvailableAsync() As Task(Of Boolean)
        Dim Token = If(_Cancellation?.Token, CancellationToken.None)
        For Each Address In _Addresses
            Try
                Using Request As New HttpRequestMessage(HttpMethod.Head, Address)
                    Using Response = Await _Client.SendAsync(Request, HttpCompletionOption.ResponseHeadersRead, Token).ConfigureAwait(False)
                        If Response.IsSuccessStatusCode Then Return True
                        If Response.StatusCode = HttpStatusCode.MethodNotAllowed Then
                            Using GetResponse = Await _Client.GetAsync(Address, HttpCompletionOption.ResponseHeadersRead, Token).ConfigureAwait(False)
                                If GetResponse.IsSuccessStatusCode Then Return True
                            End Using
                        End If
                    End Using
                End Using
            Catch ex As OperationCanceledException
                If Token.IsCancellationRequested Then Throw
            Catch ex As HttpRequestException
            Catch ex As Exception
            End Try
        Next
        Return False
    End Function
    ''' <summary>
    ''' Synchronously checks whether an internet connection is available.
    ''' </summary>
    ''' <returns><c>True</c> if the internet is available; otherwise, <c>False</c>.</returns>
    Public Function IsAvailable() As Boolean
        Return IsAvailableAsync().GetAwaiter().GetResult()
    End Function
    ''' <summary>
    ''' Releases the resources used by this instance, stopping any ongoing monitoring.
    ''' </summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        StopMonitoring()
        GC.SuppressFinalize(Me)
    End Sub
End Class
''' <summary>
''' Provides data for events related to changes in connectivity status.
''' </summary>
Public Class ConnectivityEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Gets a value indicating whether an internet connection is available.
    ''' </summary>
    Public ReadOnly Property InternetAvailable As Boolean
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ConnectivityEventArgs"/> class.
    ''' </summary>
    ''' <param name="InternetAvailable">Indicates whether the internet was available at the time of the event.</param>
    Public Sub New(InternetAvailable As Boolean)
        Me.InternetAvailable = InternetAvailable
    End Sub
End Class