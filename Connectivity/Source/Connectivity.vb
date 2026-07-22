Imports System.Diagnostics
Imports System.Net
Imports System.Net.Http
Imports System.Threading

Public Class Connectivity
    Implements IDisposable
    Public Event ConnectivityChanged(sender As Object, e As ConnectivityEventArgs)
    Public Event Connected(sender As Object, e As EventArgs)
    Public Event Disconnected(sender As Object, e As EventArgs)
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
    Public Property MonitoringInterval As TimeSpan = TimeSpan.FromSeconds(1)
    Public ReadOnly Property IsMonitoring As Boolean
        Get
            Return _IsMonitoring
        End Get
    End Property
    Public ReadOnly Property IsConnected As Boolean
        Get
            Return If(_HasStatus, _LastStatus, False)
        End Get
    End Property
    Public Sub StartMonitoring()
        If _IsMonitoring Then Return
        _Cancellation?.Dispose()
        _Cancellation = New CancellationTokenSource()
        _IsMonitoring = True
        Dim Task As Task = MonitorConnectivity()
    End Sub
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
    Public Sub StopMonitoring()
        If Not _IsMonitoring Then Return
        _IsMonitoring = False
        Dim cancellation = _Cancellation
        _Cancellation = Nothing
        cancellation?.Cancel()
        cancellation?.Dispose()
    End Sub
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
    Public Function IsAvailable() As Boolean
        Return IsAvailableAsync().GetAwaiter().GetResult()
    End Function
    Public Sub Dispose() Implements IDisposable.Dispose
        StopMonitoring()
        GC.SuppressFinalize(Me)
    End Sub
End Class
Public Class ConnectivityEventArgs
    Inherits EventArgs
    Public ReadOnly Property InternetAvailable As Boolean
    Public Sub New(InternetAvailable As Boolean)
        Me.InternetAvailable = InternetAvailable
    End Sub
End Class