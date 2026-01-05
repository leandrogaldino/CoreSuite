Imports System.Net
Imports System.Net.Http

Public Class Connectivity
    Public Event ConnectivityChanged(sender As Object, e As ConnectivityEventArgs)
    Public Event Connected(sender As Object, e As EventArgs)
    Public Event Disconnected(sender As Object, e As EventArgs)
    Private ReadOnly _Addresses() As String = {
        "https://www.google.com/",
        "https://www.microsoft.com/",
        "https://www.facebook.com/",
        "https://aws.amazon.com/",
        "https://www.cloudflare.com/",
        "https://github.com/",
        "https://www.youtube.com/"
    }

    Private _lastStatus As Boolean
    Private _isMonitoring As Boolean

    Public Sub New()
        _isMonitoring = False
    End Sub

    Public Async Sub StartMonitoring()
        _lastStatus = Await IsAvailableAsync()
        _isMonitoring = True
        Await MonitorConnectivity()
    End Sub

    Public Sub StopMonitoring()
        _isMonitoring = False
    End Sub

    Private Async Function MonitorConnectivity() As Task
        While _isMonitoring
            Dim currentStatus As Boolean = Await IsAvailableAsync()
            CheckStatusChange(currentStatus)
            Await Task.Delay(5000)
        End While
    End Function

    Private Sub CheckStatusChange(currentStatus As Boolean)
        If currentStatus <> _lastStatus Then
            _lastStatus = currentStatus
            Dim args As New ConnectivityEventArgs(currentStatus)

            RaiseEvent ConnectivityChanged(Me, args)

            If currentStatus Then
                RaiseEvent Connected(Me, EventArgs.Empty)
            Else
                RaiseEvent Disconnected(Me, EventArgs.Empty)
            End If
        End If
    End Sub

    Public Async Function IsAvailableAsync() As Task(Of Boolean)
        Using client As New HttpClient()
            For Each address In _Addresses
                Try
                    Dim response = Await client.GetAsync(address)
                    If response.IsSuccessStatusCode Then
                        Return True
                    End If
                Catch ex As Exception
                    Console.WriteLine($"Erro ao acessar {address}: {ex.Message}")
                End Try
            Next
        End Using
        Return False
    End Function

    Public Function IsAvailable() As Boolean
        For Each Address In _Addresses
            Try
                Using Client As New WebClient()
                    Using Stream = Client.OpenRead(Address)
                        Return True
                    End Using
                End Using
            Catch ex As Exception
                Console.WriteLine($"Erro ao acessar {Address}: {ex.Message}")
                Exit For
            End Try
        Next
        Return False
    End Function
End Class

Public Class ConnectivityEventArgs
    Inherits EventArgs
    Public ReadOnly Property InternetAvailable As Boolean
    Public Sub New(internetAvailable As Boolean)
        Me.InternetAvailable = internetAvailable
    End Sub
End Class
