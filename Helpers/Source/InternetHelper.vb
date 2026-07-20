Imports System.Net
Imports System.Net.Http

''' <summary>
''' Provides helper methods for checking internet connectivity
''' using HTTP-based reachability tests.
''' </summary>
Public Class InternetHelper

    Private Shared ReadOnly Client As New HttpClient With {
        .Timeout = TimeSpan.FromSeconds(3)
    }


    ''' <summary>
    ''' List of well-known public endpoints used to test connectivity.
    ''' </summary>
    Private Shared ReadOnly Addresses As String() = {
        "https://www.google.com/",
        "https://www.microsoft.com/",
        "https://www.facebook.com/",
        "https://aws.amazon.com/",
        "https://www.cloudflare.com/",
        "https://github.com/",
        "https://www.youtube.com/"
    }

    ''' <summary>
    ''' Determines whether an active internet connection is available
    ''' by attempting to access well-known public endpoints.
    ''' </summary>
    ''' <returns>
    ''' <c>True</c> if at least one endpoint is reachable; otherwise, <c>False</c>.
    ''' </returns>
    Public Shared Function IsInternetAvailable() As Boolean
        For Each Address In Addresses
            Try
                Using Request As New HttpRequestMessage(HttpMethod.Head, Address)
                    Using Response = Client.Send(Request)
                        If Response.IsSuccessStatusCode Then
                            Return True
                        End If
                    End Using
                End Using
            Catch ex As Exception
            End Try
        Next Address
        Return False
    End Function

    ''' <summary>
    ''' Asynchronously determines whether an active internet connection is available
    ''' by performing HTTP requests to well-known public endpoints.
    ''' </summary>
    ''' <returns>
    ''' A task that returns <c>True</c> if at least one endpoint responds successfully;
    ''' otherwise, <c>False</c>.
    ''' </returns>
    Public Shared Async Function IsInternetAvailableAsync() As Task(Of Boolean)
        For Each Address In Addresses
            Try
                Using Response = Await Client.GetAsync(Address, HttpCompletionOption.ResponseHeadersRead)
                    If Response.IsSuccessStatusCode Then
                        Return True
                    End If
                End Using
            Catch ex As Exception
            End Try
        Next Address
        Return False
    End Function
End Class