Imports System.Net
Imports System.Net.Http

''' <summary>
''' Provides helper methods for checking internet connectivity
''' using HTTP-based reachability tests.
''' </summary>
Public Class InternetHelper

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
                Using Client As New WebClient()
                    Using Stream = Client.OpenRead(Address)
                        Return True
                    End Using
                End Using
            Catch ex As Exception
                ' continua tentando o próximo endereço
            End Try
        Next

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
        Using Client As New HttpClient()
            For Each Address In Addresses
                Try
                    Dim Response = Await Client.GetAsync(Address)
                    If Response.IsSuccessStatusCode Then
                        Return True
                    End If
                Catch ex As Exception
                    ' continua tentando o próximo endereço
                End Try
            Next
        End Using

        Return False
    End Function

End Class