Imports System.Net.Http
Imports System.Net.Http.Headers
Public Class FirebaseClient
    Friend Property ApiKey As String
    Friend Property ProjectID As String
    Friend Property StorageBucket As String
    Friend Property Token As String
    Friend ReadOnly Http As HttpClient
    Friend Sub New(Api As String, ProjectID As String, Bucket As String)
        ApiKey = Api
        Me.ProjectID = ProjectID
        StorageBucket = Bucket
        Http = New HttpClient With {
            .Timeout = TimeSpan.FromSeconds(30)
        }
    End Sub
    Friend Function CreateRequest(Method As HttpMethod, Url As String) As HttpRequestMessage
        Dim Request As New HttpRequestMessage(Method, Url)
        If Not String.IsNullOrEmpty(Token) Then
            Request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", Token)
        End If
        Return Request
    End Function
End Class