Imports System.Net.Http
Imports System.Net.Http.Headers
''' <summary>
''' Represents the internal client that manages configuration, authentication state,
''' and HTTP communication for interacting with Firebase services.
''' </summary>
Public Class FirebaseClient
    ''' <summary>
    ''' Gets or sets the API key used to authenticate requests to Firebase services.
    ''' </summary>
    Friend Property ApiKey As String
    ''' <summary>
    ''' Gets or sets the Firebase project identifier.
    ''' </summary>
    Friend Property ProjectID As String
    ''' <summary>
    ''' Gets or sets the name of the Firebase Storage bucket associated with the project.
    ''' </summary>
    Friend Property StorageBucket As String
    ''' <summary>
    ''' Gets or sets the authentication manager responsible for handling login sessions.
    ''' </summary>
    Friend Property Auth As FirebaseAuth
    ''' <summary>
    ''' Gets or sets the current authentication token used to authorize requests.
    ''' </summary>
    Friend Property Token As String
    ''' <summary>
    ''' Gets the underlying <see cref="HttpClient"/> instance used to perform HTTP requests.
    ''' </summary>
    Friend ReadOnly Http As HttpClient
    ''' <summary>
    ''' Initializes a new instance of the <see cref="FirebaseClient"/> class with the specified
    ''' API key, project ID, and storage bucket.
    ''' </summary>
    ''' <param name="Api">The Firebase API key.</param>
    ''' <param name="ProjectID">The Firebase project identifier.</param>
    ''' <param name="Bucket">The Firebase Storage bucket name.</param>
    Friend Sub New(Api As String, ProjectID As String, Bucket As String)
        ApiKey = Api
        Me.ProjectID = ProjectID
        StorageBucket = Bucket
        Http = New HttpClient With {
            .Timeout = TimeSpan.FromSeconds(30)
        }
    End Sub
    ''' <summary>
    ''' Creates an HTTP request message for the specified method and URL, automatically
    ''' attaching the authentication token as a Bearer header when available.
    ''' </summary>
    ''' <param name="Method">The HTTP method to use for the request.</param>
    ''' <param name="Url">The target URL of the request.</param>
    ''' <returns>A configured <see cref="HttpRequestMessage"/> instance.</returns>
    Friend Function CreateRequest(Method As HttpMethod, Url As String) As HttpRequestMessage
        Dim Request As New HttpRequestMessage(Method, Url)
        If Not String.IsNullOrEmpty(Token) Then
            Request.Headers.Authorization = New AuthenticationHeaderValue("Bearer", Token)
        End If
        Return Request
    End Function
End Class