Imports System.Net
Imports System.Net.Http
Imports System.Text.Json
Imports System.Threading
''' <summary>
''' Represents an error returned by Firebase or produced while communicating with a Firebase service.
''' </summary>
Public Class FirebaseException
    Inherits Exception
    ''' <summary>
    ''' Gets the Firebase service area in which the error occurred.
    ''' </summary>
    Public ReadOnly Property ServiceArea As FirebaseServiceArea
    ''' <summary>
    ''' Gets the HTTP status code returned by Firebase, or <see langword="Nothing"/> when no response was received.
    ''' </summary>
    Public ReadOnly Property StatusCode As HttpStatusCode?
    ''' <summary>
    ''' Gets the Firebase error code when one was supplied by the service.
    ''' </summary>
    Public ReadOnly Property ErrorCode As String
    Friend Sub New(Message As String, ServiceArea As FirebaseServiceArea, StatusCode As HttpStatusCode?, ErrorCode As String, Optional InnerException As Exception = Nothing)
        MyBase.New(Message, InnerException)
        Me.ServiceArea = ServiceArea
        Me.StatusCode = StatusCode
        Me.ErrorCode = ErrorCode
    End Sub
    Friend Shared Async Function FromResponseAsync(ServiceArea As FirebaseServiceArea, Response As HttpResponseMessage, CancellationToken As CancellationToken) As Task(Of FirebaseException)
        Dim ResponseBody As String = Await Response.Content.ReadAsStringAsync(CancellationToken)
        Dim ErrorCode As String = Nothing
        Dim ErrorMessage As String = Nothing
        If Not String.IsNullOrWhiteSpace(ResponseBody) Then
            Try
                Using Document As JsonDocument = JsonDocument.Parse(ResponseBody)
                    Dim ErrorElement As JsonElement = Document.RootElement
                    Dim NestedError As JsonElement
                    If ErrorElement.ValueKind = JsonValueKind.Object AndAlso ErrorElement.TryGetProperty("error", NestedError) Then ErrorElement = NestedError
                    If ErrorElement.ValueKind = JsonValueKind.Object Then
                        Dim StatusElement As JsonElement
                        Dim MessageElement As JsonElement
                        If ErrorElement.TryGetProperty("status", StatusElement) AndAlso StatusElement.ValueKind = JsonValueKind.String Then ErrorCode = StatusElement.GetString()
                        If ErrorElement.TryGetProperty("message", MessageElement) AndAlso MessageElement.ValueKind = JsonValueKind.String Then ErrorMessage = MessageElement.GetString()
                    End If
                End Using
            Catch ex As JsonException
                ErrorMessage = Nothing
            End Try
        End If
        If String.IsNullOrWhiteSpace(ErrorCode) AndAlso IsLikelyErrorCode(ErrorMessage) Then ErrorCode = ErrorMessage
        If String.IsNullOrWhiteSpace(ErrorMessage) Then ErrorMessage = Response.ReasonPhrase
        If String.IsNullOrWhiteSpace(ErrorMessage) Then ErrorMessage = "Firebase rejected the request."
        Dim Message As String = $"Firebase {ServiceArea} request failed ({CInt(Response.StatusCode)} {Response.StatusCode}): {ErrorMessage}"
        Return New FirebaseException(Message, ServiceArea, Response.StatusCode, ErrorCode)
    End Function
    Friend Shared Function CreateNetwork(ServiceArea As FirebaseServiceArea, InnerException As HttpRequestException) As FirebaseException
        Return New FirebaseException($"A network error occurred while communicating with Firebase {ServiceArea}.", ServiceArea, Nothing, "NETWORK_ERROR", InnerException)
    End Function
    Friend Shared Function CreateTimeout(ServiceArea As FirebaseServiceArea, InnerException As OperationCanceledException) As FirebaseException
        Return New FirebaseException($"The Firebase {ServiceArea} operation timed out.", ServiceArea, Nothing, "TIMEOUT", InnerException)
    End Function
    Friend Shared Function CreateInvalidResponse(ServiceArea As FirebaseServiceArea, Message As String, Optional InnerException As Exception = Nothing) As FirebaseException
        Return New FirebaseException(Message, ServiceArea, Nothing, "INVALID_RESPONSE", InnerException)
    End Function
    Friend Shared Function CreateAuthenticationRequired() As FirebaseException
        Return New FirebaseException("An authenticated Firebase session is required. Call LoginAsync before using this operation.", FirebaseServiceArea.Authentication, Nothing, "AUTHENTICATION_REQUIRED")
    End Function
    Private Shared Function IsLikelyErrorCode(Value As String) As Boolean
        If String.IsNullOrWhiteSpace(Value) Then Return False
        For Each Character As Char In Value
            If Not Char.IsUpper(Character) AndAlso Not Char.IsDigit(Character) AndAlso Character <> "_"c AndAlso Character <> "-"c AndAlso Character <> ":"c Then Return False
        Next Character
        Return True
    End Function
End Class
