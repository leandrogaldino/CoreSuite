Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text.Json
Imports System.Threading
''' <summary>
''' Provides authenticated Cloud Storage for Firebase upload, download and deletion operations.
''' </summary>
Public NotInheritable Class FirebaseStorage
    Private Const BufferSize As Integer = 81920
    Private ReadOnly _Client As FirebaseClient
    ''' <summary>
    ''' Occurs when the progress of an upload changes.
    ''' </summary>
    Public Event UploadProgressChanged As EventHandler(Of FirebaseTransferProgressChangedEventArgs)
    ''' <summary>
    ''' Occurs when the progress of a download changes.
    ''' </summary>
    Public Event DownloadProgressChanged As EventHandler(Of FirebaseTransferProgressChangedEventArgs)
    Friend Sub New(Client As FirebaseClient)
        _Client = Client
    End Sub
    ''' <summary>
    ''' Uploads a local file to Cloud Storage for Firebase.
    ''' </summary>
    ''' <param name="LocalPath">The path of the local source file.</param>
    ''' <param name="RemotePath">The destination object path inside the configured bucket.</param>
    ''' <param name="ContentType">The media type to assign to the object, or <see langword="Nothing"/> for <c>application/octet-stream</c>.</param>
    ''' <param name="CancellationToken">A token that can cancel the upload.</param>
    ''' <returns>The download token returned in the uploaded object metadata, or an empty string when no token is returned.</returns>
    Public Async Function UploadFileAsync(LocalPath As String, RemotePath As String, Optional ContentType As String = Nothing, Optional CancellationToken As CancellationToken = Nothing) As Task(Of String)
        If String.IsNullOrWhiteSpace(LocalPath) Then Throw New ArgumentException("The local file path cannot be empty.", NameOf(LocalPath))
        Dim FullLocalPath As String = Path.GetFullPath(LocalPath)
        If Not File.Exists(FullLocalPath) Then Throw New FileNotFoundException("The upload source file was not found.", FullLocalPath)
        Dim NormalizedRemotePath As String = FirebasePathHelper.NormalizeStoragePath(RemotePath, NameOf(RemotePath))
        Dim EffectiveContentType As String = If(String.IsNullOrWhiteSpace(ContentType), "application/octet-stream", ContentType.Trim())
        Dim ParsedContentType As MediaTypeHeaderValue = Nothing
        If Not MediaTypeHeaderValue.TryParse(EffectiveContentType, ParsedContentType) Then Throw New ArgumentException("The content type is invalid.", NameOf(ContentType))
        Dim EncodedPath As String = Uri.EscapeDataString(NormalizedRemotePath)
        Dim Url As String = $"https://firebasestorage.googleapis.com/v0/b/{Uri.EscapeDataString(_Client.Options.StorageBucket)}/o?uploadType=media&name={EncodedPath}"
        Using OperationSource As CancellationTokenSource = _Client.CreateOperationCancellationSource(True, CancellationToken)
            Try
                Await _Client.Auth.EnsureValidTokenAsync(OperationSource.Token).ConfigureAwait(False)
                Dim SourceStream As New FileStream(FullLocalPath, FileMode.Open, FileAccess.Read, FileShare.Read, BufferSize, FileOptions.Asynchronous Or FileOptions.SequentialScan)
                Using Content As New ProgressableStreamContent(SourceStream, BufferSize, AddressOf ReportUploadProgress)
                    Content.Headers.ContentType = ParsedContentType
                    RaiseEvent UploadProgressChanged(Me, New FirebaseTransferProgressChangedEventArgs(0, SourceStream.Length, False))
                    Using Request As HttpRequestMessage = _Client.CreateRequest(HttpMethod.Post, Url)
                        Request.Content = Content
                        Using Response As HttpResponseMessage = Await _Client.SendAsync(Request, FirebaseServiceArea.Storage, HttpCompletionOption.ResponseContentRead, OperationSource.Token).ConfigureAwait(False)
                            Await FirebaseClient.EnsureSuccessAsync(Response, FirebaseServiceArea.Storage, OperationSource.Token).ConfigureAwait(False)
                            Dim ResponseBody As String = Await Response.Content.ReadAsStringAsync(OperationSource.Token).ConfigureAwait(False)
                            Dim UploadResponse As FirebaseStorageUploadResponse
                            Try
                                UploadResponse = JsonSerializer.Deserialize(Of FirebaseStorageUploadResponse)(ResponseBody)
                            Catch ex As JsonException
                                Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Storage, "Cloud Storage returned malformed upload metadata.", ex)
                            End Try
                            RaiseEvent UploadProgressChanged(Me, New FirebaseTransferProgressChangedEventArgs(SourceStream.Length, SourceStream.Length, True))
                            Return If(UploadResponse?.DownloadTokens, String.Empty)
                        End Using
                    End Using
                End Using
            Catch ex As OperationCanceledException When Not CancellationToken.IsCancellationRequested
                Throw FirebaseException.CreateTimeout(FirebaseServiceArea.Storage, ex)
            End Try
        End Using
    End Function
    ''' <summary>
    ''' Downloads a Cloud Storage object to a local file using an atomic destination replacement.
    ''' </summary>
    ''' <param name="RemotePath">The object path inside the configured bucket.</param>
    ''' <param name="LocalPath">The local destination path.</param>
    ''' <param name="CancellationToken">A token that can cancel the download.</param>
    ''' <remarks>The destination file is replaced only after the complete response has been written successfully.</remarks>
    Public Async Function DownloadFileAsync(RemotePath As String, LocalPath As String, Optional CancellationToken As CancellationToken = Nothing) As Task
        Dim NormalizedRemotePath As String = FirebasePathHelper.NormalizeStoragePath(RemotePath, NameOf(RemotePath))
        If String.IsNullOrWhiteSpace(LocalPath) Then Throw New ArgumentException("The local file path cannot be empty.", NameOf(LocalPath))
        Dim FullLocalPath As String = Path.GetFullPath(LocalPath)
        Dim DestinationDirectory As String = Path.GetDirectoryName(FullLocalPath)
        If String.IsNullOrWhiteSpace(DestinationDirectory) Then Throw New ArgumentException("The local destination path is invalid.", NameOf(LocalPath))
        Directory.CreateDirectory(DestinationDirectory)
        Dim TemporaryPath As String = Path.Combine(DestinationDirectory, $".{Path.GetFileName(FullLocalPath)}.{Guid.NewGuid():N}.tmp")
        Dim EncodedPath As String = Uri.EscapeDataString(NormalizedRemotePath)
        Dim Url As String = $"https://firebasestorage.googleapis.com/v0/b/{Uri.EscapeDataString(_Client.Options.StorageBucket)}/o/{EncodedPath}?alt=media"
        Using OperationSource As CancellationTokenSource = _Client.CreateOperationCancellationSource(True, CancellationToken)
            Try
                Await _Client.Auth.EnsureValidTokenAsync(OperationSource.Token).ConfigureAwait(False)
                Using Request As HttpRequestMessage = _Client.CreateRequest(HttpMethod.Get, Url)
                    Using Response As HttpResponseMessage = Await _Client.SendAsync(Request, FirebaseServiceArea.Storage, HttpCompletionOption.ResponseHeadersRead, OperationSource.Token).ConfigureAwait(False)
                        Await FirebaseClient.EnsureSuccessAsync(Response, FirebaseServiceArea.Storage, OperationSource.Token).ConfigureAwait(False)
                        Dim TotalBytes As Long? = Response.Content.Headers.ContentLength
                        RaiseEvent DownloadProgressChanged(Me, New FirebaseTransferProgressChangedEventArgs(0, TotalBytes, False))
                        Dim DownloadedBytes As Long = 0
                        Using ResponseStream As Stream = Await Response.Content.ReadAsStreamAsync(OperationSource.Token).ConfigureAwait(False)
                            Using DestinationStream As New FileStream(TemporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None, BufferSize, FileOptions.Asynchronous Or FileOptions.SequentialScan)
                                Dim Buffer(BufferSize - 1) As Byte
                                Do
                                    Dim BytesRead As Integer = Await ResponseStream.ReadAsync(Buffer.AsMemory(0, Buffer.Length), OperationSource.Token).ConfigureAwait(False)
                                    If BytesRead = 0 Then Exit Do
                                    Await DestinationStream.WriteAsync(Buffer.AsMemory(0, BytesRead), OperationSource.Token).ConfigureAwait(False)
                                    DownloadedBytes += BytesRead
                                    RaiseEvent DownloadProgressChanged(Me, New FirebaseTransferProgressChangedEventArgs(DownloadedBytes, TotalBytes, False))
                                Loop
                                Await DestinationStream.FlushAsync(OperationSource.Token).ConfigureAwait(False)
                                If TotalBytes.HasValue AndAlso DownloadedBytes <> TotalBytes.Value Then Throw FirebaseException.CreateInvalidResponse(FirebaseServiceArea.Storage, "Cloud Storage ended the download before the declared content length was received.")
                            End Using
                        End Using
                        File.Move(TemporaryPath, FullLocalPath, True)
                        RaiseEvent DownloadProgressChanged(Me, New FirebaseTransferProgressChangedEventArgs(DownloadedBytes, If(TotalBytes, DownloadedBytes), True))
                    End Using
                End Using
            Catch ex As OperationCanceledException When Not CancellationToken.IsCancellationRequested
                Throw FirebaseException.CreateTimeout(FirebaseServiceArea.Storage, ex)
            Finally
                If File.Exists(TemporaryPath) Then File.Delete(TemporaryPath)
            End Try
        End Using
    End Function
    ''' <summary>
    ''' Deletes an object from Cloud Storage for Firebase.
    ''' </summary>
    ''' <param name="RemotePath">The object path inside the configured bucket.</param>
    ''' <param name="CancellationToken">A token that can cancel the operation.</param>
    ''' <returns><see langword="True"/> when the object was deleted; <see langword="False"/> when it did not exist.</returns>
    Public Async Function DeleteFileAsync(RemotePath As String, Optional CancellationToken As CancellationToken = Nothing) As Task(Of Boolean)
        Dim NormalizedRemotePath As String = FirebasePathHelper.NormalizeStoragePath(RemotePath, NameOf(RemotePath))
        Dim EncodedPath As String = Uri.EscapeDataString(NormalizedRemotePath)
        Dim Url As String = $"https://firebasestorage.googleapis.com/v0/b/{Uri.EscapeDataString(_Client.Options.StorageBucket)}/o/{EncodedPath}"
        Using OperationSource As CancellationTokenSource = _Client.CreateOperationCancellationSource(True, CancellationToken)
            Try
                Await _Client.Auth.EnsureValidTokenAsync(OperationSource.Token).ConfigureAwait(False)
                Using Request As HttpRequestMessage = _Client.CreateRequest(HttpMethod.Delete, Url)
                    Using Response As HttpResponseMessage = Await _Client.SendAsync(Request, FirebaseServiceArea.Storage, HttpCompletionOption.ResponseContentRead, OperationSource.Token).ConfigureAwait(False)
                        If Response.StatusCode = HttpStatusCode.NotFound Then Return False
                        Await FirebaseClient.EnsureSuccessAsync(Response, FirebaseServiceArea.Storage, OperationSource.Token).ConfigureAwait(False)
                        Return True
                    End Using
                End Using
            Catch ex As OperationCanceledException When Not CancellationToken.IsCancellationRequested
                Throw FirebaseException.CreateTimeout(FirebaseServiceArea.Storage, ex)
            End Try
        End Using
    End Function
    Private Sub ReportUploadProgress(BytesTransferred As Long, TotalBytes As Long)
        RaiseEvent UploadProgressChanged(Me, New FirebaseTransferProgressChangedEventArgs(BytesTransferred, TotalBytes, False))
    End Sub
End Class
