Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports CoreSuite.Helpers
''' <summary>
''' Provides access to Firebase Storage for file upload, download, and deletion.
''' </summary>
''' <remarks>
''' Upload and download progress is reported through the
''' <see cref="UploadProgressChanged"/> and <see cref="DownloadProgressChanged"/> events.
''' 
''' Consumers should subscribe to these events before calling
''' <see cref="UploadFile"/> or <see cref="DownloadFile"/> to receive
''' progress notifications ranging from 0 to 100.
''' </remarks>
Public Class FirebaseStorage
    Private ReadOnly _Client As FirebaseClient
    ''' <summary>
    ''' Occurs when upload progress changes.
    ''' </summary>
    Public Event UploadProgressChanged(percentage As Double)
    ''' <summary>
    ''' Occurs when download progress changes.
    ''' </summary>
    Public Event DownloadProgressChanged(percentage As Double)
    Friend Sub New(Client As FirebaseClient)
        _Client = Client
    End Sub
    ''' <summary>
    ''' Uploads a local file to Firebase Storage.
    ''' </summary>
    ''' <param name="LocalPath">Local file path.</param>
    ''' <param name="RemotePath">Remote storage path.</param>
    ''' <returns>A download token for the uploaded file.</returns>
    Public Async Function UploadFile(LocalPath As String, RemotePath As String) As Task(Of String)
        Try
            Dim Bucket = _Client.StorageBucket.Replace("gs://", "").TrimEnd("/"c)
            Dim EncodedPath = Uri.EscapeDataString(RemotePath.TrimStart("/"c))
            Dim Url = $"https://firebasestorage.googleapis.com/v0/b/{Bucket}/o" & $"?uploadType=media&name={EncodedPath}"
            _Client.Http.DefaultRequestHeaders.Clear()
            _Client.Http.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", _Client.Token)
            Dim FileStream = File.OpenRead(LocalPath)
            Dim Content = New ProgressableStreamContent(FileStream, 81920, Sub(Sent, Total)
                                                                               Dim Percent = CInt((Sent * 100L) / Total)
                                                                               RaiseEvent UploadProgressChanged(Percent)
                                                                           End Sub)
            Content.Headers.ContentType = New MediaTypeHeaderValue("application/octet-stream")
            RaiseEvent UploadProgressChanged(0)
            Dim Response = Await _Client.Http.PostAsync(Url, Content)
            If Response.IsSuccessStatusCode Then
                RaiseEvent UploadProgressChanged(100)
                Dim Json = Await Response.Content.ReadAsStringAsync()
                Return TextHelper.ExtractJsonValue(Json, "downloadTokens")
            Else
                Dim [Error] = Await Response.Content.ReadAsStringAsync()
                Throw New Exception($"Erro no Storage: {[Error]}")
            End If
        Catch ex As Exception
            Throw New Exception($"Falha no Upload: {ex.Message}")
        End Try
    End Function

    ''' <summary>
    ''' Downloads a file from Firebase Storage.
    ''' </summary>
    ''' <param name="RemotePath">Remote storage path.</param>
    ''' <param name="LocalPath">Local file destination.</param>
    Public Async Function DownloadFile(RemotePath As String, LocalPath As String) As Task
        Try
            Dim Bucket = _Client.StorageBucket.Replace("gs://", "").TrimEnd("/"c)
            Dim EncodedPath = Uri.EscapeDataString(RemotePath.TrimStart("/"c))
            Dim Url = $"https://firebasestorage.googleapis.com/v0/b/{Bucket}/o/{EncodedPath}?alt=media"
            _Client.Http.DefaultRequestHeaders.Clear()
            _Client.Http.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", _Client.Token)
            Dim Response = Await _Client.Http.GetAsync(Url, HttpCompletionOption.ResponseHeadersRead)
            If Not Response.IsSuccessStatusCode Then
                Dim [Error] = Await Response.Content.ReadAsStringAsync()
                Throw New Exception($"Erro no download: {[Error]}")
            End If
            Dim TotalBytes As Long = If(Response.Content.Headers.ContentLength, -1)
            Dim Downloaded As Long = 0
            Dim Buffer(81920 - 1) As Byte
            Using httpStream = Await Response.Content.ReadAsStreamAsync(), FileStream = New FileStream(LocalPath, FileMode.Create, FileAccess.Write, FileShare.None, Buffer.Length, True)
                RaiseEvent DownloadProgressChanged(0)
                Dim BytesRead As Integer
                Do
                    BytesRead = Await httpStream.ReadAsync(Buffer, 0, Buffer.Length)
                    If BytesRead = 0 Then Exit Do
                    Await FileStream.WriteAsync(Buffer, 0, BytesRead)
                    Downloaded += BytesRead
                    If TotalBytes > 0 Then
                        Dim Percent = CInt((Downloaded * 100L) / TotalBytes)
                        RaiseEvent DownloadProgressChanged(Percent)
                    End If
                Loop
            End Using
            RaiseEvent DownloadProgressChanged(100)
        Catch ex As Exception
            Throw New Exception($"Falha no Download: {ex.Message}")
        End Try
    End Function
    ''' <summary>
    ''' Deletes a file from Firebase Storage.
    ''' </summary>
    ''' <param name="RemotePath">Remote storage path.</param>
    ''' <returns>True if deletion was successful.</returns>
    Public Async Function DeleteFileAsync(RemotePath As String) As Task(Of Boolean)
        Try
            Dim Bucket = _Client.StorageBucket.Replace("gs://", "").TrimEnd("/"c)
            Dim EncodedPath = Uri.EscapeDataString(RemotePath.TrimStart("/c"))
            Dim Url = $"https://firebasestorage.googleapis.com/v0/b/{Bucket}/o/{EncodedPath}"
            _Client.Http.DefaultRequestHeaders.Clear()
            _Client.Http.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", _Client.Token)
            Dim Response = Await _Client.Http.DeleteAsync(Url)
            If Response.IsSuccessStatusCode Then
                Return True
            ElseIf Response.StatusCode = HttpStatusCode.NotFound Then
                Return False
            Else
                Dim [Error] = Await Response.Content.ReadAsStringAsync()
                Throw New Exception("Erro ao deletar arquivo no Storage: " & [Error])
            End If
        Catch ex As HttpRequestException
            Throw New Exception("Erro de rede ao deletar arquivo do Storage.", ex)
        Catch ex As TaskCanceledException
            Throw New Exception("Timeout ao tentar deletar arquivo do Storage.")
        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Class ProgressableStreamContent
        Inherits HttpContent
        Private ReadOnly _Stream As Stream
        Private ReadOnly _BufferSize As Integer
        Private ReadOnly _Progress As Action(Of Long, Long)
        Public Sub New(Stream As Stream, BufferSize As Integer, Progress As Action(Of Long, Long))
            _Stream = Stream
            _BufferSize = BufferSize
            _Progress = Progress
            Headers.ContentLength = Stream.Length
        End Sub
        Protected Overrides Async Function SerializeToStreamAsync(TargetStream As Stream, Context As TransportContext) As Task
            Dim Buffer(_BufferSize - 1) As Byte
            Dim Uploaded As Long = 0
            Dim Total As Long = _Stream.Length
            Using _Stream
                Dim BytesRead As Integer
                Do
                    BytesRead = Await _Stream.ReadAsync(Buffer, 0, Buffer.Length)
                    If BytesRead = 0 Then Exit Do
                    Await TargetStream.WriteAsync(Buffer, 0, BytesRead)
                    Uploaded += BytesRead
                    _Progress?.Invoke(Uploaded, Total)
                Loop
            End Using
        End Function
        Protected Overrides Function TryComputeLength(ByRef Length As Long) As Boolean
            Length = _Stream.Length
            Return True
        End Function
    End Class
End Class