Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Threading
Friend NotInheritable Class ProgressableStreamContent
    Inherits HttpContent
    Private ReadOnly _Source As Stream
    Private ReadOnly _BufferSize As Integer
    Private ReadOnly _CancellationToken As CancellationToken
    Private ReadOnly _Progress As Action(Of Long, Long)
    Public Sub New(Source As Stream, BufferSize As Integer, Progress As Action(Of Long, Long), Optional CancellationToken As CancellationToken = Nothing)
        If Source Is Nothing Then Throw New ArgumentNullException("Source")
        If Not Source.CanRead Then Throw New ArgumentException("The source stream must be readable.", NameOf(Source))
        If BufferSize <= 0 Then Throw New ArgumentOutOfRangeException(NameOf(BufferSize), BufferSize, "The value must be positive.")
        _Source = Source
        _CancellationToken = CancellationToken
        _BufferSize = BufferSize
        _Progress = Progress
        Headers.ContentLength = Source.Length
    End Sub
    Protected Overrides Function SerializeToStreamAsync(TargetStream As Stream, Context As TransportContext) As Task
        Return SerializeToStreamCoreAsync(TargetStream, _CancellationToken)
    End Function
    Private Async Function SerializeToStreamCoreAsync(TargetStream As Stream, CancellationToken As CancellationToken) As Task
        Dim Buffer(_BufferSize - 1) As Byte
        Dim UploadedBytes As Long = 0
        Dim TotalBytes As Long = _Source.Length
        Do
            Dim BytesRead As Integer = Await _Source.ReadAsync(Buffer, 0, Buffer.Length, CancellationToken).ConfigureAwait(False)
            If BytesRead = 0 Then Exit Do
            Await TargetStream.WriteAsync(Buffer, 0, BytesRead, CancellationToken).ConfigureAwait(False)
            UploadedBytes += BytesRead
            If _Progress IsNot Nothing Then _Progress(UploadedBytes, TotalBytes)
        Loop
    End Function
    Protected Overrides Function TryComputeLength(ByRef Length As Long) As Boolean
        Length = _Source.Length
        Return True
    End Function
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then _Source.Dispose()
        MyBase.Dispose(Disposing)
    End Sub
End Class
