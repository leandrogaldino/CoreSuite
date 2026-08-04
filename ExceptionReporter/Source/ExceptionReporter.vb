Imports System.IO
Imports System.Text
Imports System.Text.Json
Imports System.Threading
Imports MailKit.Net.Smtp
Imports MailKit.Security
Imports MimeKit

''' <summary>
''' Captures, serializes, saves and sends structured exception reports without depending on a user interface.
''' </summary>
Public Class ExceptionReporter
    Private ReadOnly _JsonOptions As JsonSerializerOptions
    ''' <summary>
    ''' Initializes a new reporter with the default indented JSON configuration.
    ''' </summary>
    Public Sub New()
        Me.New(New JsonSerializerOptions With {.WriteIndented = True})
    End Sub
    ''' <summary>
    ''' Initializes a new reporter with custom JSON serialization options.
    ''' </summary>
    ''' <param name="JsonOptions">The JSON options used to serialize reports.</param>
    ''' <exception cref="ArgumentNullException">Thrown when <paramref name="JsonOptions"/> is <see langword="Nothing"/>.</exception>
    Public Sub New(JsonOptions As JsonSerializerOptions)
        ArgumentNullException.ThrowIfNull(JsonOptions)
        _JsonOptions = New JsonSerializerOptions(JsonOptions)
    End Sub
    ''' <summary>
    ''' Creates a structured report from an exception and optional application context.
    ''' </summary>
    ''' <param name="Exception">The exception to capture.</param>
    ''' <param name="Title">An optional report title.</param>
    ''' <param name="Message">An optional user-friendly error message.</param>
    ''' <param name="AdditionalInformations">Optional application-specific diagnostic values.</param>
    ''' <param name="UserSteps">Optional steps or description supplied by the user.</param>
    ''' <returns>A populated exception report.</returns>
    ''' <exception cref="ArgumentNullException">Thrown when <paramref name="Exception"/> is <see langword="Nothing"/>.</exception>
    Public Shared Function Capture(Exception As Exception, Optional Title As String = Nothing, Optional Message As String = Nothing, Optional AdditionalInformations As IDictionary(Of String, Object) = Nothing, Optional UserSteps As String = Nothing) As ExceptionReport
        ArgumentNullException.ThrowIfNull(Exception)
        Dim Report As New ExceptionReport With {
            .Title = Title,
            .Message = Message,
            .ExceptionMessage = Exception.Message,
            .ExceptionInnerMessage = Exception.InnerException?.Message,
            .StackTrace = Exception.StackTrace,
            .UserSteps = UserSteps,
            .ExceptionDetails = CreateExceptionDetails(Exception),
            .ExceptionDate = Date.Now
        }
        If AdditionalInformations IsNot Nothing Then
            For Each AdditionalInformation As KeyValuePair(Of String, Object) In AdditionalInformations
                Report.AdditionalInformations(AdditionalInformation.Key) = AdditionalInformation.Value
            Next AdditionalInformation
        End If
        Return Report
    End Function
    ''' <summary>
    ''' Serializes an exception report to JSON.
    ''' </summary>
    ''' <param name="Report">The report to serialize.</param>
    ''' <returns>The JSON representation of the report.</returns>
    ''' <exception cref="ArgumentNullException">Thrown when <paramref name="Report"/> is <see langword="Nothing"/>.</exception>
    Public Function Serialize(Report As ExceptionReport) As String
        ArgumentNullException.ThrowIfNull(Report)
        Return JsonSerializer.Serialize(Report, _JsonOptions)
    End Function
    ''' <summary>
    ''' Saves a report to a UTF-8 JSON file using an atomic replacement operation.
    ''' </summary>
    ''' <param name="Report">The report to save.</param>
    ''' <param name="FilePath">The destination file path.</param>
    Public Sub Save(Report As ExceptionReport, FilePath As String)
        ValidateFilePath(FilePath)
        Dim Content As String = Serialize(Report)
        Dim FullPath As String = Path.GetFullPath(FilePath)
        Dim TemporaryPath As String = CreateTemporaryPath(FullPath)
        Try
            File.WriteAllText(TemporaryPath, Content, New UTF8Encoding(False))
            File.Move(TemporaryPath, FullPath, True)
        Finally
            DeleteTemporaryFile(TemporaryPath)
        End Try
    End Sub
    ''' <summary>
    ''' Asynchronously saves a report to a UTF-8 JSON file using an atomic replacement operation.
    ''' </summary>
    ''' <param name="Report">The report to save.</param>
    ''' <param name="FilePath">The destination file path.</param>
    ''' <param name="CancellationToken">A token used to cancel the operation.</param>
    ''' <returns>A task representing the save operation.</returns>
    Public Async Function SaveAsync(Report As ExceptionReport, FilePath As String, Optional CancellationToken As CancellationToken = Nothing) As Task
        ValidateFilePath(FilePath)
        Dim Content As String = Serialize(Report)
        Dim FullPath As String = Path.GetFullPath(FilePath)
        Dim TemporaryPath As String = CreateTemporaryPath(FullPath)
        Try
            Await File.WriteAllTextAsync(TemporaryPath, Content, New UTF8Encoding(False), CancellationToken).ConfigureAwait(False)
            File.Move(TemporaryPath, FullPath, True)
        Finally
            DeleteTemporaryFile(TemporaryPath)
        End Try
    End Function
    ''' <summary>
    ''' Asynchronously sends a structured exception report by email.
    ''' </summary>
    ''' <param name="Report">The report to send.</param>
    ''' <param name="EmailOptions">The SMTP and message configuration.</param>
    ''' <param name="CancellationToken">A token used to cancel the operation.</param>
    ''' <returns><see langword="True"/> when the report was sent; otherwise, <see langword="False"/> when connectivity was unavailable.</returns>
    Public Function SendEmailAsync(Report As ExceptionReport, EmailOptions As ExceptionEmailOptions, Optional CancellationToken As CancellationToken = Nothing) As Task(Of Boolean)
        Return SendEmailAsync(Serialize(Report), EmailOptions, CancellationToken)
    End Function
    ''' <summary>
    ''' Asynchronously sends preformatted exception-report content by email.
    ''' </summary>
    ''' <param name="Content">The report content to send.</param>
    ''' <param name="EmailOptions">The SMTP and message configuration.</param>
    ''' <param name="CancellationToken">A token used to cancel the operation.</param>
    ''' <returns><see langword="True"/> when the report was sent; otherwise, <see langword="False"/> when connectivity was unavailable.</returns>
    Public Shared Async Function SendEmailAsync(Content As String, EmailOptions As ExceptionEmailOptions, Optional CancellationToken As CancellationToken = Nothing) As Task(Of Boolean)
        ArgumentNullException.ThrowIfNull(Content)
        ValidateEmailOptions(EmailOptions)
        If EmailOptions.CheckConnectivity Then
            Using ConnectivityService As New Connectivity()
                If Not Await ConnectivityService.IsAvailableAsync().ConfigureAwait(False) Then Return False
            End Using
        End If
        Dim Message As New MimeMessage()
        Message.From.Add(New MailboxAddress(EmailOptions.FromName, EmailOptions.FromEmail))
        Message.To.Add(New MailboxAddress(EmailOptions.ToName, EmailOptions.ToEmail))
        Message.Subject = EmailOptions.Subject
        Message.Body = New TextPart("plain") With {.Text = Content}
        Using Client As New SmtpClient()
            Await Client.ConnectAsync(EmailOptions.Host, EmailOptions.Port, ConvertSecureSocket(EmailOptions.SecureSocket), CancellationToken).ConfigureAwait(False)
            If EmailOptions.UseAuthentication Then
                Dim UserName As String = If(String.IsNullOrWhiteSpace(EmailOptions.UserName), EmailOptions.FromEmail, EmailOptions.UserName)
                Await Client.AuthenticateAsync(UserName, EmailOptions.Password, CancellationToken).ConfigureAwait(False)
            End If
            Await Client.SendAsync(Message, CancellationToken).ConfigureAwait(False)
            Await Client.DisconnectAsync(True, CancellationToken).ConfigureAwait(False)
        End Using
        Return True
    End Function
    Private Shared Function CreateExceptionDetails(Exception As Exception) As ExceptionReportException
        Dim Details As New ExceptionReportException With {
            .Type = Exception.GetType().FullName,
            .Message = Exception.Message,
            .Source = Exception.Source,
            .StackTrace = Exception.StackTrace,
            .HResult = Exception.HResult,
            .HelpLink = Exception.HelpLink
        }
        For Each Key As Object In Exception.Data.Keys
            Dim KeyText As String = Convert.ToString(Key, Globalization.CultureInfo.InvariantCulture)
            Dim ValueText As String = Convert.ToString(Exception.Data(Key), Globalization.CultureInfo.InvariantCulture)
            Details.Data(KeyText) = ValueText
        Next Key
        If Exception.InnerException IsNot Nothing Then Details.InnerException = CreateExceptionDetails(Exception.InnerException)
        Return Details
    End Function
    Private Shared Function CreateTemporaryPath(FullPath As String) As String
        Dim DirectoryPath As String = Path.GetDirectoryName(FullPath)
        If String.IsNullOrEmpty(DirectoryPath) Then DirectoryPath = Directory.GetCurrentDirectory()
        Directory.CreateDirectory(DirectoryPath)
        Return Path.Combine(DirectoryPath, $".{Path.GetFileName(FullPath)}.{Guid.NewGuid():N}.tmp")
    End Function
    Private Shared Sub DeleteTemporaryFile(TemporaryPath As String)
        If File.Exists(TemporaryPath) Then File.Delete(TemporaryPath)
    End Sub
    Private Shared Sub ValidateFilePath(FilePath As String)
        If String.IsNullOrWhiteSpace(FilePath) Then Throw New ArgumentException("A destination file path is required.", NameOf(FilePath))
    End Sub
    Private Shared Sub ValidateEmailOptions(EmailOptions As ExceptionEmailOptions)
        ArgumentNullException.ThrowIfNull(EmailOptions)
        If String.IsNullOrWhiteSpace(EmailOptions.FromEmail) Then Throw New ArgumentException("A sender email address is required.", NameOf(EmailOptions))
        If String.IsNullOrWhiteSpace(EmailOptions.ToEmail) Then Throw New ArgumentException("A recipient email address is required.", NameOf(EmailOptions))
        If String.IsNullOrWhiteSpace(EmailOptions.Host) Then Throw New ArgumentException("An SMTP host is required.", NameOf(EmailOptions))
        If EmailOptions.Port < 1 OrElse EmailOptions.Port > 65535 Then Throw New ArgumentOutOfRangeException(NameOf(EmailOptions), "The SMTP port must be between 1 and 65535.")
        If EmailOptions.UseAuthentication AndAlso EmailOptions.Password Is Nothing Then Throw New ArgumentException("An SMTP password is required when authentication is enabled.", NameOf(EmailOptions))
    End Sub
    Private Shared Function ConvertSecureSocket(SecureSocket As ExceptionReporterSecureSocket) As SecureSocketOptions
        Select Case SecureSocket
            Case ExceptionReporterSecureSocket.Auto
                Return SecureSocketOptions.Auto
            Case ExceptionReporterSecureSocket.StartTls
                Return SecureSocketOptions.StartTls
            Case ExceptionReporterSecureSocket.StartTlsWhenAvailable
                Return SecureSocketOptions.StartTlsWhenAvailable
            Case ExceptionReporterSecureSocket.SslOnConnect
                Return SecureSocketOptions.SslOnConnect
            Case ExceptionReporterSecureSocket.None
                Return SecureSocketOptions.None
            Case Else
                Throw New ArgumentOutOfRangeException(NameOf(SecureSocket))
        End Select
    End Function
End Class
