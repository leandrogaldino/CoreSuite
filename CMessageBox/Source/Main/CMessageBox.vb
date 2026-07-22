Imports System.Text.Json
Imports MimeKit
Imports MailKit.Net.Smtp

Public Class CMessageBox
    ''' <summary>
    ''' Gets or sets the configuration options used by the CMessageBox.
    ''' </summary>
    ''' <remarks>
    ''' These options control the behavior, appearance and additional features of the
    ''' message box, including exception handling and email notifications.
    ''' </remarks>
    Public Shared Property Options As CMessageBoxOptions = New CMessageBoxOptions()
    ''' <summary>
    ''' Provides the JSON serialization settings used to format exception information.
    ''' </summary>
    ''' <remarks>
    ''' The serializer is configured to generate indented JSON output to improve
    ''' readability when displaying or sending exception details.
    ''' </remarks>
    Private Shared ReadOnly _JsonSerializer As New JsonSerializerOptions With {.WriteIndented = True}
    ''' <summary>
    ''' Displays a message box with the specified message using the default information type.
    ''' </summary>
    ''' <param name="Message">
    ''' The text to display in the message box body.
    ''' </param>
    ''' <returns>
    ''' Returns the result of the user's interaction with the message box.
    ''' </returns>
    Public Shared Function Show(Message As String) As DialogResult
        Return Show(Message, Nothing, CMessageBoxType.Information)
    End Function
    ''' <summary>
    ''' Displays a message box with the specified message and custom title using the default information type.
    ''' </summary>
    ''' <param name="Message">
    ''' The text to display in the message box body.
    ''' </param>
    ''' <param name="Title">
    ''' The text displayed in the message box title bar.
    ''' </param>
    ''' <returns>
    ''' Returns the result of the user's interaction with the message box.
    ''' </returns>
    Public Shared Function Show(Message As String, Title As String) As DialogResult
        Return Show(Message, Title, CMessageBoxType.Information)
    End Function
    ''' <summary>
    ''' Displays a message box with the specified message type.
    ''' </summary>
    ''' <remarks>
    ''' The message type determines the visual appearance, icon and purpose of the message.
    ''' </remarks>
    ''' <param name="Message">
    ''' The text to display in the message box body.
    ''' </param>
    ''' <param name="MessageType">
    ''' The type of message to display.
    ''' </param>
    ''' <returns>
    ''' Returns the result of the user's interaction with the message box.
    ''' </returns>
    Public Shared Function Show(Message As String, MessageType As CMessageBoxType) As DialogResult
        Return Show(Message, Nothing, MessageType)
    End Function
    ''' <summary>
    ''' Displays an error message box associated with the specified exception.
    ''' </summary>
    ''' <remarks>
    ''' The provided exception details are used to display additional information about
    ''' the error that occurred.
    ''' </remarks>
    ''' <param name="Message">
    ''' The message displayed to the user describing the error.
    ''' </param>
    ''' <param name="Exception">
    ''' The exception associated with the error.
    ''' </param>
    ''' <returns>
    ''' Returns the result of the user's interaction with the message box.
    ''' </returns>
    Public Shared Function Show(Message As String, Exception As Exception) As DialogResult
        Return Show(Message, Nothing, CMessageBoxType.Error, Exception)
    End Function
    ''' <summary>
    ''' Displays an error message box with a custom title and an associated exception.
    ''' </summary>
    ''' <remarks>
    ''' The provided exception details are used to display additional information about
    ''' the error that occurred.
    ''' </remarks>
    ''' <param name="Message">
    ''' The message displayed to the user describing the error.
    ''' </param>
    ''' <param name="Title">
    ''' The text displayed in the message box title bar.
    ''' </param>
    ''' <param name="Exception">
    ''' The exception associated with the error.
    ''' </param>
    ''' <returns>
    ''' Returns the result of the user's interaction with the message box.
    ''' </returns>
    Public Shared Function Show(Message As String, Title As String, Exception As Exception) As DialogResult
        Return Show(Message, Title, CMessageBoxType.Error, Exception)
    End Function
    ''' <summary>
    ''' Displays a customized message box with the specified message type, title and optional exception details.
    ''' </summary>
    ''' <remarks>
    ''' This method is the main implementation used to display the CMessageBox.
    '''
    ''' The specified message type controls the visual appearance, icon and behavior of the message box.
    '''
    ''' When an exception is provided, the message type must be set to
    ''' <see cref="CMessageBoxType.Error"/>. The exception details can be displayed to the user
    ''' and optionally sent by email according to the configured options.
    ''' </remarks>
    ''' <param name="Message">
    ''' The text to display in the message box body.
    ''' </param>
    ''' <param name="Title">
    ''' The text displayed in the message box title bar.
    ''' </param>
    ''' <param name="MessageType">
    ''' The type of message to display, defining its appearance and behavior.
    ''' </param>
    ''' <param name="Exception">
    ''' The exception associated with the error message. This parameter can only be specified
    ''' when <paramref name="MessageType"/> is <see cref="CMessageBoxType.Error"/>.
    ''' </param>
    ''' <returns>
    ''' Returns the result of the user's interaction with the message box.
    ''' </returns>
    ''' <exception cref="ArgumentException">
    ''' Thrown when an exception is provided while the message type is different from
    ''' <see cref="CMessageBoxType.Error"/>.
    ''' </exception>
    Public Shared Function Show(Message As String, Title As String, MessageType As CMessageBoxType, Optional Exception As Exception = Nothing) As DialogResult
        If Exception IsNot Nothing AndAlso MessageType <> CMessageBoxType.Error Then
            Throw New ArgumentException("The 'Exception' parameter can only be specified when 'MessageType' is 'Error'.", NameOf(Exception))
        End If
        Using Frm As New FrmMessageBox(Options)
            Using Uc As New UcException()
                If MessageType = CMessageBoxType.Error AndAlso Exception IsNot Nothing Then
                    Dim CMessageBoxException As New CMessageBoxException With {
                        .Title = Title,
                        .Message = Message,
                        .ExceptionMessage = Exception.Message,
                        .StackTrace = Exception.StackTrace,
                        .ExceptionDate = Now.ToString("yyyy-MM-dd HH:mm:ss")
                    }
                    If Exception.InnerException IsNot Nothing Then
                        CMessageBoxException.ExceptionInnerMessage = Exception.InnerException.Message
                    End If
                    If Options.AdditionalInformations IsNot Nothing Then
                        For Each AdditionalInformation In Options.AdditionalInformations
                            CMessageBoxException.AdditionalInformations.Add(AdditionalInformation.Key, AdditionalInformation.Value)
                        Next AdditionalInformation
                    End If
                    Dim Json As String = JsonSerializer.Serialize(CMessageBoxException, _JsonSerializer)
                    Uc.TxtExceptionBody.Font = Options.MessageFont
                    Uc.TxtExceptionBody.ForeColor = Options.MessageForeColor
                    Uc.LblExceptionTitle.Font = Options.TitleFont
                    Uc.LblExceptionTitle.ForeColor = Options.TitleForeColor
                    Uc.TxtExceptionBody.Text = Json
                    If Options.ExceptionEmail Is Nothing Then
                        Uc.TlpContainer.RowStyles(2).SizeType = SizeType.Absolute
                        Uc.TlpContainer.RowStyles(2).Height = 0
                    End If


                    If Options.ShowExceptionDetails Then
                        Frm.CcException.DropDownControl = Uc
                    End If
                    If Options.ExceptionEmail IsNot Nothing Then
                        Task.Run(Async Function()
                                     Try
                                         Await SendEmailAsync(Options.ExceptionEmail, Json)
                                     Catch ex As Exception
                                     End Try
                                 End Function)
                    End If
                End If
                Frm.LblTitle.Text = Title
                If String.IsNullOrEmpty(Title) Then
                    Frm.TlpBody.RowStyles(0).SizeType = SizeType.Absolute
                    Frm.TlpBody.RowStyles(0).Height = 0
                End If
                Frm.LblMessage.Text = Message
                Frm.AllocateButtons(MessageType)
                Frm.SetMessageIcon(MessageType)
                Return Frm.ShowDialog()
            End Using
        End Using
    End Function
    ''' <summary>
    ''' Sends exception details by email using the configured SMTP settings.
    ''' </summary>
    ''' <remarks>
    ''' This method sends the serialized exception information through an SMTP server.
    '''
    ''' Before attempting to send the email, the method checks whether an active network
    ''' connection is available. If no connection is available, the email will not be sent.
    '''
    ''' Any exceptions that occur during the email sending process are handled internally
    ''' to prevent communication failures from affecting the main application flow.
    ''' </remarks>
    ''' <param name="ExceptionEmail">
    ''' Contains the SMTP configuration and email information required to send the message,
    ''' including sender, recipient, server, port, authentication credentials and security mode.
    ''' </param>
    ''' <param name="Json">
    ''' The serialized JSON content containing the exception details to be sent.
    ''' </param>
    ''' <returns>
    ''' Returns a task representing the asynchronous email sending operation.
    ''' </returns>
    Public Shared Async Function SendEmailAsync(ExceptionEmail As CMessageBoxExceptionEmail, Json As String) As Task
        Try
            Using Connectivity As New Connectivity()
                If Await Connectivity.IsAvailableAsync() Then
                    Dim Message As New MimeMessage()
                    Message.From.Add(New MailboxAddress(ExceptionEmail.FromName, ExceptionEmail.FromEmail))
                    Message.To.Add(New MailboxAddress(ExceptionEmail.ToName, ExceptionEmail.ToEmail))
                    Message.Subject = "Aviso de Exceção na Aplicação"
                    Message.Body = New TextPart("plain") With {.Text = Json}
                    Using Client As New SmtpClient()
                        Await Client.ConnectAsync(ExceptionEmail.Host, ExceptionEmail.Port, CType(ExceptionEmail.SecureSocket, MailKit.Security.SecureSocketOptions))
                        Await Client.AuthenticateAsync(ExceptionEmail.FromEmail, ExceptionEmail.Password)
                        Await Client.SendAsync(Message)
                        Await Client.DisconnectAsync(True)
                    End Using
                End If
            End Using
        Catch ex As Exception
        End Try
    End Function
End Class
