Imports System.Text.Json
Imports MimeKit
Imports MailKit.Net.Smtp

Public Class CMessageBox
    Public Shared Property Options As CMessageBoxOptions = New CMessageBoxOptions()
    Private Shared ReadOnly _JsonSerializer As New JsonSerializerOptions With {.WriteIndented = True}
    Public Shared Function Show(Message As String) As DialogResult
        Return Show(Message, Nothing, CMessageBoxType.Information)
    End Function
    Public Shared Function Show(Message As String, Title As String) As DialogResult
        Return Show(Message, Title, CMessageBoxType.Information)
    End Function
    Public Shared Function Show(Message As String, MessageType As CMessageBoxType) As DialogResult
        Return Show(Message, Nothing, MessageType)
    End Function
    Public Shared Function Show(Message As String, Exception As Exception) As DialogResult
        Return Show(Message, Nothing, CMessageBoxType.Error, Exception)
    End Function
    Public Shared Function Show(Message As String, Title As String, Exception As Exception) As DialogResult
        Return Show(Message, Title, CMessageBoxType.Error, Exception)
    End Function
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
                    Uc.LblExceptionTitle.Font = Options.TitleFont
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
