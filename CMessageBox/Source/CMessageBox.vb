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
        Task.Run(Async Function()
                     Try
                         Await SendEmailAsync()
                     Catch ex As Exception
                         ' Não deixe falha de email derrubar a aplicação
                         MsgBox(ex.Message)
                     End Try
                 End Function)
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
                        .ExceptionDate = Now
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
                    Uc.TxtExceptionBody.Text = Json
                    If Options.ShowExceptionDetails Then
                        Frm.CcException.DropDownControl = Uc
                    End If
                    If Options.SendEmailOnException AndAlso Options.Email IsNot Nothing Then


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

    Public shared Async Function SendEmailAsync() As Task
        Dim Message As New MimeMessage()
        Message.From.Add(New MailboxAddress("My Application", "compras@reicolservice.com.br"))
        Message.To.Add(New MailboxAddress("Leandro", "compras@reicolservice.com.br"))
        Message.Subject = "Test email"
        Message.Body = New TextPart("plain") With {.Text = "This is a test email."}
        Using Client As New SmtpClient()
            Await Client.ConnectAsync("email-ssl.com.br", 465, MailKit.Security.SecureSocketOptions.Auto)
            Await Client.AuthenticateAsync("compras@reicolservice.com.br", "7y9$ORz5!WxP")
            Await Client.SendAsync(Message)
            Await Client.DisconnectAsync(True)
        End Using
    End Function
End Class
