Public Class CMessageBox
    Public Shared Function Show(Message As String) As DialogResult
        Return Show(Message, Nothing, CMessageBoxType.Information)
    End Function
    Public Shared Function Show(Message As String, Title As String) As DialogResult
        Return Show(Message, Title, CMessageBoxType.Information)
    End Function
    Public Shared Function Show(Message As String, MessageType As CMessageBoxType) As DialogResult
        Return Show(Message, Nothing, MessageType)
    End Function

    Public Shared Function Show(Message As String, Title As String, MessageType As CMessageBoxType) As DialogResult
        Using Frm As New FrmMessageBox()
            Frm.LblTitle.Text = Title
            If String.IsNullOrEmpty(Title) Then
                Frm.TlpBody.RowStyles(0).SizeType = SizeType.Absolute
                Frm.TlpBody.RowStyles(0).Height = 0
            End If
            Frm.LblMessage.Text = Message
            Frm.AllocateButtons(MessageType)
            Return Frm.ShowDialog()
        End Using
    End Function
End Class
