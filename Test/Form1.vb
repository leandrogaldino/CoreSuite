Imports CoreSuite
Imports MySql.Data.MySqlClient

Friend Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim DbService As New MySqlService("localhost", "manager", "root", "123456")
        Dim Con = DbService.Client.CreateDatabaseConnection()
        QueriedBox.Connection = Con

        'CMessageBox.Options = New CMessageBoxOptions With {
        '    .ShowExceptionDetails = True,
        '    .ExceptionEmail = New CMessageBoxExceptionEmail With {
        '        .FromEmail = "nexor.sup@gmail.com",
        '        .FromName = "Nexor Support",
        '        .ToEmail = "nexor.sup@gmail.com",
        '        .ToName = "Nexor Support",
        '        .Host = "smtp.gmail.com",
        '        .Port = 587,
        '        .SecureSocket = CMessageBoxSecureSocket.StartTls,
        '        .Password = "yxsy ywwz ozka ncuj"
        '    }
        '}
        'Try
        '    Dim s As String = Nothing
        '    Dim x = s.Length
        'Catch ex As Exception
        '    Dim result = CMessageBox.Show("Leandro Galdino", "ERRO SC001", ex)
        'End Try

    End Sub
End Class
