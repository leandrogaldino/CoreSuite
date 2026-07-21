Imports CoreSuite
Friend Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CMessageBox.Options = New CMessageBoxOptions With {
            .ShowExceptionDetails = True
        }
        Try
            Dim s As String = Nothing
            Dim x = s.Length
        Catch ex As Exception
            Dim result = CMessageBox.Show("Leandro Galdino", "ERRO SC001", ex)
        End Try
    End Sub
End Class
