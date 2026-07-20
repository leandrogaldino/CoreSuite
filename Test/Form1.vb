Imports CoreSuite.CMessageBox

Friend Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim result = CMessageBox.Show("Leandro Galdino", CMessageBoxType.Error)
        MsgBox(result.ToString)
    End Sub
End Class
