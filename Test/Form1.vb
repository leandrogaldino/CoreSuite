Imports CoreSuite

Friend Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load



    End Sub

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Service As New FirebaseService("AIzaSyD0C8aIjG7X1VHIsGYth_eHOhFY8PSdCF0", "manager-mobile-7de2b", "manager-mobile-7de2b.appspot.com")
        Await Service.Auth.LoginAsync("tecnico@manager.com", "123456")
        If (Service.Auth.IsLoggedIn) Then
            Dim Evaluations = Await Service.Firestore.GetAllDocumentsAsync("evaluations")

            For Each Evaluation In Evaluations

                Evaluation.Remove("oiltypeid")


                Await Service.Storage.UploadFile("C:\Users\Leandro\Desktop\cotovelo.jpeg", "cotovelo.jpeg")

            Next Evaluation

            MsgBox(Evaluations.Count)
        End If

    End Sub
End Class
