Imports CoreSuite.Services

Public Class Form1
    Private Async Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Service As New FirebaseService("", "", "")
        Await Service.Auth.LoginAsync("", "")

        Dim Docs = Await Service.Firestore.GetAllDocumentsAsync("evaluations")

        For Each Doc In Docs

            Dim replacedProducts = TryCast(Doc("replacedproducts"), IList)

            Dim info = Doc("info")



            Dim hasProducts As Boolean = False
            If replacedProducts IsNot Nothing Then
                hasProducts = replacedProducts.Count > 0
            End If

            Doc("greasing") = 0
            info("hasreplacedproducts") = hasProducts
            info("requestprocessed") = True

            Await Service.Firestore.SaveDocumentAsync("evaluations", Doc("id"), Doc)

        Next Doc


    End Sub
End Class