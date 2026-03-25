Imports CoreSuite.Services

Public Class Form1
    Private Async Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Service As New FirebaseService("", "", "")
        Await Service.Auth.LoginAsync("reicol@manager.com", "123456")

        Dim Docs = Await Service.Firestore.GetAllDocumentsAsync("evaluations")


        Service = New FirebaseService("", "", "")
        Await Service.Auth.LoginAsync("leandro@manager.com", "123456")

        For Each Doc In Docs

            'If Doc("id") = "4a0393b63-3ee7-4219-ac1b-cd74b35eb55f_23032026_125339214" Then Continue For

            'ProcessRequests(Doc)
            'ProcessEvaluations(Doc)

            Await Service.Firestore.SaveDocumentAsync("evaluations", Doc("id"), Doc)

        Next Doc


    End Sub

    Private Function ProcessEvaluations(Doc As Dictionary(Of String, Object)) As Dictionary(Of String, Object)
        Dim info = Doc("info")

        info("importedby") = "LUIZ"
        info("importeddate") = "10/03/2026 00:00:00"

        Return Doc
    End Function



    Private Function ProcessRequests(Doc As Dictionary(Of String, Object)) As Dictionary(Of String, Object)

        Dim replacedProducts = TryCast(Doc("replacedproducts"), IList)

        Dim info = Doc("info")



        Dim hasProducts As Boolean = False
        If replacedProducts IsNot Nothing Then
            hasProducts = replacedProducts.Count > 0
        End If

        Doc("greasing") = 0
        info("hasreplacedproducts") = hasProducts
        info("requestprocessed") = True

        Return Doc
    End Function
End Class