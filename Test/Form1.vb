Imports CoreSuite.Services

Public Class Form1
    Private Async Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Remote As New FirebaseService("", "", "")
        Await Remote.Auth.LoginAsync("josias@manager.com", "123456")
        Dim Docs = Await Remote.Firestore.GetAllDocumentsAsync("evaluations")
        Dim Local As New MySqlService("localhost", "manager", "root", "123456")
        Dim EvaluationID As Long
        For Each Doc In Docs

            Dim LocalResult = Await Local.Request.ExecuteSelectAsync("evaluation", New MySqlSelectOptions() With {
                .Columns = {"id"}.ToList,
                .Where = $"cloudid = '{Doc("id")}'"
            })

            If LocalResult.HasData Then
                EvaluationID = LocalResult.Data(0)("id")
            Else
                EvaluationID = 0
            End If
            Dim info As New Dictionary(Of String, Object) From {
                {"hasreplacedproducts", False},
                {"importedby", If(EvaluationID > 0, "LUIZ", Nothing)},
                {"importeddate", If(EvaluationID > 0, "29/03/2026 00:00:00", Nothing)},
                {"importedid", If(EvaluationID > 0, EvaluationID, Nothing)},
                {"importingby", Nothing},
                {"importingdate", Nothing},
                {"requestprocessed", True},
                {"visitscheduleid", Nothing}
            }
            Doc.Add("info", info)
            'If Doc("id") = "4a0393b63-3ee7-4219-ac1b-cd74b35eb55f_23032026_125339214" Then Continue For
            'ProcessRequests(Doc)
            'ProcessEvaluations(Doc)
            Await Remote.Firestore.SaveDocumentAsync("evaluations", Doc("id"), Doc)
        Next Doc
    End Sub
End Class