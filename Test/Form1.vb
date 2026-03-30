Imports CoreSuite.Services

Public Class Form1
    Private Async Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Remote As New FirebaseService("", "", "")
        Await Remote.Auth.LoginAsync("josias@manager.com", "123456")
        Dim Docs = Await Remote.Firestore.GetAllDocumentsAsync("evaluations")
        Dim Local As New MySqlService("", "", "", "")
        Dim EvaluationID As Long
        Dim info As Dictionary(Of String, Object)
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





            If Doc.ContainsKey("info") Then
                info = Doc("info")
                'Debug.Print("hasreplacedproducts: " & info("hasreplacedproducts"))
                Debug.Print("importedby: " & info("importedby"))
                Debug.Print("importeddate: " & info("importeddate"))
                'Debug.Print("importedid: " & info("importedid"))
                'Debug.Print("importingby: " & info("importingby"))
                'Debug.Print("importingdate: " & info("importingdate"))
                'Debug.Print("requestprocessed: " & info("requestprocessed"))
                'Debug.Print("visitscheduleid: " & info("visitscheduleid"))







                info("hasreplacedproducts") = False
                info("importedby") = If(EvaluationID > 0, "LUIZ", Nothing)
                info("importeddate") = If(EvaluationID > 0, "29/03/2026 00:00:00", Nothing)
                info("importedid") = If(EvaluationID > 0, EvaluationID, Nothing)
                info("importingby") = Nothing
                info("importingdate") = Nothing
                info("requestprocessed") = True
                info("visitscheduleid") = Nothing
            Else

                info = New Dictionary(Of String, Object) From {
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
            End If


            'Await Remote.Firestore.SaveDocumentAsync("evaluations", Doc("id"), Doc)
        Next Doc
    End Sub
End Class