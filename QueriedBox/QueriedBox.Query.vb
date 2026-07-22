Imports System.Data.Common

Partial Public Class QueriedBox
    Public Sub DebugQuery()
        Debug.Print(GetQuery())
    End Sub
    Private Function GetQuery() As String
        Dim Query As String
        Query = String.Format("SELECT {0}{1}{2}{3}.{4} AS 'mainid',{1}{2}{8}{5}.{6}{9} AS '{7}'",
                                  If(Distinct, "DISTINCT", Nothing), '0
                                  Environment.NewLine, '1
                                  vbTab, '2
                                  If(MainTableAlias = Nothing, MainTableName, MainTableAlias), '3
                                  MainReturnFieldName, '4
                                  If(DisplayTableAlias = Nothing, DisplayTableName, DisplayTableAlias), '5
                                  DisplayFieldName, '6
                                  If(DisplayFieldAlias = Nothing, DisplayFieldName, DisplayFieldAlias), '7
                                  If(IfNull <> Nothing, "IFNULL(", Nothing), '8
                                  If(IfNull <> Nothing, ", " & IfNull & ") ", Nothing) '9
                              )
        For Each o As QueriedBoxField In OtherFields
            Query += String.Format(",{0}{1}{5}{2}.{3}{6} AS '{4}'",
                                       Environment.NewLine, '0
                                       vbTab, '1
                                       If(o.DisplayTableAlias = Nothing, o.DisplayTableName, o.DisplayTableAlias), '2
                                       o.DisplayFieldName, '3
                                       If(o.DisplayFieldAlias = Nothing, o.DisplayFieldName, o.DisplayFieldAlias), '4
                                       If(o.IfNull <> Nothing, "IFNULL(", Nothing),'5
                                       If(o.IfNull <> Nothing, ", " & o.IfNull & ") ", Nothing)
                                   )
        Next o
        Query += String.Format("{0}FROM {1} AS {2}",
                               Environment.NewLine, '0
                               MainTableName, '1 
                               If(MainTableAlias = Nothing, MainTableName, MainTableAlias)) '2)
        For Each r As QueriedBoxRelation In Relations
            Query += String.Format("{0}{1} JOIN {2} AS {3} ON {3}.{4} {5} {6}.{7}",
                                       Environment.NewLine, '0
                                       r.RelationType, '1
                                       r.RelateTableName, '2
                                       If(r.RelateTableAlias = Nothing, r.RelateTableName, r.RelateTableAlias), '3
                                       r.RelateFieldName, '4
                                       r.Operator, '5
                                       If(r.WithTableAlias = Nothing, r.WithTableName, r.WithTableAlias), '6
                                       r.WithFieldName '7
                                   )
            For Each c As QueriedBoxCondition In r.Conditions
                Query += String.Format(" AND{0}{1}{2}.{3} {4} {5}",
                                           Environment.NewLine,
                                           vbTab,
                                           c.TableNameOrAlias,
                                           c.FieldName,
                                           c.Operator,
                                           c.Value
                                       )
            Next c
        Next r
        Return Query
    End Function
    Private Function ExecuteQuery(Query As String, Optional Parameters As Dictionary(Of String, Object) = Nothing) As DataTable
        Dim Table As New DataTable
        Dim Par As DbParameter
        Dim Factory As DbProviderFactory = DbProviderFactories.GetFactory(Connection)
        Using Cmd As IDbCommand = Connection.CreateCommand
            Cmd.CommandText = Query
            If Parameters IsNot Nothing Then
                For Each P In Parameters
                    Par = Cmd.CreateParameter
                    Par.ParameterName = P.Key
                    Par.Value = P.Value
                    Cmd.Parameters.Add(Par)
                Next P
            End If
            If DebugOnTextChanged Then DatabaseHelper.DebugQuery(Cmd)
            Using Adp As DbDataAdapter = Factory.CreateDataAdapter()
                Adp.SelectCommand = Cmd
                Connection.Open()
                Try
                    Adp.Fill(Table)
                Catch ex As Exception
                    If DropDownResultsForm IsNot Nothing Then
                        DropDownResultsForm.Close()
                        DropDownResultsForm = Nothing
                    End If
                    Throw
                Finally
                    Connection.Close()
                End Try
                Return Table
            End Using
        End Using
    End Function
    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles Timer.Tick
        Dim Query As String
        Dim ParameterList As Dictionary(Of String, Object)
        Dim TableResult As DataTable
        Dim ValueParameter As String
        Timer.Interval = QueryInterval
        Timer.Stop()
        ValidateTick()
        Query = GetQuery()
        ValueParameter = "@" & Guid.NewGuid.ToString("N")
        Query += String.Format("{0}WHERE{0}({0}{1}{2}.{3} LIKE {4}",
                                    Environment.NewLine,
                                    vbTab,
                                    If(DisplayTableAlias = Nothing, DisplayTableName, DisplayTableAlias),
                                    DisplayFieldName,
                                    ValueParameter
                               )
        For Each o As QueriedBoxField In OtherFields
            Query += String.Format(" OR{0}{1}{2}.{3} LIKE {4}",
                                   Environment.NewLine,
                                   vbTab,
                                   If(o.DisplayTableAlias = Nothing, o.DisplayTableName, o.DisplayTableAlias),
                                   o.DisplayFieldName,
                                   ValueParameter
                               )
        Next o
        Query += String.Format("{0})", Environment.NewLine)
        If Conditions.Count > 0 Then
            Query += String.Format(" AND{0}(", Environment.NewLine)
            For Each c As QueriedBoxCondition In Conditions
                Query += String.Format("{0}{1}{2}.{3} {4} {5} AND",
                                       Environment.NewLine,
                                       vbTab,
                                       c.TableNameOrAlias,
                                       c.FieldName,
                                       c.Operator,
                                       c.Value
                                   )
            Next c
            Query = Strings.Left(Query, Query.Length - 4)
            Query += String.Format("{0})", Environment.NewLine)
        End If
        If Limit > 0 Then
            Query += $"LIMIT {Limit};"
        Else
            Query += ";"
        End If
        ParameterList = New Dictionary(Of String, Object) From {
            {ValueParameter, "%" & Text & "%"}
        }
        For Each p As QueriedBoxParameter In Parameters
            ParameterList.Add(p.ParameterName, p.ParameterValue)
        Next p
        Try
            TableResult = New DataTable
            TableResult = ExecuteQuery(Query, ParameterList)
            If DropDownResultsForm IsNot Nothing Then
                DropDownResultsForm.DgvResults.DataSource = TableResult
                DropDownResultsForm.DgvResults.Columns("mainid").Visible = False
                For Each o As QueriedBoxField In OtherFields
                    If Not o.Display Then
                        If DropDownResultsForm.DgvResults.Columns.Contains(If(String.IsNullOrEmpty(o.DisplayFieldAlias), o.DisplayFieldName, o.DisplayFieldAlias)) Then
                            DropDownResultsForm.DgvResults.Columns(If(String.IsNullOrEmpty(o.DisplayFieldAlias), o.DisplayFieldName, o.DisplayFieldAlias)).Visible = False
                        End If
                    End If
                Next o
                DropDownResultsForm.DgvResults.Columns.Cast(Of DataGridViewColumn).First(Function(c) c.Name = If(String.IsNullOrEmpty(DisplayFieldAlias), DisplayFieldName, DisplayFieldAlias)).AutoSizeMode = DisplayFieldAutoSizeColumnMode
                For Each OtherField In OtherFields
                    DropDownResultsForm.DgvResults.Columns.Cast(Of DataGridViewColumn).First(Function(c) c.Name = If(String.IsNullOrEmpty(OtherField.DisplayFieldAlias), OtherField.DisplayFieldName, OtherField.DisplayFieldAlias)).AutoSizeMode = OtherField.DisplayFieldAutoSizeColumnMode
                Next OtherField
                If DropDownAutoStretchRight Then
                    For Each c In DropDownResultsForm.DgvResults.Controls
                        If c.GetType() Is GetType(HScrollBar) Then
                            Dim vbar As HScrollBar = DirectCast(c, HScrollBar)
                            If vbar.Visible = True AndAlso DropDownResultsForm.DgvResults.Rows.Count > 0 Then
                                Do Until vbar.Visible = False
                                    DropDownResultsForm.Width += 10
                                Loop
                            End If
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            CloseDropDown()
            Throw
        End Try
    End Sub
    Private Sub ValidateTick()
        Dim FieldHeaders As New List(Of String)
        Dim ParametersNames As New List(Of String)
        If Connection Is Nothing Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("A propriedade Connection não foi definida.")
        End If
        If MainTableName = Nothing Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("A propriedade MainTableName não foi definida.")
        End If
        If MainReturnFieldName = Nothing Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("A propriedade MainReturnFieldName não foi definida.")
        End If
        If DisplayTableName = Nothing Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("A propriedade DisplayTableName não foi definida.")
        End If
        If DisplayFieldName = Nothing Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("A propriedade DisplayFieldName não foi definida.")
        End If
        If DisplayMainFieldName = Nothing Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("A propriedade DisplayMainFieldName  não foi definida.")
        End If
        If DisplayFieldAlias <> Nothing Then
            FieldHeaders.Add(DisplayFieldAlias)
        End If
        If FieldHeaders.Count <> FieldHeaders.Distinct.Count Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("Existe mais de um controle com o mesmo valor para a propriedade FieldHeader.")
        End If
        For i = 0 To OtherFields.Count - 1
            If OtherFields(i).DisplayTableName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe OtherField com a propriedade TableName não definida.")
            End If
            If OtherFields(i).DisplayFieldName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe OtherField com a propriedade FieldName não definida.")
            End If
        Next i
        For i = 0 To Parameters.Count - 1
            If Parameters(i).ParameterName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                ParametersNames.Clear()
                Throw New Exception("Existe Parameter com a propriedade ParameterName não definida.")
            Else
                ParametersNames.Add(Parameters(i).ParameterName)
            End If
        Next
        If ParametersNames.Count <> ParametersNames.Distinct.Count Then
            DropDownResultsForm.Close() : DropDownResultsForm = Nothing
            Throw New Exception("Existe Parameter com a propriedade ParameterName duplicada.")
            Exit Sub
        End If
        For i = 0 To Conditions.Count - 1
            If Conditions(i).TableNameOrAlias = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Condition com a propriedade TableName não definida.")
                Exit Sub
            End If
            If Conditions(i).FieldName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Condition com a propriedade FieldName não definida.")
                Exit Sub
            End If
            If Conditions(i).Operator = "BETWEEN" And Conditions(i).Value <> Nothing AndAlso Conditions(i).Value.Split(";").Length <> 2 Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Condition com a propriedade Value não definida para o operador BETWEEN.")
                Exit Sub
            End If
        Next i
        For i = 0 To Relations.Count - 1
            If Relations(i).RelationType = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Relation com a propriedade RelationType não definida.")
                Exit Sub
            End If
            If Relations(i).Operator = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Relation com a propriedade Operator não definida.")
                Exit Sub
            End If
            If Relations(i).RelateTableName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Relation com a propriedade TableNameX não definida.")
                Exit Sub
            End If
            If Relations(i).WithTableName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Relation com a propriedade TableNameY não definida.")
                Exit Sub
            End If
            If Relations(i).RelateFieldName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Relation com a propriedade FieldNameX não definida.")
                Exit Sub
            End If
            If Relations(i).WithFieldName = Nothing Then
                DropDownResultsForm.Close() : DropDownResultsForm = Nothing
                Throw New Exception("Existe Relation com a propriedade FieldNameY não definida.")
                Exit Sub
            End If
        Next i
    End Sub
End Class
