Imports System.Data.Common
Imports CoreSuite.Helpers
Partial Public Class QueriedBox
    Public Sub DebugBaseQuery()
        Debug.Print(GetBaseQuery())
    End Sub
    Private Function GetBaseQuery() As String
        Dim SelectSql As String = $"SELECT {Query.GetSelect()}"
        Dim JoinsSql As String = Query.GetJoins()
        Dim FullSql As String = $"{SelectSql} {JoinsSql}"
        Return FullSql
    End Function
    Private Function ExecuteQuery(Query As String, Optional Parameters As Dictionary(Of String, Object) = Nothing) As DataTable
        Dim Table As New DataTable
        Dim Par As DbParameter
        Using Connection As DbConnection = ConnectionFactory.Invoke()
            Dim Factory As DbProviderFactory = DbProviderFactories.GetFactory(Connection)
            Using Cmd As IDbCommand = Connection.CreateCommand()
                Cmd.CommandText = Query
                If Parameters IsNot Nothing Then
                    For Each P In Parameters
                        Par = Cmd.CreateParameter()
                        Par.ParameterName = P.Key
                        Par.Value = If(P.Value, DBNull.Value)
                        Cmd.Parameters.Add(Par)
                    Next P
                End If
                If Diagnostics.DebugOnTextChanged Then DatabaseHelper.DebugQuery(Cmd)
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
                End Using
            End Using
        End Using

        Return Table
    End Function
    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles Timer.Tick
        Dim FullQuery As String
        Dim ParameterList As Dictionary(Of String, Object)
        Dim TableResult As DataTable
        Dim ValueParameter As String
        Timer.Interval = Search.Interval
        Timer.Stop()
        ValidateQueryConfiguration()
        ValueParameter = SqlDialect.GetParameterPrefix() & "p" & Guid.NewGuid.ToString("N")
        _PrimaryKeyAlias = "pk" & Guid.NewGuid.ToString("N")
        Dim Conditions As New List(Of String)
        For Each c In Query.Columns.Where(Function(x) x.Options.Searchable)
            Conditions.Add($"LOWER({c.ColumnName}) LIKE LOWER({ValueParameter})")
        Next c
        Dim SearchWhere As String = $"({String.Join(" OR ", Conditions)})"
        FullQuery = Query.ToStringWithAdditions($"{Query.PrimaryKeyColumnName} AS {_PrimaryKeyAlias}", SearchWhere)
        ParameterList = New Dictionary(Of String, Object) From {
            {ValueParameter, "%" & Text & "%"}
        }
        For Each p As QueryParameter In Query.Parameters
            ParameterList.Add(p.ParameterName, p.ParameterValue)
        Next p
        Try
            TableResult = New DataTable
            TableResult = ExecuteQuery(FullQuery, ParameterList)
            If DropDownResultsForm IsNot Nothing Then
                Dim Dgv As DataGridView = DropDownResultsForm.DgvResults
                Dgv.DataSource = TableResult
                Dgv.Columns(_PrimaryKeyAlias).Visible = False
                For Each c In Query.Columns
                    Dim ColumnText As String = If(String.IsNullOrEmpty(c.ColumnAlias), c.ColumnName, c.ColumnAlias)
                    If Not c.Options.Display Then
                        If Dgv.Columns.Contains(ColumnText) Then
                            Dgv.Columns(ColumnText).Visible = False
                        End If
                    End If
                    Dgv.Columns.Cast(Of DataGridViewColumn).First(Function(DgvC) UCase(DgvC.Name) = UCase(ColumnText)).AutoSizeMode = c.Options.SizeColumnMode
                Next c
                If DropDown.AutoStretchRight Then
                    For Each c In Dgv.Controls
                        If c.GetType() Is GetType(HScrollBar) Then
                            Dim vbar As HScrollBar = DirectCast(c, HScrollBar)
                            If vbar.Visible = True AndAlso Dgv.Rows.Count > 0 Then
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
    Private Sub ValidateQueryConfiguration()
        If Search.Interval < 300 Then
            ThrowValidationException($"The {NameOf(Search.Interval)} property must be greater than or equal to 300 milliseconds. Current value: {Search.Interval}.")
        End If
        If ConnectionFactory Is Nothing Then
            ThrowValidationException($"The {NameOf(ConnectionFactory)} property was not configured.")
        End If
        If Query Is Nothing Then
            ThrowValidationException($"The {NameOf(Query)} property was not configured.")
        End If
        If Query.Table Is Nothing OrElse String.IsNullOrWhiteSpace(Query.Table.TableName) Then
            ThrowValidationException($"The Query table was not configured.")
        End If
        If String.IsNullOrWhiteSpace(Query.PrimaryKeyColumnName) Then
            ThrowValidationException($"The {NameOf(Query.PrimaryKeyColumnName)} property was not configured.")
        End If
        If Query.Columns Is Nothing OrElse Query.Columns.Count = 0 Then
            ThrowValidationException($"The {NameOf(QueriedBox)} must have at least one column configured.")
        End If
        If Query.Columns.All(Function(c) Not c.Options.Freeze) Then
            ThrowValidationException($"At least one column must have the {NameOf(QueryColumnOptions.Freeze)} property set to True.")
        End If
        If Query.Columns.All(Function(c) Not c.options.Display) Then
            ThrowValidationException($"At least one column must have the {NameOf(QueryColumnOptions.Display)} property set to True.")
        End If
        Dim ColumnNames = Query.Columns.Select(Function(c) If(String.IsNullOrWhiteSpace(c.ColumnAlias), c.ColumnName, c.ColumnAlias)).ToList()
        If ColumnNames.Count <> ColumnNames.Distinct().Count() Then
            ThrowValidationException($"There are Columns with duplicated names or aliases.")
        End If
        For Each c In Query.Columns
            If c Is Nothing OrElse String.IsNullOrWhiteSpace(c.ColumnName) Then
                ThrowValidationException($"There is a Column with the {NameOf(QueryColumn.ColumnName)} property not configured.")
            End If
            If c Is Nothing OrElse String.IsNullOrWhiteSpace(c.ColumnAlias) Then
                ThrowValidationException($"There is a Column with the {NameOf(QueryColumn.ColumnAlias)} property not configured.")
            End If
        Next c
        Dim ParameterNames As New List(Of String)
        If Query.Parameters IsNot Nothing Then
            For Each p In Query.Parameters
                If p Is Nothing Then
                    ThrowValidationException("There is an invalid Parameter.")
                End If
                If String.IsNullOrWhiteSpace(p.ParameterName) Then
                    ThrowValidationException($"There is a Parameter with the {NameOf(QueryParameter.ParameterName)} property not configured.")
                End If
                ParameterNames.Add(p.ParameterName)
            Next p

            If ParameterNames.Count <> ParameterNames.Distinct().Count() Then
                ThrowValidationException("There is a Parameter with a duplicated ParameterName.")
            End If
        End If
        If Query.Conditions IsNot Nothing Then
            For Each c In Query.Conditions
                If c Is Nothing Then
                    ThrowValidationException("There is an invalid Condition.")
                End If
                If c.Column Is Nothing OrElse String.IsNullOrWhiteSpace(c.Column.ColumnName) Then
                    ThrowValidationException("There is a Condition with an undefined ColumnName.")
                End If
                If c.Values Is Nothing OrElse c.Values.Length = 0 Then
                    ThrowValidationException($"The Condition for column '{c.Column.ColumnName}' does not have any values.")
                End If

                Select Case c.Operator
                    Case QueryConditionOperator.Between
                        If c.Values.Length <> 2 Then
                            ThrowValidationException($"The BETWEEN condition for column '{c.Column.ColumnName}' requires exactly two values.")
                        End If
                End Select
            Next c
        End If
        If Query.Joins IsNot Nothing Then
            For Each j In Query.Joins
                If j Is Nothing Then
                    ThrowValidationException("There is an invalid Join.")
                End If
                If j.Table Is Nothing OrElse String.IsNullOrWhiteSpace(j.Table.TableName) Then
                    ThrowValidationException("There is a Join without a configured table.")
                End If
                If j.Conditions Is Nothing OrElse j.Conditions.Count = 0 Then
                    ThrowValidationException($"The Join for table '{j.Table.TableName}' does not have any conditions.")
                End If
                For Each jc In j.Conditions
                    If jc Is Nothing Then
                        ThrowValidationException($"The Join for table '{j.Table.TableName}' contains an invalid JoinCondition.")
                    End If
                    If jc.LeftColumn Is Nothing OrElse String.IsNullOrWhiteSpace(jc.LeftColumn.ColumnName) Then
                        ThrowValidationException("There is a JoinCondition without a LeftColumn.")
                    End If
                    If jc.RightColumn Is Nothing OrElse String.IsNullOrWhiteSpace(jc.RightColumn.ColumnName) Then
                        ThrowValidationException("There is a JoinCondition without a RightColumn.")
                    End If
                    Select Case jc.Operator
                        Case QueryConditionOperator.Between, QueryConditionOperator.In, QueryConditionOperator.NotIn
                            ThrowValidationException($"The operator '{jc.Operator.GetSqlValue()}' cannot be used in a JoinCondition with only one RightColumn.")
                    End Select
                Next jc
            Next j
        End If
    End Sub
    Private Sub ThrowValidationException(Message As String)
        If DropDownResultsForm IsNot Nothing Then
            DropDownResultsForm.Close()
            DropDownResultsForm = Nothing
        End If
        Throw New InvalidOperationException(Message)
    End Sub
End Class
