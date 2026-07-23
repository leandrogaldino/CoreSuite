Partial Public Class QueriedBox
    Public Sub Freeze(ID As Long)
        Dim OldPrimaryKey As Long = _FreezedPrimaryKey
        Dim Query As String
        Dim TableResults As DataTable
        Dim MainValue As String
        Dim OtherValue As String
        Dim FullValue As String = String.Empty
        Dim ParameterList As Dictionary(Of String, Object)
        Dim TableName As String = If(MainTableAlias = Nothing, MainTableName, MainTableAlias)
        If QueryEnabled Then
            _Freezing = True
            Query = GetQuery()
            Query &= $"{Environment.NewLine}WHERE{Environment.NewLine}{vbTab}{TableName}.{MainReturnFieldName} = {ID}"
            ParameterList = New Dictionary(Of String, Object)
            For Each p As QueriedBoxParameter In Parameters
                ParameterList.Add(p.ParameterName, p.ParameterValue)
            Next p
            TableResults = ExecuteQuery(Query, ParameterList)
            If TableResults.Rows.Count = 1 Then
                QueryEnabled = False
                MainValue = TableResults.Rows(0).Item(If(DisplayFieldAlias = Nothing, DisplayFieldName, DisplayFieldAlias)).ToString
                _RawFreezedValues.Add((DisplayTableName, DisplayFieldName, MainValue))
                If Not String.IsNullOrEmpty(MainValue) Then
                    FullValue &= Prefix & MainValue & Suffix
                End If
                For Each o As QueriedBoxField In OtherFields
                    OtherValue = TableResults.Rows(0).Item(If(o.DisplayFieldAlias = Nothing, o.DisplayFieldName, o.DisplayFieldAlias)).ToString
                    _RawFreezedValues.Add((o.DisplayTableName, o.DisplayFieldName, OtherValue))
                    If o.Freeze Then
                        If Not String.IsNullOrEmpty(OtherValue) Then
                            FullValue &= o.Prefix & OtherValue & o.Suffix
                        End If
                    End If
                Next o
                If OldPrimaryKey <> ID Then
                    RaiseEvent FreezedPrimaryKeyChanging(Me, EventArgs.Empty)
                End If
                Text = FullValue
                ForeColor = FreezeColor
                _FreezedValue = FullValue
                _FreezedPrimaryKey = ID
                _IsFreezed = True
                If _FreezedPrimaryKey <> OldPrimaryKey Then
                    RaiseEvent FreezedPrimaryKeyChanged(Me, EventArgs.Empty)
                End If
                If ShowStartOnFreeze Then
                    Me.Select(0, 0)
                Else
                    Me.Select(Me.TextLength, 0)
                End If
                _CtrlHyperlink = True
                QueryEnabled = True
                CloseDropDown()
                _Freezing = False
            Else
                Unfreeze()
            End If
        End If
    End Sub
    Public Sub Freeze(Table As String, Field As String, ID As Long)
        Dim OldPrimaryKey As Long = _FreezedPrimaryKey
        Dim Query As String
        If QueryEnabled Then
            _Freezing = True
            Query = $"SELECT {Field} FROM {Table} WHERE ID = @ID"
            Dim TableResults As DataTable = ExecuteQuery(Query, New Dictionary(Of String, Object) From {{"@ID", ID}})
            If TableResults.Rows.Count = 1 Then
                QueryEnabled = False
                If ShowStartOnFreeze Then
                    Me.Select(0, 0)
                Else
                    Me.Select(Me.TextLength, 0)
                End If
                _CtrlHyperlink = True
                If OldPrimaryKey <> ID Then
                    RaiseEvent FreezedPrimaryKeyChanging(Me, EventArgs.Empty)
                End If
                Text = TableResults.Rows(0).Item(Field).ToString
                ForeColor = FreezeColor
                _IsFreezed = True
                _FreezedPrimaryKey = ID
                If _FreezedPrimaryKey <> OldPrimaryKey Then
                    RaiseEvent FreezedPrimaryKeyChanged(Me, EventArgs.Empty)
                End If
                _RawFreezedValues.Add((Table, Field, TableResults.Rows(0).Item(Field).ToString))
                QueryEnabled = True
                CloseDropDown()
                _Freezing = False
            Else
                Unfreeze()
            End If
        End If
    End Sub
    Public Sub Unfreeze()
        Dim OldPrimaryKey As Long
        If QueryEnabled Then
            QueryEnabled = False
            OldPrimaryKey = _FreezedPrimaryKey
            If OldPrimaryKey <> 0 Then
                RaiseEvent FreezedPrimaryKeyChanging(Me, EventArgs.Empty)
            End If
            Text = Nothing
            ForeColor = UnFreezeColor
            _FreezedValue = Nothing
            _FreezedPrimaryKey = 0
            _IsFreezed = False
            If _FreezedPrimaryKey <> OldPrimaryKey Then
                RaiseEvent FreezedPrimaryKeyChanged(Me, EventArgs.Empty)
            End If
            _RawFreezedValues = New List(Of (String, String, Object))
            QueryEnabled = True
            CloseDropDown()
        End If
    End Sub
    Private Sub AutoFreeze()
        Dim OldPrimaryKey As Long
        Dim MainValue As String
        Dim OtherValue As String
        Dim FullValue As String = String.Empty
        If QueryEnabled Then
            _Freezing = True
            OldPrimaryKey = _FreezedPrimaryKey
            If DropDownResultsForm IsNot Nothing AndAlso DropDownResultsForm.DgvResults.SelectedRows.Count = 1 Then
                QueryEnabled = False
                MainValue = DropDownResultsForm.DgvResults.SelectedRows(0).Cells(If(DisplayFieldAlias = Nothing, DisplayFieldName, DisplayFieldAlias)).Value.ToString
                _RawFreezedValues.Add((DisplayTableName, DisplayFieldName, MainValue))
                If Not String.IsNullOrEmpty(MainValue) Then
                    FullValue &= Prefix & MainValue & Suffix
                End If
                For Each o As QueriedBoxField In OtherFields
                    OtherValue = DropDownResultsForm.DgvResults.SelectedRows(0).Cells(If(o.DisplayFieldAlias = Nothing, o.DisplayFieldName, o.DisplayFieldAlias)).Value.ToString
                    _RawFreezedValues.Add((o.DisplayTableName, o.DisplayFieldName, OtherValue))
                    If o.Freeze Then
                        If Not String.IsNullOrEmpty(OtherValue) Then
                            FullValue &= o.Prefix & OtherValue & o.Suffix
                        End If
                    End If
                Next o
                If OldPrimaryKey <> DropDownResultsForm.DgvResults.SelectedRows(0).Cells("mainid").Value Then
                    RaiseEvent FreezedPrimaryKeyChanging(Me, EventArgs.Empty)
                End If
                Text = FullValue
                ForeColor = FreezeColor
                _FreezedValue = FullValue
                If Not IsNumeric(DropDownResultsForm.DgvResults.SelectedRows(0).Cells("mainid").Value) Then
                    Throw New Exception("No primary key was returned, check the relationships.")
                End If
                _FreezedPrimaryKey = DropDownResultsForm.DgvResults.SelectedRows(0).Cells("mainid").Value
                _IsFreezed = True
                If _FreezedPrimaryKey <> OldPrimaryKey Then
                    RaiseEvent FreezedPrimaryKeyChanged(Me, EventArgs.Empty)
                End If
                If ShowStartOnFreeze Then
                    Me.Select(0, 0)
                Else
                    Me.Select(Me.TextLength, 0)
                End If
                _CtrlHyperlink = True
                QueryEnabled = True
                CloseDropDown()
            End If
            _Freezing = False
        End If
    End Sub
    Private Sub AutoUnfreeze()
        Dim OldPrimaryKey As Long
        If QueryEnabled Then
            OldPrimaryKey = _FreezedPrimaryKey
            QueryEnabled = False
            If ClearOnUnfreeze Then
                Text = Nothing
            End If
            If OldPrimaryKey <> 0 Then
                RaiseEvent FreezedPrimaryKeyChanging(Me, EventArgs.Empty)
            End If
            ForeColor = UnFreezeColor
            _FreezedValue = Nothing
            _IsFreezed = False
            _FreezedPrimaryKey = 0
            If _FreezedPrimaryKey <> OldPrimaryKey Then
                RaiseEvent FreezedPrimaryKeyChanged(Me, EventArgs.Empty)
            End If
            _RawFreezedValues = New List(Of (String, String, Object))
            _CtrlHyperlink = False
            QueryEnabled = True
        End If
    End Sub
    Public Function GetRawFreezedValueOf(TableName As String, FieldName As String) As Object
        Dim Match = _RawFreezedValues.Find(Function(t) t.Item1 = TableName AndAlso t.Item2 = FieldName)
        If String.IsNullOrEmpty(Match.Item1) Then
            Throw New KeyNotFoundException($"Table '{TableName}' is not configured.")
        End If
        If String.IsNullOrEmpty(Match.Item2) Then
            Throw New KeyNotFoundException($"Field '{FieldName}' of table '{TableName}' is not configured.")
        End If
        Return Match.Item3
    End Function
End Class
