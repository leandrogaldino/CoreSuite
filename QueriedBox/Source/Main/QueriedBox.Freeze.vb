Partial Public Class QueriedBox
    ''' <summary>
    ''' Freezes the control value using the specified primary key and retrieves the corresponding query data.
    ''' </summary>
    ''' <param name="PrimaryKeyValue">The primary key value used to locate the record.</param>
    ''' <param name="PrimaryKeyColumnName">The name of the primary key column used in the query condition.</param>
    Public Sub Freeze(PrimaryKeyValue As Object, PrimaryKeyColumnName As String)
        Dim FullQuery As String
        Dim ParameterList As Dictionary(Of String, Object)
        Dim TableResults As DataTable
        Dim OldPrimaryKey As Object = Frozen.FrozenPrimaryKey
        Dim FullValue As String = String.Empty
        Dim ColumnName As String
        Dim PrimaryKeyParameter As String = "@" & Guid.NewGuid.ToString("N")
        If Search.Enabled Then
            If String.IsNullOrWhiteSpace(PrimaryKeyColumnName) Then
                Throw New ArgumentException("The primary key column name cannot be null, empty, or whitespace.", NameOf(PrimaryKeyColumnName))
            End If
            _Freezing = True
            Try
                _PrimaryKeyAlias = Nothing
                FullQuery = Query.ToStringWithAdditionalWhereCondition("#Placeholder")
                FullQuery = FullQuery.Replace("#Placeholder", $"{PrimaryKeyColumnName} = {PrimaryKeyParameter}")
                ParameterList = New Dictionary(Of String, Object) From {{PrimaryKeyParameter, PrimaryKeyValue}}
                For Each p As QueryParameter In Query.Parameters
                    ParameterList.Add(p.ParameterName, p.ParameterValue)
                Next p
                TableResults = ExecuteQuery(FullQuery, ParameterList)
                If TableResults.Rows.Count = 1 Then
                    Search.Enabled = False
                    _RawFrozenValues.Clear()
                    For Each Column In Query.Columns
                        ColumnName = Column.ColumnAlias
                        Dim Value As Object = TableResults.Rows(0).Item(ColumnName)
                        Dim RawValue As Object = If(Value Is Nothing OrElse IsDBNull(Value), Nothing, Value)
                        Dim DisplayValue As String = If(RawValue Is Nothing, String.Empty, RawValue.ToString())
                        _RawFrozenValues.Add((ColumnName, RawValue))
                        If Not String.IsNullOrEmpty(DisplayValue) AndAlso Column.Options.Freeze Then
                            FullValue &= $"{Column.Options.Prefix}{DisplayValue}{Column.Options.Suffix}"
                        End If
                    Next Column
                    If Not Equals(OldPrimaryKey, PrimaryKeyValue) Then
                        OnFrozenPrimaryKeyChanging(OldPrimaryKey, PrimaryKeyValue)
                    End If
                    Text = FullValue
                    ForeColor = frozen.FrozenColor
                    Frozen.SetFrozenValue(FullValue)
                    Frozen.SetFrozenPrimaryKey(PrimaryKeyValue)
                    Frozen.SetIsFrozen(True)
                    If Not Equals(OldPrimaryKey, PrimaryKeyValue) Then
                        OnFrozenPrimaryKeyChanged(OldPrimaryKey, PrimaryKeyValue)
                    End If
                    If Frozen.ShowStartOnFreeze Then
                        Me.Select(0, 0)
                    Else
                        Me.Select(Me.TextLength, 0)
                    End If
                    _CtrlHyperlink = True
                    CloseDropDown()
                Else
                    Unfreeze()
                End If
            Finally
                Search.Enabled = True
                _Freezing = False
            End Try
        End If
    End Sub
    ''' <summary>
    ''' Removes the current frozen value from the control and clears the associated primary key.
    ''' </summary>
    Public Sub Unfreeze()
        Dim OldPrimaryKey As Object = Frozen.FrozenPrimaryKey
        If Search.Enabled Then
            Search.Enabled = False
            Try
                If Not Equals(OldPrimaryKey, Nothing) Then
                    OnFrozenPrimaryKeyChanging(OldPrimaryKey, Nothing)
                End If
                Text = Nothing
                ForeColor = Frozen.UnFrozenColor
                Frozen.SetFrozenValue(Nothing)
                Frozen.SetFrozenPrimaryKey(Nothing)
                Frozen.SetIsFrozen(False)
                _RawFrozenValues.Clear()
                _PrimaryKeyAlias = Nothing
                If Not Equals(Frozen.FrozenPrimaryKey, OldPrimaryKey) Then
                    OnFrozenPrimaryKeyChanged(OldPrimaryKey, Nothing)
                End If
                CloseDropDown()
            Finally
                Search.Enabled = True
            End Try
        End If
    End Sub
    ''' <summary>
    ''' Freezes the control using the currently selected row from the results dropdown.
    ''' </summary>
    Friend Sub AutoFreeze()
        Dim FrozenPrimaryKey As Object
        Dim OldPrimaryKey As Object
        Dim FullValue As String = String.Empty
        Dim ColumnName As String
        If Search.Enabled Then
            _Freezing = True
            Try
                OldPrimaryKey = Frozen.FrozenPrimaryKey
                If DropDownResultsForm IsNot Nothing AndAlso DropDownResultsForm.DgvResults.SelectedRows.Count = 1 Then
                    Search.Enabled = False
                    _RawFrozenValues.Clear()
                    For Each Column In Query.Columns
                        ColumnName = Column.ColumnAlias
                        Dim Value As Object = DropDownResultsForm.DgvResults.SelectedRows(0).Cells(ColumnName).Value
                        Dim RawValue As Object = If(Value Is Nothing OrElse IsDBNull(Value), Nothing, Value)
                        Dim DisplayValue As String = If(RawValue Is Nothing, String.Empty, RawValue.ToString())
                        _RawFrozenValues.Add((ColumnName, RawValue))
                        If Not String.IsNullOrEmpty(DisplayValue) AndAlso Column.Options.Freeze Then
                            FullValue &= $"{Column.Options.Prefix}{DisplayValue}{Column.Options.Suffix}"
                        End If
                    Next Column
                    FrozenPrimaryKey = DropDownResultsForm.DgvResults.SelectedRows(0).Cells(_PrimaryKeyAlias).Value
                    If FrozenPrimaryKey Is Nothing OrElse IsDBNull(FrozenPrimaryKey) Then
                        Throw New InvalidOperationException("No primary key was returned. Check the query configuration.")
                    End If
                    If Not Object.Equals(OldPrimaryKey, FrozenPrimaryKey) Then
                        OnFrozenPrimaryKeyChanging(OldPrimaryKey, FrozenPrimaryKey)
                    End If
                    Text = FullValue
                    ForeColor = Frozen.FrozenColor
                    Frozen.SetFrozenValue(FullValue)
                    Frozen.SetFrozenPrimaryKey(FrozenPrimaryKey)
                    Frozen.SetIsFrozen(True)
                    If Not Object.Equals(OldPrimaryKey, FrozenPrimaryKey) Then
                        OnFrozenPrimaryKeyChanged(OldPrimaryKey, FrozenPrimaryKey)
                    End If
                    If Frozen.ShowStartOnFreeze Then
                        Me.Select(0, 0)
                    Else
                        Me.Select(TextLength, 0)
                    End If
                    _CtrlHyperlink = True
                    CloseDropDown()
                End If
            Finally
                Search.Enabled = True
                _Freezing = False
            End Try
        End If
    End Sub
    ''' <summary>
    ''' Removes the frozen state automatically and restores the control to its editable state.
    ''' </summary>
    Private Sub AutoUnfreeze()
        Dim OldPrimaryKey As Object = Frozen.FrozenPrimaryKey
        If Search.Enabled Then
            Search.Enabled = False
            Try
                If Frozen.ClearOnUnfreeze Then
                    Text = Nothing
                End If
                If OldPrimaryKey IsNot Nothing Then
                    OnFrozenPrimaryKeyChanging(OldPrimaryKey, Nothing)
                End If
                ForeColor = Frozen.UnFrozenColor
                Frozen.SetFrozenValue(Nothing)
                Frozen.SetFrozenPrimaryKey(Nothing)
                Frozen.SetIsFrozen(False)
                _PrimaryKeyAlias = Nothing
                If Not Equals(Frozen.FrozenPrimaryKey, OldPrimaryKey) Then
                    OnFrozenPrimaryKeyChanged(OldPrimaryKey, Nothing)
                End If
                _RawFrozenValues.Clear()
                _CtrlHyperlink = False
            Finally
                Search.Enabled = True
            End Try
        End If
    End Sub
    ''' <summary>
    ''' Freezes the selected result automatically when the current text matches a result item.
    ''' </summary>
    Friend Sub AutoFreezeIfMatched()
        If DropDownResultsForm IsNot Nothing AndAlso DropDownResultsForm.DgvResults.SelectedRows.Count = 1 Then
            Dim Row As DataGridViewRow = DropDownResultsForm.DgvResults.SelectedRows(0)
            Dim Matched As Boolean = Query.Columns.Any(Function(Column)
                                                           Dim ColumnName = If(String.IsNullOrWhiteSpace(Column.ColumnAlias), Column.ColumnName, Column.ColumnAlias)
                                                           Dim Value = Row.Cells(ColumnName).Value?.ToString()
                                                           Return String.Equals(Text, Value, StringComparison.OrdinalIgnoreCase)
                                                       End Function)
            If Matched Then
                AutoFreeze()
            End If
        End If
        CloseDropDown()
    End Sub
    ''' <summary>
    ''' Gets the raw frozen value stored for the specified column.
    ''' </summary>
    ''' <param name="ColumnAlias">The alias of the column whose frozen value should be retrieved.</param>
    ''' <returns>The stored raw value associated with the specified column.</returns>
    ''' <exception cref="KeyNotFoundException">Thrown when the specified column is not configured or was not frozen.</exception>
    Public Function GetRawFrozenValueOf(ColumnAlias As String) As Object
        Dim Match = _RawFrozenValues.Find(Function(t) t.Item1 = ColumnAlias)
        If String.IsNullOrEmpty(Match.Item1) Then
            Throw New KeyNotFoundException($"Column '{ColumnAlias}' was not found in the control.")
        End If
        Return Match.Item2
    End Function
End Class
