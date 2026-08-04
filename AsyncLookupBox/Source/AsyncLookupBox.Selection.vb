Imports System.ComponentModel
Imports System.Globalization
Partial Public Class AsyncLookupBox
    Friend Function GetDisplayText(Item As Object) As String
        If Item Is Nothing Then Return String.Empty
        Dim DisplayValue As Object = If(String.IsNullOrWhiteSpace(DisplayMember), Item, GetMemberValue(Item, DisplayMember))
        If DisplayValue Is Nothing OrElse Convert.IsDBNull(DisplayValue) Then Return String.Empty
        Return Convert.ToString(DisplayValue, CultureInfo.CurrentCulture)
    End Function
    Friend Function GetMemberValue(Item As Object, PropertyPath As String) As Object
        If Item Is Nothing Then Return Nothing
        If String.IsNullOrWhiteSpace(PropertyPath) Then Return Item
        Dim CurrentValue As Object = Item
        For Each MemberName As String In PropertyPath.Split("."c)
            If CurrentValue Is Nothing OrElse Convert.IsDBNull(CurrentValue) Then Return Nothing
            Dim TrimmedMemberName As String = MemberName.Trim()
            If TrimmedMemberName.Length = 0 Then Return Nothing
            Dim RowView As DataRowView = TryCast(CurrentValue, DataRowView)
            If RowView IsNot Nothing Then
                If Not RowView.DataView.Table.Columns.Contains(TrimmedMemberName) Then Return Nothing
                CurrentValue = RowView(TrimmedMemberName)
                Continue For
            End If
            Dim Row As DataRow = TryCast(CurrentValue, DataRow)
            If Row IsNot Nothing Then
                If Not Row.Table.Columns.Contains(TrimmedMemberName) Then Return Nothing
                CurrentValue = Row(TrimmedMemberName)
                Continue For
            End If
            Dim Dictionary As IDictionary = TryCast(CurrentValue, IDictionary)
            If Dictionary IsNot Nothing Then
                Dim DictionaryKey As Object = FindDictionaryKey(Dictionary, TrimmedMemberName)
                If DictionaryKey Is Nothing Then Return Nothing
                CurrentValue = Dictionary(DictionaryKey)
                Continue For
            End If
            Dim Descriptor As PropertyDescriptor = TypeDescriptor.GetProperties(CurrentValue).Find(TrimmedMemberName, True)
            If Descriptor Is Nothing Then Return Nothing
            CurrentValue = Descriptor.GetValue(CurrentValue)
        Next
        Return CurrentValue
    End Function
    Private Function ResolveValue(Item As Object) As Object
        If Item Is Nothing Then Return Nothing
        If String.IsNullOrWhiteSpace(ValueMember) Then Return Item
        Return GetMemberValue(Item, ValueMember)
    End Function
    Private Sub SelectItemCore(Item As Object)
        If Item Is Nothing Then Return
        Dim OldItem As Object = _SelectedItem
        Dim OldValue As Object = _SelectedValue
        _SelectedItem = Item
        _SelectedValue = ResolveValue(Item)
        CancelActiveSearch()
        SetTextWithoutSearching(GetDisplayText(Item))
        ApplySelectionAppearance()
        SelectionStart = TextLength
        SelectionLength = 0
        CloseDropDown()
        If Not ReferenceEquals(OldItem, _SelectedItem) OrElse Not Object.Equals(OldValue, _SelectedValue) Then RaiseEvent SelectionChanged(Me, New AsyncLookupSelectionChangedEventArgs(OldItem, _SelectedItem, OldValue, _SelectedValue))
    End Sub
    Private Sub ClearSelectionCore(ClearText As Boolean)
        Dim OldItem As Object = _SelectedItem
        Dim OldValue As Object = _SelectedValue
        _SelectedItem = Nothing
        _SelectedValue = Nothing
        CancelActiveSearch()
        _Results = Array.Empty(Of Object)()
        If ClearText Then SetTextWithoutSearching(String.Empty)
        ApplySelectionAppearance()
        CloseDropDown()
        UpdateActionButtonLayout()
        If OldItem IsNot Nothing OrElse OldValue IsNot Nothing Then RaiseEvent SelectionChanged(Me, New AsyncLookupSelectionChangedEventArgs(OldItem, Nothing, OldValue, Nothing))
    End Sub
    Private Sub SetTextWithoutSearching(Value As String)
        _SuppressTextChanged = True
        Try
            Text = If(Value, String.Empty)
        Finally
            _SuppressTextChanged = False
        End Try
        UpdateActionButtonLayout()
    End Sub
    Private Shared Function FindDictionaryKey(Dictionary As IDictionary, MemberName As String) As Object
        If Dictionary.Contains(MemberName) Then Return MemberName
        For Each Entry As DictionaryEntry In Dictionary
            If TypeOf Entry.Key Is String AndAlso String.Equals(CStr(Entry.Key), MemberName, StringComparison.OrdinalIgnoreCase) Then Return Entry.Key
        Next
        Return Nothing
    End Function
End Class
