Imports System.ComponentModel
Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading
Imports System.Drawing
Imports System.Windows.Forms


Partial Public Class LookupBox
    ''' <summary>When selected text is edited, clears the text as well as unselecting the record.</summary>
    <Category("LookupBox - Selection")>
    <DefaultValue(False)>
    Public Property ClearOnEdit As Boolean
        Get
            Return _ClearOnEdit
        End Get
        Set(value As Boolean)
            _ClearOnEdit = value
        End Set
    End Property

    ''' <summary>When focus leaves, selects an item only if exactly one existing result has matching text.</summary>
    <Category("LookupBox - Selection")>
    <DefaultValue(True)>
    Public Property AutoSelectMatchedResultOnLeave As Boolean
        Get
            Return _AutoSelectMatchedResultOnLeave
        End Get
        Set(value As Boolean)
            _AutoSelectMatchedResultOnLeave = value
        End Set
    End Property

    <Category("LookupBox - Selection")>
    <DefaultValue(True)>
    Public Property SelectFromStartOnLeave As Boolean
        Get
            Return _SelectFromStartOnLeave
        End Get
        Set(value As Boolean)
            _SelectFromStartOnLeave = value
        End Set
    End Property

    <Category("LookupBox - Selection")>
    <DefaultValue(False)>
    Public Property ShowStartOnSelect As Boolean
        Get
            Return _ShowStartOnSelect
        End Get
        Set(value As Boolean)
            _ShowStartOnSelect = value
        End Set
    End Property

    <Category("LookupBox - Selection")>
    <DefaultValue(False)>
    Public Property AllowSelectionLink As Boolean
        Get
            Return _AllowSelectionLink
        End Get
        Set(value As Boolean)
            _AllowSelectionLink = value
            If Not value Then ResetSelectionLinkAppearance()
        End Set
    End Property

    ''' <summary>Additional width at the left side of the drop-down.</summary>
    <Category("LookupBox - DropDown")>
    <DefaultValue(0)>
    Public Property DropDownExtendLeft As Integer
        Get
            Return _DropDownExtendLeft
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value))
            _DropDownExtendLeft = value
            RefreshDropDownAppearance()
        End Set
    End Property

    <Category("LookupBox - DropDown")>
    <DefaultValue(0)>
    Public Property DropDownExtendRight As Integer
        Get
            Return _DropDownExtendRight
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value))
            _DropDownExtendRight = value
            RefreshDropDownAppearance()
        End Set
    End Property

    <Category("LookupBox - DropDown")>
    <DefaultValue(False)>
    Public Property AutoStretchDropDown As Boolean
        Get
            Return _AutoStretchDropDown
        End Get
        Set(value As Boolean)
            _AutoStretchDropDown = value
            RefreshDropDownAppearance()
        End Set
    End Property

    <Category("LookupBox - Appearance")>
    <DefaultValue(GetType(Color), "Control")>
    Public Property GridHeaderBackColor As Color
        Get
            Return _GridHeaderBackColor
        End Get
        Set(value As Color)
            _GridHeaderBackColor = value
            RefreshDropDownAppearance()
        End Set
    End Property

    <Category("LookupBox - Appearance")>
    <DefaultValue(GetType(Color), "ControlText")>
    Public Property GridHeaderForeColor As Color
        Get
            Return _GridHeaderForeColor
        End Get
        Set(value As Color)
            _GridHeaderForeColor = value
            RefreshDropDownAppearance()
        End Set
    End Property

    <Category("LookupBox - Appearance")>
    <DefaultValue(False)>
    Public Property GridHeadersBold As Boolean
        Get
            Return _GridHeadersBold
        End Get
        Set(value As Boolean)
            _GridHeadersBold = value
            RefreshDropDownAppearance()
        End Set
    End Property

    <Category("LookupBox - Appearance")>
    <DefaultValue(True)>
    Public Property GridShowVerticalLines As Boolean
        Get
            Return _GridShowVerticalLines
        End Get
        Set(value As Boolean)
            _GridShowVerticalLines = value
            RefreshDropDownAppearance()
        End Set
    End Property

    <Category("LookupBox - Appearance")>
    <DefaultValue(GetType(Color), "White")>
    Public Property MessageBackColor As Color
        Get
            Return _MessageBackColor
        End Get
        Set(value As Color)
            _MessageBackColor = value
            RefreshDropDownAppearance()
        End Set
    End Property

    <Category("LookupBox - Appearance")>
    <DefaultValue(GetType(Color), "ControlText")>
    Public Property MessageForeColor As Color
        Get
            Return _MessageForeColor
        End Get
        Set(value As Color)
            _MessageForeColor = value
            RefreshDropDownAppearance()
        End Set
    End Property

    ''' <summary>Gets a property or dictionary field from the currently selected object, or Nothing.</summary>
    Public Function GetSelectedMemberValue(PropertyPath As String) As Object
        If Not HasSelection Then Return Nothing
        Dim RawValue As Object = GetMemberValue(_SelectedItem, PropertyPath)
        Return If(Convert.IsDBNull(RawValue), Nothing, RawValue)
    End Function

    ''' <summary>Resolves a record by key through the synchronous resolver, independently of search results.</summary>
    ''' <param name="Value">Identifier to resolve through ValueMember.</param>
    ''' <returns>True if a matching record was resolved and selected.</returns>
    ''' <exception cref="InvalidOperationException">ResolveItem is not configured.</exception>
    Public Function SelectByValue(Value As Object) As Boolean
        If Value Is Nothing OrElse Convert.IsDBNull(Value) OrElse _IsDisposing OrElse IsDisposed Then Return False
        Dim Resolver As Func(Of Object, Object) = ResolveItem
        If Resolver Is Nothing Then Throw New InvalidOperationException("ResolveItem is not configured. Assign a synchronous resolver or use SelectByValueAsync with ResolveItemAsync.")

        ' Each selection attempt replaces any pending asynchronous selection.
        _SelectionResolveVersion += 1
        Dim RequestVersion As Long = _SelectionResolveVersion
        Dim CurrentSearchVersion As Long = _SearchVersion
        Dim CurrentTextVersion As Long = _TextVersion
        Dim Item As Object = Resolver(Value)
        If _IsDisposing OrElse IsDisposed OrElse RequestVersion <> _SelectionResolveVersion OrElse CurrentSearchVersion <> _SearchVersion OrElse CurrentTextVersion <> _TextVersion Then Return False
        If Item Is Nothing OrElse Convert.IsDBNull(Item) OrElse Not Object.Equals(ResolveValue(Item), Value) Then Return False
        SelectItemCore(Item)
        Return True
    End Function

    ''' <summary>Application-provided synchronous resolver; receives the key and returns the corresponding record, or Nothing.</summary>
    ''' <remarks>Use for fast in-memory lookups. Long synchronous database operations can block the WinForms UI thread.</remarks>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property ResolveItem As Func(Of Object, Object)

    ''' <summary>Application-provided asynchronous resolver; receives the key and a cancellation token, then returns a record or Nothing.</summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property ResolveItemAsync As Func(Of Object, CancellationToken, Task(Of Object))

    ''' <summary>Resolves and selects a record using ResolveItemAsync, independently of previous dropdown searches.</summary>
    ''' <param name="Value">Identifier corresponding to ValueMember.</param>
    ''' <param name="CancellationToken">Cancellation token passed to the application's resolver.</param>
    ''' <returns>True only if the resolved item's key matches Value and no newer operation superseded the resolution.</returns>
    ''' <exception cref="InvalidOperationException">Thrown when ResolveItemAsync is not configured or returns Nothing instead of a Task.</exception>
    Public Async Function SelectByValueAsync(Value As Object, Optional CancellationToken As CancellationToken = Nothing) As Task(Of Boolean)
        If Value Is Nothing OrElse Convert.IsDBNull(Value) OrElse _IsDisposing OrElse IsDisposed Then Return False
        CancellationToken.ThrowIfCancellationRequested()
        Dim Resolver As Func(Of Object, CancellationToken, Task(Of Object)) = ResolveItemAsync
        If Resolver Is Nothing Then Throw New InvalidOperationException("ResolveItemAsync is not configured. Assign an asynchronous resolver or use SelectByValue with ResolveItem.")

        _SelectionResolveVersion += 1
        Dim RequestVersion As Long = _SelectionResolveVersion
        Dim CurrentSearchVersion As Long = _SearchVersion
        Dim CurrentTextVersion As Long = _TextVersion
        Dim ResolutionTask As Task(Of Object) = Resolver(Value, CancellationToken)
        If ResolutionTask Is Nothing Then Throw New InvalidOperationException("ResolveItemAsync returned Nothing; it must return a Task(Of Object).")
        Dim Item As Object = Await ResolutionTask
        CancellationToken.ThrowIfCancellationRequested()
        If _IsDisposing OrElse IsDisposed OrElse RequestVersion <> _SelectionResolveVersion OrElse CurrentSearchVersion <> _SearchVersion OrElse CurrentTextVersion <> _TextVersion Then Return False
        If Item Is Nothing OrElse Convert.IsDBNull(Item) OrElse Not Object.Equals(ResolveValue(Item), Value) Then Return False
        SelectItemCore(Item)
        Return True
    End Function

    Friend Function GetResultColumns() As IEnumerable(Of LookupColumn)
        Return Columns
    End Function

    Private Function GetCompositeDisplayText(Item As Object) As String
        ' Explicit composite columns take precedence over DisplayMember.
        ' DisplayMember remains the fallback when no columns are opted in.
        Dim Parts As New List(Of String)
        Dim CompositeColumns As List(Of LookupColumn) = Columns.Where(Function(C) C.IncludeInSelectedText).ToList()
        If CompositeColumns.Count > 0 Then
            For Each Column As LookupColumn In CompositeColumns
                Dim Raw As Object = GetMemberValue(Item, Column.PropertyName)
                If Raw Is Nothing OrElse Convert.IsDBNull(Raw) Then Continue For
                Dim Value As String = Convert.ToString(Raw, Globalization.CultureInfo.CurrentCulture)
                If Value.Length > 0 Then Parts.Add(Column.SelectedTextPrefix & Value & Column.SelectedTextSuffix)
            Next
            Return String.Concat(Parts)
        End If
        Return Nothing
    End Function

    Private Sub SelectUniqueMatchingResult()
        If _Results Is Nothing OrElse _Results.Count = 0 OrElse String.IsNullOrWhiteSpace(Text) Then Return
        Dim TypedText As String = Text.Trim()
        Dim Matches As New List(Of Object)
        For Each Item As Object In _Results
            Dim IsMatch As Boolean = String.Equals(GetDisplayText(Item).Trim(), TypedText, StringComparison.CurrentCultureIgnoreCase)
            If Not IsMatch Then
                For Each Column As LookupColumn In GetResultColumns()
                    Dim Raw As Object = GetMemberValue(Item, Column.PropertyName)
                    If Raw IsNot Nothing AndAlso Not Convert.IsDBNull(Raw) AndAlso String.Equals(Convert.ToString(Raw, Globalization.CultureInfo.CurrentCulture).Trim(), TypedText, StringComparison.CurrentCultureIgnoreCase) Then
                        IsMatch = True
                        Exit For
                    End If
                Next
            End If
            If IsMatch Then Matches.Add(Item)
            If Matches.Count > 1 Then Exit For
        Next
        If Matches.Count = 1 Then SelectItemCore(Matches(0))
    End Sub

    Private Sub ApplySelectionLinkAppearance()
        If _NormalLinkFont IsNot Nothing OrElse Not AllowSelectionLink OrElse Not HasSelection Then Return
        _NormalLinkFont = Font
        _LinkFont = New Font(Font, Font.Style Or FontStyle.Underline)
        Font = _LinkFont
        Cursor = Cursors.Hand
    End Sub

    Private Sub ResetSelectionLinkAppearance()
        If _NormalLinkFont Is Nothing Then Return
        Dim Original As Font = _NormalLinkFont
        _NormalLinkFont = Nothing
        Font = Original
        Cursor = Cursors.IBeam
        If _LinkFont IsNot Nothing Then
            _LinkFont.Dispose()
            _LinkFont = Nothing
        End If
    End Sub
End Class
