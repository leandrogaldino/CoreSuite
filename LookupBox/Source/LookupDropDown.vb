Imports System.ComponentModel
Imports System.Drawing
Imports System.Reflection
Imports System.Windows.Forms
''' <summary>
''' Represents the non-public drop-down used to display asynchronous lookup results and status messages.
''' </summary>
<DesignerCategory("Code")>
Friend Class LookupDropDown
    Inherits Form
    Private Const WS_EX_NOACTIVATE As Integer = &H8000000
    Private Const WS_EX_TOOLWINDOW As Integer = &H80
    Private ReadOnly _Owner As LookupBox
    Private ReadOnly _ContainerPanel As Panel
    Private ReadOnly _ResultsGrid As DataGridView
    Private ReadOnly _TemplateList As LookupTemplateList
    Private _ActiveViewMode As LookupResultsViewMode = LookupResultsViewMode.Grid
    Private _ShowingResults As Boolean
    Private ReadOnly _StatusLabel As Label
    Private _HeaderBoldFont As Font
    Friend Event ItemActivated As EventHandler(Of LookupItemActivatedEventArgs)
    Friend Event PopupOpened As EventHandler
    Friend Event PopupClosed As EventHandler
    ''' <summary>
    ''' Gets the native window parameters used to keep the lookup text box active while the drop-down is displayed.
    ''' </summary>
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim Parameters As CreateParams = MyBase.CreateParams
            Parameters.ExStyle = Parameters.ExStyle Or WS_EX_NOACTIVATE Or WS_EX_TOOLWINDOW
            Return Parameters
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating that the popup must be shown without activating it.
    ''' </summary>
    Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
        Get
            Return True
        End Get
    End Property
    ''' <summary>
    ''' Initializes a new instance of the <see cref="LookupDropDown"/> class.
    ''' </summary>
    ''' <param name="Owner">The lookup box that owns the drop-down.</param>
    Public Sub New(Owner As LookupBox)
        If Owner Is Nothing Then Throw New ArgumentNullException("Owner")
        _Owner = Owner
        AutoScaleMode = AutoScaleMode.None
        AutoSize = False
        ControlBox = False
        FormBorderStyle = FormBorderStyle.None
        MaximizeBox = False
        MinimizeBox = False
        Padding = New Padding(1)
        ShowIcon = False
        ShowInTaskbar = False
        SizeGripStyle = SizeGripStyle.Hide
        StartPosition = FormStartPosition.Manual
        _ContainerPanel = New Panel With {.Dock = DockStyle.Fill, .Margin = System.Windows.Forms.Padding.Empty, .Padding = System.Windows.Forms.Padding.Empty}
        _ResultsGrid = CreateResultsGrid()
        _TemplateList = New LookupTemplateList(_Owner) With {.Dock = DockStyle.Fill, .Visible = False, .TabStop = False}
        _StatusLabel = New Label With {.AutoSize = False, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleCenter, .Visible = False}
        _ContainerPanel.Controls.Add(_ResultsGrid)
        _ContainerPanel.Controls.Add(_TemplateList)
        _ContainerPanel.Controls.Add(_StatusLabel)
        Controls.Add(_ContainerPanel)
        AddHandler _ResultsGrid.CellMouseClick, AddressOf ResultsGridCellMouseClick
        AddHandler _ResultsGrid.KeyDown, AddressOf ResultsGridKeyDown
        AddHandler _TemplateList.ItemActivated, AddressOf TemplateItemActivated
    End Sub
    ''' <summary>
    ''' Configures drop-down size and appearance from the owning control.
    ''' </summary>
    Friend Sub ApplyOwnerSettings()
        Dim ContentWidth As Integer = If(_Owner.DropDownWidth > 0, _Owner.DropDownWidth, Math.Max(_Owner.Width, 240)) + _Owner.DropDownExtendLeft + _Owner.DropDownExtendRight
        Dim ContentHeight As Integer = _Owner.DropDownHeight
        ClientSize = New Size(ContentWidth + Padding.Horizontal, ContentHeight + Padding.Vertical)
        BackColor = _Owner.DropDownBorderColor
        _ContainerPanel.BackColor = _Owner.DropDownBackColor
        _ResultsGrid.BackgroundColor = _Owner.DropDownBackColor
        _ResultsGrid.DefaultCellStyle.BackColor = _Owner.DropDownBackColor
        _ResultsGrid.DefaultCellStyle.ForeColor = _Owner.DropDownForeColor
        _ResultsGrid.DefaultCellStyle.SelectionBackColor = _Owner.SelectionBackColor
        _ResultsGrid.DefaultCellStyle.SelectionForeColor = _Owner.SelectionForeColor
        _ResultsGrid.ColumnHeadersVisible = _Owner.ShowColumnHeaders
        _ResultsGrid.ColumnHeadersDefaultCellStyle.BackColor = _Owner.GridHeaderBackColor
        _ResultsGrid.ColumnHeadersDefaultCellStyle.ForeColor = _Owner.GridHeaderForeColor
        _ResultsGrid.ColumnHeadersDefaultCellStyle.Font = _Owner.Font
        If _HeaderBoldFont IsNot Nothing Then
            _HeaderBoldFont.Dispose()
            _HeaderBoldFont = Nothing
        End If
        If _Owner.GridHeadersBold Then _HeaderBoldFont = New Font(_Owner.Font, _Owner.Font.Style Or FontStyle.Bold)
        If _Owner.GridHeadersBold Then _ResultsGrid.ColumnHeadersDefaultCellStyle.Font = _HeaderBoldFont
        _ResultsGrid.CellBorderStyle = If(_Owner.GridShowVerticalLines, DataGridViewCellBorderStyle.Single, DataGridViewCellBorderStyle.SingleHorizontal)
        _ResultsGrid.Font = _Owner.Font
        _StatusLabel.Font = _Owner.Font
        _StatusLabel.BackColor = _Owner.MessageBackColor
        _StatusLabel.ForeColor = _Owner.MessageForeColor
        _TemplateList.ApplyOwnerSettings()
    End Sub
    ''' <summary>
    ''' Shows or repositions the popup below the owning lookup box without taking keyboard focus.
    ''' </summary>
    Friend Sub ShowPopup()
        ApplyOwnerSettings()
        If _Owner.AutoStretchDropDown AndAlso _ShowingResults AndAlso _ActiveViewMode = LookupResultsViewMode.Grid Then
            Dim NeededWidth As Integer = Padding.Horizontal + 2
            For Each GridColumn As DataGridViewColumn In _ResultsGrid.Columns
                Dim SuggestedWidth As Integer = TextRenderer.MeasureText(GridColumn.HeaderText, _Owner.Font).Width + 24
                For Each Row As DataGridViewRow In _ResultsGrid.Rows
                    Dim Value As Object = Row.Cells(GridColumn.Index).FormattedValue
                    Dim CellText As String = If(Value Is Nothing, String.Empty, Convert.ToString(Value))
                    SuggestedWidth = Math.Max(SuggestedWidth, TextRenderer.MeasureText(CellText, _Owner.Font).Width + 24)
                Next
                NeededWidth += Math.Max(GridColumn.Width, SuggestedWidth)
            Next
            Dim Area As Rectangle = Screen.FromControl(_Owner).WorkingArea
            Width = Math.Min(Area.Width, Math.Max(Width, NeededWidth))
        End If
        Location = CalculatePopupLocation()
        If Visible Then Return
        Dim OwnerForm As Form = _Owner.FindForm()
        If OwnerForm Is Nothing Then
            Show()
        Else
            Show(OwnerForm)
        End If
        RaiseEvent PopupOpened(Me, EventArgs.Empty)
    End Sub
    ''' <summary>
    ''' Hides the popup while keeping it available for the next lookup.
    ''' </summary>
    Friend Sub HidePopup()
        If Not Visible Then Return
        Hide()
        RaiseEvent PopupClosed(Me, EventArgs.Empty)
    End Sub
    ''' <summary>
    ''' Displays a status message instead of result rows.
    ''' </summary>
    ''' <param name="Message">The message displayed in the drop-down.</param>
    Friend Sub ShowStatus(Message As String)
        _ShowingResults = False
        _ResultsGrid.Visible = False
        _TemplateList.Visible = False
        _StatusLabel.Text = If(Message, String.Empty)
        _StatusLabel.Visible = True
        _StatusLabel.BringToFront()
    End Sub
    ''' <summary>
    ''' Displays the supplied result objects using the owner column configuration.
    ''' </summary>
    ''' <param name="Results">The objects displayed in the grid.</param>
    Friend Sub ShowResults(Results As IReadOnlyList(Of Object))
        If _Owner.ResultsViewMode = LookupResultsViewMode.Template Then
            _ActiveViewMode = LookupResultsViewMode.Template
            _TemplateList.ShowItems(Results)
            _ResultsGrid.Visible = False
            _StatusLabel.Visible = False
            _TemplateList.Visible = True
            _TemplateList.BringToFront()
            _ShowingResults = True
            Return
        End If
        _ActiveViewMode = LookupResultsViewMode.Grid
        _TemplateList.Visible = False
        _ResultsGrid.SuspendLayout()
        Try
            _ResultsGrid.Rows.Clear()
            _ResultsGrid.Columns.Clear()
            Dim VisibleColumns As List(Of LookupColumn) = _Owner.GetResultColumns().Where(Function(Column) Column.Visible AndAlso Not String.IsNullOrWhiteSpace(Column.PropertyName)).ToList()
            If VisibleColumns.Count = 0 Then
                AddDefaultColumn()
            Else
                For Each Column As LookupColumn In VisibleColumns
                    AddConfiguredColumn(Column)
                Next
            End If
            For Each Item As Object In Results
                Dim RowIndex As Integer = _ResultsGrid.Rows.Add()
                Dim Row As DataGridViewRow = _ResultsGrid.Rows(RowIndex)
                Row.Tag = Item
                If VisibleColumns.Count = 0 Then
                    Row.Cells(0).Value = _Owner.GetDisplayText(Item)
                Else
                    For ColumnIndex As Integer = 0 To VisibleColumns.Count - 1
                        Row.Cells(ColumnIndex).Value = _Owner.GetMemberValue(Item, VisibleColumns(ColumnIndex).PropertyName)
                    Next
                End If
            Next
            If _ResultsGrid.Rows.Count > 0 Then
                _ResultsGrid.ClearSelection()
                _ResultsGrid.Rows(0).Selected = True
                _ResultsGrid.CurrentCell = _ResultsGrid.Rows(0).Cells(0)
            End If
            _StatusLabel.Visible = False
            _ResultsGrid.Visible = True
            _ResultsGrid.BringToFront()
            _ShowingResults = True
        Finally
            _ResultsGrid.ResumeLayout()
        End Try
    End Sub
    ''' <summary>
    ''' Moves the current grid selection by the specified number of rows.
    ''' </summary>
    ''' <param name="Offset">A positive or negative row offset.</param>
    Friend Sub MoveSelection(Offset As Integer)
        If Not _ShowingResults Then Return
        If _ActiveViewMode = LookupResultsViewMode.Template Then
            _TemplateList.MoveSelection(Offset)
            Return
        End If
        If _ResultsGrid.Rows.Count = 0 Then Return
        Dim CurrentIndex As Integer = If(_ResultsGrid.CurrentRow Is Nothing, 0, _ResultsGrid.CurrentRow.Index)
        Dim TargetIndex As Integer = Math.Max(0, Math.Min(_ResultsGrid.Rows.Count - 1, CurrentIndex + Offset))
        MoveGridSelection(TargetIndex)
    End Sub
    ''' <summary>
    ''' Gets the result object associated with the current row.
    ''' </summary>
    ''' <returns>The current result object, or <see langword="Nothing"/> when no row is selected.</returns>
    Friend Function GetSelectedItem() As Object
        If Not _ShowingResults Then Return Nothing
        If _ActiveViewMode = LookupResultsViewMode.Template Then Return _TemplateList.SelectedItem
        If _ResultsGrid.CurrentRow Is Nothing Then Return Nothing
        Return _ResultsGrid.CurrentRow.Tag
    End Function

    ''' <summary>Moves to the first or final result, in either display mode.</summary>
    Friend Sub JumpSelection(ToLast As Boolean)
        If Not _ShowingResults Then Return
        If _ActiveViewMode = LookupResultsViewMode.Template Then
            _TemplateList.JumpSelection(ToLast)
        ElseIf _ResultsGrid.Rows.Count > 0 Then
            Dim Target As Integer = If(ToLast, _ResultsGrid.Rows.Count - 1, 0)
            MoveGridSelection(Target)
        End If
    End Sub

    ''' <summary>Moves the selection by one visible page.</summary>
    Friend Sub PageSelection(Down As Boolean)
        If Not _ShowingResults Then Return
        If _ActiveViewMode = LookupResultsViewMode.Template Then
            _TemplateList.PageSelection(Down)
        ElseIf _ResultsGrid.Rows.Count > 0 Then
            Dim Current As Integer = If(_ResultsGrid.CurrentRow Is Nothing, 0, _ResultsGrid.CurrentRow.Index)
            Dim Count As Integer = Math.Max(1, _ResultsGrid.DisplayedRowCount(False))
            MoveGridSelection(Math.Max(0, Math.Min(_ResultsGrid.Rows.Count - 1, Current + If(Down, Count, -Count))))
        End If
    End Sub

    ''' <summary>Scrolls the custom-item viewport without modifying its highlight.</summary>
    Friend Sub ScrollResults(Delta As Integer)
        If _ShowingResults AndAlso _ActiveViewMode = LookupResultsViewMode.Template Then _TemplateList.ScrollByWheel(Delta)
    End Sub

    Private Sub MoveGridSelection(Index As Integer)
        _ResultsGrid.ClearSelection()
        _ResultsGrid.Rows(Index).Selected = True
        _ResultsGrid.CurrentCell = _ResultsGrid.Rows(Index).Cells(0)
        If Not _ResultsGrid.Rows(Index).Displayed Then _ResultsGrid.FirstDisplayedScrollingRowIndex = Index
    End Sub

    Private Sub TemplateItemActivated(Sender As Object, E As LookupItemActivatedEventArgs)
        RaiseEvent ItemActivated(Me, E)
    End Sub
    ''' <summary>
    ''' Releases grid event subscriptions and hosted controls.
    ''' </summary>
    ''' <param name="Disposing"><see langword="True"/> to release managed resources; otherwise, <see langword="False"/>.</param>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            _ResultsGrid.ColumnHeadersDefaultCellStyle.Font = _Owner.Font
            If _HeaderBoldFont IsNot Nothing Then _HeaderBoldFont.Dispose()
            RemoveHandler _ResultsGrid.CellMouseClick, AddressOf ResultsGridCellMouseClick
            RemoveHandler _ResultsGrid.KeyDown, AddressOf ResultsGridKeyDown
            RemoveHandler _TemplateList.ItemActivated, AddressOf TemplateItemActivated
        End If
        MyBase.Dispose(Disposing)
    End Sub
    Private Function CreateResultsGrid() As DataGridView
        Dim Grid As New DataGridView With {
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AllowUserToOrderColumns = False,
            .AllowUserToResizeRows = False,
            .AutoGenerateColumns = False,
            .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.None,
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
            .Dock = DockStyle.Fill,
            .EditMode = DataGridViewEditMode.EditProgrammatically,
            .EnableHeadersVisualStyles = False,
            .MultiSelect = False,
            .ReadOnly = True,
            .RowHeadersVisible = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .TabStop = False,
            .Visible = False
        }
        Dim DoubleBufferedProperty As PropertyInfo = GetType(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
        If DoubleBufferedProperty IsNot Nothing Then DoubleBufferedProperty.SetValue(Grid, True)
        Return Grid
    End Function
    Private Sub AddDefaultColumn()
        Dim HeaderText As String = If(String.IsNullOrWhiteSpace(_Owner.ResultColumnHeaderText), _Owner.DisplayMember, _Owner.ResultColumnHeaderText)
        If String.IsNullOrWhiteSpace(HeaderText) Then HeaderText = "Result"
        _ResultsGrid.Columns.Add(New DataGridViewTextBoxColumn With {.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, .HeaderText = HeaderText, .MinimumWidth = 5, .SortMode = DataGridViewColumnSortMode.NotSortable})
    End Sub
    Private Sub AddConfiguredColumn(Column As LookupColumn)
        Dim GridColumn As New DataGridViewTextBoxColumn With {
            .AutoSizeMode = Column.AutoSizeMode,
            .FillWeight = Column.FillWeight,
            .HeaderText = If(String.IsNullOrWhiteSpace(Column.HeaderText), Column.PropertyName, Column.HeaderText),
            .MinimumWidth = Column.MinimumWidth,
            .SortMode = DataGridViewColumnSortMode.NotSortable,
            .Width = Math.Max(Column.MinimumWidth, Column.Width)
        }
        GridColumn.DefaultCellStyle.Format = Column.Format
        GridColumn.DefaultCellStyle.NullValue = Column.NullValue
        _ResultsGrid.Columns.Add(GridColumn)
    End Sub
    Private Sub ResultsGridCellMouseClick(Sender As Object, E As DataGridViewCellMouseEventArgs)
        If E.Button <> MouseButtons.Left OrElse E.RowIndex < 0 Then Return
        RaiseEvent ItemActivated(Me, New LookupItemActivatedEventArgs(_ResultsGrid.Rows(E.RowIndex).Tag))
    End Sub
    Private Sub ResultsGridKeyDown(Sender As Object, E As KeyEventArgs)
        If E.KeyCode = Keys.Enter Then
            Dim Item As Object = GetSelectedItem()
            If Item IsNot Nothing Then RaiseEvent ItemActivated(Me, New LookupItemActivatedEventArgs(Item))
            E.Handled = True
            E.SuppressKeyPress = True
        ElseIf E.KeyCode = Keys.Escape Then
            HidePopup()
            _Owner.Focus()
            E.Handled = True
            E.SuppressKeyPress = True
        End If
    End Sub
    Protected Overrides Sub OnMouseWheel(E As MouseEventArgs)
        If _ShowingResults AndAlso _ActiveViewMode = LookupResultsViewMode.Template Then
            _TemplateList.ScrollByWheel(E.Delta)
            Dim Handled As HandledMouseEventArgs = TryCast(E, HandledMouseEventArgs)
            If Handled IsNot Nothing Then Handled.Handled = True
            Return
        End If
        MyBase.OnMouseWheel(E)
    End Sub

    Private Function CalculatePopupLocation() As Point
        Dim OwnerTopLeft As Point = _Owner.PointToScreen(Point.Empty)
        Dim BelowLocation As Point = _Owner.PointToScreen(New Point(0, _Owner.Height))
        Dim WorkingArea As Rectangle = Screen.FromControl(_Owner).WorkingArea
        Dim PopupX As Integer = If(_Owner.RightToLeft = RightToLeft.Yes, OwnerTopLeft.X + _Owner.Width - Width, BelowLocation.X - _Owner.DropDownExtendLeft)
        Dim PopupY As Integer = BelowLocation.Y
        If PopupY + Height > WorkingArea.Bottom AndAlso OwnerTopLeft.Y - Height >= WorkingArea.Top Then PopupY = OwnerTopLeft.Y - Height
        PopupX = Math.Max(WorkingArea.Left, Math.Min(PopupX, WorkingArea.Right - Width))
        PopupY = Math.Max(WorkingArea.Top, Math.Min(PopupY, WorkingArea.Bottom - Height))
        Return New Point(PopupX, PopupY)
    End Function
End Class
