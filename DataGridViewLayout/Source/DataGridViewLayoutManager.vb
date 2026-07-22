Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports System.Text.Json
''' <summary>
''' Manages DataGridView layout persistence, including column visibility, order,
''' formatting, sorting state, and restoration using JSON storage.
''' </summary>
''' <remarks>
''' This class provides automatic handling for:
''' - Column visibility and display order
''' - Column width and autosize mode
''' - Cell and header formatting
''' - Sorting state persistence
''' - Layout saving and restoration via JSON files
''' </remarks>
Public Class DataGridViewLayoutManager
    Inherits Component
    ''' <summary>
    ''' Occurs when the layout has been successfully loaded and applied.
    ''' </summary>
    Public Event Loaded(sender As Object, e As EventArgs)
    Private Shared ReadOnly JsonOptions As New JsonSerializerOptions With {.WriteIndented = True}
    Private _SaveLayout As Boolean
    Private _LoadLayout As Boolean
    Private _CmsPoint As Point
    Private _DataGridView As New DataGridView
    Private _CurrentLayout As DataGridViewLayout
    Private ReadOnly _CmsColumns As New ContextMenuStrip
    ''' <summary>
    ''' Directory where layout JSON files are stored.
    ''' </summary>
    Public Shared Property LayoutDirectory As String
    ''' <summary>
    ''' Default font used in the column context menu.
    ''' </summary>
    Public Shared Property Font As Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Regular)
    ''' <summary>
    ''' Default layout used when creating or resetting a layout file.
    ''' </summary>
    <Browsable(False)>
    Public Property DefaultLayout As DataGridViewLayout
    ''' <summary>
    ''' Gets or sets the DataGridView managed by this layout manager.
    ''' Attaching a DataGridView automatically hooks required events.
    ''' </summary>
    Public Property DataGridView As DataGridView
        Get
            Return _DataGridView
        End Get
        Set(value As DataGridView)
            If _DataGridView IsNot Nothing Then
                RemoveHandler DataGridView.MouseDown, AddressOf DataGridView_MouseDown
                RemoveHandler DataGridView.MouseUp, AddressOf DataGridView_MouseUp
                RemoveHandler _CmsColumns.Closing, AddressOf CmsColumns_Closing
            End If

            _DataGridView = value

            AddHandler DataGridView.MouseDown, AddressOf DataGridView_MouseDown
            AddHandler DataGridView.MouseUp, AddressOf DataGridView_MouseUp
            AddHandler _CmsColumns.Closing, AddressOf CmsColumns_Closing
        End Set
    End Property
    ''' <summary>
    ''' Creates the initial JSON layout file using the default model.
    ''' </summary>
    Private Sub CreateJsonFile()
        Dim Json = JsonSerializer.Serialize(DefaultLayout, JsonOptions)
        File.WriteAllText(
        Path.Combine(LayoutDirectory, $"{DefaultLayout.Routine}.json"), Json, Encoding.UTF8)
        _CurrentLayout = DefaultLayout.Clone()
    End Sub

    ''' <summary>
    ''' Saves the current layout state into a JSON file.
    ''' </summary>
    Private Sub SaveJsonFile()
        Dim Json = JsonSerializer.Serialize(_CurrentLayout, JsonOptions)
        File.WriteAllText(
        Path.Combine(LayoutDirectory, $"{DefaultLayout.Routine}.json"), Json, Encoding.UTF8)
    End Sub
    ''' <summary>
    ''' Reads the layout JSON file and deserializes it into a layout object.
    ''' </summary>
    ''' <returns>The deserialized DataGridView layout.</returns>
    Private Function ReadJsonFile() As DataGridViewLayout
        Dim Json = File.ReadAllText(Path.Combine(LayoutDirectory, $"{DefaultLayout.Routine}.json"), Encoding.UTF8)
        Return JsonSerializer.Deserialize(Of DataGridViewLayout)(Json)
    End Function
    ''' <summary>
    ''' Clears any sorting applied to the DataGridView and its underlying data source.
    ''' </summary>
    Private Sub ClearSort()
        For Each Column As DataGridViewColumn In _DataGridView.Columns
            Column.HeaderCell.SortGlyphDirection = SortOrder.None
        Next

        Dim BindingSource = TryCast(_DataGridView.DataSource, BindingSource)
        If BindingSource IsNot Nothing Then
            BindingSource.RemoveSort()
            Exit Sub
        End If

        Dim DataView As DataView = TryCast(_DataGridView.DataSource, DataView)
        If DataView IsNot Nothing Then
            DataView.Sort = ""
            Exit Sub
        End If

        Dim DataTable = TryCast(_DataGridView.DataSource, DataTable)
        If DataTable IsNot Nothing Then
            DataTable.DefaultView.Sort = ""
            Exit Sub
        End If
    End Sub
    ''' <summary>
    ''' Loads the saved layout and applies all configuration to the DataGridView.
    ''' </summary>
    Public Sub Load()
        Dim Index As Integer
        Dim Button As ToolStripMenuItem
        Dim SelectedRow As Long = 0
        Dim FirstRow As Long = 0
        If DefaultLayout Is Nothing Then
            Throw New InvalidOperationException("DefaultLayout must be defined before calling Load().")
        End If
        _CmsColumns.Font = Font
        If Not File.Exists(LayoutDirectory & "\" & DefaultLayout.Routine & ".json") OrElse HasNewVersion() Then
            CreateJsonFile()
        End If
        _CurrentLayout = ReadJsonFile()
        For Each Column In _CurrentLayout.Columns
            Index = _CurrentLayout.Columns.IndexOf(Column)
            _DataGridView.Columns(Index).Visible = Column.VisibleInGrid
            _DataGridView.Columns(Index).DisplayIndex = Column.DisplayIndex
            _DataGridView.Columns(Index).HeaderText = Column.DisplayName
            If Column.WidthSizeMode = DataGridViewAutoSizeColumnMode.None Then
                _DataGridView.Columns(Index).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                _DataGridView.Columns(Index).Width = Column.Width
            Else
                _DataGridView.Columns(Index).AutoSizeMode = Column.WidthSizeMode
            End If
            _DataGridView.Columns(Index).DefaultCellStyle.Alignment = Column.CellAlignment
            _DataGridView.Columns(Index).HeaderCell.Style.Alignment = Column.HeaderAlignment
            _DataGridView.Columns(Index).DefaultCellStyle.Format = Column.Format
        Next Column
        If _DataGridView IsNot Nothing Then
            If _DataGridView.SelectedRows.Count = 1 Then
                SelectedRow = _DataGridView.SelectedRows(0).Index
            End If
            FirstRow = _DataGridView.FirstDisplayedScrollingRowIndex
        End If
        If _CurrentLayout.SortDirection = SortOrder.Ascending Then
            _DataGridView.Sort(_DataGridView.Columns(_CurrentLayout.SortedColumn), 0)
        End If
        If _CurrentLayout.SortDirection = SortOrder.Descending Then
            _DataGridView.Sort(_DataGridView.Columns(_CurrentLayout.SortedColumn), 1)
        End If
        If _CurrentLayout.SortDirection = SortOrder.None Then ClearSort()
        If _DataGridView.Rows.Count > 0 Then
            If SelectedRow < _DataGridView.Rows.Count Then
                _DataGridView.Rows(SelectedRow).Selected = True
            Else
                _DataGridView.Rows(_DataGridView.Rows.Count - 1).Selected = True
            End If
            If FirstRow >= 0 Then
                If _DataGridView.Rows.Count >= FirstRow Then
                    _DataGridView.FirstDisplayedScrollingRowIndex = FirstRow
                Else
                    _DataGridView.FirstDisplayedScrollingRowIndex = _DataGridView.Rows.Count - 1
                End If
            End If
        End If
        _CmsColumns.Items.Clear()
        Button = New ToolStripMenuItem With {
            .Name = "BtnRestoreGridLayout",
            .Text = "Restaurar Colunas"
        }
        AddHandler Button.Click, AddressOf BtnRestoreGridLayout_Click
        _CmsColumns.Items.Add(Button)

        Button = New ToolStripMenuItem With {
            .Name = "BtnRemoveSort",
            .Text = "Remover Classificação"
        }
        AddHandler Button.Click, AddressOf BtnRemoveSort_Click
        _CmsColumns.Items.Add(Button)
        _CmsColumns.Items.Add(New ToolStripSeparator)
        For Each Column In _DataGridView.Columns
            Dim LayoutColumn As DataGridViewLayoutColumn =
                _CurrentLayout.Columns.FirstOrDefault(Function(c) c.DisplayName = Column.HeaderText)
            If LayoutColumn.VisibleInContext Then
                Button = New ToolStripMenuItem With {
                    .CheckOnClick = True,
                    .Text = LayoutColumn.DisplayName,
                    .Checked = LayoutColumn.VisibleInGrid
                }
                AddHandler Button.CheckedChanged, AddressOf BtnColumn_CheckedChanged
                _CmsColumns.Items.Add(Button)
            End If
        Next Column
        RaiseEvent Loaded(Me, EventArgs.Empty)
    End Sub
    ''' <summary>
    ''' Saves the current DataGridView state into the layout model and persists it to disk.
    ''' </summary>
    Public Sub Save()
        Dim Index As Integer

        If File.Exists(LayoutDirectory & "\" & DefaultLayout.Routine & ".json") Then
            For Each Column In _CurrentLayout.Columns
                Index = _CurrentLayout.Columns.IndexOf(Column)
                Column.VisibleInGrid = DataGridView.Columns(Index).Visible
                Column.DisplayIndex = DataGridView.Columns(Index).DisplayIndex
                Column.DisplayName = DataGridView.Columns(Index).HeaderText
                Column.Width = DataGridView.Columns(Index).Width
                Column.WidthSizeMode = DataGridView.Columns(Index).AutoSizeMode
                Column.CellAlignment = DataGridView.Columns(Index).DefaultCellStyle.Alignment
                Column.HeaderAlignment = DataGridView.Columns(Index).HeaderCell.Style.Alignment
                Column.Format = DataGridView.Columns(Index).DefaultCellStyle.Format
            Next Column
            _CurrentLayout.SortedColumn = If(DataGridView.SortedColumn IsNot Nothing, DataGridView.SortedColumn.Index, -1)
            _CurrentLayout.SortDirection = If(DataGridView.SortOrder = SortOrder.Descending, SortOrder.Descending, SortOrder.Ascending)
        End If
        SaveJsonFile()
    End Sub
    ''' <summary>
    ''' Handles mouse down events on the DataGridView to detect column header interactions.
    ''' </summary>
    <DebuggerStepThrough>
    Private Sub DataGridView_MouseDown(sender As Object, e As MouseEventArgs)
        Dim Click As DataGridView.HitTestInfo = DataGridView.HitTest(e.X, e.Y)
        If Click.Type = DataGridViewHitTestType.ColumnHeader And e.Button = MouseButtons.Right Then
            _SaveLayout = False
            _LoadLayout = True
            _CmsPoint = e.Location
        ElseIf Click.Type = DataGridViewHitTestType.ColumnHeader And e.Button = MouseButtons.Left Then
            _SaveLayout = True
            _LoadLayout = False
        End If
    End Sub
    ''' <summary>
    ''' Handles mouse up events, triggering save or context menu display actions.
    ''' </summary>
    <DebuggerStepThrough>
    Private Sub DataGridView_MouseUp(sender As Object, e As MouseEventArgs)
        If _SaveLayout Then Save()
        If _LoadLayout Then
            _CmsColumns.Show(DataGridView.PointToScreen(_CmsPoint))
        End If
        _SaveLayout = False
        _LoadLayout = False
    End Sub
    ''' <summary>
    ''' Prevents the context menu from closing when a menu item is clicked.
    ''' </summary>
    Private Sub CmsColumns_Closing(sender As Object, e As ToolStripDropDownClosingEventArgs)
        If e.CloseReason = ToolStripDropDownCloseReason.ItemClicked Then
            e.Cancel = True
        End If
    End Sub

    Private Sub BtnRestoreGridLayout_Click(sender As Object, e As EventArgs)
        _CmsColumns.Close()
        If File.Exists(LayoutDirectory & "\" & _CurrentLayout.Routine & ".json") Then
            CreateJsonFile()
            Load()
        End If
    End Sub
    Private Sub BtnRemoveSort_Click(sender As Object, e As EventArgs)
        _CmsColumns.Close()
        _CurrentLayout = ReadJsonFile()
        _CurrentLayout.SortedColumn = -1
        _CurrentLayout.SortDirection = SortOrder.None
        SaveJsonFile()
        Load()
    End Sub
    Private Sub BtnColumn_CheckedChanged(sender As Object, e As EventArgs)
        Dim Menu As ContextMenuStrip =
            CType(CType(sender, ToolStripMenuItem).GetCurrentParent, ContextMenuStrip)
        If Not Menu.Items.OfType(Of ToolStripMenuItem).Any(Function(x) x.Checked) Then
            If _CurrentLayout.Columns.All(Function(x) x.VisibleInContext) Then
                sender.checked = True
            End If
        End If
        DataGridView.Columns.Cast(Of DataGridViewColumn).
            Single(Function(x) x.HeaderText = sender.Text).Visible = sender.Checked
        Save()
    End Sub
    ''' <summary>
    ''' Determines whether a newer version of the default layout exists compared to the saved one.
    ''' </summary>
    ''' <returns>True if the default model version is newer; otherwise false.</returns>
    Private Function HasNewVersion() As Boolean
        Dim Current = ReadJsonFile()
        Return DefaultLayout.Version > Current.Version
    End Function
End Class