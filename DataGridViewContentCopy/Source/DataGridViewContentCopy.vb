Imports System.ComponentModel
''' <summary>
''' Provides copy commands for <see cref="DataGridView"/> cells and rows through a context menu.
''' </summary>
Public Class DataGridViewContentCopy
    Inherits Component
    Private ReadOnly _ToolStripMenuItemCopyCell As ToolStripMenuItem
    Private ReadOnly _ToolStripMenuItemCopyRow As ToolStripMenuItem
    Private _ContextMenuStrip As ContextMenuStrip
    Private _DataGridView As DataGridView
    Private _ShowContext As Boolean
    Private _ContextPoint As Point
    Private _ClickedColumnIndex As Integer
    Private _ClickedRowIndex As Integer
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DataGridViewContentCopy"/> class.
    ''' </summary>
    Public Sub New()
        _ToolStripMenuItemCopyCell = New ToolStripMenuItem() With {.Text = "Copiar Celula", .DisplayStyle = ToolStripItemDisplayStyle.ImageAndText, .Image = My.Resources.ImageResources.CellCopy}
        _ToolStripMenuItemCopyRow = New ToolStripMenuItem() With {.Text = "Copiar Linha", .DisplayStyle = ToolStripItemDisplayStyle.ImageAndText, .Image = My.Resources.ImageResources.RowCopy}
        AddHandler _ToolStripMenuItemCopyCell.Click, AddressOf CopyCellClick
        AddHandler _ToolStripMenuItemCopyRow.Click, AddressOf CopyRowClick
        CreateDefaultCms()
    End Sub
    ''' <summary>
    ''' Gets or sets a value indicating whether images are displayed in the context menu items.
    ''' </summary>
    Public Property ShowImages As Boolean
        Get
            Return _ToolStripMenuItemCopyCell.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
        End Get
        Set(value As Boolean)
            If value Then
                _ToolStripMenuItemCopyCell.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
                _ToolStripMenuItemCopyRow.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
            Else
                _ToolStripMenuItemCopyCell.DisplayStyle = ToolStripItemDisplayStyle.Text
                _ToolStripMenuItemCopyRow.DisplayStyle = ToolStripItemDisplayStyle.Text
            End If
        End Set
    End Property
    ''' <summary>
    ''' Creates the default context menu used by the component.
    ''' </summary>
    Private Sub CreateDefaultCms()
        _ContextMenuStrip = New ContextMenuStrip
        _ContextMenuStrip.Items.AddRange({_ToolStripMenuItemCopyCell, _ToolStripMenuItemCopyRow})
    End Sub
    ''' <summary>
    ''' Gets or sets the context menu displayed when the user right-clicks a cell.
    ''' </summary>
    ''' <remarks>
    ''' If set to <see langword="Nothing"/>, a default context menu is created automatically.
    ''' </remarks>
    Public Property ContextMenuStrip As ContextMenuStrip
        Get
            Return _ContextMenuStrip
        End Get
        Set(value As ContextMenuStrip)
            If value Is Nothing Then
                CreateDefaultCms()
            Else
                If _ContextMenuStrip IsNot Nothing Then
                    RemoveHandler _ContextMenuStrip.Opening, AddressOf OnContextMenuOpening
                End If
                _ContextMenuStrip = value
                AddHandler _ContextMenuStrip.Opening, AddressOf OnContextMenuOpening
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the <see cref="Windows.Forms.DataGridView"/> associated with this component.
    ''' </summary>
    Public Property DataGridView As DataGridView
        Get
            Return _DataGridView
        End Get
        Set(value As DataGridView)
            If _DataGridView IsNot Nothing Then
                RemoveHandler _DataGridView.MouseDown, AddressOf Dgv_MouseDown
                RemoveHandler _DataGridView.MouseUp, AddressOf Dgv_MouseUp
            End If
            _DataGridView = value
            If _DataGridView IsNot Nothing Then
                AddHandler _DataGridView.MouseDown, AddressOf Dgv_MouseDown
                AddHandler _DataGridView.MouseUp, AddressOf Dgv_MouseUp
            End If
        End Set
    End Property
    ''' <summary>
    ''' Handles the mouse down event to identify the clicked cell and prepare the context menu.
    ''' </summary>
    ''' <param name="sender">The event source.</param>
    ''' <param name="e">Mouse event data.</param>
    Private Sub Dgv_MouseDown(sender As Object, e As MouseEventArgs)
        Dim Click As DataGridView.HitTestInfo = _DataGridView.HitTest(e.X, e.Y)
        If Click.Type = DataGridViewHitTestType.Cell And e.Button = MouseButtons.Right Then
            _DataGridView.Rows(Click.RowIndex).Selected = True
            _ShowContext = True
            _ContextPoint = e.Location
            _ClickedColumnIndex = Click.ColumnIndex
            _ClickedRowIndex = Click.RowIndex
        End If
    End Sub
    ''' <summary>
    ''' Displays the context menu after a valid right-click operation.
    ''' </summary>
    ''' <param name="sender">The event source.</param>
    ''' <param name="e">Mouse event data.</param>
    Private Sub Dgv_MouseUp(sender As Object, e As MouseEventArgs)
        If _ShowContext Then
            _ContextMenuStrip.Show(_DataGridView.PointToScreen(_ContextPoint))
            _ShowContext = False
        End If
    End Sub
    ''' <summary>
    ''' Copies the selected cell content to the clipboard.
    ''' </summary>
    ''' <param name="sender">The event source.</param>
    ''' <param name="e">Event data.</param>
    Private Sub CopyCellClick(sender As Object, e As EventArgs)
        Dim Cell As DataGridViewCell = _DataGridView.Rows(_ClickedRowIndex).Cells(_ClickedColumnIndex)
        Dim Text As String
        If IncludeHeaderTextInCellCopy Then
            Text = $"{_DataGridView.Columns(Cell.ColumnIndex).HeaderText}: {Convert.ToString(Cell.Value)}"
        Else
            Text = Convert.ToString(Cell.Value)
        End If
        Clipboard.SetText(Text)
    End Sub
    ''' <summary>
    ''' Copies the selected row content to the clipboard.
    ''' </summary>
    ''' <param name="sender">The event source.</param>
    ''' <param name="e">Event data.</param>
    Private Sub CopyRowClick(sender As Object, e As EventArgs)
        Dim Row As DataGridViewRow = _DataGridView.Rows(_ClickedRowIndex)
        Dim RowCellsText As New List(Of String)
        Dim Text As String
        If IncludeHeaderTextInRowCopy Then
            For Each Cell As DataGridViewCell In Row.Cells
                RowCellsText.Add($"{_DataGridView.Columns(Cell.ColumnIndex).HeaderText}: {Convert.ToString(Cell.Value)}")
            Next Cell
        Else
            For Each Cell As DataGridViewCell In Row.Cells
                RowCellsText.Add(Convert.ToString(Cell.Value))
            Next Cell
        End If
        Text = String.Join(" | ", RowCellsText)
        Clipboard.SetText(Text)
    End Sub
    ''' <summary>
    ''' Ensures the copy commands are available in the context menu before it is displayed.
    ''' </summary>
    ''' <param name="sender">The event source.</param>
    ''' <param name="e">Provides data for the opening event.</param>
    Private Sub OnContextMenuOpening(sender As Object, e As CancelEventArgs)
        Dim menu = DirectCast(sender, ContextMenuStrip)
        If Not menu.Items.Contains(_ToolStripMenuItemCopyCell) And Not DesignMode Then
            menu.Items.Add(New ToolStripSeparator())
            menu.Items.Add(_ToolStripMenuItemCopyCell)
            menu.Items.Add(_ToolStripMenuItemCopyRow)
        End If
    End Sub
    ''' <summary>
    ''' Gets or sets a value indicating whether column headers are included when copying a row.
    ''' </summary>
    Public Property IncludeHeaderTextInRowCopy As Boolean
    ''' <summary>
    ''' Gets or sets a value indicating whether the column header is included when copying a cell.
    ''' </summary>
    Public Property IncludeHeaderTextInCellCopy As Boolean
End Class
