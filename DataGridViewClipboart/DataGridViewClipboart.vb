Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms

Public Class DataGridViewClipboart
    Inherits Component
    Private _ContextMenuStrip As ContextMenuStrip
    Private _DataGridView As DataGridView
    Private ReadOnly _ToolStripMenuItemCopyCell As ToolStripMenuItem
    Private ReadOnly _ToolStripMenuItemCopyRow As ToolStripMenuItem
    Private _ShowContext As Boolean
    Private _ContextPoint As Point
    Private _ClickedColumnIndex As Integer
    Private _ClickedRowIndex As Integer
    Public Sub New()
        _ToolStripMenuItemCopyCell = New ToolStripMenuItem() With {.Text = "Copiar Celula", .DisplayStyle = ToolStripItemDisplayStyle.ImageAndText, .Image = My.Resources.CellCopy}
        _ToolStripMenuItemCopyRow = New ToolStripMenuItem() With {.Text = "Copiar Linha", .DisplayStyle = ToolStripItemDisplayStyle.ImageAndText, .Image = My.Resources.RowCopy}
        AddHandler _ToolStripMenuItemCopyCell.Click, AddressOf CopyCellClick
        AddHandler _ToolStripMenuItemCopyRow.Click, AddressOf CopyRowClick
        CreateDefaultCms()
    End Sub

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

    Private Sub CreateDefaultCms()
        _ContextMenuStrip = New ContextMenuStrip
        _ContextMenuStrip.Items.AddRange({_ToolStripMenuItemCopyCell, _ToolStripMenuItemCopyRow})
    End Sub

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
    Public Property DataGridView As DataGridView
        Get
            Return _DataGridView
        End Get
        Set(value As DataGridView)
            RemoveHandler value.MouseDown, AddressOf Dgv_MouseDown
            RemoveHandler value.MouseUp, AddressOf Dgv_MouseUp
            _DataGridView = value
            If _DataGridView IsNot Nothing Then
                AddHandler value.MouseDown, AddressOf Dgv_MouseDown
                AddHandler value.MouseUp, AddressOf Dgv_MouseUp
            End If
        End Set
    End Property
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
    Private Sub Dgv_MouseUp(sender As Object, e As MouseEventArgs)
        If _ShowContext Then
            _ContextMenuStrip.Show(_DataGridView.PointToScreen(_ContextPoint))
            _ShowContext = False
        End If
    End Sub

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

    Private Sub OnContextMenuOpening(sender As Object, e As CancelEventArgs)
        Dim menu = DirectCast(sender, ContextMenuStrip)

        ' Evita duplicar caso o evento seja chamado mais de uma vez
        If Not menu.Items.Contains(_ToolStripMenuItemCopyCell) And Not DesignMode Then
            menu.Items.Add(New ToolStripSeparator())
            menu.Items.Add(_ToolStripMenuItemCopyCell)
            menu.Items.Add(_ToolStripMenuItemCopyRow)
        End If
    End Sub

    Public Property IncludeHeaderTextInRowCopy As Boolean
    Public Property IncludeHeaderTextInCellCopy As Boolean
End Class
