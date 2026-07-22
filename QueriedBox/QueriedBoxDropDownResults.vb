Imports System.Reflection

Friend Class QueriedBoxDropDownResults
    Inherits Form
    Public Textbox As Control
    Private ReadOnly _QueriedBox As QueriedBox
    Public Sub New(ByVal SearchBox As QueriedBox)
        SuspendLayout()
        _QueriedBox = SearchBox
        InitializeComponent()
        BackColor = _QueriedBox.DropDownBorderColor
        Font = SearchBox.Font
        FormBorderStyle = FormBorderStyle.None
        Padding = New Padding(1)
        Size = New Size(300, 120)
        DoubleBuffered = True
        TopMost = True
        KeyPreview = True
        ResumeLayout(True)
    End Sub
    Private Sub InitializeComponent()
        Dim sg As Boolean = _QueriedBox.ShowVerticalGridLines
        PanelContainer = New Panel With {
            .Dock = DockStyle.Fill
        }
        DgvResults = New DataGridView With {
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .AllowUserToResizeColumns = True,
            .AllowUserToResizeRows = False,
            .AllowUserToOrderColumns = True,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
            .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.None,
            .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            .CellBorderStyle = If(_QueriedBox.ShowVerticalGridLines, DataGridViewCellBorderStyle.Single, DataGridViewCellBorderStyle.SingleHorizontal),
            .ColumnHeadersBorderStyle = If(_QueriedBox.ShowVerticalGridLines, DataGridViewHeaderBorderStyle.Raised, DataGridViewCellBorderStyle.None),
            .ColumnHeadersVisible = _QueriedBox.GridHeaderVisible,
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .SelectionBackColor = _QueriedBox.GridSelectionBackColor,
                .SelectionForeColor = _QueriedBox.GridSelectionForeColor,
                .BackColor = _QueriedBox.GridBackColor,
                .ForeColor = _QueriedBox.GridForeColor
            },
            .ColumnHeadersDefaultCellStyle = New DataGridViewCellStyle With {
                .BackColor = _QueriedBox.GridHeaderBackColor,
                .ForeColor = _QueriedBox.GridHeaderForeColor
            },
            .Dock = DockStyle.Fill,
            .MultiSelect = False,
            .[ReadOnly] = True,
            .RowHeadersVisible = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .Visible = False,
            .EnableHeadersVisualStyles = False
        }
        DgvResults.ColumnHeadersDefaultCellStyle.Font = New Font(_QueriedBox.Font, If(_QueriedBox.GridHeadersBold, FontStyle.Bold, FontStyle.Regular))
        EnableDoubleBuffered(DgvResults)
        LblCharsRemaining = New Label With {
            .AutoSize = False,
            .Dock = DockStyle.Fill,
            .TextAlign = ContentAlignment.MiddleCenter,
            .BackColor = _QueriedBox.LabelBackColor,
            .ForeColor = _QueriedBox.LabelForeColor,
            .Visible = False
        }
        PanelContainer.Controls.AddRange({DgvResults, LblCharsRemaining})
        Controls.Add(PanelContainer)
    End Sub
    Private Sub DgvResults_DataSourceChanged(sender As Object, e As EventArgs) Handles DgvResults.DataSourceChanged
        If DgvResults.Rows.Count = 0 Then
            DgvResults.Visible = False
            LblCharsRemaining.Text = "Nenhum Registro Encontrado."
            LblCharsRemaining.Visible = True
        End If
    End Sub
    Private Sub DropDownResults_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Application.AddMessageFilter(New PopupWindowHelperMessageFilter(Me, Textbox))
    End Sub
    Private Shared Sub EnableDoubleBuffered(ByVal dgv As DataGridView)
        Dim DgvType As Type = dgv.[GetType]()
        Dim pi As PropertyInfo = DgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
        pi.SetValue(dgv, True, Nothing)
    End Sub
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim ret As CreateParams = MyBase.CreateParams
            ret.Style = CInt(Flags.WindowStyles.WS_SYSMENU) Or CInt(Flags.WindowStyles.WS_CHILD)
            ret.ExStyle = ret.ExStyle Or CInt(Flags.WindowStyles.WS_EX_NOACTIVATE) Or CInt(Flags.WindowStyles.WS_EX_TOOLWINDOW)
            ret.X = Me.Location.X
            ret.Y = Me.Location.Y
            Return ret
        End Get
    End Property
    Private Sub DataGridView_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles DgvResults.PreviewKeyDown
        If e.KeyCode = Keys.Tab Then
            Close()
            Me.Select()
        End If
    End Sub
    Private Sub DataGridView_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles DgvResults.MouseDoubleClick
        Dim Click As DataGridView.HitTestInfo = DgvResults.HitTest(e.X, e.Y)
        If Click.Type = DataGridViewHitTestType.Cell Then
            _QueriedBox.AutoFreeze()
            _QueriedBox.Focus()
            Close()
        End If
    End Sub
    Friend WithEvents DgvResults As DataGridView
    Friend WithEvents LblCharsRemaining As Label
    Friend WithEvents PanelContainer As Panel
End Class
