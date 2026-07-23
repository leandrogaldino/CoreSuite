Imports System.Reflection
Partial Public Class QueriedBox
    Private Class FormDropDownResults
        Inherits Form
        Public Textbox As Control
        Private ReadOnly _QueriedBox As QueriedBox
        Public Sub New(Control As QueriedBox)
            SuspendLayout()
            _QueriedBox = Control
            InitializeComponent()
            BackColor = _QueriedBox.DropDownBorderColor
            Font = Control.Font
            FormBorderStyle = FormBorderStyle.None
            Padding = New Padding(1)
            Size = New Size(300, 120)
            DoubleBuffered = True
            TopMost = True
            KeyPreview = True
            ResumeLayout(True)
        End Sub
        Private Sub InitializeComponent()
            Dim ShowVGridLines As Boolean = _QueriedBox.ShowVerticalGridLines
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
        Private Sub DropDownResults_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            Application.AddMessageFilter(New PopupWindowHelperMessageFilter(Me, Textbox))
        End Sub
        Private Shared Sub EnableDoubleBuffered(dgv As DataGridView)
            Dim DgvType As Type = dgv.[GetType]()
            Dim Info As PropertyInfo = DgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
            Info.SetValue(dgv, True, Nothing)
        End Sub
        Protected Overrides ReadOnly Property CreateParams As CreateParams
            Get
                Dim Rect As CreateParams = MyBase.CreateParams
                Rect.Style = CInt(Flags.WindowStyles.WS_SYSMENU) Or CInt(Flags.WindowStyles.WS_CHILD)
                Rect.ExStyle = Rect.ExStyle Or CInt(Flags.WindowStyles.WS_EX_NOACTIVATE) Or CInt(Flags.WindowStyles.WS_EX_TOOLWINDOW)
                Rect.X = Me.Location.X
                Rect.Y = Me.Location.Y
                Return Rect
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
    Private Class PopupWindowHelperMessageFilter
        Implements IMessageFilter
        Private Const WM_LBUTTONDOWN As Integer = &H201
        Private Const WM_RBUTTONDOWN As Integer = &H204
        Private Const WM_MBUTTONDOWN As Integer = &H207
        Private Const WM_NCLBUTTONDOWN As Integer = &HA1
        Private Const WM_NCRBUTTONDOWN As Integer = &HA4
        Private Const WM_NCMBUTTONDOWN As Integer = &HA7
        Private TextBox As Control = Nothing
        Public Property Popup As Form = Nothing
        Public Sub New(popupW As Form, textbox As Control)
            Popup = popupW
            Me.TextBox = textbox
        End Sub
        <DebuggerStepThrough>
        Private Sub OnMouseDown()
            Dim CursorPos As Point = Cursor.Position
            If TextBox.Parent Is Nothing Then Exit Sub
            If Not Popup.Bounds.Contains(CursorPos) Then
                If Not TextBox.Bounds.Contains(TextBox.Parent.PointToClient(CursorPos)) Then
                    If CType(TextBox, QueriedBox).DropDownResultsForm IsNot Nothing AndAlso CType(TextBox, QueriedBox).DropDownResultsForm.DgvResults.SelectedRows.Count = 1 Then
                        If UCase(CType(TextBox, QueriedBox).Text) = UCase(CType(TextBox, QueriedBox).DropDownResultsForm.DgvResults.SelectedRows(0).Cells(If(CType(TextBox, QueriedBox).DisplayFieldAlias = Nothing, CType(TextBox, QueriedBox).DisplayFieldName, CType(TextBox, QueriedBox).DisplayFieldAlias)).Value.ToString) Then
                            CType(TextBox, QueriedBox).AutoFreeze()
                        End If
                    End If
                    Application.RemoveMessageFilter(Me)
                    Popup.Close()
                End If
            End If
        End Sub
        <DebuggerStepThrough>
        Private Function IMessageFilter_PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage
            If Popup IsNot Nothing Then
                Select Case m.Msg
                    Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN, WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_NCMBUTTONDOWN
                        OnMouseDown()
                End Select
            End If
            Return False
        End Function
    End Class
    Private Class Flags
        <Flags()>
        Public Enum WindowStyles As UInteger
            WS_OVERLAPPED = 0
            WS_POPUP = 2147483648
            WS_CHILD = 1073741824
            WS_MINIMIZE = 536870912
            WS_VISIBLE = 268435456
            WS_DISABLED = 134217728
            WS_CLIPSIBLINGS = 67108864
            WS_CLIPCHILDREN = 33554432
            WS_MAXIMIZE = 16777216
            WS_BORDER = 8388608
            WS_DLGFRAME = 4194304
            WS_VSCROLL = 2097152
            WS_HSCROLL = 1048576
            WS_SYSMENU = 524288
            WS_THICKFRAME = 262144
            WS_GROUP = 131072
            WS_TABSTOP = 65536
            WS_MINIMIZEBOX = 131072
            WS_MAXIMIZEBOX = 65536
            WS_CAPTION = WS_BORDER Or WS_DLGFRAME
            WS_TILED = WS_OVERLAPPED
            WS_ICONIC = WS_MINIMIZE
            WS_SIZEBOX = WS_THICKFRAME
            WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
            WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
            WS_POPUPWINDOW = WS_POPUP Or WS_BORDER Or WS_SYSMENU
            WS_CHILDWINDOW = WS_CHILD
            WS_EX_DLGMODALFRAME = 1
            WS_EX_NOPARENTNOTIFY = 4
            WS_EX_TOPMOST = 8
            WS_EX_ACCEPTFILES = 16
            WS_EX_TRANSPARENT = 32
            WS_EX_MDICHILD = 64
            WS_EX_TOOLWINDOW = 128
            WS_EX_WINDOWEDGE = 256
            WS_EX_CLIENTEDGE = 512
            WS_EX_CONTEXTHELP = 1024
            WS_EX_RIGHT = 4096
            WS_EX_LEFT = 0
            WS_EX_RTLREADING = 8192
            WS_EX_LTRREADING = 0
            WS_EX_LEFTSCROLLBAR = 16384
            WS_EX_RIGHTSCROLLBAR = 0
            WS_EX_CONTROLPARENT = 65536
            WS_EX_STATICEDGE = 131072
            WS_EX_APPWINDOW = 262144
            WS_EX_OVERLAPPEDWINDOW = WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE
            WS_EX_PALETTEWINDOW = WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
            WS_EX_LAYERED = 524288
            WS_EX_NOINHERITLAYOUT = 1048576
            WS_EX_LAYOUTRTL = 4194304
            WS_EX_COMPOSITED = 33554432
            WS_EX_NOACTIVATE = 134217728
        End Enum
    End Class
End Class
