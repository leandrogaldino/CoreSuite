Imports System.ComponentModel
Imports System.Reflection
''' <summary>
''' Represents the popup window used to display query results for the <see cref="QueriedBox"/> control.
''' </summary>
<DesignerCategory("Code")>
Friend Class DropDownResultsPopup
    Inherits Form
    Private ReadOnly _QueriedBox As QueriedBox
    Private _MessageFilter As PopupWindowHelperMessageFilter
    ''' <summary>
    ''' Gets the data grid used to display query results.
    ''' </summary>
    Friend WithEvents DgvResults As DataGridView
    ''' <summary>
    ''' Gets the label used to display informational messages.
    ''' </summary>
    Friend WithEvents LblCharsRemaining As Label
    ''' <summary>
    ''' Gets the container panel used to organize popup controls.
    ''' </summary>
    Friend WithEvents PanelContainer As Panel
    ''' <summary>
    ''' Gets or sets the control that owns this popup window.
    ''' </summary>
    Public Textbox As Control
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DropDownResultsPopup"/> class.
    ''' </summary>
    ''' <param name="Control">
    ''' The <see cref="QueriedBox"/> control associated with this popup.
    ''' </param>
    Public Sub New(Control As QueriedBox)
        SuspendLayout()
        _QueriedBox = Control
        InitializeComponent()
        BackColor = _QueriedBox.DropDown.BorderColor
        Font = Control.Font
        FormBorderStyle = FormBorderStyle.None
        Padding = New Padding(1)
        Size = New Size(300, 120)
        DoubleBuffered = True
        TopMost = True
        KeyPreview = True
        ResumeLayout(True)
    End Sub
    ''' <summary>
    ''' Initializes the child controls contained in the popup window.
    ''' </summary>
    Private Sub InitializeComponent()
        Dim ShowVGridLines As Boolean = _QueriedBox.Grid.ShowVerticalLines
        PanelContainer = New Panel With {.Dock = DockStyle.Fill}
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
            .CellBorderStyle = If(_QueriedBox.Grid.ShowVerticalLines, DataGridViewCellBorderStyle.Single, DataGridViewCellBorderStyle.SingleHorizontal),
            .ColumnHeadersBorderStyle = If(_QueriedBox.Grid.ShowVerticalLines, DataGridViewHeaderBorderStyle.Raised, DataGridViewCellBorderStyle.None),
            .ColumnHeadersVisible = _QueriedBox.Grid.HeaderVisible,
            .DefaultCellStyle = New DataGridViewCellStyle With {
                .SelectionBackColor = _QueriedBox.Grid.SelectionBackColor,
                .SelectionForeColor = _QueriedBox.Grid.SelectionForeColor,
                .BackColor = _QueriedBox.Grid.BackColor,
                .ForeColor = _QueriedBox.Grid.ForeColor
            },
            .ColumnHeadersDefaultCellStyle = New DataGridViewCellStyle With {
                .BackColor = _QueriedBox.Grid.HeaderBackColor,
                .ForeColor = _QueriedBox.Grid.HeaderForeColor
            },
            .Dock = DockStyle.Fill,
            .MultiSelect = False,
            .[ReadOnly] = True,
            .RowHeadersVisible = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .Visible = False,
            .EnableHeadersVisualStyles = False
        }
        DgvResults.ColumnHeadersDefaultCellStyle.Font = New Font(_QueriedBox.Font, If(_QueriedBox.Grid.HeadersBold, FontStyle.Bold, FontStyle.Regular))
        EnableDoubleBuffered(DgvResults)
        LblCharsRemaining = New Label With {
            .AutoSize = False,
            .Dock = DockStyle.Fill,
            .TextAlign = ContentAlignment.MiddleCenter,
            .BackColor = _QueriedBox.Messages.BackColor,
            .ForeColor = _QueriedBox.Messages.ForeColor,
            .Visible = False
        }
        PanelContainer.Controls.AddRange({DgvResults, LblCharsRemaining})
        Controls.Add(PanelContainer)
    End Sub
    ''' <summary>
    ''' Handles changes to the data source of the results grid.
    ''' Displays the no results message when the grid contains no rows.
    ''' </summary>
    ''' <param name="sender">The object that raised the event.</param>
    ''' <param name="e">The event data.</param>
    Private Sub DgvResults_DataSourceChanged(sender As Object, e As EventArgs) Handles DgvResults.DataSourceChanged
        If DgvResults.Rows.Count = 0 Then
            DgvResults.Visible = False
            LblCharsRemaining.Text = _QueriedBox.Messages.NoResultsText
            LblCharsRemaining.Visible = True
        End If
    End Sub
    ''' <summary>
    ''' Handles keyboard preview events from the results grid.
    ''' Closes the popup when the TAB key is pressed.
    ''' </summary>
    ''' <param name="sender">The object that raised the event.</param>
    ''' <param name="e">The keyboard event data.</param>
    Private Sub DataGridView_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles DgvResults.PreviewKeyDown
        If e.KeyCode = Keys.Tab Then
            Close()
            Me.Select()
        End If
    End Sub
    ''' <summary>
    ''' Handles double-click events on result rows.
    ''' Selects the associated item and closes the popup.
    ''' </summary>
    ''' <param name="sender">The object that raised the event.</param>
    ''' <param name="e">The mouse event data.</param>
    Private Sub DataGridView_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles DgvResults.MouseDoubleClick
        Dim Click As DataGridView.HitTestInfo = DgvResults.HitTest(e.X, e.Y)
        If Click.Type = DataGridViewHitTestType.Cell Then
            _QueriedBox.AutoFreeze()
            _QueriedBox.Focus()
            Close()
        End If
    End Sub
    ''' <summary>
    ''' Unregisters the message filter associated with this popup window.
    ''' </summary>
    Private Sub RemoveMessageFilter()
        If _MessageFilter Is Nothing Then Exit Sub
        Application.RemoveMessageFilter(_MessageFilter)
        _MessageFilter = Nothing
    End Sub
    ''' <summary>
    ''' Enables double buffering on a <see cref="DataGridView"/> control to reduce rendering flicker.
    ''' </summary>
    ''' <param name="dgv">The grid control to configure.</param>
    Private Shared Sub EnableDoubleBuffered(dgv As DataGridView)
        Dim DgvType As Type = dgv.[GetType]()
        Dim Info As PropertyInfo = DgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
        Info.SetValue(dgv, True, Nothing)
    End Sub
    ''' <summary>
    ''' Registers the message filter responsible for monitoring mouse clicks
    ''' outside the popup window.
    ''' </summary>
    ''' <param name="e">
    ''' The event data associated with the load event.
    ''' </param>
    Protected Overrides Sub OnLoad(e As EventArgs)
        MyBase.OnLoad(e)
        If _MessageFilter Is Nothing Then
            _MessageFilter = New PopupWindowHelperMessageFilter(Me, Textbox)
            Application.AddMessageFilter(_MessageFilter)
        End If
    End Sub
    ''' <summary>
    ''' Gets the parameters used when creating the native window handle.
    ''' </summary>
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim Rect As CreateParams = MyBase.CreateParams
            Rect.Style = CInt(WindowStyles.WS_SYSMENU) Or CInt(WindowStyles.WS_CHILD)
            Rect.ExStyle = Rect.ExStyle Or CInt(WindowStyles.WS_EX_NOACTIVATE) Or CInt(WindowStyles.WS_EX_TOOLWINDOW)
            Rect.X = Me.Location.X
            Rect.Y = Me.Location.Y
            Return Rect
        End Get
    End Property
    ''' <summary>
    ''' Removes the registered message filter when the popup window is closed.
    ''' </summary>
    ''' <param name="e">
    ''' The event data associated with the form closed event.
    ''' </param>
    Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
        RemoveMessageFilter()
        MyBase.OnFormClosed(e)
    End Sub
    ''' <summary>
    ''' Releases the resources used by the popup window and unregisters
    ''' the associated message filter.
    ''' </summary>
    ''' <param name="disposing">
    ''' <see langword="true"/> to release managed resources; otherwise,
    ''' <see langword="false"/>.
    ''' </param>
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then
            RemoveMessageFilter()
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class