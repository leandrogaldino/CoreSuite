Imports System.ComponentModel

''' <summary>
''' Represents the non-activating window that hosts the action buttons at run time.
''' </summary>
<DesignerCategory("Code")>
Friend NotInheritable Class TextBoxActionPopup
    Inherits Form
    Private Const WS_EX_NOACTIVATE As Integer = &H8000000
    Private Const WS_EX_TOOLWINDOW As Integer = &H80
    Private Const WM_MOUSEACTIVATE As Integer = &H21
    Private Const MA_NOACTIVATE As Integer = 3
    Private Shared ReadOnly TransparentColor As Color = Color.FromArgb(1, 2, 3)
    Private ReadOnly _LayoutPanel As TableLayoutPanel
    Private _BorderColor As Color = SystemColors.ControlDark
    Private _ShowBorder As Boolean
    Private _HasActions As Boolean
    Friend Event ActionClick As EventHandler(Of TextBoxActionPopupEventArgs)
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TextBoxActionPopup"/> class.
    ''' </summary>
    Public Sub New()
        AutoScaleMode = AutoScaleMode.Dpi
        FormBorderStyle = FormBorderStyle.None
        ShowInTaskbar = False
        StartPosition = FormStartPosition.Manual
        ControlBox = False
        MinimizeBox = False
        MaximizeBox = False
        AllowTransparency = True
        TransparencyKey = TransparentColor
        BackColor = TransparentColor
        Padding = System.Windows.Forms.Padding.Empty
        DoubleBuffered = True
        _LayoutPanel = New TableLayoutPanel With {.ColumnCount = 1, .RowCount = 1, .Dock = DockStyle.Fill, .GrowStyle = TableLayoutPanelGrowStyle.FixedSize, .Margin = System.Windows.Forms.Padding.Empty}
        Controls.Add(_LayoutPanel)
    End Sub
    ''' <summary>
    ''' Gets a value indicating whether the popup contains at least one visible action.
    ''' </summary>
    ''' <value><see langword="True"/> when a visible action is present; otherwise, <see langword="False"/>.</value>
    Friend ReadOnly Property HasActions As Boolean
        Get
            Return _HasActions
        End Get
    End Property
    ''' <summary>
    ''' Gets a value indicating that the popup should be displayed without activating its window.
    ''' </summary>
    Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
        Get
            Return True
        End Get
    End Property
    ''' <summary>
    ''' Adds the extended styles that prevent the popup from appearing in task switching or taking activation.
    ''' </summary>
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim Parameters As CreateParams = MyBase.CreateParams
            Parameters.ExStyle = Parameters.ExStyle Or WS_EX_NOACTIVATE Or WS_EX_TOOLWINDOW
            Return Parameters
        End Get
    End Property
    ''' <summary>
    ''' Rebuilds the internal table layout from the current action collection and appearance settings.
    ''' </summary>
    ''' <param name="Actions">The ordered action collection.</param>
    ''' <param name="ButtonSize">The width and height of each action button.</param>
    ''' <param name="ButtonSpacing">The space between adjacent buttons.</param>
    ''' <param name="ContentPadding">The space between the buttons and popup border.</param>
    ''' <param name="TransparentBackground">Indicates whether the popup background is transparent.</param>
    ''' <param name="ShowBorder">Indicates whether a one-pixel popup border is drawn.</param>
    ''' <param name="PanelBackColor">The panel background color used when transparency is disabled.</param>
    ''' <param name="BorderColor">The popup border color.</param>
    ''' <param name="ButtonBackColor">The normal button background color.</param>
    ''' <param name="ButtonHoverBackColor">The background color used while the pointer is over a button.</param>
    ''' <param name="ButtonPressedBackColor">The background color used while a button is pressed.</param>
    Friend Sub Rebuild(Actions As IEnumerable(Of TextBoxAction), ButtonSize As Integer, ButtonSpacing As Integer, ContentPadding As Integer, TransparentBackground As Boolean, ShowBorder As Boolean, PanelBackColor As Color, BorderColor As Color, ButtonBackColor As Color, ButtonHoverBackColor As Color, ButtonPressedBackColor As Color)
        Dim VisibleActions As New List(Of TextBoxAction)
        For Each Action As TextBoxAction In Actions
            If Action.Visible Then VisibleActions.Add(Action)
        Next
        _HasActions = VisibleActions.Count > 0
        _BorderColor = BorderColor
        _ShowBorder = ShowBorder
        Dim BackgroundColor As Color = If(TransparentBackground, TransparentColor, PanelBackColor)
        AllowTransparency = TransparentBackground
        TransparencyKey = If(TransparentBackground, TransparentColor, Color.Empty)
        BackColor = BackgroundColor
        Padding = If(ShowBorder, New Padding(1), System.Windows.Forms.Padding.Empty)
        _LayoutPanel.SuspendLayout()
        SuspendLayout()
        Try
            While _LayoutPanel.Controls.Count > 0
                Dim ExistingControl As Control = _LayoutPanel.Controls(0)
                _LayoutPanel.Controls.RemoveAt(0)
                ExistingControl.Dispose()
            End While
            _LayoutPanel.ColumnStyles.Clear()
            _LayoutPanel.RowStyles.Clear()
            _LayoutPanel.BackColor = BackgroundColor
            _LayoutPanel.Padding = New Padding(ContentPadding)
            If Not _HasActions Then
                _LayoutPanel.ColumnCount = 1
                _LayoutPanel.RowCount = 1
                _LayoutPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 1.0F))
                _LayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 1.0F))
                ClientSize = New Size(1, 1)
                Return
            End If
            _LayoutPanel.ColumnCount = VisibleActions.Count * 2 - 1
            _LayoutPanel.RowCount = 1
            For ColumnIndex As Integer = 0 To _LayoutPanel.ColumnCount - 1
                Dim ColumnWidth As Integer = If(ColumnIndex Mod 2 = 0, ButtonSize, ButtonSpacing)
                _LayoutPanel.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, ColumnWidth))
            Next
            _LayoutPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, ButtonSize))
            For ActionIndex As Integer = 0 To VisibleActions.Count - 1
                Dim Action As TextBoxAction = VisibleActions(ActionIndex)
                Dim Button As TextBoxActionButton = CreateButton(Action, ButtonSize, ButtonBackColor, ButtonHoverBackColor, ButtonPressedBackColor)
                Dim ButtonColumn As Integer = (VisibleActions.Count - 1 - ActionIndex) * 2
                _LayoutPanel.Controls.Add(Button, ButtonColumn, 0)
            Next
            Dim ContentWidth As Integer = VisibleActions.Count * ButtonSize + Math.Max(0, VisibleActions.Count - 1) * ButtonSpacing
            ClientSize = New Size(ContentWidth + ContentPadding * 2 + Padding.Horizontal, ButtonSize + ContentPadding * 2 + Padding.Vertical)
        Finally
            ResumeLayout(True)
            _LayoutPanel.ResumeLayout(True)
            Invalidate()
        End Try
    End Sub
    Private Function CreateButton(Action As TextBoxAction, ButtonSize As Integer, ButtonBackColor As Color, ButtonHoverBackColor As Color, ButtonPressedBackColor As Color) As TextBoxActionButton
        Dim Button As New TextBoxActionButton(Action) With {.Size = New Size(ButtonSize, ButtonSize), .Margin = System.Windows.Forms.Padding.Empty, .Anchor = AnchorStyles.None, .Enabled = Action.Enabled, .Image = Action.Image, .TooltipText = Action.ToolTipText, .BackColor = ButtonBackColor}
        Button.FlatAppearance.BorderSize = 0
        Button.FlatAppearance.MouseOverBackColor = ButtonHoverBackColor
        Button.FlatAppearance.MouseDownBackColor = ButtonPressedBackColor
        Button.AccessibleName = GetAccessibleName(Action)
        Button.AccessibleDescription = Action.ToolTipText
        AddHandler Button.ActionClick, AddressOf Button_ActionClick
        Return Button
    End Function
    Private Shared Function GetAccessibleName(Action As TextBoxAction) As String
        If Not String.IsNullOrWhiteSpace(Action.AccessibleName) Then Return Action.AccessibleName
        If Not String.IsNullOrWhiteSpace(Action.ToolTipText) Then Return Action.ToolTipText
        Return Action.Key
    End Function
    Private Sub Button_ActionClick(Sender As Object, E As EventArgs)
        Dim Button As TextBoxActionButton = TryCast(Sender, TextBoxActionButton)
        If Button Is Nothing OrElse Not Button.Enabled Then Return
        RaiseEvent ActionClick(Me, New TextBoxActionPopupEventArgs(Button.Action))
    End Sub
    ''' <summary>
    ''' Prevents mouse interaction with the popup from activating it and taking focus from the target text box.
    ''' </summary>
    ''' <param name="M">The Windows message to process.</param>
    Protected Overrides Sub WndProc(ByRef M As Message)
        If M.Msg = WM_MOUSEACTIVATE Then
            M.Result = New IntPtr(MA_NOACTIVATE)
            Return
        End If
        MyBase.WndProc(M)
    End Sub
    ''' <summary>
    ''' Draws the one-pixel border around the popup.
    ''' </summary>
    ''' <param name="E">The paint event data.</param>
    Protected Overrides Sub OnPaint(E As PaintEventArgs)
        MyBase.OnPaint(E)
        If Not _ShowBorder OrElse ClientSize.Width <= 0 OrElse ClientSize.Height <= 0 Then Return
        Using BorderPen As New Pen(_BorderColor)
            E.Graphics.DrawRectangle(BorderPen, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1)
        End Using
    End Sub
End Class
''' <summary>
''' Provides the action selected inside a <see cref="TextBoxActionPopup"/>.
''' </summary>
Friend NotInheritable Class TextBoxActionPopupEventArgs
    Inherits EventArgs
    Private ReadOnly _Action As TextBoxAction
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TextBoxActionPopupEventArgs"/> class.
    ''' </summary>
    ''' <param name="Action">The selected action.</param>
    Public Sub New(Action As TextBoxAction)
        _Action = Action
    End Sub
    ''' <summary>
    ''' Gets the selected action.
    ''' </summary>
    Friend ReadOnly Property Action As TextBoxAction
        Get
            Return _Action
        End Get
    End Property
End Class
