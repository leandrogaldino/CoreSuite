Imports System.ComponentModel
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports CoreSuite.Controls.My.Resources

''' <summary>
''' Provides a reusable password text box with an embedded action button for revealing,
''' hiding, or replacing a password without requiring an existing password value to be loaded.
''' </summary>
<DefaultEvent("PasswordValueChanged")>
<DesignerCategory("Code")>
Public Class PasswordBox
    Inherits TextBox

    Private Const DefaultButtonWidth As Integer = 30
    Private Const DefaultMaskLength As Integer = 8
    Private Const MinimumButtonWidth As Integer = 24
    Private Const MinimumMaskLength As Integer = 1
    Private Const ButtonMargin As Integer = 2
    Private Const EmSetMargins As Integer = &HD3
    Private Const EcRightMargin As Integer = &H2

    Private ReadOnly _ActionButton As PasswordActionButton
    Private ReadOnly _ToolTip As ToolTip
    Private _PasswordDefined As Boolean
    Private _IsEditingPassword As Boolean
    Private _PasswordChanged As Boolean
    Private _IsPasswordVisible As Boolean
    Private _SuppressTextChanged As Boolean
    Private _ReadOnly As Boolean
    Private _AllowPasswordReveal As Boolean = True
    Private _ShowActionButton As Boolean = True
    Private _ButtonWidth As Integer = DefaultButtonWidth
    Private _DefinedPasswordMaskLength As Integer = DefaultMaskLength
    Private _ShowPasswordImage As Image = Images.ShowPassword
    Private _HidePasswordImage As Image = Images.HidePassword
    Private _ChangePasswordImage As Image = Images.ChangePassword
    Private _ChangePasswordToolTipText As String = "Change password"
    Private _ShowPasswordToolTipText As String = "Show password"
    Private _HidePasswordToolTipText As String = "Hide password"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="PasswordBox"/> class.
    ''' </summary>
    Public Sub New()
        MyBase.UseSystemPasswordChar = True
        _ActionButton = New PasswordActionButton With {
            .TabStop = False,
            .Cursor = Cursors.Hand
        }
        _ToolTip = New ToolTip()
        Controls.Add(_ActionButton)
        AddHandler _ActionButton.Click, AddressOf ActionButton_Click
        UpdateVisualState()
    End Sub

    ''' <summary>
    ''' Occurs when the replacement password value is modified by the user.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Occurs when the replacement password value is modified by the user.")>
    Public Event PasswordValueChanged As EventHandler

    ''' <summary>
    ''' Occurs when the control enters password replacement mode.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Occurs when the control enters password replacement mode.")>
    Public Event PasswordEditStarted As EventHandler

    ''' <summary>
    ''' Occurs when a pending password replacement is canceled.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Occurs when a pending password replacement is canceled.")>
    Public Event PasswordEditCanceled As EventHandler

    ''' <summary>
    ''' Occurs when the visibility of the replacement password changes.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Occurs when the visibility of the replacement password changes.")>
    Public Event PasswordVisibilityChanged As EventHandler

    ''' <summary>
    ''' Gets or sets a value indicating whether a password is already stored by the application.
    ''' </summary>
    ''' <remarks>
    ''' When set to <see langword="True"/>, the control displays a fixed-length placeholder and
    ''' does not require the existing password value. The action button starts password replacement mode.
    ''' </remarks>
    <Category("PasswordBox")>
    <Description("Indicates whether a password is already stored by the application.")>
    <DefaultValue(False)>
    Public Property PasswordDefined As Boolean
        Get
            Return _PasswordDefined
        End Get
        Set(value As Boolean)
            If _PasswordDefined = value Then Return
            _PasswordDefined = value
            _IsEditingPassword = False
            _PasswordChanged = False
            _IsPasswordVisible = False
            ClearTextBoxValue()
            UpdateVisualState()
        End Set
    End Property

    ''' <summary>
    ''' Gets the replacement password currently entered by the user.
    ''' </summary>
    ''' <remarks>
    ''' If an existing password is defined but has not been placed into replacement mode,
    ''' this property returns an empty string because the existing password is intentionally not loaded.
    ''' </remarks>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property Password As String
        Get
            If _PasswordDefined AndAlso Not _IsEditingPassword Then Return String.Empty
            Return MyBase.Text
        End Get
    End Property

    ''' <summary>
    ''' Gets a value indicating whether the user has modified the replacement password value.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property PasswordChanged As Boolean
        Get
            Return _PasswordChanged
        End Get
    End Property

    ''' <summary>
    ''' Gets a value indicating whether the control is currently accepting a replacement password.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsEditingPassword As Boolean
        Get
            Return _IsEditingPassword
        End Get
    End Property

    ''' <summary>
    ''' Gets a value indicating whether the current replacement password is visible as plain text.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsPasswordVisible As Boolean
        Get
            Return _IsPasswordVisible
        End Get
    End Property

    ''' <summary>
    ''' Gets or sets a value indicating whether password editing and replacement are disabled.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Indicates whether password editing and replacement are disabled.")>
    <DefaultValue(False)>
    Public Shadows Property [ReadOnly] As Boolean
        Get
            Return _ReadOnly
        End Get
        Set(value As Boolean)
            If _ReadOnly = value Then Return
            _ReadOnly = value
            If _ReadOnly Then
                _IsPasswordVisible = False
            End If
            UpdateVisualState()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets a value indicating whether the user can reveal the replacement password.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Indicates whether the replacement password can be revealed by the action button.")>
    <DefaultValue(True)>
    Public Property AllowPasswordReveal As Boolean
        Get
            Return _AllowPasswordReveal
        End Get
        Set(value As Boolean)
            If _AllowPasswordReveal = value Then Return
            _AllowPasswordReveal = value
            If Not value Then
                _IsPasswordVisible = False
            End If
            UpdateVisualState()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets a value indicating whether the embedded action button is displayed.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Indicates whether the embedded password action button is displayed.")>
    <DefaultValue(True)>
    Public Property ShowActionButton As Boolean
        Get
            Return _ShowActionButton
        End Get
        Set(value As Boolean)
            If _ShowActionButton = value Then Return
            _ShowActionButton = value
            _ActionButton.Visible = value
            UpdateActionButtonBounds()
            UpdateTextMargin()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the width, in pixels, of the embedded action button.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Specifies the width, in pixels, of the embedded password action button.")>
    <DefaultValue(DefaultButtonWidth)>
    Public Property ActionButtonWidth As Integer
        Get
            Return _ButtonWidth
        End Get
        Set(value As Integer)
            If value < MinimumButtonWidth Then
                Throw New ArgumentOutOfRangeException(NameOf(value), $"The button width must be at least {MinimumButtonWidth} pixels.")
            End If
            If _ButtonWidth = value Then Return
            _ButtonWidth = value
            UpdateActionButtonBounds()
            UpdateTextMargin()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the image displayed when the action button can reveal the password.
    ''' </summary>
    ''' <remarks>
    ''' The control does not dispose this image because ownership remains with the caller or designer resources.
    ''' </remarks>
    <Category("PasswordBox")>
    <Description("Specifies the image displayed by the action button when the password can be revealed.")>
    Public Property ShowPasswordImage As Image
        Get
            Return _ShowPasswordImage
        End Get
        Set(value As Image)
            If ReferenceEquals(_ShowPasswordImage, value) Then Return
            _ShowPasswordImage = value
            UpdateActionButtonImage()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the image displayed when the action button can hide the visible password.
    ''' </summary>
    ''' <remarks>
    ''' The control does not dispose this image because ownership remains with the caller or designer resources.
    ''' </remarks>
    <Category("PasswordBox")>
    <Description("Specifies the image displayed by the action button when the visible password can be hidden.")>
    Public Property HidePasswordImage As Image
        Get
            Return _HidePasswordImage
        End Get
        Set(value As Image)
            If ReferenceEquals(_HidePasswordImage, value) Then Return
            _HidePasswordImage = value
            UpdateActionButtonImage()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the image displayed when the action button can start password replacement mode.
    ''' </summary>
    ''' <remarks>
    ''' The control does not dispose this image because ownership remains with the caller or designer resources.
    ''' </remarks>
    <Category("PasswordBox")>
    <Description("Specifies the image displayed by the action button when an existing password can be replaced.")>
    Public Property ChangePasswordImage As Image
        Get
            Return _ChangePasswordImage
        End Get
        Set(value As Image)
            If ReferenceEquals(_ChangePasswordImage, value) Then Return
            _ChangePasswordImage = value
            UpdateActionButtonImage()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the alignment of the state image within the embedded action button.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Specifies the alignment of the state image within the embedded action button.")>
    <DefaultValue(GetType(ContentAlignment), "MiddleCenter")>
    Public Property ActionButtonImageAlign As ContentAlignment
        Get
            Return _ActionButton.ImageAlign
        End Get
        Set(value As ContentAlignment)
            _ActionButton.ImageAlign = value
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the number of mask characters displayed when an existing password is defined.
    ''' </summary>
    ''' <remarks>
    ''' The displayed length is intentionally independent of the stored password length.
    ''' </remarks>
    <Category("PasswordBox")>
    <Description("Specifies the fixed number of mask characters displayed for an existing password.")>
    <DefaultValue(DefaultMaskLength)>
    Public Property DefinedPasswordMaskLength As Integer
        Get
            Return _DefinedPasswordMaskLength
        End Get
        Set(value As Integer)
            If value < MinimumMaskLength Then
                Throw New ArgumentOutOfRangeException(NameOf(value), $"The mask length must be at least {MinimumMaskLength}.")
            End If
            If _DefinedPasswordMaskLength = value Then Return
            _DefinedPasswordMaskLength = value
            If _PasswordDefined AndAlso Not _IsEditingPassword Then
                UpdateDefinedPasswordPlaceholder()
            End If
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the tooltip text displayed when the action button starts password replacement mode.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Specifies the tooltip text used by the change-password action.")>
    <DefaultValue("Change password")>
    <Localizable(True)>
    Public Property ChangePasswordToolTipText As String
        Get
            Return _ChangePasswordToolTipText
        End Get
        Set(value As String)
            _ChangePasswordToolTipText = If(value, String.Empty)
            UpdateActionButtonToolTip()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the tooltip text displayed when the action button can reveal the password.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Specifies the tooltip text used by the show-password action.")>
    <DefaultValue("Show password")>
    <Localizable(True)>
    Public Property ShowPasswordToolTipText As String
        Get
            Return _ShowPasswordToolTipText
        End Get
        Set(value As String)
            _ShowPasswordToolTipText = If(value, String.Empty)
            UpdateActionButtonToolTip()
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the tooltip text displayed when the action button can hide the password.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Specifies the tooltip text used by the hide-password action.")>
    <DefaultValue("Hide password")>
    <Localizable(True)>
    Public Property HidePasswordToolTipText As String
        Get
            Return _HidePasswordToolTipText
        End Get
        Set(value As String)
            _HidePasswordToolTipText = If(value, String.Empty)
            UpdateActionButtonToolTip()
        End Set
    End Property

    ''' <summary>
    ''' Hides the inherited text property because password values should be accessed through <see cref="Password"/>.
    ''' </summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows Property Text As String
        Get
            Return MyBase.Text
        End Get
        Set(value As String)
            MyBase.Text = If(value, String.Empty)
        End Set
    End Property

    ''' <summary>
    ''' Hides the inherited password character property because password masking is managed internally.
    ''' </summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows Property UseSystemPasswordChar As Boolean
        Get
            Return MyBase.UseSystemPasswordChar
        End Get
        Set(value As Boolean)
            MyBase.UseSystemPasswordChar = value
        End Set
    End Property

    ''' <summary>
    ''' Starts password replacement mode and clears any visual placeholder for an existing password.
    ''' </summary>
    ''' <returns><see langword="True"/> when replacement mode is active; otherwise, <see langword="False"/>.</returns>
    Public Function BeginPasswordChange() As Boolean
        If _ReadOnly OrElse Not Enabled Then Return False
        If Not _PasswordDefined Then
            _IsEditingPassword = True
            UpdateVisualState()
            Focus()
            Return True
        End If
        If _IsEditingPassword Then
            Focus()
            Return True
        End If
        _IsEditingPassword = True
        _PasswordChanged = False
        _IsPasswordVisible = False
        ClearTextBoxValue()
        UpdateVisualState()
        Focus()
        RaiseEvent PasswordEditStarted(Me, EventArgs.Empty)
        Return True
    End Function

    ''' <summary>
    ''' Cancels a pending password replacement and restores the defined-password placeholder when applicable.
    ''' </summary>
    Public Sub CancelPasswordChange()
        Dim WasEditing As Boolean = _IsEditingPassword OrElse _PasswordChanged
        _IsEditingPassword = False
        _PasswordChanged = False
        _IsPasswordVisible = False
        ClearTextBoxValue()
        UpdateVisualState()
        If WasEditing Then
            RaiseEvent PasswordEditCanceled(Me, EventArgs.Empty)
        End If
    End Sub

    ''' <summary>
    ''' Marks the current replacement password as successfully persisted by the application.
    ''' </summary>
    ''' <param name="PasswordDefined">
    ''' Indicates whether a password remains defined after the application persists the change.
    ''' </param>
    ''' <remarks>
    ''' This method clears the replacement password from the text box and resets change tracking.
    ''' It does not persist or protect the password itself.
    ''' </remarks>
    Public Sub AcceptPasswordChange(Optional PasswordDefined As Boolean = True)
        _PasswordDefined = PasswordDefined
        _IsEditingPassword = False
        _PasswordChanged = False
        _IsPasswordVisible = False
        ClearTextBoxValue()
        UpdateVisualState()
    End Sub

    ''' <summary>
    ''' Sets a replacement password programmatically and optionally marks it as changed.
    ''' </summary>
    ''' <param name="Password">The replacement password to place in the editor.</param>
    ''' <param name="MarkAsChanged">
    ''' Indicates whether <see cref="PasswordChanged"/> should be set to <see langword="True"/>.
    ''' </param>
    ''' <remarks>
    ''' This method is intended for replacement values. It should not be used to load an existing stored password
    ''' unless the caller explicitly accepts keeping that password in the control's managed memory.
    ''' </remarks>
    Public Sub SetPassword(Password As String, Optional MarkAsChanged As Boolean = True)
        If _ReadOnly Then
            Throw New InvalidOperationException("The password cannot be set while the control is read-only.")
        End If
        _IsEditingPassword = True
        _IsPasswordVisible = False
        _SuppressTextChanged = True
        Try
            MyBase.Text = If(Password, String.Empty)
            SelectionStart = TextLength
        Finally
            _SuppressTextChanged = False
        End Try
        _PasswordChanged = MarkAsChanged
        UpdateVisualState()
        If MarkAsChanged Then
            RaiseEvent PasswordValueChanged(Me, EventArgs.Empty)
        End If
    End Sub

    ''' <summary>
    ''' Clears the current replacement password from the editor without changing <see cref="PasswordDefined"/>.
    ''' </summary>
    Public Sub ClearPassword()
        If _ReadOnly Then Return
        If _PasswordDefined AndAlso Not _IsEditingPassword Then
            BeginPasswordChange()
        End If
        Clear()
    End Sub

    ''' <summary>
    ''' Selects all characters in the replacement password editor when the editor is active.
    ''' </summary>
    Public Sub SelectAllPassword()
        If _PasswordDefined AndAlso Not _IsEditingPassword Then Return
        SelectAll()
    End Sub

    ''' <summary>
    ''' Gives input focus to the password editor.
    ''' </summary>
    ''' <returns><see langword="True"/> if focus was assigned; otherwise, <see langword="False"/>.</returns>
    Public Function FocusPassword() As Boolean
        Return Focus()
    End Function

    ''' <summary>
    ''' Releases the unmanaged resources used by the control and optionally releases the managed resources.
    ''' </summary>
    ''' <param name="Disposing">
    ''' <see langword="True"/> to release both managed and unmanaged resources; otherwise, <see langword="False"/>.
    ''' </param>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            RemoveHandler _ActionButton.Click, AddressOf ActionButton_Click
            _ToolTip.Dispose()
        End If
        MyBase.Dispose(Disposing)
    End Sub

    ''' <summary>
    ''' Raises the <see cref="Control.HandleCreated"/> event and initializes the embedded button layout.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        UpdateActionButtonBounds()
        UpdateTextMargin()
    End Sub

    ''' <summary>
    ''' Raises the <see cref="Control.Resize"/> event and repositions the embedded action button.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        UpdateActionButtonBounds()
        UpdateTextMargin()
    End Sub

    ''' <summary>
    ''' Raises the <see cref="Control.FontChanged"/> event and recalculates the embedded action button layout.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnFontChanged(e As EventArgs)
        MyBase.OnFontChanged(e)
        UpdateActionButtonBounds()
        UpdateTextMargin()
    End Sub

    ''' <summary>
    ''' Raises the <see cref="Control.BackColorChanged"/> event and updates the embedded action button background.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnBackColorChanged(e As EventArgs)
        MyBase.OnBackColorChanged(e)
        If _ActionButton IsNot Nothing Then
            _ActionButton.BackColor = BackColor
        End If
    End Sub

    ''' <summary>
    ''' Raises the <see cref="Control.ForeColorChanged"/> event and updates the embedded action button foreground.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnForeColorChanged(e As EventArgs)
        MyBase.OnForeColorChanged(e)
        If _ActionButton IsNot Nothing Then
            _ActionButton.ForeColor = ForeColor
        End If
    End Sub

    ''' <summary>
    ''' Raises the <see cref="Control.EnabledChanged"/> event and refreshes the visual state.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnEnabledChanged(e As EventArgs)
        MyBase.OnEnabledChanged(e)
        If Not Enabled Then
            _IsPasswordVisible = False
        End If
        UpdateVisualState()
    End Sub

    ''' <summary>
    ''' Processes keyboard input used to start or cancel password replacement.
    ''' </summary>
    ''' <param name="e">The keyboard event data.</param>
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        If e.KeyCode = Keys.F2 AndAlso _PasswordDefined AndAlso Not _IsEditingPassword AndAlso Not _ReadOnly Then
            BeginPasswordChange()
            e.SuppressKeyPress = True
            Return
        End If
        If e.KeyCode = Keys.Escape AndAlso _PasswordDefined AndAlso _IsEditingPassword Then
            CancelPasswordChange()
            e.SuppressKeyPress = True
            Return
        End If
        MyBase.OnKeyDown(e)
    End Sub

    ''' <summary>
    ''' Raises the <see cref="TextBoxBase.TextChanged"/> event and updates password change tracking.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnTextChanged(e As EventArgs)
        If _SuppressTextChanged Then Return
        MyBase.OnTextChanged(e)
        If _PasswordDefined AndAlso Not _IsEditingPassword Then Return
        _PasswordChanged = True
        RaiseEvent PasswordValueChanged(Me, EventArgs.Empty)
    End Sub

    ''' <summary>
    ''' Executes the context-sensitive action associated with the embedded button.
    ''' </summary>
    ''' <param name="sender">The embedded action button.</param>
    ''' <param name="e">The event data.</param>
    Private Sub ActionButton_Click(sender As Object, e As EventArgs)
        If _ReadOnly OrElse Not Enabled Then Return
        If _PasswordDefined AndAlso Not _IsEditingPassword Then
            BeginPasswordChange()
            Return
        End If
        If Not _AllowPasswordReveal Then Return
        TogglePasswordVisibility()
    End Sub

    ''' <summary>
    ''' Toggles the visibility of the currently entered replacement password.
    ''' </summary>
    Private Sub TogglePasswordVisibility()
        _IsPasswordVisible = Not _IsPasswordVisible
        UpdateVisualState()
        Focus()
        SelectionStart = TextLength
        RaiseEvent PasswordVisibilityChanged(Me, EventArgs.Empty)
    End Sub

    ''' <summary>
    ''' Clears the text box while preventing the operation from being tracked as a password change.
    ''' </summary>
    Private Sub ClearTextBoxValue()
        _SuppressTextChanged = True
        Try
            MyBase.Clear()
        Finally
            _SuppressTextChanged = False
        End Try
    End Sub

    ''' <summary>
    ''' Writes the fixed-length placeholder used to represent an existing password.
    ''' </summary>
    Private Sub UpdateDefinedPasswordPlaceholder()
        _SuppressTextChanged = True
        Try
            MyBase.UseSystemPasswordChar = False
            MyBase.Text = New String("•"c, _DefinedPasswordMaskLength)
            SelectionStart = 0
            SelectionLength = 0
        Finally
            _SuppressTextChanged = False
        End Try
    End Sub

    ''' <summary>
    ''' Applies the current control state to the text box, action button, and tooltip.
    ''' </summary>
    Private Sub UpdateVisualState()
        If _ActionButton Is Nothing Then Return
        Dim ShowingDefinedPlaceholder As Boolean = _PasswordDefined AndAlso Not _IsEditingPassword
        MyBase.ReadOnly = _ReadOnly OrElse ShowingDefinedPlaceholder
        _ActionButton.Enabled = Enabled AndAlso Not _ReadOnly
        _ActionButton.Visible = _ShowActionButton
        If ShowingDefinedPlaceholder Then
            UpdateDefinedPasswordPlaceholder()
            _ActionButton.ActionKind = PasswordActionKind.Change
        Else
            MyBase.UseSystemPasswordChar = Not _IsPasswordVisible
            _ActionButton.ActionKind = If(_IsPasswordVisible, PasswordActionKind.Hide, PasswordActionKind.Show)
            If Not _AllowPasswordReveal Then
                _ActionButton.Enabled = False
            End If
        End If
        _ActionButton.BackColor = If(Enabled, BackColor, SystemColors.Control)
        _ActionButton.ForeColor = If(Enabled, ForeColor, SystemColors.GrayText)
        UpdateActionButtonImage()
        UpdateActionButtonToolTip()
        UpdateActionButtonBounds()
        UpdateTextMargin()
    End Sub

    ''' <summary>
    ''' Updates the image displayed by the embedded action button according to the current action.
    ''' </summary>
    Private Sub UpdateActionButtonImage()
        If _ActionButton Is Nothing Then Return
        Select Case _ActionButton.ActionKind
            Case PasswordActionKind.Change
                _ActionButton.Image = _ChangePasswordImage
            Case PasswordActionKind.Hide
                _ActionButton.Image = _HidePasswordImage
            Case Else
                _ActionButton.Image = _ShowPasswordImage
        End Select
    End Sub

    ''' <summary>
    ''' Updates the tooltip text of the embedded action button according to the current action.
    ''' </summary>
    Private Sub UpdateActionButtonToolTip()
        If _ToolTip Is Nothing OrElse _ActionButton Is Nothing Then Return
        Dim ToolTipText As String
        Select Case _ActionButton.ActionKind
            Case PasswordActionKind.Change
                ToolTipText = _ChangePasswordToolTipText
            Case PasswordActionKind.Hide
                ToolTipText = _HidePasswordToolTipText
            Case Else
                ToolTipText = _ShowPasswordToolTipText
        End Select
        _ToolTip.SetToolTip(_ActionButton, ToolTipText)
        _ActionButton.AccessibleName = ToolTipText
        _ActionButton.AccessibleDescription = ToolTipText
    End Sub

    ''' <summary>
    ''' Positions the embedded action button at the right edge of the text box client area.
    ''' </summary>
    Private Sub UpdateActionButtonBounds()
        If _ActionButton Is Nothing Then Return
        Dim ButtonWidth As Integer = If(_ShowActionButton, Math.Min(_ButtonWidth, ClientSize.Width), 0)
        Dim ButtonLeft As Integer = Math.Max(0, ClientSize.Width - ButtonWidth)
        _ActionButton.Bounds = New Rectangle(ButtonLeft, 0, ButtonWidth, ClientSize.Height)
    End Sub

    ''' <summary>
    ''' Updates the native right text margin so entered text does not overlap the embedded action button.
    ''' </summary>
    Private Sub UpdateTextMargin()
        If Not IsHandleCreated Then Return
        Dim RightMargin As Integer = If(_ShowActionButton, _ButtonWidth + ButtonMargin, 0)
        Dim MarginValue As Integer = (RightMargin And &HFFFF) << 16
        SendMessage(Handle, EmSetMargins, New IntPtr(EcRightMargin), New IntPtr(MarginValue))
    End Sub

    ''' <summary>
    ''' Sends a native Windows message to the text box control.
    ''' </summary>
    ''' <param name="hWnd">The handle of the target window.</param>
    ''' <param name="Msg">The message identifier.</param>
    ''' <param name="wParam">The first message parameter.</param>
    ''' <param name="lParam">The second message parameter.</param>
    ''' <returns>The result returned by the native window procedure.</returns>
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    ''' <summary>
    ''' Identifies the action represented by the embedded password button.
    ''' </summary>
    Private Enum PasswordActionKind

        ''' <summary>
        ''' Reveals the replacement password.
        ''' </summary>
        Show

        ''' <summary>
        ''' Hides the replacement password.
        ''' </summary>
        Hide

        ''' <summary>
        ''' Starts replacement of an existing password.
        ''' </summary>
        Change

    End Enum

    ''' <summary>
    ''' Provides the embedded image button used by <see cref="PasswordBox"/>.
    ''' </summary>
    Private NotInheritable Class PasswordActionButton
        Inherits Button

        Private _ActionKind As PasswordActionKind

        ''' <summary>
        ''' Initializes a new instance of the <see cref="PasswordActionButton"/> class.
        ''' </summary>
        Public Sub New()
            FlatStyle = FlatStyle.Flat
            FlatAppearance.BorderSize = 0
            FlatAppearance.MouseOverBackColor = SystemColors.ControlLight
            FlatAppearance.MouseDownBackColor = SystemColors.ControlLightLight
            UseVisualStyleBackColor = False
            Text = String.Empty
            ImageAlign = ContentAlignment.MiddleCenter
            TabStop = False
            SetStyle(ControlStyles.Selectable, False)
        End Sub

        ''' <summary>
        ''' Gets or sets the action currently represented by the button.
        ''' </summary>
        Public Property ActionKind As PasswordActionKind
            Get
                Return _ActionKind
            End Get
            Set(value As PasswordActionKind)
                If _ActionKind = value Then Return
                _ActionKind = value
                Invalidate()
            End Set
        End Property

    End Class

End Class
