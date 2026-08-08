Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports CoreSuite.Controls.My.Resources

''' <summary>
''' Provides a reusable password-entry control with an embedded action button for revealing,
''' hiding, or replacing a password without requiring an existing password value to be loaded.
''' </summary>
<DefaultEvent("PasswordValueChanged")>
<DesignerCategory("Code")>
Public Class PasswordBox
    Inherits UserControl
    Private Const DefaultControlHeight As Integer = 27
    Private Const DefaultButtonWidth As Integer = 30
    Private Const DefaultMaskLength As Integer = 8
    Private Const MinimumButtonWidth As Integer = 24
    Private Const MinimumMaskLength As Integer = 1
    Private Const HorizontalTextPadding As Integer = 6
    Private Const BorderThickness As Integer = 1
    Private ReadOnly _TextBox As TextBox
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
    Private _BorderColor As Color = SystemColors.ControlDark
    Private _FocusedBorderColor As Color = SystemColors.Highlight
    Private _DisabledBorderColor As Color = SystemColors.ControlDark
    Private _ChangePasswordToolTipText As String = "Change password"
    Private _ShowPasswordToolTipText As String = "Show password"
    Private _HidePasswordToolTipText As String = "Hide password"

    ''' <summary>
    ''' Initializes a new instance of the <see cref="PasswordBox"/> class.
    ''' </summary>
    Public Sub New()
        SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.OptimizedDoubleBuffer Or ControlStyles.ResizeRedraw Or ControlStyles.UserPaint, True)
        BackColor = SystemColors.Window
        ForeColor = SystemColors.WindowText
        MinimumSize = New Size(80, DefaultControlHeight)
        Size = New Size(180, DefaultControlHeight)
        TabStop = False
        _TextBox = New TextBox With {
            .BorderStyle = BorderStyle.None,
            .UseSystemPasswordChar = True,
            .TabStop = True
        }
        _ActionButton = New PasswordActionButton With {
            .TabStop = False,
            .Cursor = Cursors.Hand
        }
        _ToolTip = New ToolTip()
        Controls.Add(_TextBox)
        Controls.Add(_ActionButton)
        AddHandler _TextBox.TextChanged, AddressOf TextBox_TextChanged
        AddHandler _TextBox.Enter, AddressOf ChildControl_Enter
        AddHandler _TextBox.Leave, AddressOf ChildControl_Leave
        AddHandler _TextBox.KeyDown, AddressOf TextBox_KeyDown
        AddHandler _ActionButton.Click, AddressOf ActionButton_Click
        AddHandler _ActionButton.Enter, AddressOf ChildControl_Enter
        AddHandler _ActionButton.Leave, AddressOf ChildControl_Leave
        UpdateVisualState()
        PerformLayout()
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
            Return _TextBox.Text
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
    ''' Gets or sets a value indicating whether the password text can be edited.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Indicates whether password editing and replacement are disabled.")>
    <DefaultValue(False)>
    Public Property [ReadOnly] As Boolean
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
            PerformLayout()
            Invalidate()
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
            PerformLayout()
            Invalidate()
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
    ''' Gets or sets the maximum number of characters the user can enter as a replacement password.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Specifies the maximum number of characters allowed in the replacement password.")>
    <DefaultValue(32767)>
    Public Property MaxLength As Integer
        Get
            Return _TextBox.MaxLength
        End Get
        Set(value As Integer)
            If value < 0 Then
                Throw New ArgumentOutOfRangeException(NameOf(value), "MaxLength cannot be negative.")
            End If
            _TextBox.MaxLength = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the horizontal alignment of the replacement password text.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Specifies the horizontal alignment of the replacement password text.")>
    <DefaultValue(GetType(HorizontalAlignment), "Left")>
    Public Property TextAlign As HorizontalAlignment
        Get
            Return _TextBox.TextAlign
        End Get
        Set(value As HorizontalAlignment)
            _TextBox.TextAlign = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the internal text box processes standard shortcut keys.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Indicates whether standard text editing shortcut keys are enabled.")>
    <DefaultValue(True)>
    Public Property ShortcutsEnabled As Boolean
        Get
            Return _TextBox.ShortcutsEnabled
        End Get
        Set(value As Boolean)
            _TextBox.ShortcutsEnabled = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the placeholder text displayed when no password is defined and the field is empty.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Specifies the placeholder text displayed when no password is defined and the field is empty.")>
    <DefaultValue("")>
    <Localizable(True)>
    Public Property PlaceholderText As String
        Get
            Return _TextBox.PlaceholderText
        End Get
        Set(value As String)
            _TextBox.PlaceholderText = If(value, String.Empty)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the border color used while the control does not contain focus.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Specifies the border color used while the control does not contain focus.")>
    Public Property BorderColor As Color
        Get
            Return _BorderColor
        End Get
        Set(value As Color)
            If _BorderColor = value Then Return
            _BorderColor = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the border color used while the control contains focus.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Specifies the border color used while the control contains focus.")>
    Public Property FocusedBorderColor As Color
        Get
            Return _FocusedBorderColor
        End Get
        Set(value As Color)
            If _FocusedBorderColor = value Then Return
            _FocusedBorderColor = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the border color used while the control is disabled.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Specifies the border color used while the control is disabled.")>
    Public Property DisabledBorderColor As Color
        Get
            Return _DisabledBorderColor
        End Get
        Set(value As Color)
            If _DisabledBorderColor = value Then Return
            _DisabledBorderColor = value
            Invalidate()
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
    ''' Gets or sets the background color of the password field.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Specifies the background color of the password field.")>
    Public Overrides Property BackColor As Color
        Get
            Return MyBase.BackColor
        End Get
        Set(value As Color)
            MyBase.BackColor = value
            If _TextBox IsNot Nothing Then _TextBox.BackColor = value
            If _ActionButton IsNot Nothing Then _ActionButton.BackColor = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the foreground color of the password text and action glyph.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Specifies the foreground color of the password text and action glyph.")>
    Public Overrides Property ForeColor As Color
        Get
            Return MyBase.ForeColor
        End Get
        Set(value As Color)
            MyBase.ForeColor = value
            If _TextBox IsNot Nothing Then _TextBox.ForeColor = value
            If _ActionButton IsNot Nothing Then _ActionButton.ForeColor = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the font used by the password text box.
    ''' </summary>
    <Category("PasswordBox")>
    <Description("Specifies the font used by the password field.")>
    <Localizable(True)>
    Public Overrides Property Font As Font
        Get
            Return MyBase.Font
        End Get
        Set(value As Font)
            MyBase.Font = value
            If _TextBox IsNot Nothing Then _TextBox.Font = value
            PerformLayout()
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
            MyBase.Text = value
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
            _TextBox.Focus()
            Return True
        End If
        If _IsEditingPassword Then
            _TextBox.Focus()
            Return True
        End If
        _IsEditingPassword = True
        _PasswordChanged = False
        _IsPasswordVisible = False
        ClearTextBoxValue()
        UpdateVisualState()
        _TextBox.Focus()
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
            _TextBox.Text = If(Password, String.Empty)
            _TextBox.SelectionStart = _TextBox.TextLength
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
        _TextBox.Clear()
    End Sub
    ''' <summary>
    ''' Selects all characters in the replacement password editor when the editor is active.
    ''' </summary>
    Public Sub SelectAllPassword()
        If _PasswordDefined AndAlso Not _IsEditingPassword Then Return
        _TextBox.SelectAll()
    End Sub
    ''' <summary>
    ''' Gives input focus to the internal password editor.
    ''' </summary>
    ''' <returns><see langword="True"/> if focus was assigned; otherwise, <see langword="False"/>.</returns>
    Public Function FocusPassword() As Boolean
        Return _TextBox.Focus()
    End Function
    ''' <summary>
    ''' Releases the unmanaged resources used by the control and optionally releases the managed resources.
    ''' </summary>
    ''' <param name="Disposing">
    ''' <see langword="True"/> to release both managed and unmanaged resources; otherwise, <see langword="False"/>.
    ''' </param>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            RemoveHandler _TextBox.TextChanged, AddressOf TextBox_TextChanged
            RemoveHandler _TextBox.Enter, AddressOf ChildControl_Enter
            RemoveHandler _TextBox.Leave, AddressOf ChildControl_Leave
            RemoveHandler _TextBox.KeyDown, AddressOf TextBox_KeyDown
            RemoveHandler _ActionButton.Click, AddressOf ActionButton_Click
            RemoveHandler _ActionButton.Enter, AddressOf ChildControl_Enter
            RemoveHandler _ActionButton.Leave, AddressOf ChildControl_Leave
            _ToolTip.Dispose()
        End If
        MyBase.Dispose(Disposing)
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
    ''' Raises the <see cref="Control.PaddingChanged"/> event and recalculates child control bounds.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnPaddingChanged(e As EventArgs)
        MyBase.OnPaddingChanged(e)
        PerformLayout()
    End Sub
    ''' <summary>
    ''' Raises the <see cref="Control.Layout"/> event and positions the embedded action button and text box.
    ''' </summary>
    ''' <param name="Levent">The layout event data.</param>
    Protected Overrides Sub OnLayout(Levent As LayoutEventArgs)
        MyBase.OnLayout(Levent)
        If _TextBox Is Nothing OrElse _ActionButton Is Nothing Then Return
        Dim InnerLeft As Integer = BorderThickness + Padding.Left
        Dim InnerTop As Integer = BorderThickness + Padding.Top
        Dim InnerHeight As Integer = Math.Max(1, ClientSize.Height - (BorderThickness * 2) - Padding.Vertical)
        Dim ButtonWidth As Integer = If(_ShowActionButton, Math.Min(_ButtonWidth, Math.Max(0, ClientSize.Width - BorderThickness * 2)), 0)
        Dim ButtonLeft As Integer = ClientSize.Width - BorderThickness - Padding.Right - ButtonWidth
        _ActionButton.Bounds = New Rectangle(ButtonLeft, InnerTop, ButtonWidth, InnerHeight)
        Dim TextLeft As Integer = InnerLeft + HorizontalTextPadding
        Dim TextRight As Integer = If(ButtonWidth > 0, ButtonLeft - HorizontalTextPadding, ClientSize.Width - BorderThickness - Padding.Right - HorizontalTextPadding)
        Dim TextWidth As Integer = Math.Max(0, TextRight - TextLeft)
        Dim PreferredTextHeight As Integer = _TextBox.PreferredHeight
        Dim TextTop As Integer = Math.Max(InnerTop, InnerTop + ((InnerHeight - PreferredTextHeight) \ 2))
        _TextBox.Bounds = New Rectangle(TextLeft, TextTop, TextWidth, PreferredTextHeight)
    End Sub

    ''' <summary>
    ''' Paints the outer border according to the current focus and enabled state.
    ''' </summary>
    ''' <param name="e">The paint event data.</param>
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        Dim BorderColor As Color
        If Not Enabled Then
            BorderColor = _DisabledBorderColor
        ElseIf ContainsFocus Then
            BorderColor = _FocusedBorderColor
        Else
            BorderColor = _BorderColor
        End If
        Using BorderPen As New Pen(BorderColor)
            Dim BorderRectangle As New Rectangle(0, 0, Math.Max(0, ClientSize.Width - 1), Math.Max(0, ClientSize.Height - 1))
            e.Graphics.DrawRectangle(BorderPen, BorderRectangle)
        End Using
    End Sub
    ''' <summary>
    ''' Processes keyboard input used to cancel an active password replacement operation.
    ''' </summary>
    ''' <param name="sender">The internal text box that raised the event.</param>
    ''' <param name="e">The keyboard event data.</param>
    Private Sub TextBox_KeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.F2 AndAlso _PasswordDefined AndAlso Not _IsEditingPassword AndAlso Not _ReadOnly Then
            BeginPasswordChange()
            e.SuppressKeyPress = True
            Return
        End If
        If e.KeyCode = Keys.Escape AndAlso _PasswordDefined AndAlso _IsEditingPassword Then
            CancelPasswordChange()
            e.SuppressKeyPress = True
        End If
    End Sub
    ''' <summary>
    ''' Handles changes to the replacement password and updates change tracking.
    ''' </summary>
    ''' <param name="sender">The internal text box that raised the event.</param>
    ''' <param name="e">The event data.</param>
    Private Sub TextBox_TextChanged(sender As Object, e As EventArgs)
        If _SuppressTextChanged Then Return
        If _PasswordDefined AndAlso Not _IsEditingPassword Then Return
        _PasswordChanged = True
        RaiseEvent PasswordValueChanged(Me, EventArgs.Empty)
    End Sub
    ''' <summary>
    ''' Refreshes the border when a child control receives focus.
    ''' </summary>
    ''' <param name="sender">The child control that raised the event.</param>
    ''' <param name="e">The event data.</param>
    Private Sub ChildControl_Enter(sender As Object, e As EventArgs)
        Invalidate()
    End Sub
    ''' <summary>
    ''' Refreshes the border when a child control loses focus.
    ''' </summary>
    ''' <param name="sender">The child control that raised the event.</param>
    ''' <param name="e">The event data.</param>
    Private Sub ChildControl_Leave(sender As Object, e As EventArgs)
        BeginInvoke(Sub() Invalidate())
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
        _TextBox.Focus()
        _TextBox.SelectionStart = _TextBox.TextLength
        RaiseEvent PasswordVisibilityChanged(Me, EventArgs.Empty)
    End Sub
    ''' <summary>
    ''' Clears the internal text box while preventing the operation from being tracked as a password change.
    ''' </summary>
    Private Sub ClearTextBoxValue()
        _SuppressTextChanged = True
        Try
            _TextBox.Clear()
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
            _TextBox.UseSystemPasswordChar = False
            _TextBox.Text = New String("•"c, _DefinedPasswordMaskLength)
            _TextBox.SelectionStart = 0
            _TextBox.SelectionLength = 0
        Finally
            _SuppressTextChanged = False
        End Try
    End Sub
    ''' <summary>
    ''' Applies the current control state to the internal text box, action button, and tooltip.
    ''' </summary>
    Private Sub UpdateVisualState()
        If _TextBox Is Nothing OrElse _ActionButton Is Nothing Then Return
        Dim ShowingDefinedPlaceholder As Boolean = _PasswordDefined AndAlso Not _IsEditingPassword
        _TextBox.ReadOnly = _ReadOnly OrElse ShowingDefinedPlaceholder
        _TextBox.Enabled = Enabled
        _ActionButton.Enabled = Enabled AndAlso Not _ReadOnly
        _ActionButton.Visible = _ShowActionButton
        If ShowingDefinedPlaceholder Then
            UpdateDefinedPasswordPlaceholder()
            _ActionButton.ActionKind = PasswordActionKind.Change
        Else
            _TextBox.UseSystemPasswordChar = Not _IsPasswordVisible
            _ActionButton.ActionKind = If(_IsPasswordVisible, PasswordActionKind.Hide, PasswordActionKind.Show)
            If Not _AllowPasswordReveal Then
                _ActionButton.Enabled = False
            End If
        End If
        _TextBox.BackColor = If(Enabled, BackColor, SystemColors.Control)
        _TextBox.ForeColor = If(Enabled, ForeColor, SystemColors.GrayText)
        _ActionButton.BackColor = _TextBox.BackColor
        _ActionButton.ForeColor = _TextBox.ForeColor
        UpdateActionButtonImage()
        UpdateActionButtonToolTip()
        PerformLayout()
        Invalidate()
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