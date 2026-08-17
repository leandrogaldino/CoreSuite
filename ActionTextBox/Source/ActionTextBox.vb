Imports System.ComponentModel
Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices
''' <summary>
''' Specifies the side on which the action button is displayed.
''' </summary>
Public Enum ActionButtonPosition
    ''' <summary>
    ''' Displays the action button on the left side of the text box.
    ''' </summary>
    Left
    ''' <summary>
    ''' Displays the action button on the right side of the text box.
    ''' </summary>
    Right
End Enum
''' <summary>
''' Represents a standard Windows Forms text box with an integrated configurable action button.
''' </summary>
<DefaultEvent("ActionButtonClick")>
<ToolboxItem(True)>
<Designer(GetType(ActionTextBoxControlDesigner))>
Public Class ActionTextBox
    Inherits TextBox
    Private Const EM_SETMARGINS As Integer = &HD3
    Private Const EC_LEFTMARGIN As Integer = &H1
    Private Const EC_RIGHTMARGIN As Integer = &H2
    Private WithEvents ActionButton As New PictureBox
    Private ReadOnly _ToolTip As New ToolTip
    Private _ActionButtonImage As Image
    Private _ActionButtonPosition As ActionButtonPosition = ActionButtonPosition.Right
    Private _ActionButtonVisible As Boolean = True
    Private _ActionButtonEnabled As Boolean = True
    Private _ActionButtonWidth As Integer = 25
    Private _ActionButtonPadding As Integer = 4
    Private _ActionButtonToolTipText As String = String.Empty
    Private _IsButtonHot As Boolean
    Private _IsButtonPressed As Boolean
    Private _ActionButtonTextSpacing As Integer = 4

    ''' <summary>
    ''' Occurs when the integrated action button is clicked.
    ''' </summary>
    <Category("ActionTextBox")>
    <Description("Occurs when the integrated action button is clicked.")>
    Public Event ActionButtonClick As EventHandler
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ActionTextBox"/> class.
    ''' </summary>
    Public Sub New()
        ActionButton.Cursor = Cursors.Hand
        ActionButton.TabStop = False
        ActionButton.BackColor = BackColor
        ActionButton.SizeMode = PictureBoxSizeMode.Normal
        Controls.Add(ActionButton)
        ActionButton.BringToFront()
        UpdateActionButtonState()
        UpdateActionButtonLayout()
    End Sub
    ''' <summary>
    ''' Gets or sets the image displayed by the action button.
    ''' </summary>
    <Category("ActionTextBox")>
    <Description("Specifies the image displayed by the action button.")>
    <DefaultValue(GetType(Image), Nothing)>
    Public Property ActionButtonImage As Image
        Get
            Return _ActionButtonImage
        End Get
        Set(value As Image)
            If ReferenceEquals(_ActionButtonImage, value) Then Return
            _ActionButtonImage = value
            ActionButton.Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Indicates whether the action button image should be serialized by the Windows Forms designer.
    ''' </summary>
    Private Function ShouldSerializeActionButtonImage() As Boolean
        Return _ActionButtonImage IsNot Nothing
    End Function
    ''' <summary>
    ''' Resets the action button image to its default value.
    ''' </summary>
    Private Sub ResetActionButtonImage()
        ActionButtonImage = Nothing
    End Sub
    ''' <summary>
    ''' Gets or sets the side on which the action button is displayed.
    ''' </summary>
    <Category("ActionTextBox")>
    <Description("Specifies the side on which the action button is displayed.")>
    <DefaultValue(ActionButtonPosition.Right)>
    Public Property ActionButtonPosition As ActionButtonPosition
        Get
            Return _ActionButtonPosition
        End Get
        Set(value As ActionButtonPosition)
            If _ActionButtonPosition = value Then Return
            _ActionButtonPosition = value
            UpdateActionButtonLayout()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the action button is visible.
    ''' </summary>
    <Category("ActionTextBox")>
    <Description("Specifies whether the action button is visible.")>
    <DefaultValue(True)>
    Public Property ActionButtonVisible As Boolean
        Get
            Return _ActionButtonVisible
        End Get
        Set(value As Boolean)
            If _ActionButtonVisible = value Then Return
            _ActionButtonVisible = value
            ActionButton.Visible = value
            UpdateActionButtonLayout()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the action button can be clicked.
    ''' </summary>
    <Category("ActionTextBox")>
    <Description("Specifies whether the action button can be clicked.")>
    <DefaultValue(True)>
    Public Property ActionButtonEnabled As Boolean
        Get
            Return _ActionButtonEnabled
        End Get
        Set(value As Boolean)
            If _ActionButtonEnabled = value Then Return
            _ActionButtonEnabled = value
            UpdateActionButtonState()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the width of the action button in logical pixels.
    ''' </summary>
    ''' <exception cref="ArgumentOutOfRangeException">The assigned value is less than 1.</exception>
    <Category("ActionTextBox")>
    <Description("Specifies the width of the action button in logical pixels.")>
    <DefaultValue(25)>
    Public Property ActionButtonWidth As Integer
        Get
            Return _ActionButtonWidth
        End Get
        Set(value As Integer)
            If value < 1 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "ActionButtonWidth must be greater than zero.")
            If _ActionButtonWidth = value Then Return
            _ActionButtonWidth = value
            UpdateActionButtonLayout()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the padding applied around the action button image in logical pixels.
    ''' </summary>
    ''' <exception cref="ArgumentOutOfRangeException">The assigned value is less than zero.</exception>
    <Category("ActionTextBox")>
    <Description("Specifies the padding applied around the action button image in logical pixels.")>
    <DefaultValue(4)>
    Public Property ActionButtonPadding As Integer
        Get
            Return _ActionButtonPadding
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "ActionButtonPadding cannot be negative.")
            If _ActionButtonPadding = value Then Return
            _ActionButtonPadding = value
            ActionButton.Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the spacing between the text and the action button in logical pixels.
    ''' </summary>
    <Category("ActionTextBox")>
    <Description("Specifies the spacing between the text and the action button in logical pixels.")>
    <DefaultValue(4)>
    Public Property ActionButtonTextSpacing As Integer
        Get
            Return _ActionButtonTextSpacing
        End Get
        Set(value As Integer)
            If value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(value), value, "ActionButtonTextSpacing cannot be negative.")
            If _ActionButtonTextSpacing = value Then Return
            _ActionButtonTextSpacing = value
            UpdateTextMargins()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the tooltip text displayed for the action button.
    ''' </summary>
    <Category("ActionTextBox")>
    <Description("Specifies the tooltip text displayed for the action button.")>
    <DefaultValue("")>
    Public Property ActionButtonToolTipText As String
        Get
            Return _ActionButtonToolTipText
        End Get
        Set(value As String)
            Dim newValue As String = If(value, String.Empty)
            If _ActionButtonToolTipText = newValue Then Return
            _ActionButtonToolTipText = newValue
            _ToolTip.SetToolTip(ActionButton, newValue)
        End Set
    End Property
    ''' <summary>
    ''' Programmatically invokes the action button when it is available.
    ''' </summary>
    Public Sub PerformActionButtonClick()
        If Not Enabled OrElse Not ActionButtonVisible OrElse Not ActionButtonEnabled Then Return
        OnActionButtonClick(EventArgs.Empty)
    End Sub
    ''' <summary>
    ''' Raises the <see cref="ActionButtonClick"/> event.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overridable Sub OnActionButtonClick(e As EventArgs)
        RaiseEvent ActionButtonClick(Me, e)
    End Sub
    ''' <summary>
    ''' Configures the action button and text margins after the native text box handle is created.
    ''' </summary>
    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        UpdateActionButtonLayout()
    End Sub
    ''' <summary>
    ''' Repositions the action button when the text box size changes.
    ''' </summary>
    Protected Overrides Sub OnSizeChanged(e As EventArgs)
        MyBase.OnSizeChanged(e)
        UpdateActionButtonLayout()
    End Sub
    ''' <summary>
    ''' Repositions the action button when the border style changes.
    ''' </summary>
    Protected Overrides Sub OnBorderStyleChanged(e As EventArgs)
        MyBase.OnBorderStyleChanged(e)
        UpdateActionButtonLayout()
    End Sub
    ''' <summary>
    ''' Rescales the action button when the control DPI changes.
    ''' </summary>
    Protected Overrides Sub OnDpiChangedAfterParent(e As EventArgs)
        MyBase.OnDpiChangedAfterParent(e)
        UpdateActionButtonLayout()
    End Sub
    ''' <summary>
    ''' Synchronizes the action button background with the text box background.
    ''' </summary>
    Protected Overrides Sub OnBackColorChanged(e As EventArgs)
        MyBase.OnBackColorChanged(e)
        UpdateActionButtonAppearance()
    End Sub
    ''' <summary>
    ''' Synchronizes the action button enabled state with the text box enabled state.
    ''' </summary>
    Protected Overrides Sub OnEnabledChanged(e As EventArgs)
        MyBase.OnEnabledChanged(e)
        UpdateActionButtonState()
    End Sub
    ''' <summary>
    ''' Releases resources owned by the control.
    ''' </summary>
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing Then _ToolTip.Dispose()
        MyBase.Dispose(disposing)
    End Sub
    ''' <summary>
    ''' Raises the action event when the internal button is clicked.
    ''' </summary>
    Private Sub ActionButton_Click(sender As Object, e As EventArgs) Handles ActionButton.Click
        If Not Enabled OrElse Not ActionButtonEnabled Then Return
        OnActionButtonClick(e)
    End Sub
    ''' <summary>
    ''' Tracks the hot state of the action button.
    ''' </summary>
    Private Sub ActionButton_MouseEnter(sender As Object, e As EventArgs) Handles ActionButton.MouseEnter
        If Not Enabled OrElse Not ActionButtonEnabled Then Return
        _IsButtonHot = True
        UpdateActionButtonAppearance()
    End Sub
    ''' <summary>
    ''' Clears the hot and pressed states when the pointer leaves the action button.
    ''' </summary>
    Private Sub ActionButton_MouseLeave(sender As Object, e As EventArgs) Handles ActionButton.MouseLeave
        _IsButtonHot = False
        _IsButtonPressed = False
        UpdateActionButtonAppearance()
    End Sub
    ''' <summary>
    ''' Tracks the pressed state of the action button.
    ''' </summary>
    Private Sub ActionButton_MouseDown(sender As Object, e As MouseEventArgs) Handles ActionButton.MouseDown
        If e.Button <> MouseButtons.Left OrElse Not Enabled OrElse Not ActionButtonEnabled Then Return
        _IsButtonPressed = True
        UpdateActionButtonAppearance()
    End Sub
    ''' <summary>
    ''' Clears the pressed state of the action button.
    ''' </summary>
    Private Sub ActionButton_MouseUp(sender As Object, e As MouseEventArgs) Handles ActionButton.MouseUp
        If e.Button <> MouseButtons.Left Then Return
        _IsButtonPressed = False
        UpdateActionButtonAppearance()
    End Sub
    ''' <summary>
    ''' Draws the action button image inside the dedicated child control.
    ''' </summary>
    Private Sub ActionButton_Paint(sender As Object, e As PaintEventArgs) Handles ActionButton.Paint
        If ActionButtonImage Is Nothing Then Return
        Dim Padding As Integer = GetScaledValue(ActionButtonPadding)
        Dim ImageBounds As Rectangle = Rectangle.Inflate(ActionButton.ClientRectangle, -Padding, -Padding)
        If ImageBounds.Width <= 0 OrElse ImageBounds.Height <= 0 Then Return
        Dim ImageSize As Size = GetContainedImageSize(ActionButtonImage.Size, ImageBounds.Size)
        If ImageSize.Width <= 0 OrElse ImageSize.Height <= 0 Then Return
        Dim ImageRectangle As New Rectangle(ImageBounds.X + (ImageBounds.Width - ImageSize.Width) \ 2, ImageBounds.Y + (ImageBounds.Height - ImageSize.Height) \ 2, ImageSize.Width, ImageSize.Height)
        e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic
        If Enabled AndAlso ActionButtonEnabled Then
            e.Graphics.DrawImage(ActionButtonImage, ImageRectangle)
        Else
            Using DisabledImage As New Bitmap(ImageRectangle.Width, ImageRectangle.Height)
                Using ImageGraphics As Graphics = Graphics.FromImage(DisabledImage)
                    ImageGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic
                    ImageGraphics.DrawImage(ActionButtonImage, New Rectangle(Point.Empty, DisabledImage.Size))
                End Using
                ControlPaint.DrawImageDisabled(e.Graphics, DisabledImage, ImageRectangle.X, ImageRectangle.Y, ActionButton.BackColor)
            End Using
        End If
    End Sub
    ''' <summary>
    ''' Updates the child action button position, size and native text margins.
    ''' </summary>
    Private Sub UpdateActionButtonLayout()
        If ActionButton Is Nothing Then Return
        Dim buttonWidth As Integer = GetScaledValue(ActionButtonWidth)
        Dim overlap As Integer = If(BorderStyle = BorderStyle.None, 0, 1)
        ActionButton.Size = New Size(buttonWidth, Math.Max(1, ClientSize.Height + overlap * 2))
        If ActionButtonPosition = ActionButtonPosition.Left Then
            ActionButton.Location = New Point(-overlap, -overlap)
        Else
            ActionButton.Location = New Point(ClientSize.Width - ActionButton.Width + overlap, -overlap)
        End If
        ActionButton.Visible = ActionButtonVisible
        ActionButton.BringToFront()
        UpdateTextMargins()
        ActionButton.Invalidate()
    End Sub
    ''' <summary>
    ''' Updates the child action button enabled state and visual state.
    ''' </summary>
    Private Sub UpdateActionButtonState()
        If ActionButton Is Nothing Then Return
        ActionButton.Enabled = Enabled AndAlso ActionButtonEnabled
        If Not ActionButton.Enabled Then
            _IsButtonHot = False
            _IsButtonPressed = False
        End If
        UpdateActionButtonAppearance()
        ActionButton.Invalidate()
    End Sub
    ''' <summary>
    ''' Updates the child action button background for its current interaction state.
    ''' </summary>
    Private Sub UpdateActionButtonAppearance()
        If ActionButton Is Nothing Then Return
        If Not Enabled OrElse Not ActionButtonEnabled Then
            ActionButton.BackColor = SystemColors.Control
        ElseIf _IsButtonPressed Then
            ActionButton.BackColor = SystemColors.ControlDark
        ElseIf _IsButtonHot Then
            ActionButton.BackColor = SystemColors.ControlLight
        Else
            ActionButton.BackColor = BackColor
        End If
        ActionButton.Invalidate()
    End Sub
    ''' <summary>
    ''' Reserves text space for the visible action button through the native edit control margins.
    ''' </summary>
    Private Sub UpdateTextMargins()
        If Not IsHandleCreated Then Return
        Dim ButtonWidth As Integer = If(ActionButtonVisible, ActionButton.Width, 0)
        Dim Spacing As Integer = If(ActionButtonVisible, GetScaledValue(ActionButtonTextSpacing), 0)
        Dim LeftMargin As Integer = If(ActionButtonVisible AndAlso ActionButtonPosition = ActionButtonPosition.Left, ButtonWidth + Spacing, 0)
        Dim RightMargin As Integer = If(ActionButtonVisible AndAlso ActionButtonPosition = ActionButtonPosition.Right, ButtonWidth + Spacing, 0)
        Dim MarginValue As Integer = (RightMargin << 16) Or (LeftMargin And &HFFFF)
        SendMessage(Handle, EM_SETMARGINS, New IntPtr(EC_LEFTMARGIN Or EC_RIGHTMARGIN), New IntPtr(MarginValue))
    End Sub
    ''' <summary>
    ''' Returns a value scaled for the current control DPI.
    ''' </summary>
    Private Function GetScaledValue(value As Integer) As Integer
        Return Math.Max(0, CInt(Math.Round(value * DeviceDpi / 96.0R)))
    End Function
    ''' <summary>
    ''' Returns the largest image size that fits within the specified bounds while preserving aspect ratio.
    ''' </summary>
    Private Shared Function GetContainedImageSize(SourceSize As Size, targetSize As Size) As Size
        If SourceSize.Width <= 0 OrElse SourceSize.Height <= 0 OrElse targetSize.Width <= 0 OrElse targetSize.Height <= 0 Then Return Size.Empty
        Dim Scale As Double = Math.Min(targetSize.Width / CDbl(SourceSize.Width), targetSize.Height / CDbl(SourceSize.Height))
        Return New Size(Math.Max(1, CInt(Math.Round(SourceSize.Width * Scale))), Math.Max(1, CInt(Math.Round(SourceSize.Height * Scale))))
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function
End Class
