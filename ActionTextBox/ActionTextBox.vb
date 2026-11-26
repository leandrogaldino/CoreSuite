Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices
Imports System.Windows.Forms

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
    Inherits CTextBox

    Private _ButtonMouseOverBackColor As Color

    <Category("Appearance"), DefaultValue(GetType(Color), ""), Description("The button background while the pointer is over it. Empty uses the default color.")>
    Public Property ButtonMouseOverBackColor As Color
        Get
            Return _ButtonMouseOverBackColor
        End Get
        Set(value As Color)
            If _ButtonMouseOverBackColor = value Then Return
            _ButtonMouseOverBackColor = value
            UpdateActionButtonAppearance()
        End Set
    End Property

    Private _ButtonMouseDownBackColor As Color

    <Category("Appearance"), DefaultValue(GetType(Color), ""), Description("The button background while it is pressed. Empty uses the default color.")>
    Public Property ButtonMouseDownBackColor As Color
        Get
            Return _ButtonMouseDownBackColor
        End Get
        Set(value As Color)
            If _ButtonMouseDownBackColor = value Then Return
            _ButtonMouseDownBackColor = value
            UpdateActionButtonAppearance()
        End Set
    End Property

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
    Private _ButtonTopExtension As Integer
    Private _ExtensionPressed As Boolean
    Private _ExtensionHot As Boolean
    Private Const WM_PAINT As Integer = &HF
    Private Const WM_NCPAINT As Integer = &H85
    Private Const WM_NCACTIVATE As Integer = &H86
    Private Const WM_SETFOCUS As Integer = &H7
    Private Const WM_KILLFOCUS As Integer = &H8
    Private Const WM_SETCURSOR As Integer = &H20
    Private Const WM_MOUSEMOVE As Integer = &H200
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202
    Private Const WM_LBUTTONDBLCLK As Integer = &H203
    Private Const WM_CAPTURECHANGED As Integer = &H215
    Private Const WM_MOUSELEAVE As Integer = &H2A3

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
            PaintButtonExtension()
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
            PaintButtonExtension()
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

    Protected Overrides Sub OnClientSizeChanged(e As EventArgs)
        MyBase.OnClientSizeChanged(e)
        UpdateActionButtonLayout()
    End Sub

    Protected Overrides Sub OnFontChanged(e As EventArgs)
        MyBase.OnFontChanged(e)
        UpdateActionButtonLayout()
    End Sub

    Protected Overrides Sub OnMultilineChanged(e As EventArgs)
        MyBase.OnMultilineChanged(e)
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
        _ExtensionHot = IsOverActionButton(System.Windows.Forms.Cursor.Position, True) AndAlso Enabled AndAlso ActionButtonEnabled
        _IsButtonHot = _ExtensionHot
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
        Dim Bounds As New Rectangle(0, -_ButtonTopExtension, ActionButton.ClientSize.Width,
                                    ActionButton.ClientSize.Height + _ButtonTopExtension)
        DrawActionImage(e.Graphics, Bounds)
    End Sub

    Private Sub DrawActionImage(Graphics As Graphics, Bounds As Rectangle)
        If ActionButtonImage Is Nothing Then Return
        Dim Padding As Integer = GetScaledValue(ActionButtonPadding)
        Dim ImageBounds As Rectangle = Rectangle.Inflate(Bounds, -Padding, -Padding)
        If ImageBounds.Width <= 0 OrElse ImageBounds.Height <= 0 Then Return
        Dim ImageSize As Size = GetContainedImageSize(ActionButtonImage.Size, ImageBounds.Size)
        If ImageSize.Width <= 0 OrElse ImageSize.Height <= 0 Then Return
        Dim ImageRectangle As New Rectangle(ImageBounds.X + (ImageBounds.Width - ImageSize.Width) \ 2, ImageBounds.Y + (ImageBounds.Height - ImageSize.Height) \ 2, ImageSize.Width, ImageSize.Height)
        Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic
        If Enabled AndAlso ActionButtonEnabled Then
            Graphics.DrawImage(ActionButtonImage, ImageRectangle)
        Else
            Using DisabledImage As New Bitmap(ImageRectangle.Width, ImageRectangle.Height)
                Using ImageGraphics As Graphics = Graphics.FromImage(DisabledImage)
                    ImageGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic
                    ImageGraphics.DrawImage(ActionButtonImage, New Rectangle(Point.Empty, DisabledImage.Size))
                End Using
                ControlPaint.DrawImageDisabled(Graphics, DisabledImage, ImageRectangle.X, ImageRectangle.Y, ActionButton.BackColor)
            End Using
        End If
    End Sub
    ''' <summary>
    ''' Updates the child action button position, size and native text margins.
    ''' </summary>
    Private Sub UpdateActionButtonLayout()
        If ActionButton Is Nothing Then Return
        Dim ButtonWidth As Integer = GetScaledValue(ActionButtonWidth)
        Dim ButtonBounds As Rectangle
        _ButtonTopExtension = 0
        If Not Multiline AndAlso BorderStyle = BorderStyle.FixedSingle AndAlso TryGetFullActionButtonBounds(ButtonWidth, ButtonBounds) Then
            ActionButton.Bounds = ButtonBounds
            ActionButton.Visible = ActionButtonVisible
            ActionButton.BringToFront()
            UpdateTextMargins()
            ActionButton.Invalidate()
            RedrawWindow(Handle, IntPtr.Zero, IntPtr.Zero, &H401UI)
            PaintButtonExtension()
            Return
        End If

        Dim BorderInset As Integer
        Select Case BorderStyle
            Case BorderStyle.FixedSingle
                BorderInset = 1
            Case BorderStyle.Fixed3D
                ' ClientSize already excludes the native 3D frame.
                BorderInset = 0
            Case Else
                BorderInset = 0
        End Select
        ActionButton.Size = New Size(ButtonWidth, Math.Max(1, ClientSize.Height - BorderInset * 2))
        If ActionButtonPosition = ActionButtonPosition.Left Then
            ActionButton.Location = New Point(BorderInset, BorderInset)
        Else
            ActionButton.Location = New Point(ClientSize.Width - ActionButton.Width - BorderInset, BorderInset)
        End If
        ActionButton.Visible = ActionButtonVisible
        ActionButton.BringToFront()
        UpdateTextMargins()
        ActionButton.Invalidate()
    End Sub
    ' The child fills the EDIT client area. The strip above it is part of
    ' CTextBox's non-client frame and must be painted on the parent window.
    Private Function TryGetFullActionButtonBounds(ButtonWidth As Integer, ByRef Bounds As Rectangle) As Boolean
        If Not IsHandleCreated OrElse IsDisposed OrElse Disposing Then Return False
        Dim Info As New NativeWindowInfo With {.Size = CUInt(Marshal.SizeOf(GetType(NativeWindowInfo)))}
        If Not GetWindowInfo(Handle, Info) Then Return False
        Dim ClientWidth As Integer = Info.ClientBounds.Right - Info.ClientBounds.Left
        Dim ClientHeight As Integer = Info.ClientBounds.Bottom - Info.ClientBounds.Top
        If ClientWidth <= 0 OrElse ClientHeight <= 0 Then Return False
        ' CTextBox owns a one-pixel FixedSingle frame. Do not paint over it.
        _ButtonTopExtension = Math.Max(0, Info.ClientBounds.Top - Info.WindowBounds.Top - 1)
        ButtonWidth = Math.Max(1, Math.Min(ButtonWidth, ClientWidth))
        Dim ButtonLeft As Integer = If(ActionButtonPosition = ActionButtonPosition.Left, 0, ClientWidth - ButtonWidth)
        Bounds = New Rectangle(ButtonLeft, 0, ButtonWidth, ClientHeight)
        Return True
    End Function

    Private Function HasButtonExtension() As Boolean
        Return _ButtonTopExtension > 0 AndAlso Not Multiline AndAlso BorderStyle = BorderStyle.FixedSingle AndAlso
               ActionButton IsNot Nothing AndAlso ActionButtonVisible AndAlso ActionButton.Visible AndAlso
               IsHandleCreated AndAlso Not IsDisposed AndAlso Not Disposing
    End Function

    Private Sub PaintButtonExtension()
        If Not HasButtonExtension() Then Return
        Dim Info As New NativeWindowInfo With {.Size = CUInt(Marshal.SizeOf(GetType(NativeWindowInfo)))}
        If Not GetWindowInfo(Handle, Info) Then Return
        Dim X As Integer = Info.ClientBounds.Left - Info.WindowBounds.Left + ActionButton.Left
        Dim Y As Integer = Info.ClientBounds.Top - Info.WindowBounds.Top - _ButtonTopExtension
        If Y < 1 OrElse ActionButton.Width <= 0 Then Return
        Dim DC As IntPtr = GetWindowDC(Handle)
        If DC = IntPtr.Zero Then Return
        Try
            Using G As Graphics = Graphics.FromHdc(DC)
                Dim Strip As New Rectangle(X, Y, ActionButton.Width, _ButtonTopExtension)
                G.SetClip(Strip)
                Using Brush As New SolidBrush(ActionButton.BackColor)
                    G.FillRectangle(Brush, Strip)
                End Using
                DrawActionImage(G, New Rectangle(X, Y, ActionButton.Width, ActionButton.Height + _ButtonTopExtension))
            End Using
        Finally
            ReleaseDC(Handle, DC)
        End Try
    End Sub

    Private Function IsOverActionButton(ScreenPoint As Point, ExtensionOnly As Boolean) As Boolean
        If Not HasButtonExtension() Then Return False
        Dim P As Point = PointToClient(ScreenPoint)
        Dim Bounds As New Rectangle(ActionButton.Left, -_ButtonTopExtension, ActionButton.Width,
                                    If(ExtensionOnly, _ButtonTopExtension, ActionButton.Height + _ButtonTopExtension))
        Return Bounds.Contains(P)
    End Function

    Protected Overrides Sub WndProc(ByRef M As Message)
        ' The base maps the added non-client strip to HTCLIENT. Mouse messages
        ' therefore reach this EDIT with a negative client Y coordinate.
        Select Case M.Msg
            Case WM_SETCURSOR
                If IsOverActionButton(System.Windows.Forms.Cursor.Position, True) AndAlso Enabled AndAlso ActionButtonEnabled Then
                    System.Windows.Forms.Cursor.Current = Cursors.Hand
                    M.Result = New IntPtr(1)
                    Return
                End If
            Case WM_LBUTTONDOWN, WM_LBUTTONDBLCLK
                If IsOverActionButton(System.Windows.Forms.Cursor.Position, True) Then
                    If Enabled AndAlso ActionButtonEnabled Then
                        _ExtensionPressed = True
                        _ExtensionHot = True
                        _IsButtonHot = True
                        _IsButtonPressed = True
                        Capture = True
                        UpdateActionButtonAppearance()
                    End If
                    M.Result = IntPtr.Zero
                    Return
                End If
            Case WM_LBUTTONUP
                If _ExtensionPressed Then
                    Dim ShouldInvoke As Boolean = IsOverActionButton(System.Windows.Forms.Cursor.Position, False) AndAlso Enabled AndAlso ActionButtonEnabled
                    _ExtensionPressed = False
                    _IsButtonPressed = False
                    Capture = False
                    _ExtensionHot = IsOverActionButton(System.Windows.Forms.Cursor.Position, True)
                    _IsButtonHot = IsOverActionButton(System.Windows.Forms.Cursor.Position, False)
                    UpdateActionButtonAppearance()
                    If ShouldInvoke Then OnActionButtonClick(EventArgs.Empty)
                    M.Result = IntPtr.Zero
                    Return
                End If
            Case WM_MOUSEMOVE
                If _ExtensionPressed Then
                    _IsButtonHot = IsOverActionButton(System.Windows.Forms.Cursor.Position, False)
                    _IsButtonPressed = _IsButtonHot
                    UpdateActionButtonAppearance()
                    M.Result = IntPtr.Zero
                    Return
                End If
            Case WM_CAPTURECHANGED
                If _ExtensionPressed Then
                    _ExtensionPressed = False
                    _IsButtonPressed = False
                    _IsButtonHot = IsOverActionButton(System.Windows.Forms.Cursor.Position, False)
                    UpdateActionButtonAppearance()
                End If
        End Select

        MyBase.WndProc(M)
        Select Case M.Msg
            Case WM_PAINT, WM_NCPAINT, WM_NCACTIVATE, WM_SETFOCUS, WM_KILLFOCUS
                PaintButtonExtension()
            Case WM_MOUSEMOVE, WM_MOUSELEAVE
                Dim Hot As Boolean = IsOverActionButton(System.Windows.Forms.Cursor.Position, True) AndAlso Enabled AndAlso ActionButtonEnabled
                If _ExtensionHot <> Hot Then
                    _ExtensionHot = Hot
                    _IsButtonHot = Hot OrElse IsOverActionButton(System.Windows.Forms.Cursor.Position, False)
                    UpdateActionButtonAppearance()
                End If
        End Select
    End Sub

    <StructLayout(LayoutKind.Sequential)>
    Private Structure NativeRect
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Private Structure NativeWindowInfo
        Public Size As UInteger
        Public WindowBounds As NativeRect
        Public ClientBounds As NativeRect
        Public Style As UInteger
        Public ExStyle As UInteger
        Public WindowStatus As UInteger
        Public WindowBorderWidth As UInteger
        Public WindowBorderHeight As UInteger
        Public WindowType As UShort
        Public CreatorVersion As UShort
    End Structure

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowInfo(WindowHandle As IntPtr, ByRef Info As NativeWindowInfo) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function RedrawWindow(WindowHandle As IntPtr, UpdateRect As IntPtr, UpdateRegion As IntPtr, Flags As UInteger) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowDC(WindowHandle As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ReleaseDC(WindowHandle As IntPtr, DC As IntPtr) As Integer
    End Function

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
            ActionButton.BackColor = If(ButtonMouseDownBackColor.IsEmpty, SystemColors.ControlDark, ButtonMouseDownBackColor)
        ElseIf _IsButtonHot Then
            ActionButton.BackColor = If(ButtonMouseOverBackColor.IsEmpty, SystemColors.ControlLight, ButtonMouseOverBackColor)
        Else
            ActionButton.BackColor = BackColor
        End If
        ActionButton.Invalidate()
        PaintButtonExtension()
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
        SetTextMarginReservation(LeftMargin, RightMargin)
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
End Class
