Imports System.ComponentModel
Imports System.Drawing.Drawing2D
''' <summary>
''' Represents a button that combines a standard button action with an optional drop-down menu.
''' </summary>
<DefaultEvent("Click")>
<DefaultProperty(NameOf(Text))>
Public Class SplitButton
    Inherits NoFocusCueButton
    Private Const MinimumDropDownAreaWidth As Integer = 16
    Private _DropDownMenu As ContextMenuStrip
    Private _Mode As SplitButtonMode = SplitButtonMode.Split
    Private _DropDownAreaWidth As Integer = 24
    Private _ArrowColor As Color = Color.Empty
    Private _SeparatorColor As Color = SystemColors.ControlDark
    Private _ShowSeparator As Boolean = True
    Private _ContentPadding As Padding = System.Windows.Forms.Padding.Empty
    Private _IsMenuOpen As Boolean
    Private _IsMenuOpening As Boolean
    Private _IsDisposing As Boolean
    Private _DropDownPressed As Boolean
    Private _PendingMouseClick As Boolean
    Private _MouseDownInDropDownArea As Boolean
    Private _PendingMouseDropDown As Boolean
    Private _SuppressPendingClick As Boolean
    ''' <summary>
    ''' Occurs before the drop-down menu is displayed.
    ''' </summary>
    ''' <remarks>
    ''' Set the <see cref="CancelEventArgs.Cancel"/> property to <see langword="True"/> to prevent the menu from opening.
    ''' </remarks>
    Public Event DropDownOpening As CancelEventHandler
    ''' <summary>
    ''' Occurs after the drop-down menu has been displayed.
    ''' </summary>
    Public Event DropDownOpened As EventHandler
    ''' <summary>
    ''' Occurs after the drop-down menu has been closed.
    ''' </summary>
    Public Event DropDownClosed As EventHandler
    ''' <summary>
    ''' Initializes a new instance of the <see cref="SplitButton"/> class.
    ''' </summary>
    Public Sub New()
        SetStyle(ControlStyles.ResizeRedraw, True)
        UpdateEffectivePadding()
    End Sub
    ''' <summary>
    ''' Gets or sets the context menu displayed by the drop-down portion of the button.
    ''' </summary>
    <Category("SplitButton")>
    <DefaultValue(CType(Nothing, ContextMenuStrip))>
    <Description("Gets or sets the context menu displayed by the drop-down portion of the button.")>
    Public Property DropDownMenu As ContextMenuStrip
        Get
            Return _DropDownMenu
        End Get
        Set(value As ContextMenuStrip)
            If ReferenceEquals(_DropDownMenu, value) Then Return
            If _DropDownMenu IsNot Nothing Then
                If _IsMenuOpen AndAlso Not _DropDownMenu.IsDisposed Then _DropDownMenu.Close(ToolStripDropDownCloseReason.CloseCalled)
                RemoveHandler _DropDownMenu.Opened, AddressOf DropDownMenu_Opened
                RemoveHandler _DropDownMenu.Closed, AddressOf DropDownMenu_Closed
            End If
            _DropDownMenu = value
            _IsMenuOpen = False
            _IsMenuOpening = False
            If _DropDownMenu IsNot Nothing Then
                AddHandler _DropDownMenu.Opened, AddressOf DropDownMenu_Opened
                AddHandler _DropDownMenu.Closed, AddressOf DropDownMenu_Closed
            End If
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets how the button responds when activated.
    ''' </summary>
    <Category("SplitButton")>
    <DefaultValue(SplitButtonMode.Split)>
    <Description("Gets or sets how the button responds when activated.")>
    Public Property Mode As SplitButtonMode
        Get
            Return _Mode
        End Get
        Set(value As SplitButtonMode)
            If Not System.Enum.IsDefined(GetType(SplitButtonMode), value) Then Throw New InvalidEnumArgumentException(NameOf(value), CInt(value), GetType(SplitButtonMode))
            If _Mode = value Then Return
            _Mode = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the logical width of the drop-down area.
    ''' </summary>
    <Category("SplitButton")>
    <DefaultValue(24)>
    <Description("Gets or sets the logical width of the drop-down area.")>
    Public Property DropDownAreaWidth As Integer
        Get
            Return _DropDownAreaWidth
        End Get
        Set(value As Integer)
            If value <MinimumDropDownAreaWidth Then Throw New ArgumentOutOfRangeException(NameOf(value), value, $"The drop-down area width must be at least {MinimumDropDownAreaWidth}.")
            If _DropDownAreaWidth = value Then Return
            _DropDownAreaWidth = value
            UpdateEffectivePadding()
            PerformLayout()
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the color used to draw the drop-down arrow.
    ''' </summary>
    ''' <remarks>
    ''' When set to <see cref="Color.Empty"/>, the button's foreground color is used.
    ''' </remarks>
    <Category("SplitButton")>
    <DefaultValue(GetType(Color), "Empty")>
    <Description("Gets or sets the color used to draw the drop-down arrow. An empty color uses the button foreground color.")>
    Public Property ArrowColor As Color
        Get
            Return _ArrowColor
        End Get
        Set(value As Color)
            If _ArrowColor = value Then Return
            _ArrowColor = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the color used to draw the separator between the main and drop-down areas.
    ''' </summary>
    <Category("SplitButton")>
    <DefaultValue(GetType(Color), "ControlDark")>
    <Description("Gets or sets the color used to draw the separator between the main and drop-down areas.")>
    Public Property SeparatorColor As Color
        Get
            Return _SeparatorColor
        End Get
        Set(value As Color)
            If _SeparatorColor = value Then Return
            _SeparatorColor = value
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether a separator is drawn between the main and drop-down areas.
    ''' </summary>
    <Category("SplitButton")>
    <DefaultValue(True)>
    <Description("Gets or sets whether a separator is drawn between the main and drop-down areas.")>
    Public Property ShowSeparator As Boolean
        Get
            Return _ShowSeparator
        End Get
        Set(value As Boolean)
            If _ShowSeparator = value Then Return
            _ShowSeparator = value
            Invalidate()
        End Set
    End Property
    Public Shadows Property Padding As Padding
        Get
            Return _ContentPadding
        End Get
        Set(value As Padding)
            If _ContentPadding.Equals(value) Then Return
            _ContentPadding = value
            UpdateEffectivePadding()
            PerformLayout()
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets the bounds occupied by the drop-down area.
    ''' </summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Advanced)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property DropDownRectangle As Rectangle
        Get
            Dim AreaWidth As Integer = Math.Min(GetScaledDropDownAreaWidth(), ClientSize.Width)
            If RightToLeft = RightToLeft.Yes Then Return New Rectangle(0, 0, AreaWidth, ClientSize.Height)
            Return New Rectangle(Math.Max(0, ClientSize.Width - AreaWidth), 0, AreaWidth, ClientSize.Height)
        End Get
    End Property
    ''' <summary>
    ''' Displays the associated drop-down menu.
    ''' </summary>
    ''' <returns><see langword="True"/> if the menu was displayed; otherwise, <see langword="False"/>.</returns>
    Public Function ShowDropDown() As Boolean
        If Not CanShowDropDown() Then Return False
        Dim Args As New CancelEventArgs()
        OnDropDownOpening(Args)
        If Args.Cancel Then Return False
        If CanFocus Then Focus()
        Dim Location As Point
        Dim Direction As ToolStripDropDownDirection
        If RightToLeft = RightToLeft.Yes Then
            Location = New Point(ClientSize.Width, ClientSize.Height)
            Direction = ToolStripDropDownDirection.BelowLeft
        Else
            Location = New Point(0, ClientSize.Height)
            Direction = ToolStripDropDownDirection.BelowRight
        End If
        _IsMenuOpening = True
        Invalidate()
        Try
            _DropDownMenu.Show(Me, Location, Direction)
        Catch
            _IsMenuOpening = False
            _IsMenuOpen = False
            Invalidate()
            Throw
        End Try
        If Not _DropDownMenu.Visible Then
            _IsMenuOpening = False
            _IsMenuOpen = False
            Invalidate()
            Return False
        End If
        Return True
    End Function
    ''' <summary>
    ''' Raises the <see cref="DropDownOpening"/> event.
    ''' </summary>
    ''' <param name="E">A <see cref="CancelEventArgs"/> containing the event data.</param>
    Protected Overridable Sub OnDropDownOpening(E As CancelEventArgs)
        RaiseEvent DropDownOpening(Me, E)
    End Sub
    ''' <summary>
    ''' Raises the <see cref="DropDownOpened"/> event.
    ''' </summary>
    ''' <param name="E">An <see cref="EventArgs"/> containing the event data.</param>
    Protected Overridable Sub OnDropDownOpened(E As EventArgs)
        RaiseEvent DropDownOpened(Me, E)
    End Sub
    ''' <summary>
    ''' Raises the <see cref="DropDownClosed"/> event.
    ''' </summary>
    ''' <param name="E">An <see cref="EventArgs"/> containing the event data.</param>
    Protected Overridable Sub OnDropDownClosed(E As EventArgs)
        RaiseEvent DropDownClosed(Me, E)
    End Sub
    ''' <summary>
    ''' Gets the default size of the control.
    ''' </summary>
    ''' <value>
    ''' A <see cref="Size"/> representing the default width and height of the control.
    ''' </value>
    Protected Overrides ReadOnly Property DefaultSize As Size
        Get
            Return New Size(125, 30)
        End Get
    End Property
    ''' <summary>
    ''' Raises the <see cref="Control.Click"/> event or displays the associated drop-down menu according to the current interaction state and button mode.
    ''' </summary>
    ''' <param name="E">An <see cref="EventArgs"/> containing the event data.</param>
    Protected Overrides Sub OnClick(E As EventArgs)
        If _PendingMouseClick Then
            Dim OpenDropDown As Boolean = _PendingMouseDropDown
            Dim SuppressClick As Boolean = _SuppressPendingClick
            ClearPendingMouseState()
            If SuppressClick Then Return
            If OpenDropDown Then
                If CanShowDropDown() Then
                    ShowDropDown()
                ElseIf _Mode = SplitButtonMode.DropDown Then
                    MyBase.OnClick(E)
                End If
                Return
            End If
        End If
        If _Mode = SplitButtonMode.DropDown Then
            If CanShowDropDown() Then
                ShowDropDown()
            Else
                MyBase.OnClick(E)
            End If
            Return
        End If
        MyBase.OnClick(E)
    End Sub
    ''' <summary>
    ''' Processes a mouse button press and records whether the interaction started within the drop-down area.
    ''' </summary>
    ''' <param name="E">A <see cref="MouseEventArgs"/> containing the mouse event data.</param>
    Protected Overrides Sub OnMouseDown(E As MouseEventArgs)
        If E.Button = MouseButtons.Left Then
            _PendingMouseClick = True
            _MouseDownInDropDownArea = IsDropDownArea(E.Location)
            _PendingMouseDropDown = _MouseDownInDropDownArea
            _SuppressPendingClick = False
            _DropDownPressed = _MouseDownInDropDownArea
            Invalidate(DropDownRectangle)
        End If
        MyBase.OnMouseDown(E)
    End Sub
    ''' <summary>
    ''' Updates the pressed state of the drop-down area while the pointer moves during a mouse interaction.
    ''' </summary>
    ''' <param name="E">A <see cref="MouseEventArgs"/> containing the mouse event data.</param>
    Protected Overrides Sub OnMouseMove(E As MouseEventArgs)
        MyBase.OnMouseMove(E)
        If Not _PendingMouseClick OrElse E.Button <> MouseButtons.Left Then Return
        Dim CurrentAreaIsDropDown As Boolean = IsDropDownArea(E.Location)
        Dim Pressed As Boolean = _MouseDownInDropDownArea AndAlso CurrentAreaIsDropDown
        If _DropDownPressed = Pressed Then Return
        _DropDownPressed = Pressed
        Invalidate(DropDownRectangle)
    End Sub
    ''' <summary>
    ''' Processes a mouse button release and determines whether the pending interaction should open the drop-down menu.
    ''' </summary>
    ''' <param name="E">A <see cref="MouseEventArgs"/> containing the mouse event data.</param>
    Protected Overrides Sub OnMouseUp(E As MouseEventArgs)
        If E.Button = MouseButtons.Left AndAlso _PendingMouseClick Then
            Dim ReleasedInDropDownArea As Boolean = IsDropDownArea(E.Location)
            _PendingMouseDropDown = _MouseDownInDropDownArea AndAlso ReleasedInDropDownArea
            _SuppressPendingClick = _MouseDownInDropDownArea <> ReleasedInDropDownArea
        End If
        _DropDownPressed = False
        MyBase.OnMouseUp(E)
        Invalidate(DropDownRectangle)
        QueueMouseStateReset()
    End Sub
    ''' <summary>
    ''' Clears the visual pressed state when the mouse pointer leaves the control.
    ''' </summary>
    ''' <param name="E">An <see cref="EventArgs"/> containing the event data.</param>
    Protected Overrides Sub OnMouseLeave(E As EventArgs)
        MyBase.OnMouseLeave(E)
        If Not _DropDownPressed Then Return
        _DropDownPressed = False
        Invalidate(DropDownRectangle)
    End Sub
    ''' <summary>
    ''' Processes command keys used to display the associated drop-down menu.
    ''' </summary>
    ''' <param name="Msg">A <see cref="Message"/> representing the window message to process.</param>
    ''' <param name="KeyData">A combination of <see cref="Keys"/> values representing the key and modifier keys.</param>
    ''' <returns>
    ''' <see langword="True"/> if the command key was processed by the control; otherwise, <see langword="False"/>.
    ''' </returns>
    Protected Overrides Function ProcessCmdKey(ByRef Msg As Message, KeyData As Keys) As Boolean
        If KeyData = (Keys.Alt Or Keys.Down) OrElse KeyData = Keys.F4 Then
            If CanShowDropDown() Then Return ShowDropDown()
        End If
        Return MyBase.ProcessCmdKey(Msg, KeyData)
    End Function
    ''' <summary>
    ''' Paints the button and its drop-down area.
    ''' </summary>
    ''' <param name="E">A <see cref="PaintEventArgs"/> containing the painting data.</param>
    Protected Overrides Sub OnPaint(E As PaintEventArgs)
        MyBase.OnPaint(E)
        DrawDropDownArea(E.Graphics)
    End Sub
    ''' <summary>
    ''' Updates the visual state of the control when its enabled state changes.
    ''' </summary>
    ''' <param name="E">An <see cref="EventArgs"/> containing the event data.</param>
    Protected Overrides Sub OnEnabledChanged(E As EventArgs)
        MyBase.OnEnabledChanged(E)
        Invalidate()
    End Sub
    ''' <summary>
    ''' Updates the visual state of the control when its foreground color changes.
    ''' </summary>
    ''' <param name="E">An <see cref="EventArgs"/> containing the event data.</param>
    Protected Overrides Sub OnForeColorChanged(E As EventArgs)
        MyBase.OnForeColorChanged(E)
        Invalidate()
    End Sub
    ''' <summary>
    ''' Updates the effective padding and layout when the right-to-left setting changes.
    ''' </summary>
    ''' <param name="E">An <see cref="EventArgs"/> containing the event data.</param>
    Protected Overrides Sub OnRightToLeftChanged(E As EventArgs)
        MyBase.OnRightToLeftChanged(E)
        UpdateEffectivePadding()
        PerformLayout()
        Invalidate()
    End Sub
    ''' <summary>
    ''' Updates DPI-dependent layout values after the DPI of the parent control changes.
    ''' </summary>
    ''' <param name="E">An <see cref="EventArgs"/> containing the event data.</param>
    Protected Overrides Sub OnDpiChangedAfterParent(E As EventArgs)
        MyBase.OnDpiChangedAfterParent(E)
        UpdateEffectivePadding()
        PerformLayout()
        Invalidate()
    End Sub
    ''' <summary>
    ''' Scales the control and its content padding by the specified scaling factors.
    ''' </summary>
    ''' <param name="Factor">A <see cref="SizeF"/> containing the horizontal and vertical scaling factors.</param>
    ''' <param name="Specified">A combination of <see cref="BoundsSpecified"/> values indicating which bounds should be scaled.</param>
    Protected Overrides Sub ScaleControl(Factor As SizeF, Specified As BoundsSpecified)
        _ContentPadding = New Padding(ScalePaddingValue(_ContentPadding.Left, Factor.Width), ScalePaddingValue(_ContentPadding.Top, Factor.Height), ScalePaddingValue(_ContentPadding.Right, Factor.Width), ScalePaddingValue(_ContentPadding.Bottom, Factor.Height))
        MyBase.ScaleControl(Factor, Specified)
        UpdateEffectivePadding()
    End Sub
    ''' <summary>
    ''' Releases the managed resources used by the control.
    ''' </summary>
    ''' <param name="Disposing">
    ''' <see langword="True"/> to release managed and unmanaged resources; <see langword="False"/> to release only unmanaged resources.
    ''' </param>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            _IsDisposing = True
            If _DropDownMenu IsNot Nothing Then
                If _IsMenuOpen AndAlso Not _DropDownMenu.IsDisposed Then _DropDownMenu.Close(ToolStripDropDownCloseReason.CloseCalled)
                RemoveHandler _DropDownMenu.Opened, AddressOf DropDownMenu_Opened
                RemoveHandler _DropDownMenu.Closed, AddressOf DropDownMenu_Closed
                _DropDownMenu = Nothing
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    ''' <summary>
    ''' Determines whether the associated drop-down menu can currently be displayed.
    ''' </summary>
    ''' <returns>
    ''' <see langword="True"/> if the menu is available and the control is in a valid state to display it; otherwise, <see langword="False"/>.
    ''' </returns>
    Private Function CanShowDropDown() As Boolean
        Return Enabled AndAlso IsHandleCreated AndAlso Visible AndAlso _DropDownMenu IsNot Nothing AndAlso Not _DropDownMenu.IsDisposed AndAlso Not _IsMenuOpen AndAlso Not _IsMenuOpening
    End Function
    ''' <summary>
    ''' Determines whether the specified client location belongs to the drop-down area.
    ''' </summary>
    ''' <param name="Location">The client coordinates to evaluate.</param>
    ''' <returns>
    ''' <see langword="True"/> if the location belongs to the drop-down area; otherwise, <see langword="False"/>.
    ''' </returns>
    Private Function IsDropDownArea(Location As Point) As Boolean
        If Not ClientRectangle.Contains(Location) Then Return False
        If _Mode = SplitButtonMode.DropDown Then Return True
        Return DropDownRectangle.Contains(Location)
    End Function
    ''' <summary>
    ''' Draws the separator and arrow contained in the drop-down area.
    ''' </summary>
    ''' <param name="Graphics">The graphics surface used to draw the drop-down area.</param>
    Private Sub DrawDropDownArea(Graphics As Graphics)
        Dim Bounds As Rectangle = DropDownRectangle
        If Bounds.Width <= 0 OrElse Bounds.Height <= 0 Then Return
        If _Mode = SplitButtonMode.Split AndAlso _ShowSeparator Then DrawSeparator(Graphics, Bounds)
        DrawArrow(Graphics, Bounds)
    End Sub
    ''' <summary>
    ''' Draws the separator between the main button area and the drop-down area.
    ''' </summary>
    ''' <param name="Graphics">The graphics surface used to draw the separator.</param>
    ''' <param name="Bounds">The bounds of the drop-down area.</param>
    Private Sub DrawSeparator(Graphics As Graphics, Bounds As Rectangle)
        Dim LineX As Integer
        If RightToLeft = RightToLeft.Yes Then
            LineX = Bounds.Right - 1
        Else
            LineX = Bounds.Left
        End If
        Dim VerticalMargin As Integer = GetScaledValue(5)
        Dim Top As Integer = VerticalMargin
        Dim Bottom As Integer = ClientSize.Height - VerticalMargin - 1
        If Bottom <= Top Then Return
        Dim Color As Color = If(Enabled, _SeparatorColor, SystemColors.ControlDark)
        Using SeparatorPen As New Pen(Color)
            Graphics.DrawLine(SeparatorPen, LineX, Top, LineX, Bottom)
        End Using
    End Sub
    ''' <summary>
    ''' Draws the drop-down arrow within the specified bounds.
    ''' </summary>
    ''' <param name="Graphics">The graphics surface used to draw the arrow.</param>
    ''' <param name="Bounds">The bounds in which the arrow is drawn.</param>
    Private Sub DrawArrow(Graphics As Graphics, Bounds As Rectangle)
        Dim ArrowWidth As Integer = GetScaledValue(8)
        Dim ArrowHeight As Integer = GetScaledValue(5)
        Dim PressedOffset As Integer = If(_DropDownPressed OrElse _IsMenuOpen OrElse _IsMenuOpening, GetScaledValue(1), 0)
        Dim CenterX As Single = Bounds.Left + Bounds.Width / 2.0F
        Dim CenterY As Single = Bounds.Top + Bounds.Height / 2.0F + PressedOffset
        Dim HalfWidth As Single = ArrowWidth / 2.0F
        Dim HalfHeight As Single = ArrowHeight / 2.0F
        Dim Points As PointF() = {New PointF(CenterX - HalfWidth, CenterY - HalfHeight), New PointF(CenterX + HalfWidth, CenterY - HalfHeight), New PointF(CenterX, CenterY + HalfHeight)}
        Dim Color As Color
        If Not Enabled Then
            Color = SystemColors.GrayText
        ElseIf _ArrowColor.IsEmpty Then
            Color = ForeColor
        Else
            Color = _ArrowColor
        End If
        Dim State As GraphicsState = Graphics.Save()
        Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Using ArrowBrush As New SolidBrush(Color)
            Graphics.FillPolygon(ArrowBrush, Points)
        End Using
        Graphics.Restore(State)
    End Sub
    ''' <summary>
    ''' Updates the base padding to reserve space for the drop-down area while preserving the user-defined content padding.
    ''' </summary>
    Private Sub UpdateEffectivePadding()
        Dim AreaWidth As Integer = GetScaledDropDownAreaWidth()
        Dim EffectivePadding As Padding
        If RightToLeft = RightToLeft.Yes Then
            EffectivePadding = New Padding(_ContentPadding.Left + AreaWidth, _ContentPadding.Top, _ContentPadding.Right, _ContentPadding.Bottom)
        Else
            EffectivePadding = New Padding(_ContentPadding.Left, _ContentPadding.Top, _ContentPadding.Right + AreaWidth, _ContentPadding.Bottom)
        End If
        If MyBase.Padding.Equals(EffectivePadding) Then Return
        MyBase.Padding = EffectivePadding
    End Sub
    ''' <summary>
    ''' Gets the DPI-scaled width of the drop-down area.
    ''' </summary>
    ''' <returns>The width of the drop-down area adjusted for the current DPI.</returns>
    Private Function GetScaledDropDownAreaWidth() As Integer
        Return GetScaledValue(_DropDownAreaWidth)
    End Function
    ''' <summary>
    ''' Scales a logical pixel value according to the current control DPI.
    ''' </summary>
    ''' <param name="Value">The logical pixel value to scale.</param>
    ''' <returns>The DPI-adjusted pixel value.</returns>
    Private Function GetScaledValue(Value As Integer) As Integer
        Return Math.Max(1, CInt(Math.Round(Value * DeviceDpi / 96.0R)))
    End Function
    ''' <summary>
    ''' Scales a padding value by the specified layout scaling factor.
    ''' </summary>
    ''' <param name="Value">The padding value to scale.</param>
    ''' <param name="Factor">The scaling factor to apply.</param>
    ''' <returns>The scaled non-negative padding value.</returns>
    Private Shared Function ScalePaddingValue(Value As Integer, Factor As Single) As Integer
        Return Math.Max(0, CInt(Math.Round(Value * Factor)))
    End Function
    ''' <summary>
    ''' Handles the opening of the associated drop-down menu and updates the internal menu state.
    ''' </summary>
    ''' <param name="Sender">The object that raised the event.</param>
    ''' <param name="E">An <see cref="EventArgs"/> containing the event data.</param>
    Private Sub DropDownMenu_Opened(Sender As Object, E As EventArgs)
        If Not _IsMenuOpening Then Return
        _IsMenuOpening = False
        _IsMenuOpen = True
        Invalidate()
        If Not _IsDisposing Then OnDropDownOpened(EventArgs.Empty)
    End Sub
    ''' <summary>
    ''' Handles the closing of the associated drop-down menu and updates the internal menu state.
    ''' </summary>
    ''' <param name="Sender">The object that raised the event.</param>
    ''' <param name="E">A <see cref="ToolStripDropDownClosedEventArgs"/> containing the event data.</param>
    Private Sub DropDownMenu_Closed(Sender As Object, E As ToolStripDropDownClosedEventArgs)
        If Not _IsMenuOpen AndAlso Not _IsMenuOpening Then Return
        Dim WasOpen As Boolean = _IsMenuOpen
        _IsMenuOpening = False
        _IsMenuOpen = False
        Invalidate()
        If WasOpen AndAlso Not _IsDisposing Then OnDropDownClosed(EventArgs.Empty)
    End Sub
    ''' <summary>
    ''' Schedules the pending mouse interaction state to be cleared after the current event processing completes.
    ''' </summary>
    Private Sub QueueMouseStateReset()
        If Not IsHandleCreated OrElse IsDisposed OrElse Disposing Then Return
        Try
            BeginInvoke(New MethodInvoker(AddressOf ClearPendingMouseState))
        Catch Ex As InvalidOperationException
            ClearPendingMouseState()
        End Try
    End Sub
    ''' <summary>
    ''' Clears the state associated with the current pending mouse interaction.
    ''' </summary>
    Private Sub ClearPendingMouseState()
        _PendingMouseClick = False
        _MouseDownInDropDownArea = False
        _PendingMouseDropDown = False
        _SuppressPendingClick = False
    End Sub
    ''' <summary>
    ''' Determines whether the content padding should be serialized by the designer.
    ''' </summary>
    ''' <returns>
    ''' <see langword="True"/> if the content padding differs from its default value; otherwise, <see langword="False"/>.
    ''' </returns>
    Private Function ShouldSerializePadding() As Boolean
        Return Not _ContentPadding.Equals(Padding.Empty)
    End Function
    ''' <summary>
    ''' Restores the content padding to its default value.
    ''' </summary>
    Private Sub ResetPadding()
        Padding = Padding.Empty
    End Sub
End Class