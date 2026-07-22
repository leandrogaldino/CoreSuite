Imports System.ComponentModel
''' <summary>
''' Provides a reusable drop-down container mechanism that can host any control,
''' displaying it in a floating borderless window anchored to a host control when triggered.
''' </summary>
Public Class ControlContainer
    Inherits Component
    Private _DropDownContainer As DropDownContainer
    Private _DropDownControl As Control
    Private _ClosedWhileInControl As Boolean
    Private _DropState As ControlContainerDropDownState
    Private _HostControl As Control
    ''' <summary>
    ''' Occurs when the drop-down state changes.
    ''' </summary>
    <Category("Comportamento")>
    Public Event DropStateChanged(sender As Object)
    ''' <summary>
    ''' Occurs when the drop-down starts opening.
    ''' </summary>
    <Category("Comportamento")>
    Public Event Dropping(sender As Object)
    ''' <summary>
    ''' Occurs when the drop-down has fully opened.
    ''' </summary>
    <Category("Comportamento")>
    Public Event Dropped(sender As Object)
    ''' <summary>
    ''' Occurs when the drop-down starts closing.
    ''' </summary>
    <Category("Comportamento")>
    Public Event Closing(sender As Object)
    ''' <summary>
    ''' Occurs when the drop-down has fully closed.
    ''' </summary>
    <Category("Comportamento")>
    Public Event Closed(sender As Object)
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ControlContainer"/> class.
    ''' </summary>
    Public Sub New()
        InitializeDropDown(_DropDownControl)
    End Sub
    ''' <summary>
    ''' Gets or sets the control that hosts the drop-down and triggers it when clicked.
    ''' </summary>
    Public Property HostControl As Control
        Get
            Return _HostControl
        End Get
        Set(value As Control)
            _HostControl = value
            If _HostControl IsNot Nothing Then
                RemoveHandler _HostControl.Click, AddressOf HostControl_Click
                AddHandler _HostControl.Click, AddressOf HostControl_Click
            End If
        End Set
    End Property
    ''' <summary>
    ''' Gets a value indicating whether the drop-down can currently be opened.
    ''' </summary>
    Public Overridable ReadOnly Property CanDrop As Boolean
        Get
            If _DropDownContainer IsNot Nothing Then Return False
            If _DropDownContainer Is Nothing AndAlso _ClosedWhileInControl Then
                _ClosedWhileInControl = False
                Return False
            End If
            Return Not _ClosedWhileInControl
        End Get
    End Property
    ''' <summary>
    ''' Gets the current state of the drop-down container.
    ''' </summary>
    Public ReadOnly Property DropState As ControlContainerDropDownState
        Get
            Return _DropState
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the control that will be displayed inside the drop-down container.
    ''' </summary>
    Public Property DropDownControl As Control
        Get
            Return _DropDownControl
        End Get
        Set(value As Control)
            _DropDownControl = value
            InitializeDropDown(_DropDownControl)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the drop-down is enabled and can be opened by clicking the host control.
    ''' </summary>
    Public Property DropDownEnabled As Boolean = True
    ''' <summary>
    ''' Gets or sets the border color of the drop-down container.
    ''' </summary>
    Public Property DropDownBorderColor As Color = SystemColors.HotTrack
    ''' <summary>
    ''' Handles the host control's click event, opening the drop-down if enabled.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">An <see cref="EventArgs"/> that contains no event data.</param>
    Private Sub HostControl_Click(sender As Object, e As EventArgs)
        If DropDownEnabled Then ShowDropDown()
    End Sub
    ''' <summary>
    ''' Prepares the specified control to be used as the drop-down content, removing it
    ''' from the host control's collection if necessary and resetting the drop state.
    ''' </summary>
    ''' <param name="DropDownControl">The control to initialize as the drop-down content.</param>
    Private Sub InitializeDropDown(ByVal DropDownControl As Control)
        _DropState = ControlContainerDropDownState.Closed
        If HostControl IsNot Nothing AndAlso HostControl.Controls.Contains(DropDownControl) Then
            HostControl.Controls.Remove(DropDownControl)
        End If
        _DropDownControl = DropDownControl
    End Sub
    ''' <summary>
    ''' Closes the drop-down when the host control's parent is moved.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">An <see cref="EventArgs"/> that contains no event data.</param>
    Private Sub Parent_Move(ByVal sender As Object, ByVal e As EventArgs)
        CloseDropDown()
    End Sub
    ''' <summary>
    ''' Handles the drop-down container's closed event, cleaning up event handlers,
    ''' disposing the container, and updating the drop state.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">A <see cref="FormClosedEventArgs"/> that contains the event data.</param>
    Private Sub DropContainer_Closed(ByVal sender As Object, ByVal e As FormClosedEventArgs)
        RaiseEvent Closing(Me)
        If Not _DropDownContainer.IsDisposed Then
            RemoveHandler _DropDownContainer.DropStateChange, AddressOf DropContainer_DropStateChange
            RemoveHandler _DropDownContainer.FormClosed, AddressOf DropContainer_Closed
            RemoveHandler HostControl.Parent.Move, AddressOf Parent_Move
            _DropDownContainer.Dispose()
        End If
        _DropDownContainer = Nothing
        _ClosedWhileInControl = (HostControl.RectangleToScreen(HostControl.ClientRectangle).Contains(Cursor.Position))
        _DropState = ControlContainerDropDownState.Closed
        HostControl.Invalidate()
        RaiseEvent Closed(Me)
    End Sub
    ''' <summary>
    ''' Updates the current drop-down state when notified by the drop-down container.
    ''' </summary>
    ''' <param name="state">The new drop-down state.</param>
    Private Sub DropContainer_DropStateChange(ByVal state As ControlContainerDropDownState)
        _DropState = state
    End Sub
    ''' <summary>
    ''' Calculates the screen bounds where the drop-down container should be displayed,
    ''' positioning it relative to the host control and adjusting it to fit within the working area.
    ''' </summary>
    ''' <returns>A <see cref="Rectangle"/> representing the screen bounds for the drop-down container.</returns>
    Private Function GetDropDownBounds() As Rectangle
        Dim InflatedDropSize As New Size(_DropDownControl.Width + 2, _DropDownControl.Height + 2)
        Dim ScreenBounds As New Rectangle(HostControl.Parent.PointToScreen(New Point(HostControl.Bounds.X, HostControl.Bounds.Bottom)), InflatedDropSize)
        Dim WorkingArea As Rectangle = Screen.GetWorkingArea(ScreenBounds)
        If ScreenBounds.X < WorkingArea.X Then ScreenBounds.X = WorkingArea.X
        If ScreenBounds.Y < WorkingArea.Y Then ScreenBounds.Y = WorkingArea.Y
        If ScreenBounds.Right > WorkingArea.Right AndAlso WorkingArea.Width > ScreenBounds.Width Then ScreenBounds.X = WorkingArea.Right - ScreenBounds.Width
        If ScreenBounds.Bottom > WorkingArea.Bottom AndAlso WorkingArea.Height > ScreenBounds.Height Then ScreenBounds.Y = WorkingArea.Bottom - ScreenBounds.Height
        Return ScreenBounds
    End Function
    ''' <summary>
    ''' Displays the drop-down container anchored to the host control.
    ''' </summary>
    ''' <exception cref="Exception">
    ''' Thrown when <see cref="HostControl"/> or <see cref="DropDownControl"/> has not been set.
    ''' </exception>
    Public Sub ShowDropDown()
        If HostControl Is Nothing Then Throw New Exception("O controle hospedeiro não foi definido.")
        If DropDownControl Is Nothing Then Throw New Exception("O controle a ser hospedado não foi definido.")
        If Not CanDrop Then Return
        RaiseEvent Dropping(Me)
        _DropDownContainer = New DropDownContainer(_DropDownControl, DropDownBorderColor) With {
            .Bounds = GetDropDownBounds()
        }
        AddHandler _DropDownContainer.DropStateChange, New DropDownContainer.DropWindowArgs(AddressOf DropContainer_DropStateChange)
        AddHandler _DropDownContainer.FormClosed, AddressOf DropContainer_Closed
        AddHandler HostControl.Parent.Move, New EventHandler(AddressOf Parent_Move)
        _DropState = ControlContainerDropDownState.Dropping
        If HostControl.GetType = GetType(Button) Then
            If CType(HostControl, Button).FlatStyle = FlatStyle.Standard Or CType(HostControl, Button).FlatStyle = FlatStyle.System Then
                _DropDownContainer.Top -= 2
                _DropDownContainer.Left += 1
            Else
                _DropDownContainer.Top -= 1
                _DropDownContainer.Left += 0
            End If
        End If
        _DropDownContainer.Show(HostControl)
        _DropState = ControlContainerDropDownState.Dropped
        HostControl.Invalidate()
        RaiseEvent Dropped(Me)
    End Sub
    ''' <summary>
    ''' Closes the drop-down container if it is currently open.
    ''' </summary>
    Public Sub CloseDropDown()
        If _DropDownContainer IsNot Nothing Then
            _DropState = ControlContainerDropDownState.Closing
            _DropDownContainer.Close()
        End If
    End Sub
    ''' <summary>
    ''' Represents the floating, borderless window used internally to host and display
    ''' the drop-down content.
    ''' </summary>
    Friend NotInheritable Class DropDownContainer
        Inherits Form
        Implements IMessageFilter
        Private Const WM_LBUTTONDOWN As Integer = &H201
        Private Const WM_RBUTTONDOWN As Integer = &H204
        Private Const WM_MBUTTONDOWN As Integer = &H207
        Private Const WM_NCLBUTTONDOWN As Integer = &HA1
        Private Const WM_NCRBUTTONDOWN As Integer = &HA4
        Private Const WM_NCMBUTTONDOWN As Integer = &HA7
        Private ReadOnly _DropDownBorderColor As Color
        Private ReadOnly _DropDownItem As Control
        ''' <summary>
        ''' Occurs when the drop-down container's visibility state changes.
        ''' </summary>
        Public Event DropStateChange As DropWindowArgs
        ''' <summary>
        ''' Represents the method that handles a change in the drop-down container's state.
        ''' </summary>
        ''' <param name="State">The new drop-down state.</param>
        Public Delegate Sub DropWindowArgs(ByVal State As ControlContainerDropDownState)
        ''' <summary>
        ''' Initializes a new instance of the <see cref="DropDownContainer"/> class with the
        ''' specified content control and border color.
        ''' </summary>
        ''' <param name="DropDownItem">The control to be hosted inside the drop-down window.</param>
        ''' <param name="DropDownBorderColor">The border color of the drop-down window.</param>
        Public Sub New(DropDownItem As Control, DropDownBorderColor As Color)
            Dim DropDownMinSize As New Size(136, 39)
            _DropDownBorderColor = DropDownBorderColor
            _DropDownItem = DropDownItem
            FormBorderStyle = FormBorderStyle.None
            DropDownItem.Location = New Point(1, 1)
            Controls.Add(DropDownItem)
            StartPosition = FormStartPosition.Manual
            ShowInTaskbar = False
            KeyPreview = True
            Size = New Size(200, 200)
            If DropDownItem.Width < DropDownMinSize.Width Or DropDownItem.Height < DropDownMinSize.Height Then
                DropDownItem.Left = (DropDownMinSize.Width - DropDownItem.Width) / 2
                DropDownItem.Top = (DropDownMinSize.Height - DropDownItem.Height) / 2
            End If
            Application.AddMessageFilter(Me)
        End Sub
        ''' <summary>
        ''' Paints the border of the drop-down window.
        ''' </summary>
        ''' <param name="e">A <see cref="PaintEventArgs"/> that contains the event data.</param>
        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            MyBase.OnPaint(e)
            ControlPaint.DrawBorder(Me.CreateGraphics, ClientRectangle, _DropDownBorderColor, ButtonBorderStyle.Solid)
        End Sub
        ''' <summary>
        ''' Closes the drop-down window when the Escape key is pressed.
        ''' </summary>
        ''' <param name="e">A <see cref="KeyEventArgs"/> that contains the event data.</param>
        Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
            MyBase.OnKeyDown(e)
            If e.KeyCode = Keys.Escape Then
                Close()
            End If
        End Sub
        ''' <summary>
        ''' Removes the message filter and detaches the hosted control before the window closes.
        ''' </summary>
        ''' <param name="e">A <see cref="CancelEventArgs"/> that contains the event data.</param>
        Protected Overrides Sub OnClosing(ByVal e As CancelEventArgs)
            Application.RemoveMessageFilter(Me)
            Controls.RemoveAt(0)
            MyBase.OnClosing(e)
        End Sub
        ''' <summary>
        ''' Filters application messages to detect clicks outside the drop-down window
        ''' and close it accordingly.
        ''' </summary>
        ''' <param name="m">The message to be filtered.</param>
        ''' <returns><c>False</c> to allow the message to continue to the next filter or control.</returns>
        Private Function IMessageFilter_PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage
            Dim CursorPos As Point = Cursor.Position
            If Visible AndAlso (ActiveForm Is Nothing OrElse Not ActiveForm.Equals(Me)) Then
                Close()
            End If
            Select Case m.Msg
                Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN, WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN, WM_NCMBUTTONDOWN
                    If _DropDownItem.Parent IsNot Nothing AndAlso Not Bounds.Contains(CursorPos) Then
                        If Not _DropDownItem.Bounds.Contains(_DropDownItem.Parent.PointToClient(CursorPos)) Then
                            Application.RemoveMessageFilter(Me)
                            Close()
                        End If
                    End If
            End Select
            Return False
        End Function
    End Class
End Class

