Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Threading

''' <summary>
''' Provides a reusable drop-down container component capable of hosting any
''' Windows Forms control inside a floating drop-down.
''' </summary>
''' <remarks>
''' The container manages positioning, outside clicks, keyboard dismissal,
''' modal dialog ownership and the complete opening and closing lifecycle.
''' </remarks>
Public Class ControlContainer
    Inherits Component
    Private _DropDown As ToolStripDropDown
    Private _ControlHost As ToolStripControlHost
    Private _HostPanel As Panel
    Private _HostedControl As Control
    Private _HostControl As Control
    Private _HostParent As Control
    Private _OwnerForm As Form
    Private _ClosedWhileInControl As Boolean
    Private _DropState As ControlContainerDropDownState
    Private _AllowDropDownClose As Boolean
    Private _AutomaticCloseScheduled As Boolean
    ''' <summary>
    ''' Occurs when the drop-down state changes.
    ''' </summary>
    <Category("ControlContainer")>
    Public Event DropStateChanged(sender As Object)
    ''' <summary>
    ''' Occurs when the drop-down starts opening.
    ''' </summary>
    <Category("ControlContainer")>
    Public Event Dropping(sender As Object)
    ''' <summary>
    ''' Occurs when the drop-down has fully opened.
    ''' </summary>
    <Category("ControlContainer")>
    Public Event Dropped(sender As Object)
    ''' <summary>
    ''' Occurs when the drop-down starts closing.
    ''' </summary>
    <Category("ControlContainer")>
    Public Event Closing(sender As Object)
    ''' <summary>
    ''' Occurs when the drop-down has fully closed.
    ''' </summary>
    <Category("ControlContainer")>
    Public Event Closed(sender As Object)
    ''' <summary>
    ''' Gets or sets the control used as the anchor for the drop-down.
    ''' </summary>
    <Description("Defines the control used as the anchor for displaying the drop-down container.")>
    <Category("ControlContainer")>
    Public Property HostControl As Control
        Get
            Return _HostControl
        End Get
        Set(value As Control)
            If ReferenceEquals(_HostControl, value) Then Return
            If _HostControl IsNot Nothing Then
                RemoveHandler _HostControl.Click, AddressOf HostControl_Click
            End If
            _HostControl = value
            If _HostControl IsNot Nothing Then
                AddHandler _HostControl.Click, AddressOf HostControl_Click
            End If
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the control displayed inside the drop-down.
    ''' </summary>
    <Description("Defines the control displayed inside the drop-down container.")>
    <Category("ControlContainer")>
    Public Property HostedControl As Control
        Get
            Return _HostedControl
        End Get
        Set(value As Control)
            If ReferenceEquals(_HostedControl, value) Then Return
            CloseDropDown()
            _HostedControl = value
            SetDropState(ControlContainerDropDownState.Closed)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether clicking the host control opens the drop-down.
    ''' </summary>
    <Description("Determines whether the drop-down container can be opened by the host control.")>
    <Category("ControlContainer")>
    <DefaultValue(True)>
    Public Property DropDownEnabled As Boolean = True
    ''' <summary>
    ''' Gets or sets the border color of the drop-down.
    ''' </summary>
    <Description("Defines the border color of the drop-down container.")>
    <Category("ControlContainer")>
    Public Property DropDownBorderColor As Color = SystemColors.HotTrack
    ''' <summary>
    ''' Gets whether the drop-down can currently be opened.
    ''' </summary>
    <Description("Indicates whether the drop-down container can currently be opened.")>
    <Category("ControlContainer")>
    Public Overridable ReadOnly Property CanDrop As Boolean
        Get
            If _DropDown IsNot Nothing Then Return False
            If _ClosedWhileInControl Then
                _ClosedWhileInControl = False
                Return False
            End If
            Return True
        End Get
    End Property
    ''' <summary>
    ''' Gets the current drop-down lifecycle state.
    ''' </summary>
    <Description("Indicates the current state of the drop-down container.")>
    <Category("ControlContainer")>
    Public ReadOnly Property DropState As ControlContainerDropDownState
        Get
            Return _DropState
        End Get
    End Property
    ''' <summary>
    ''' Displays the configured hosted control below the host control.
    ''' </summary>
    ''' <exception cref="InvalidOperationException">
    ''' Thrown when the host control or hosted control has not been assigned.
    ''' </exception>
    Public Sub ShowDropDown()
        If HostControl Is Nothing Then
            Throw New InvalidOperationException("The host control has not been defined.")
        End If
        If HostedControl Is Nothing Then
            Throw New InvalidOperationException("The hosted control has not been defined.")
        End If
        If HostControl.IsDisposed OrElse HostedControl.IsDisposed Then Return
        If Not CanDrop Then Return
        RaiseEvent Dropping(Me)
        SetDropState(ControlContainerDropDownState.Dropping)
        _OwnerForm = HostControl.FindForm()
        CreateDropDown()
        _HostParent = HostControl.Parent
        If _HostParent IsNot Nothing Then
            AddHandler _HostParent.Move, AddressOf HostParent_Move
        End If
        _DropDown.Show(HostControl, New Point(0, HostControl.Height))
    End Sub
    ''' <summary>
    ''' Closes the drop-down when it is open.
    ''' </summary>
    Public Sub CloseDropDown()
        Dim DropDown As ToolStripDropDown = _DropDown
        If DropDown Is Nothing OrElse DropDown.IsDisposed Then Return
        _AllowDropDownClose = True
        Try
            DropDown.Close(ToolStripDropDownCloseReason.CloseCalled)
        Finally
            _AllowDropDownClose = False
        End Try
    End Sub
    Private Sub HostControl_Click(sender As Object, e As EventArgs)
        If DropDownEnabled Then ShowDropDown()
    End Sub
    Private Sub HostParent_Move(sender As Object, e As EventArgs)
        CloseDropDown()
    End Sub
    Private Sub CreateDropDown()
        Dim hostedSize As Size = HostedControl.Size
        _HostPanel = New Panel With {
            .AutoSize = False,
            .BackColor = HostedControl.BackColor,
            .Margin = Padding.Empty,
            .Padding = Padding.Empty,
            .Size = hostedSize
        }
        HostedControl.Location = Point.Empty
        _HostPanel.Controls.Add(HostedControl)
        _ControlHost = New ToolStripControlHost(_HostPanel) With {
            .AutoSize = False,
            .Margin = Padding.Empty,
            .Padding = Padding.Empty,
            .Size = hostedSize
        }
        _DropDown = New ToolStripDropDown With {
            .AutoClose = True,
            .AutoSize = False,
            .BackColor = DropDownBorderColor,
            .DropShadowEnabled = False,
            .LayoutStyle = ToolStripLayoutStyle.Flow,
            .Margin = Padding.Empty,
            .Padding = New Padding(1),
            .Size = New Size(hostedSize.Width + 2, hostedSize.Height + 2)
        }
        _DropDown.Items.Add(_ControlHost)
        AddHandler _DropDown.Opened, AddressOf DropDown_Opened
        AddHandler _DropDown.Closing, AddressOf DropDown_Closing
        AddHandler _DropDown.Closed, AddressOf DropDown_Closed
    End Sub
    Private Sub DropDown_Opened(sender As Object, e As EventArgs)
        SetDropState(ControlContainerDropDownState.Dropped)
        If HostControl IsNot Nothing AndAlso Not HostControl.IsDisposed Then
            HostControl.Invalidate()
        End If
        RaiseEvent Dropped(Me)
    End Sub
    Private Sub DropDown_Closing(sender As Object, e As ToolStripDropDownClosingEventArgs)
        Dim DropDown As ToolStripDropDown = TryCast(sender, ToolStripDropDown)
        If DropDown Is Nothing Then Return
        If Not _AllowDropDownClose AndAlso IsAutomaticCloseReason(e.CloseReason) Then
            If TryScheduleAutomaticClose(DropDown, e.CloseReason) Then
                e.Cancel = True
                Return
            End If
        End If
        SetDropState(ControlContainerDropDownState.Closing)
        RaiseEvent Closing(Me)
    End Sub
    Private Sub DropDown_Closed(sender As Object, e As ToolStripDropDownClosedEventArgs)
        Dim ClosedDropDown As ToolStripDropDown = TryCast(sender, ToolStripDropDown)
        Dim ClosedControlHost As ToolStripControlHost = _ControlHost
        Dim ClosedHostPanel As Panel = _HostPanel
        Dim ClosedHostedControl As Control = _HostedControl
        Dim CleanupInvoker As Control = GetInvoker()
        If _HostParent IsNot Nothing Then
            RemoveHandler _HostParent.Move, AddressOf HostParent_Move
            _HostParent = Nothing
        End If
        If ClosedDropDown IsNot Nothing Then
            RemoveHandler ClosedDropDown.Opened, AddressOf DropDown_Opened
            RemoveHandler ClosedDropDown.Closing, AddressOf DropDown_Closing
            RemoveHandler ClosedDropDown.Closed, AddressOf DropDown_Closed
        End If
        If HostControl IsNot Nothing AndAlso Not HostControl.IsDisposed Then
            Dim HostBounds As Rectangle = HostControl.RectangleToScreen(HostControl.ClientRectangle)
            _ClosedWhileInControl = HostBounds.Contains(Cursor.Position)
            HostControl.Invalidate()
        Else
            _ClosedWhileInControl = False
        End If
        If ReferenceEquals(_DropDown, ClosedDropDown) Then
            _DropDown = Nothing
            _ControlHost = Nothing
            _HostPanel = Nothing
        End If
        _OwnerForm = Nothing
        _AutomaticCloseScheduled = False
        SetDropState(ControlContainerDropDownState.Closed)
        RaiseEvent Closed(Me)
        ScheduleCleanup(CleanupInvoker, ClosedDropDown, ClosedControlHost, ClosedHostPanel, ClosedHostedControl)
    End Sub
    Private Shared Function IsAutomaticCloseReason(CloseReason As ToolStripDropDownCloseReason) As Boolean
        Return CloseReason = ToolStripDropDownCloseReason.AppClicked OrElse CloseReason = ToolStripDropDownCloseReason.AppFocusChange
    End Function
    Private Function TryScheduleAutomaticClose(dropDown As ToolStripDropDown, closeReason As ToolStripDropDownCloseReason) As Boolean
        If _AutomaticCloseScheduled Then Return True
        Dim Invoker As Control = GetInvoker()
        If Invoker Is Nothing OrElse Invoker.IsDisposed OrElse Not Invoker.IsHandleCreated Then Return False
        _AutomaticCloseScheduled = True
        Try
            Invoker.BeginInvoke(
                New MethodInvoker(
                    Sub()
                        HandleDeferredAutomaticClose(dropDown, closeReason)
                    End Sub))
            Return True
        Catch ex As InvalidOperationException
            _AutomaticCloseScheduled = False
            Return False
        End Try
    End Function
    Private Sub HandleDeferredAutomaticClose(DropDown As ToolStripDropDown, closeReason As ToolStripDropDownCloseReason)
        _AutomaticCloseScheduled = False
        If DropDown Is Nothing OrElse DropDown.IsDisposed OrElse Not DropDown.Visible Then Return
        If Not DropDown.IsHandleCreated Then Return
        Dim ForegroundWindow As IntPtr = NativeMethods.GetForegroundWindow()
        If ForegroundWindow = DropDown.Handle Then Return
        TransferForegroundWindowOwnership(ForegroundWindow, DropDown.Handle)
        _AllowDropDownClose = True
        Try
            DropDown.Close(closeReason)
        Finally
            _AllowDropDownClose = False
        End Try
    End Sub
    Private Sub TransferForegroundWindowOwnership(WindowHandle As IntPtr, dropDownHandle As IntPtr)
        If WindowHandle = IntPtr.Zero OrElse dropDownHandle = IntPtr.Zero Then Return
        If _OwnerForm Is Nothing OrElse _OwnerForm.IsDisposed OrElse Not _OwnerForm.IsHandleCreated Then Return
        If Not IsWindowOwnedBy(WindowHandle, dropDownHandle) Then Return
        NativeMethods.SetOwner(WindowHandle, _OwnerForm.Handle)
    End Sub
    Private Shared Function IsWindowOwnedBy(windowHandle As IntPtr, ownerHandle As IntPtr) As Boolean
        Dim CurrentOwner As IntPtr = NativeMethods.GetWindow(windowHandle, NativeMethods.GW_OWNER)
        Dim InspectedOwners As Integer
        While CurrentOwner <> IntPtr.Zero AndAlso InspectedOwners < 32
            If CurrentOwner = ownerHandle Then Return True
            CurrentOwner = NativeMethods.GetWindow(CurrentOwner, NativeMethods.GW_OWNER)
            InspectedOwners += 1
        End While
        Return False
    End Function
    Private Function GetInvoker() As Control
        If HostControl IsNot Nothing AndAlso
           Not HostControl.IsDisposed AndAlso
           HostControl.IsHandleCreated Then
            Return HostControl
        End If
        If _OwnerForm IsNot Nothing AndAlso
           Not _OwnerForm.IsDisposed AndAlso
           _OwnerForm.IsHandleCreated Then
            Return _OwnerForm
        End If
        Return Nothing
    End Function
    Private Shared Sub ScheduleCleanup(Invoker As Control, dropDown As ToolStripDropDown, controlHost As ToolStripControlHost, hostPanel As Panel, hostedControl As Control)
        If Invoker IsNot Nothing AndAlso
           Not Invoker.IsDisposed AndAlso
           Invoker.IsHandleCreated Then
            Try
                Invoker.BeginInvoke(
                    New MethodInvoker(
                        Sub()
                            CleanupDropDown(dropDown, controlHost, hostPanel, hostedControl)
                        End Sub))
                Return
            Catch ex As InvalidOperationException
            End Try
        End If
        Dim SynchronizationContext As SynchronizationContext = SynchronizationContext.Current
        SynchronizationContext?.Post(
                Sub(state)
                    CleanupDropDown(dropDown, controlHost, hostPanel, hostedControl)
                End Sub,
                Nothing)
    End Sub
    Private Shared Sub CleanupDropDown(DropDown As ToolStripDropDown, ControlHost As ToolStripControlHost, HostPanel As Panel, HostedControl As Control)
        If HostPanel IsNot Nothing AndAlso
           HostedControl IsNot Nothing AndAlso
           ReferenceEquals(HostedControl.Parent, HostPanel) Then
            HostPanel.Controls.Remove(HostedControl)
        End If
        If DropDown IsNot Nothing AndAlso
           ControlHost IsNot Nothing AndAlso
           Not DropDown.IsDisposed AndAlso
           DropDown.Items.Contains(ControlHost) Then

            DropDown.Items.Remove(ControlHost)
        End If
        If ControlHost IsNot Nothing AndAlso Not ControlHost.IsDisposed Then
            ControlHost.Dispose()
        End If
        If HostPanel IsNot Nothing AndAlso Not HostPanel.IsDisposed Then
            HostPanel.Dispose()
        End If
        If DropDown IsNot Nothing AndAlso Not DropDown.IsDisposed Then
            DropDown.Dispose()
        End If
    End Sub
    Private Sub SetDropState(State As ControlContainerDropDownState)
        If _DropState = State Then Return
        _DropState = State
        RaiseEvent DropStateChanged(Me)
    End Sub
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            If _HostControl IsNot Nothing Then
                RemoveHandler _HostControl.Click, AddressOf HostControl_Click
            End If
            If _HostParent IsNot Nothing Then
                RemoveHandler _HostParent.Move, AddressOf HostParent_Move
                _HostParent = Nothing
            End If
            Dim DropDown As ToolStripDropDown = _DropDown
            If DropDown IsNot Nothing AndAlso Not DropDown.IsDisposed Then
                _AllowDropDownClose = True
                Try
                    DropDown.Close(ToolStripDropDownCloseReason.CloseCalled)
                Finally
                    _AllowDropDownClose = False
                End Try
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    Private NotInheritable Class NativeMethods
        Public Const GW_OWNER As UInteger = 4UI
        Private Const GWLP_HWNDPARENT As Integer = -8
        Private Sub New()
        End Sub
        <DllImport("user32.dll")>
        Public Shared Function GetForegroundWindow() As IntPtr
        End Function
        <DllImport("user32.dll")>
        Public Shared Function GetWindow(windowHandle As IntPtr, command As UInteger) As IntPtr
        End Function
        <DllImport("user32.dll", EntryPoint:="SetWindowLongW", SetLastError:=True)>
        Private Shared Function SetWindowLong32(windowHandle As IntPtr, index As Integer, newValue As IntPtr) As IntPtr
        End Function
        <DllImport("user32.dll", EntryPoint:="SetWindowLongPtrW", SetLastError:=True)>
        Private Shared Function SetWindowLongPtr64(windowHandle As IntPtr, index As Integer, newValue As IntPtr) As IntPtr
        End Function
        Public Shared Function SetOwner(windowHandle As IntPtr, ownerHandle As IntPtr) As IntPtr
            If IntPtr.Size = 8 Then
                Return SetWindowLongPtr64(windowHandle, GWLP_HWNDPARENT, ownerHandle)
            End If
            Return SetWindowLong32(windowHandle, GWLP_HWNDPARENT, ownerHandle)
        End Function
    End Class
End Class