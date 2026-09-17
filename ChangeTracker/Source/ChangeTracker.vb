Imports System.ComponentModel
Imports System.Linq
''' <summary>Tracks Windows Forms control properties against an explicitly accepted baseline.</summary>
''' <remarks>Configure extender properties in the designer, load the form values, and call AcceptChanges. All operations must run on the thread that created the component. Only scalar values are supported.</remarks>
<DefaultEvent("HasChangesChanged")>
<Description("Tracks configured control properties and exposes a centralized Boolean change state.")>
<DesignerCategory("Component")>
<Designer(GetType(ChangeTrackerDesigner))>
<ToolboxItem(True)>
<ToolboxItemFilter("CoreSuite")>
<ProvideProperty("TrackChanges", GetType(Control))>
<ProvideProperty("TrackedProperties", GetType(Control))>
<ProvideProperty("TrackedEvents", GetType(Control))>
<ProvideProperty("AutomaticTracking", GetType(Control))>
Public Class ChangeTracker
    Inherits Component
    Implements IExtenderProvider
    Private ReadOnly _Settings As New Dictionary(Of Control, ControlTrackingSettings)
    Private ReadOnly _OwnerThreadId As Integer = Environment.CurrentManagedThreadId
    Private _IsTracking As Boolean
    Private _IsDisposed As Boolean
    Private _IsRefreshing As Boolean
    Private _IsPublishing As Boolean
    Private _RefreshPending As Boolean
    Private _SuspensionCount As Integer
    Private _HasChanges As Boolean
    Private _ChangedPropertyCount As Integer
    ''' <summary>Occurs only when HasChanges changes between False and True.</summary>
    <Category("ChangeTracker")>
    <Description("Occurs when the centralized Boolean change state changes.")>
    Public Event HasChangesChanged As EventHandler(Of HasChangesChangedEventArgs)
    ''' <summary>Occurs for every observed property value change, even while HasChanges remains True.</summary>
    <Category("ChangeTracker")>
    <Description("Occurs when a tracked property value changes and includes the centralized Boolean state.")>
    Public Event TrackedPropertyChanged As EventHandler(Of TrackedPropertyChangedEventArgs)
    ''' <summary>Initializes a new tracker on the current UI thread.</summary>
    Public Sub New()
        MyBase.New()
    End Sub
    ''' <summary>Initializes a new tracker and adds it to the specified component container.</summary>
    ''' <param name="Container">The container that owns and disposes the tracker.</param>
    Public Sub New(Container As IContainer)
        Me.New()
        ArgumentNullException.ThrowIfNull(Container)
        Container.Add(Me)
    End Sub
    ''' <summary>Gets whether any observed property differs from its accepted baseline.</summary>
    ''' <remarks>This is cached state. Manual properties require RefreshChanges. Suspended changes are reconciled when the outermost suspension ends.</remarks>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property HasChanges As Boolean
        Get
            Return _HasChanges
        End Get
    End Property
    ''' <summary>Gets the number of observed properties differing from their accepted baselines.</summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property ChangedPropertyCount As Integer
        Get
            Return _ChangedPropertyCount
        End Get
    End Property
    ''' <summary>Gets whether a baseline has been established and tracking is active.</summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsTracking As Boolean
        Get
            Return _IsTracking
        End Get
    End Property
    ''' <summary>Gets whether one or more suspension scopes remain open.</summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property IsSuspended As Boolean
        Get
            Return _SuspensionCount > 0
        End Get
    End Property
    ''' <summary>Determines whether extender properties can be offered to an object.</summary>
    ''' <param name="Extendee">The object being inspected by the designer.</param>
    ''' <returns>True for a live Windows Forms control.</returns>
    Public Function CanExtend(Extendee As Object) As Boolean Implements IExtenderProvider.CanExtend
        Dim TargetControl As Control = TryCast(Extendee, Control)
        Return Not _IsDisposed AndAlso TargetControl IsNot Nothing AndAlso Not TargetControl.IsDisposed
    End Function
    ''' <summary>Gets whether the control participates in tracking.</summary>
    ''' <param name="TargetControl">The control to inspect.</param>
    ''' <returns>The configured participation flag, or False.</returns>
    <Category("ChangeTracker")>
    <DefaultValue(False)>
    <Description("Determines whether this control participates in change tracking.")>
    Public Function GetTrackChanges(TargetControl As Control) As Boolean
        Dim Settings As ControlTrackingSettings = FindSettings(TargetControl)
        Return Settings IsNot Nothing AndAlso Settings.TrackChanges
    End Function
    ''' <summary>Sets whether the control participates in tracking.</summary>
    ''' <param name="TargetControl">The control to configure.</param>
    ''' <param name="Value">Whether to include the control.</param>
    ''' <remarks>Stop tracking before changing configuration. Configuration never accepts existing edits implicitly.</remarks>
    Public Sub SetTrackChanges(TargetControl As Control, Value As Boolean)
        GetWritableSettings(TargetControl).TrackChanges = Value
    End Sub
    ''' <summary>Gets the comma-separated property names configured for the control.</summary>
    ''' <param name="TargetControl">The control to inspect.</param>
    ''' <returns>The configured property list, or an empty string to use a built-in default.</returns>
    <Category("ChangeTracker")>
    <DefaultValue("")>
    <Description("Comma-separated scalar property names. Empty uses a built-in default for recognized controls.")>
    Public Function GetTrackedProperties(TargetControl As Control) As String
        Dim Settings As ControlTrackingSettings = FindSettings(TargetControl)
        Return If(Settings Is Nothing, String.Empty, Settings.TrackedProperties)
    End Function
    ''' <summary>Sets the comma-separated property names to track.</summary>
    ''' <param name="TargetControl">The control to configure.</param>
    ''' <param name="Value">Simple property names, such as Text or SelectedIndex, SelectedValue. Nested paths are not supported.</param>
    Public Sub SetTrackedProperties(TargetControl As Control, Value As String)
        GetWritableSettings(TargetControl).TrackedProperties = NormalizeNames(Value)
    End Sub
    ''' <summary>Gets optional notification events used instead of conventional property notifications.</summary>
    ''' <param name="TargetControl">The control to inspect.</param>
    ''' <returns>The configured event list, or an empty string for automatic discovery.</returns>
    <Category("ChangeTracker")>
    <DefaultValue("")>
    <Description("Optional comma-separated EventHandler event names, such as TextChanged. These events replace automatic property-event discovery.")>
    Public Function GetTrackedEvents(TargetControl As Control) As String
        Dim Settings As ControlTrackingSettings = FindSettings(TargetControl)
        Return If(Settings Is Nothing, String.Empty, Settings.TrackedEvents)
    End Function
    ''' <summary>Sets notification events shared by all tracked properties of the control.</summary>
    ''' <param name="TargetControl">The control to configure.</param>
    ''' <param name="Value">Comma-separated event names with the System.EventHandler delegate type.</param>
    Public Sub SetTrackedEvents(TargetControl As Control, Value As String)
        GetWritableSettings(TargetControl).TrackedEvents = NormalizeNames(Value)
    End Sub
    ''' <summary>Gets whether the tracker subscribes to change notifications for this control.</summary>
    ''' <param name="TargetControl">The control to inspect.</param>
    ''' <returns>True unless manual observation was explicitly configured.</returns>
    <Category("ChangeTracker")>
    <DefaultValue(True)>
    <Description("Subscribes to property notifications. Set False and call RefreshChanges for controls without supported change events.")>
    Public Function GetAutomaticTracking(TargetControl As Control) As Boolean
        Dim Settings As ControlTrackingSettings = FindSettings(TargetControl)
        Return Settings Is Nothing OrElse Settings.AutomaticTracking
    End Function
    ''' <summary>Enables automatic notifications or selects explicit RefreshChanges calls.</summary>
    ''' <param name="TargetControl">The control to configure.</param>
    ''' <param name="Value">True for automatic subscriptions; False for manual observation.</param>
    Public Sub SetAutomaticTracking(TargetControl As Control, Value As Boolean)
        GetWritableSettings(TargetControl).AutomaticTracking = Value
    End Sub
    ''' <summary>Validates configuration, captures initial values, and attaches notification handlers.</summary>
    ''' <remarks>Repeated calls while active refresh the current state without replacing the baseline. Design-time calls do nothing. Failed initialization releases subscriptions and leaves tracking stopped.</remarks>
    Public Sub StartTracking()
        VerifyOperation()
        If IsDesignEnvironment() Then Return
        If _IsTracking Then
            RefreshChanges()
            Return
        End If
        Try
            For Each Pair As KeyValuePair(Of Control, ControlTrackingSettings) In _Settings
                If Not Pair.Value.TrackChanges Then Continue For
                BuildProperties(Pair.Key, Pair.Value)
            Next
            For Each Pair As KeyValuePair(Of Control, ControlTrackingSettings) In _Settings
                If Pair.Value.TrackChanges AndAlso Pair.Value.AutomaticTracking Then AttachNotifications(Pair.Key, Pair.Value)
            Next
            _IsTracking = True
        Catch
            ClearRuntimeState()
            Throw
        End Try
    End Sub
    ''' <summary>Accepts current values as the new baseline and clears the centralized change state.</summary>
    ''' <remarks>Starts tracking when necessary. Call after loading a record and only after a successful save. Values are read before any baseline is replaced. This method does not restore values or persist data.</remarks>
    Public Sub AcceptChanges()
        VerifyOperation()
        If Not _IsTracking Then
            StartTracking()
            Return
        End If
        Dim Values As Dictionary(Of TrackedPropertyState, Object) = CaptureValues()
        For Each Pair As KeyValuePair(Of TrackedPropertyState, Object) In Values
            Pair.Key.OriginalValue = Pair.Value
            Pair.Key.CurrentValue = Pair.Value
        Next
        PublishState(0)
    End Sub
    ''' <summary>Detaches notifications, discards the baseline, and clears HasChanges without changing control values.</summary>
    ''' <remarks>Configuration is retained. Call this before changing extender settings at runtime.</remarks>
    Public Sub StopTracking()
        VerifyOperation()
        If IsSuspended Then Throw New InvalidOperationException("Dispose all suspension scopes before stopping tracking.")
        ClearRuntimeState()
        PublishState(0)
    End Sub
    ''' <summary>Temporarily defers observation and returns a nestable disposable scope.</summary>
    ''' <returns>A scope that refreshes changes when the outermost scope is disposed.</returns>
    ''' <remarks>Suspension does not accept changes. To load a record without a transient dirty event, call AcceptChanges inside the scope after loading. Do not keep a scope open across Await.</remarks>
    Public Function SuspendTracking() As IDisposable
        VerifyOperation()
        _SuspensionCount += 1
        Return New TrackingSuspension(Me)
    End Function
    ''' <summary>Reads all tracked values and updates the centralized state without modifying the baseline.</summary>
    ''' <remarks>Does nothing while stopped or suspended. Automatic events also use this method. Property getters must be side-effect free. Reentrant notifications are processed after the current event batch.</remarks>
    Public Sub RefreshChanges()
        VerifyAccess()
        If Not _IsTracking OrElse IsSuspended Then Return
        If _IsRefreshing OrElse _IsPublishing Then
            _RefreshPending = True
            Return
        End If
        _IsRefreshing = True
        Try
            Do
                _RefreshPending = False
                Dim Values As Dictionary(Of TrackedPropertyState, Object) = CaptureValues()
                Dim Changes As New List(Of TrackedPropertyChangedEventArgs)
                Dim Count As Integer = Values.Where(Function(Pair) Not Object.Equals(Pair.Key.OriginalValue, Pair.Value)).Count()
                For Each Pair As KeyValuePair(Of TrackedPropertyState, Object) In Values
                    Dim State As TrackedPropertyState = Pair.Key
                    If Not Object.Equals(State.CurrentValue, Pair.Value) Then Changes.Add(New TrackedPropertyChangedEventArgs(State.TargetControl, State.Descriptor.Name, State.OriginalValue, State.CurrentValue, Pair.Value, Count > 0, Count))
                    State.CurrentValue = Pair.Value
                Next
                PublishState(Count)
                For Each Change As TrackedPropertyChangedEventArgs In Changes
                    If _IsDisposed Then Exit For
                    RaiseEvent TrackedPropertyChanged(Me, Change)
                Next
            Loop While _RefreshPending AndAlso Not _IsDisposed
        Finally
            _IsRefreshing = False
        End Try
    End Sub
    ''' <summary>Returns an immutable list of the currently observed properties differing from their baselines.</summary>
    ''' <returns>A snapshot of cached differences. Call RefreshChanges first when using manual observation.</returns>
    Public Function GetChanges() As IReadOnlyList(Of ChangeTrackingResult)
        VerifyAccess()
        Dim Results As New List(Of ChangeTrackingResult)
        For Each Settings As ControlTrackingSettings In _Settings.Values
            For Each State As TrackedPropertyState In Settings.Properties
                If Not Object.Equals(State.OriginalValue, State.CurrentValue) Then Results.Add(New ChangeTrackingResult(State.TargetControl, State.Descriptor.Name, State.OriginalValue, State.CurrentValue))
            Next
        Next
        Return Results.AsReadOnly()
    End Function
    ''' <summary>Releases all control references and event subscriptions.</summary>
    ''' <param name="Disposing">True when disposing managed resources.</param>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing AndAlso Not _IsDisposed Then
            VerifyThread()
            _IsDisposed = True
            ClearRuntimeState()
            For Each TargetControl As Control In _Settings.Keys
                RemoveHandler TargetControl.Disposed, AddressOf OnControlDisposed
            Next
            _Settings.Clear()
            _HasChanges = False
            _ChangedPropertyCount = 0
            _SuspensionCount = 0
        End If
        MyBase.Dispose(Disposing)
    End Sub
    Friend Sub EndSuspension()
        If _IsDisposed Then Return
        VerifyOperation()
        If _SuspensionCount = 0 Then Return
        _SuspensionCount -= 1
        If _SuspensionCount = 0 Then RefreshChanges()
    End Sub
    Friend Sub VerifySuspensionAccess()
        If Not _IsDisposed Then VerifyOperation()
    End Sub
    Private Function FindSettings(TargetControl As Control) As ControlTrackingSettings
        VerifyAccess()
        ArgumentNullException.ThrowIfNull(TargetControl)
        Dim Settings As ControlTrackingSettings = Nothing
        _Settings.TryGetValue(TargetControl, Settings)
        Return Settings
    End Function
    Private Function GetWritableSettings(TargetControl As Control) As ControlTrackingSettings
        VerifyOperation()
        ArgumentNullException.ThrowIfNull(TargetControl)
        If TargetControl.IsDisposed Then Throw New ObjectDisposedException(TargetControl.Name)
        If _IsTracking Then Throw New InvalidOperationException("Call StopTracking before changing tracking configuration, then call AcceptChanges to establish the new baseline.")
        Dim Settings As ControlTrackingSettings = Nothing
        If Not _Settings.TryGetValue(TargetControl, Settings) Then
            Settings = New ControlTrackingSettings
            _Settings.Add(TargetControl, Settings)
            AddHandler TargetControl.Disposed, AddressOf OnControlDisposed
        End If
        Return Settings
    End Function
    Private Shared Function NormalizeNames(Value As String) As String
        If String.IsNullOrWhiteSpace(Value) Then Return String.Empty
        Dim Names As New List(Of String)
        For Each Part As String In Value.Split(","c)
            Dim Name As String = Part.Trim()
            If Name.Length = 0 Then Throw New ArgumentException("Property and event lists cannot contain empty names.", NameOf(Value))
            If Not Names.Contains(Name, StringComparer.OrdinalIgnoreCase) Then Names.Add(Name)
        Next
        Return String.Join(", ", Names)
    End Function
    Private Shared Function ResolvePropertyNames(TargetControl As Control, Settings As ControlTrackingSettings) As String
        If Settings.TrackedProperties.Length > 0 Then Return Settings.TrackedProperties
        If TypeOf TargetControl Is TextBoxBase Then Return "Text"
        If TypeOf TargetControl Is CheckBox Then Return "CheckState"
        If TypeOf TargetControl Is RadioButton Then Return "Checked"
        If TypeOf TargetControl Is ComboBox Then Return "SelectedIndex"
        If TypeOf TargetControl Is DateTimePicker Then Return "Value"
        If TypeOf TargetControl Is NumericUpDown Then Return "Value"
        If TypeOf TargetControl Is TrackBar Then Return "Value"
        Throw New InvalidOperationException($"Control '{TargetControl.Name}' requires an explicit TrackedProperties value.")
    End Function
    Private Shared Sub BuildProperties(TargetControl As Control, Settings As ControlTrackingSettings)
        Dim Descriptors As PropertyDescriptorCollection = TypeDescriptor.GetProperties(TargetControl)
        For Each Name As String In ResolvePropertyNames(TargetControl, Settings).Split(","c)
            Dim Descriptor As PropertyDescriptor = Descriptors.Find(Name.Trim(), True)
            If Descriptor Is Nothing Then Throw New InvalidOperationException($"Property '{Name.Trim()}' was not found on control '{TargetControl.Name}'.")
            Dim State As New TrackedPropertyState With {.TargetControl = TargetControl, .Descriptor = Descriptor}
            Dim Value As Object = ReadValue(State)
            State.OriginalValue = Value
            State.CurrentValue = Value
            Settings.Properties.Add(State)
        Next
    End Sub
    Private Shared Function ReadValue(State As TrackedPropertyState) As Object
        Try
            Dim Value As Object = State.Descriptor.GetValue(State.TargetControl)
            If Value Is Nothing OrElse TypeOf Value Is String OrElse Convert.IsDBNull(Value) Then Return Value
            Dim ValueType As Type = Value.GetType()
            If ValueType.IsEnum OrElse ValueType.IsPrimitive OrElse TypeOf Value Is Decimal OrElse TypeOf Value Is DateTime OrElse TypeOf Value Is DateTimeOffset OrElse TypeOf Value Is TimeSpan OrElse TypeOf Value Is Guid OrElse TypeOf Value Is DateOnly OrElse TypeOf Value Is TimeOnly Then Return Value
            Throw New NotSupportedException($"Type '{ValueType.FullName}' is not a supported scalar snapshot. Track a scalar identifier or serialized representation instead.")
        Catch Exception As Exception
            Throw New InvalidOperationException($"Cannot read tracked property '{State.Descriptor.Name}' on control '{State.TargetControl.Name}'.", Exception)
        End Try
    End Function
    Private Function CaptureValues() As Dictionary(Of TrackedPropertyState, Object)
        Dim Values As New Dictionary(Of TrackedPropertyState, Object)
        For Each Settings As ControlTrackingSettings In _Settings.Values
            For Each State As TrackedPropertyState In Settings.Properties
                Values.Add(State, ReadValue(State))
            Next
        Next
        Return Values
    End Function
    Private Sub AttachNotifications(TargetControl As Control, Settings As ControlTrackingSettings)
        Dim Handler As EventHandler = AddressOf OnValueChanged
        Dim Events As EventDescriptorCollection = TypeDescriptor.GetEvents(TargetControl)
        If Settings.TrackedEvents.Length > 0 Then
            For Each Name As String In Settings.TrackedEvents.Split(","c)
                AttachEvent(TargetControl, Settings, Events.Find(Name.Trim(), True), Name.Trim(), Handler)
            Next
            Return
        End If
        For Each State As TrackedPropertyState In Settings.Properties
            Dim Descriptor As PropertyDescriptor = State.Descriptor
            If Descriptor.SupportsChangeEvents Then
                Descriptor.AddValueChanged(TargetControl, Handler)
                Settings.Subscriptions.Add(New TrackingSubscription(Sub() Descriptor.RemoveValueChanged(TargetControl, Handler)))
            Else
                Throw New InvalidOperationException($"Property '{Descriptor.Name}' on control '{TargetControl.Name}' has no supported change notification. Configure TrackedEvents or set AutomaticTracking to False and call RefreshChanges explicitly.")
            End If
        Next
    End Sub
    Private Shared Sub AttachEvent(TargetControl As Control, Settings As ControlTrackingSettings, Descriptor As EventDescriptor, EventName As String, Handler As EventHandler)
        If Descriptor Is Nothing OrElse Descriptor.EventType IsNot GetType(EventHandler) Then Throw New InvalidOperationException($"Event '{EventName}' on control '{TargetControl.Name}' must exist and use System.EventHandler. For other delegate types, call RefreshChanges from an application event handler.")
        Descriptor.AddEventHandler(TargetControl, Handler)
        Settings.Subscriptions.Add(New TrackingSubscription(Sub() Descriptor.RemoveEventHandler(TargetControl, Handler)))
    End Sub
    Private Sub OnValueChanged(Sender As Object, E As EventArgs)
        If Not _IsDisposed Then RefreshChanges()
    End Sub
    Private Sub OnControlDisposed(Sender As Object, E As EventArgs)
        Dim TargetControl As Control = DirectCast(Sender, Control)
        Dim Settings As ControlTrackingSettings = Nothing
        If Not _Settings.TryGetValue(TargetControl, Settings) Then Return
        DetachNotifications(Settings)
        RemoveHandler TargetControl.Disposed, AddressOf OnControlDisposed
        _Settings.Remove(TargetControl)
        If Not _IsDisposed Then RefreshChanges()
    End Sub
    Private Sub PublishState(Count As Integer)
        Dim PreviousHasChanges As Boolean = _HasChanges
        _ChangedPropertyCount = Count
        _HasChanges = Count > 0
        If PreviousHasChanges = _HasChanges Then Return
        _IsPublishing = True
        Try
            RaiseEvent HasChangesChanged(Me, New HasChangesChangedEventArgs(_HasChanges, Count))
        Finally
            _IsPublishing = False
        End Try
        If _RefreshPending AndAlso Not _IsRefreshing AndAlso Not _IsDisposed Then RefreshChanges()
    End Sub
    Private Shared Sub DetachNotifications(Settings As ControlTrackingSettings)
        For Each Subscription As TrackingSubscription In Settings.Subscriptions
            Subscription.Dispose()
        Next
        Settings.Subscriptions.Clear()
        Settings.Properties.Clear()
    End Sub
    Private Sub ClearRuntimeState()
        _IsTracking = False
        _RefreshPending = False
        For Each Settings As ControlTrackingSettings In _Settings.Values
            DetachNotifications(Settings)
        Next
    End Sub
    Private Sub VerifyThread()
        If Environment.CurrentManagedThreadId <> _OwnerThreadId Then Throw New InvalidOperationException("ChangeTracker must be used on the thread that created it.")
    End Sub
    Private Sub VerifyAccess()
        ObjectDisposedException.ThrowIf(_IsDisposed, Me)
        VerifyThread()
    End Sub
    Private Sub VerifyOperation()
        VerifyAccess()
        If _IsRefreshing OrElse _IsPublishing Then Throw New InvalidOperationException("Tracking configuration and lifecycle operations cannot run inside tracker event handlers. Defer the operation with BeginInvoke.")
    End Sub
    Private Function IsDesignEnvironment() As Boolean
        Return DesignMode OrElse LicenseManager.UsageMode = LicenseUsageMode.Designtime
    End Function
End Class
