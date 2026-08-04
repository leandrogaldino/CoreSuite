Imports System.ComponentModel
Imports System.Globalization
Imports System.Text.RegularExpressions
''' <summary>
''' Provides centralized validation rules and ErrorProvider feedback for Windows Forms controls.
''' </summary>
''' <remarks>
''' Add the component to a form and configure the extender properties displayed on each control. Validation can be performed for every configured control, for a named group, or for one control at a time.
''' </remarks>
<DefaultEvent("ValidationCompleted")>
<Description("Provides centralized validation rules and ErrorProvider feedback for Windows Forms controls.")>
<DesignerCategory("Component")>
<Designer(GetType(ValidationProviderDesigner))>
<ToolboxItem(True)>
<ToolboxItemFilter("CoreSuite")>
<ProvideProperty("Required", GetType(Control))>
<ProvideProperty("CustomValidationEnabled", GetType(Control))>
<ProvideProperty("ValidationGroup", GetType(Control))>
<ProvideProperty("ValidationMessage", GetType(Control))>
<ProvideProperty("ValidationDisplayName", GetType(Control))>
<ProvideProperty("MinimumLength", GetType(Control))>
<ProvideProperty("MaximumLength", GetType(Control))>
<ProvideProperty("RegularExpression", GetType(Control))>
<ProvideProperty("CompareWith", GetType(Control))>
<ProvideProperty("ValuePropertyName", GetType(Control))>
Public Class ValidationProvider
    Inherits ErrorProvider
    Private Const DefaultRequiredMessage As String = "{0} is required."
    Private Const DefaultMinimumLengthMessage As String = "{0} must contain at least {1} characters."
    Private Const DefaultMaximumLengthMessage As String = "{0} must contain no more than {1} characters."
    Private Const DefaultInvalidFormatMessage As String = "{0} has an invalid format."
    Private Const DefaultComparisonMessage As String = "{0} must match {1}."
    Private Const DefaultCustomMessage As String = "The value in {0} is invalid."
    Private Shared ReadOnly RegularExpressionTimeout As TimeSpan = TimeSpan.FromSeconds(1)
    Private ReadOnly _Settings As New Dictionary(Of Control, ControlValidationSettings)
    Private _AutomaticValidation As Boolean = True
    Private _CancelValidationOnError As Boolean
    Private _ClearErrorOnValueChanged As Boolean = True
    Private _ValidateDisabledControls As Boolean
    Private _ValidateHiddenControls As Boolean
    Private _TrimTextValues As Boolean = True
    Private _CaseSensitiveComparison As Boolean = True
    Private _FocusFirstInvalidControl As Boolean = True
    Private _RequiredErrorMessage As String = DefaultRequiredMessage
    Private _MinimumLengthErrorMessage As String = DefaultMinimumLengthMessage
    Private _MaximumLengthErrorMessage As String = DefaultMaximumLengthMessage
    Private _InvalidFormatErrorMessage As String = DefaultInvalidFormatMessage
    Private _ComparisonErrorMessage As String = DefaultComparisonMessage
    Private _CustomErrorMessage As String = DefaultCustomMessage
    ''' <summary>
    ''' Occurs when the provider needs the value represented by a control.
    ''' </summary>
    ''' <remarks>Handle this event to support custom controls that do not expose a conventional Text, Value, or SelectedValue property.</remarks>
    <Category("ValidationProvider")>
    <Description("Occurs when the value represented by a control is being resolved and allows the application to supply a custom value.")>
    Public Event ValidationValueRequested As EventHandler(Of ValidationValueRequestedEventArgs)
    ''' <summary>
    ''' Occurs after the built-in rules are evaluated and allows the application to change the final result.
    ''' </summary>
    <Category("ValidationProvider")>
    <Description("Occurs after built-in validation and allows application-defined rules to change the final result.")>
    Public Event ValidatingControl As EventHandler(Of ValidatingControlEventArgs)
    ''' <summary>
    ''' Occurs after a configured control has been validated and its error message has been updated.
    ''' </summary>
    <Category("ValidationProvider")>
    <Description("Occurs after an individual control has been validated.")>
    Public Event ControlValidated As EventHandler(Of ControlValidatedEventArgs)
    ''' <summary>
    ''' Occurs after a call to <see cref="Validate"/> or <see cref="ValidateGroup"/> completes.
    ''' </summary>
    <Category("ValidationProvider")>
    <Description("Occurs after a complete validation operation and exposes all control results.")>
    Public Event ValidationCompleted As EventHandler(Of ValidationCompletedEventArgs)
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ValidationProvider"/> class.
    ''' </summary>
    Public Sub New()
        MyBase.New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ValidationProvider"/> class and adds it to the specified component container.
    ''' </summary>
    ''' <param name="Container">The container that owns the component.</param>
    Public Sub New(Container As IContainer)
        MyBase.New(Container)
    End Sub
    ''' <summary>
    ''' Gets or sets a value indicating whether controls are validated automatically when their Validating event occurs.
    ''' </summary>
    ''' <value><see langword="True"/> to validate when focus leaves a configured control; otherwise, <see langword="False"/>.</value>
    <Category("ValidationProvider")>
    <DefaultValue(True)>
    <Description("Determines whether configured controls are validated automatically during their Validating event.")>
    Public Property AutomaticValidation As Boolean
        Get
            Return _AutomaticValidation
        End Get
        Set(value As Boolean)
            _AutomaticValidation = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether automatic validation cancels the control's Validating event when the value is invalid.
    ''' </summary>
    ''' <value><see langword="True"/> to request that focus remain on an invalid control; otherwise, <see langword="False"/>.</value>
    <Category("ValidationProvider")>
    <DefaultValue(False)>
    <Description("Determines whether automatic validation cancels the Validating event when a control is invalid.")>
    Public Property CancelValidationOnError As Boolean
        Get
            Return _CancelValidationOnError
        End Get
        Set(value As Boolean)
            _CancelValidationOnError = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether the current error is cleared as soon as the control value changes.
    ''' </summary>
    ''' <value><see langword="True"/> to remove stale feedback while the user edits the value; otherwise, <see langword="False"/>.</value>
    <Category("ValidationProvider")>
    <DefaultValue(True)>
    <Description("Determines whether an existing validation error is cleared when the control value changes.")>
    Public Property ClearErrorOnValueChanged As Boolean
        Get
            Return _ClearErrorOnValueChanged
        End Get
        Set(value As Boolean)
            _ClearErrorOnValueChanged = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether disabled controls participate in validation.
    ''' </summary>
    <Category("ValidationProvider")>
    <DefaultValue(False)>
    <Description("Determines whether disabled controls participate in validation operations.")>
    Public Property ValidateDisabledControls As Boolean
        Get
            Return _ValidateDisabledControls
        End Get
        Set(value As Boolean)
            _ValidateDisabledControls = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether hidden controls participate in validation.
    ''' </summary>
    <Category("ValidationProvider")>
    <DefaultValue(False)>
    <Description("Determines whether hidden controls participate in validation operations.")>
    Public Property ValidateHiddenControls As Boolean
        Get
            Return _ValidateHiddenControls
        End Get
        Set(value As Boolean)
            _ValidateHiddenControls = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether leading and trailing white space is ignored by text length and comparison rules.
    ''' </summary>
    <Category("ValidationProvider")>
    <DefaultValue(True)>
    <Description("Determines whether leading and trailing white space is ignored by text length and comparison rules.")>
    Public Property TrimTextValues As Boolean
        Get
            Return _TrimTextValues
        End Get
        Set(value As Boolean)
            _TrimTextValues = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether comparison rules distinguish uppercase and lowercase characters.
    ''' </summary>
    <Category("ValidationProvider")>
    <DefaultValue(True)>
    <Description("Determines whether values compared through CompareWith use case-sensitive text comparison.")>
    Public Property CaseSensitiveComparison As Boolean
        Get
            Return _CaseSensitiveComparison
        End Get
        Set(value As Boolean)
            _CaseSensitiveComparison = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets a value indicating whether batch validation focuses the first invalid control.
    ''' </summary>
    <Category("ValidationProvider")>
    <DefaultValue(True)>
    <Description("Determines whether Validate and ValidateGroup focus the first invalid control.")>
    Public Property FocusFirstInvalidControl As Boolean
        Get
            Return _FocusFirstInvalidControl
        End Get
        Set(value As Boolean)
            _FocusFirstInvalidControl = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the default message format used by required rules.
    ''' </summary>
    ''' <remarks>Placeholder {0} is replaced by the validation display name.</remarks>
    <Category("ValidationProvider Messages")>
    <DefaultValue(DefaultRequiredMessage)>
    <Description("Defines the default required-field message. Placeholder {0} represents the control display name.")>
    Public Property RequiredErrorMessage As String
        Get
            Return _RequiredErrorMessage
        End Get
        Set(value As String)
            _RequiredErrorMessage = NormalizeMessage(value, DefaultRequiredMessage)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the default message format used by minimum-length rules.
    ''' </summary>
    ''' <remarks>Placeholder {0} is the display name and {1} is the configured minimum length.</remarks>
    <Category("ValidationProvider Messages")>
    <DefaultValue(DefaultMinimumLengthMessage)>
    <Description("Defines the minimum-length message. Placeholders {0} and {1} represent the display name and minimum length.")>
    Public Property MinimumLengthErrorMessage As String
        Get
            Return _MinimumLengthErrorMessage
        End Get
        Set(value As String)
            _MinimumLengthErrorMessage = NormalizeMessage(value, DefaultMinimumLengthMessage)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the default message format used by maximum-length rules.
    ''' </summary>
    ''' <remarks>Placeholder {0} is the display name and {1} is the configured maximum length.</remarks>
    <Category("ValidationProvider Messages")>
    <DefaultValue(DefaultMaximumLengthMessage)>
    <Description("Defines the maximum-length message. Placeholders {0} and {1} represent the display name and maximum length.")>
    Public Property MaximumLengthErrorMessage As String
        Get
            Return _MaximumLengthErrorMessage
        End Get
        Set(value As String)
            _MaximumLengthErrorMessage = NormalizeMessage(value, DefaultMaximumLengthMessage)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the default message format used by regular-expression and incomplete-mask rules.
    ''' </summary>
    <Category("ValidationProvider Messages")>
    <DefaultValue(DefaultInvalidFormatMessage)>
    <Description("Defines the invalid-format message. Placeholder {0} represents the control display name.")>
    Public Property InvalidFormatErrorMessage As String
        Get
            Return _InvalidFormatErrorMessage
        End Get
        Set(value As String)
            _InvalidFormatErrorMessage = NormalizeMessage(value, DefaultInvalidFormatMessage)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the default message format used by comparison rules.
    ''' </summary>
    ''' <remarks>Placeholder {0} is the current display name and {1} is the comparison control display name.</remarks>
    <Category("ValidationProvider Messages")>
    <DefaultValue(DefaultComparisonMessage)>
    <Description("Defines the comparison message. Placeholders {0} and {1} represent both control display names.")>
    Public Property ComparisonErrorMessage As String
        Get
            Return _ComparisonErrorMessage
        End Get
        Set(value As String)
            _ComparisonErrorMessage = NormalizeMessage(value, DefaultComparisonMessage)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the fallback message format used by custom validation rules.
    ''' </summary>
    <Category("ValidationProvider Messages")>
    <DefaultValue(DefaultCustomMessage)>
    <Description("Defines the fallback custom-validation message. Placeholder {0} represents the control display name.")>
    Public Property CustomErrorMessage As String
        Get
            Return _CustomErrorMessage
        End Get
        Set(value As String)
            _CustomErrorMessage = NormalizeMessage(value, DefaultCustomMessage)
        End Set
    End Property
    ''' <summary>
    ''' Gets whether the specified control must contain a value.
    ''' </summary>
    ''' <param name="TargetControl">The control whose required setting is returned.</param>
    ''' <returns><see langword="True"/> when the control is required; otherwise, <see langword="False"/>.</returns>
    <Category("ValidationProvider")>
    <DefaultValue(False)>
    <Description("Determines whether the control must contain a value.")>
    Public Function GetRequired(TargetControl As Control) As Boolean
        Dim Settings As ControlValidationSettings = GetSettings(TargetControl)
        Return Settings IsNot Nothing AndAlso Settings.Required
    End Function
    ''' <summary>
    ''' Sets whether the specified control must contain a value.
    ''' </summary>
    ''' <param name="TargetControl">The control to configure.</param>
    ''' <param name="Value"><see langword="True"/> to require a value; otherwise, <see langword="False"/>.</param>
    Public Sub SetRequired(TargetControl As Control, Value As Boolean)
        Dim Settings As ControlValidationSettings = GetOrCreateSettings(TargetControl)
        Settings.Required = Value
        UpdateControlRegistration(TargetControl, Settings)
    End Sub
    ''' <summary>
    ''' Gets whether the specified control is explicitly included for application-defined validation.
    ''' </summary>
    ''' <param name="TargetControl">The control whose explicit validation state is returned.</param>
    ''' <returns><see langword="True"/> when the control participates even without another configured rule; otherwise, <see langword="False"/>.</returns>
    <Category("ValidationProvider")>
    <DefaultValue(False)>
    <Description("Includes the control in validation even when it has no built-in rule, allowing custom-only validation.")>
    Public Function GetCustomValidationEnabled(TargetControl As Control) As Boolean
        Dim Settings As ControlValidationSettings = GetSettings(TargetControl)
        Return Settings IsNot Nothing AndAlso Settings.CustomValidationEnabled
    End Function
    ''' <summary>
    ''' Sets whether the specified control is explicitly included for application-defined validation.
    ''' </summary>
    ''' <param name="TargetControl">The control to configure.</param>
    ''' <param name="Value"><see langword="True"/> to include the control for custom-only rules; otherwise, <see langword="False"/>.</param>
    Public Sub SetCustomValidationEnabled(TargetControl As Control, Value As Boolean)
        Dim Settings As ControlValidationSettings = GetOrCreateSettings(TargetControl)
        Settings.CustomValidationEnabled = Value
        UpdateControlRegistration(TargetControl, Settings)
    End Sub
    ''' <summary>
    ''' Gets the validation group assigned to the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control whose group is returned.</param>
    ''' <returns>The configured validation group, or an empty string.</returns>
    <Category("ValidationProvider")>
    <DefaultValue("")>
    <Description("Defines the validation group assigned to the control.")>
    Public Function GetValidationGroup(TargetControl As Control) As String
        Dim Settings As ControlValidationSettings = GetSettings(TargetControl)
        Return If(Settings Is Nothing, String.Empty, Settings.ValidationGroup)
    End Function
    ''' <summary>
    ''' Sets the validation group assigned to the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control to configure.</param>
    ''' <param name="Value">The group name, or an empty string for no named group.</param>
    Public Sub SetValidationGroup(TargetControl As Control, Value As String)
        Dim Settings As ControlValidationSettings = GetOrCreateSettings(TargetControl)
        Settings.ValidationGroup = If(Value, String.Empty).Trim()
        UpdateControlRegistration(TargetControl, Settings)
    End Sub
    ''' <summary>
    ''' Gets the custom error message assigned to the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control whose custom message is returned.</param>
    ''' <returns>The custom error message, or an empty string when generated messages are used.</returns>
    <Category("ValidationProvider")>
    <DefaultValue("")>
    <Description("Defines a custom message that replaces the default message when the control is invalid.")>
    Public Function GetValidationMessage(TargetControl As Control) As String
        Dim Settings As ControlValidationSettings = GetSettings(TargetControl)
        Return If(Settings Is Nothing, String.Empty, Settings.ValidationMessage)
    End Function
    ''' <summary>
    ''' Sets the custom error message assigned to the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control to configure.</param>
    ''' <param name="Value">The replacement message, or an empty string to use generated messages.</param>
    Public Sub SetValidationMessage(TargetControl As Control, Value As String)
        Dim Settings As ControlValidationSettings = GetOrCreateSettings(TargetControl)
        Settings.ValidationMessage = If(Value, String.Empty)
        UpdateControlRegistration(TargetControl, Settings)
    End Sub
    ''' <summary>
    ''' Gets the friendly name used in generated messages for the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control whose display name is returned.</param>
    ''' <returns>The configured friendly name, or an empty string when the provider should infer one.</returns>
    <Category("ValidationProvider")>
    <DefaultValue("")>
    <Description("Defines the friendly control name inserted into generated validation messages.")>
    Public Function GetValidationDisplayName(TargetControl As Control) As String
        Dim Settings As ControlValidationSettings = GetSettings(TargetControl)
        Return If(Settings Is Nothing, String.Empty, Settings.ValidationDisplayName)
    End Function
    ''' <summary>
    ''' Sets the friendly name used in generated messages for the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control to configure.</param>
    ''' <param name="Value">The friendly field name, or an empty string to infer it automatically.</param>
    Public Sub SetValidationDisplayName(TargetControl As Control, Value As String)
        Dim Settings As ControlValidationSettings = GetOrCreateSettings(TargetControl)
        Settings.ValidationDisplayName = If(Value, String.Empty).Trim()
        UpdateControlRegistration(TargetControl, Settings)
    End Sub
    ''' <summary>
    ''' Gets the minimum text length required by the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control whose minimum length is returned.</param>
    ''' <returns>The minimum character count, or zero when the rule is disabled.</returns>
    <Category("ValidationProvider")>
    <DefaultValue(0)>
    <Description("Defines the minimum number of characters accepted by the control. Zero disables this rule.")>
    Public Function GetMinimumLength(TargetControl As Control) As Integer
        Dim Settings As ControlValidationSettings = GetSettings(TargetControl)
        Return If(Settings Is Nothing, 0, Settings.MinimumLength)
    End Function
    ''' <summary>
    ''' Sets the minimum text length required by the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control to configure.</param>
    ''' <param name="Value">The minimum character count, or zero to disable the rule.</param>
    ''' <exception cref="ArgumentOutOfRangeException">The value is less than zero.</exception>
    Public Sub SetMinimumLength(TargetControl As Control, Value As Integer)
        If Value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(Value), Value, "MinimumLength cannot be less than zero.")
        Dim Settings As ControlValidationSettings = GetOrCreateSettings(TargetControl)
        Settings.MinimumLength = Value
        UpdateControlRegistration(TargetControl, Settings)
    End Sub
    ''' <summary>
    ''' Gets the maximum text length accepted by the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control whose maximum length is returned.</param>
    ''' <returns>The maximum character count, or zero when the rule is disabled.</returns>
    <Category("ValidationProvider")>
    <DefaultValue(0)>
    <Description("Defines the maximum number of characters accepted by the control. Zero disables this rule.")>
    Public Function GetMaximumLength(TargetControl As Control) As Integer
        Dim Settings As ControlValidationSettings = GetSettings(TargetControl)
        Return If(Settings Is Nothing, 0, Settings.MaximumLength)
    End Function
    ''' <summary>
    ''' Sets the maximum text length accepted by the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control to configure.</param>
    ''' <param name="Value">The maximum character count, or zero to disable the rule.</param>
    ''' <exception cref="ArgumentOutOfRangeException">The value is less than zero.</exception>
    Public Sub SetMaximumLength(TargetControl As Control, Value As Integer)
        If Value < 0 Then Throw New ArgumentOutOfRangeException(NameOf(Value), Value, "MaximumLength cannot be less than zero.")
        Dim Settings As ControlValidationSettings = GetOrCreateSettings(TargetControl)
        Settings.MaximumLength = Value
        UpdateControlRegistration(TargetControl, Settings)
    End Sub
    ''' <summary>
    ''' Gets the regular expression used to validate the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control whose expression is returned.</param>
    ''' <returns>The configured regular expression, or an empty string when the rule is disabled.</returns>
    <Category("ValidationProvider")>
    <DefaultValue("")>
    <Description("Defines the regular expression used to validate the control text. An empty value disables this rule.")>
    Public Function GetRegularExpression(TargetControl As Control) As String
        Dim Settings As ControlValidationSettings = GetSettings(TargetControl)
        Return If(Settings Is Nothing, String.Empty, Settings.RegularExpression)
    End Function
    ''' <summary>
    ''' Sets the regular expression used to validate the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control to configure.</param>
    ''' <param name="Value">The regular expression, or an empty string to disable the rule.</param>
    ''' <exception cref="ArgumentException">The supplied expression is not a valid regular expression.</exception>
    Public Sub SetRegularExpression(TargetControl As Control, Value As String)
        Dim ExpressionText As String = If(Value, String.Empty)
        Dim CompiledExpression As Regex = Nothing
        If Not String.IsNullOrEmpty(ExpressionText) Then
            Try
                CompiledExpression = New Regex(ExpressionText, RegexOptions.None, RegularExpressionTimeout)
            Catch Failure As ArgumentException
                Throw New ArgumentException("RegularExpression must be a valid regular expression.", NameOf(Value), Failure)
            End Try
        End If
        Dim Settings As ControlValidationSettings = GetOrCreateSettings(TargetControl)
        Settings.RegularExpression = ExpressionText
        Settings.CompiledRegularExpression = CompiledExpression
        UpdateControlRegistration(TargetControl, Settings)
    End Sub
    ''' <summary>
    ''' Gets the control whose represented value must match the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The configured control whose comparison target is returned.</param>
    ''' <returns>The comparison control, or <see langword="Nothing"/> when the rule is disabled.</returns>
    <Category("ValidationProvider")>
    <DefaultValue(GetType(Control), Nothing)>
    <Description("Defines another control whose represented value must match this control.")>
    Public Function GetCompareWith(TargetControl As Control) As Control
        Dim Settings As ControlValidationSettings = GetSettings(TargetControl)
        Return Settings?.CompareWith
    End Function
    ''' <summary>
    ''' Sets the control whose represented value must match the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control to configure.</param>
    ''' <param name="Value">The comparison control, or <see langword="Nothing"/> to disable the rule.</param>
    ''' <exception cref="ArgumentException">The comparison target is the same control being configured.</exception>
    Public Sub SetCompareWith(TargetControl As Control, Value As Control)
        EnsureTargetControl(TargetControl)
        If ReferenceEquals(TargetControl, Value) Then Throw New ArgumentException("A control cannot be compared with itself.", NameOf(Value))
        Dim Settings As ControlValidationSettings = GetOrCreateSettings(TargetControl)
        Settings.CompareWith = Value
        UpdateControlRegistration(TargetControl, Settings)
    End Sub
    ''' <summary>
    ''' Gets the optional property path used to retrieve the value represented by the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control whose value property path is returned.</param>
    ''' <returns>The configured property path, or an empty string for automatic value resolution.</returns>
    <Category("ValidationProvider")>
    <DefaultValue("")>
    <Description("Defines an optional property path used as the validation value, such as SelectedValue or Frozen.IsFrozen.")>
    Public Function GetValuePropertyName(TargetControl As Control) As String
        Dim Settings As ControlValidationSettings = GetSettings(TargetControl)
        Return If(Settings Is Nothing, String.Empty, Settings.ValuePropertyName)
    End Function
    ''' <summary>
    ''' Sets the optional property path used to retrieve the value represented by the specified control.
    ''' </summary>
    ''' <param name="TargetControl">The control to configure.</param>
    ''' <param name="Value">A property name or dot-separated property path, or an empty string for automatic resolution.</param>
    Public Sub SetValuePropertyName(TargetControl As Control, Value As String)
        Dim Settings As ControlValidationSettings = GetOrCreateSettings(TargetControl)
        Settings.ValuePropertyName = If(Value, String.Empty).Trim()
        UpdateControlRegistration(TargetControl, Settings)
    End Sub
    ''' <summary>
    ''' Validates every configured control that is eligible under the current provider settings.
    ''' </summary>
    ''' <returns><see langword="True"/> when every evaluated control is valid; otherwise, <see langword="False"/>.</returns>
    Public Function Validate() As Boolean
        Return ValidateInternal(String.Empty, False)
    End Function
    ''' <summary>
    ''' Validates configured controls assigned to the specified validation group.
    ''' </summary>
    ''' <param name="GroupName">The case-insensitive group name to validate. An empty string validates ungrouped controls.</param>
    ''' <returns><see langword="True"/> when every evaluated control in the group is valid; otherwise, <see langword="False"/>.</returns>
    Public Function ValidateGroup(GroupName As String) As Boolean
        ArgumentNullException.ThrowIfNull(GroupName)
        Return ValidateInternal(GroupName.Trim(), True)
    End Function
    ''' <summary>
    ''' Validates one configured control and updates its ErrorProvider message.
    ''' </summary>
    ''' <param name="TargetControl">The control to validate.</param>
    ''' <returns><see langword="True"/> when the control is valid, not configured, or currently excluded; otherwise, <see langword="False"/>.</returns>
    Public Function ValidateControl(TargetControl As Control) As Boolean
        EnsureTargetControl(TargetControl)
        Dim Settings As ControlValidationSettings = GetSettings(TargetControl)
        If Settings Is Nothing Then Return True
        If Not ShouldValidateControl(TargetControl) Then
            SetError(TargetControl, String.Empty)
            Return True
        End If
        Return ValidateSingleControl(TargetControl, Settings).IsValid
    End Function
    ''' <summary>
    ''' Clears every validation message currently displayed by this provider.
    ''' </summary>
    Public Sub ClearValidation()
        MyBase.Clear()
    End Sub
    ''' <summary>
    ''' Clears the validation message displayed for one control.
    ''' </summary>
    ''' <param name="TargetControl">The control whose message is removed.</param>
    Public Sub ClearValidation(TargetControl As Control)
        EnsureTargetControl(TargetControl)
        SetError(TargetControl, String.Empty)
    End Sub
    ''' <summary>
    ''' Clears validation messages for controls assigned to the specified group.
    ''' </summary>
    ''' <param name="GroupName">The case-insensitive validation group to clear.</param>
    Public Sub ClearValidationGroup(GroupName As String)
        ArgumentNullException.ThrowIfNull(GroupName)
        Dim NormalizedGroup As String = GroupName.Trim()
        For Each Pair As KeyValuePair(Of Control, ControlValidationSettings) In New Dictionary(Of Control, ControlValidationSettings)(_Settings)
            If String.Equals(Pair.Value.ValidationGroup, NormalizedGroup, StringComparison.OrdinalIgnoreCase) AndAlso Not Pair.Key.IsDisposed Then SetError(Pair.Key, String.Empty)
        Next
    End Sub
    ''' <summary>
    ''' Raises the <see cref="ValidationValueRequested"/> event.
    ''' </summary>
    ''' <param name="E">The event data associated with the value request.</param>
    Protected Overridable Sub OnValidationValueRequested(E As ValidationValueRequestedEventArgs)
        RaiseEvent ValidationValueRequested(Me, E)
    End Sub
    ''' <summary>
    ''' Raises the <see cref="ValidatingControl"/> event.
    ''' </summary>
    ''' <param name="E">The event data associated with the control being validated.</param>
    Protected Overridable Sub OnValidatingControl(E As ValidatingControlEventArgs)
        RaiseEvent ValidatingControl(Me, E)
    End Sub
    ''' <summary>
    ''' Raises the <see cref="ControlValidated"/> event.
    ''' </summary>
    ''' <param name="E">The event data containing the completed control result.</param>
    Protected Overridable Sub OnControlValidated(E As ControlValidatedEventArgs)
        RaiseEvent ControlValidated(Me, E)
    End Sub
    ''' <summary>
    ''' Raises the <see cref="ValidationCompleted"/> event.
    ''' </summary>
    ''' <param name="E">The event data containing the completed batch results.</param>
    Protected Overridable Sub OnValidationCompleted(E As ValidationCompletedEventArgs)
        RaiseEvent ValidationCompleted(Me, E)
    End Sub
    ''' <summary>
    ''' Releases the resources used by the provider and detaches all observed control events.
    ''' </summary>
    ''' <param name="Disposing"><see langword="True"/> to release managed resources; otherwise, <see langword="False"/>.</param>
    Protected Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            For Each TargetControl As Control In New List(Of Control)(_Settings.Keys)
                RemoveControl(TargetControl)
            Next
            MyBase.Clear()
        End If
        MyBase.Dispose(Disposing)
    End Sub
    Private Function ValidateInternal(GroupName As String, FilterByGroup As Boolean) As Boolean
        Dim Results As New List(Of ValidationResult)
        For Each TargetControl As Control In New List(Of Control)(_Settings.Keys)
            Dim Settings As ControlValidationSettings = GetSettings(TargetControl)
            If Settings Is Nothing OrElse TargetControl.IsDisposed Then Continue For
            If FilterByGroup AndAlso Not String.Equals(Settings.ValidationGroup, GroupName, StringComparison.OrdinalIgnoreCase) Then Continue For
            If Not ShouldValidateControl(TargetControl) Then
                SetError(TargetControl, String.Empty)
                Continue For
            End If
            Results.Add(ValidateSingleControl(TargetControl, Settings))
        Next
        Dim CompletedEventArgs As New ValidationCompletedEventArgs(Results)
        If FocusFirstInvalidControl Then FocusFirstInvalidResult(Results)
        OnValidationCompleted(CompletedEventArgs)
        Return CompletedEventArgs.IsValid
    End Function
    Private Function ValidateSingleControl(TargetControl As Control, Settings As ControlValidationSettings) As ValidationResult
        Dim Value As Object = ResolveValidationValue(TargetControl, Settings)
        Dim FailureReason As ValidationFailureReason = ValidationFailureReason.None
        Dim ErrorMessage As String = String.Empty
        Dim DisplayName As String = ResolveDisplayName(TargetControl, Settings)
        If Settings.Required AndAlso IsEmptyValue(Value) Then
            FailureReason = ValidationFailureReason.Required
            ErrorMessage = ResolveErrorMessage(Settings, FailureReason, DisplayName, Nothing)
        ElseIf Not IsEmptyValue(Value) Then
            Dim TextValue As String = ConvertValueToText(Value)
            Dim MaskedControl As MaskedTextBox = TryCast(TargetControl, MaskedTextBox)
            If MaskedControl IsNot Nothing AndAlso Not String.IsNullOrEmpty(MaskedControl.Mask) AndAlso Not MaskedControl.MaskCompleted Then
                FailureReason = ValidationFailureReason.InvalidFormat
                ErrorMessage = ResolveErrorMessage(Settings, FailureReason, DisplayName, Nothing)
            ElseIf Settings.MinimumLength > 0 AndAlso TextValue.Length < Settings.MinimumLength Then
                FailureReason = ValidationFailureReason.MinimumLength
                ErrorMessage = ResolveErrorMessage(Settings, FailureReason, DisplayName, Nothing)
            ElseIf Settings.MaximumLength > 0 AndAlso TextValue.Length > Settings.MaximumLength Then
                FailureReason = ValidationFailureReason.MaximumLength
                ErrorMessage = ResolveErrorMessage(Settings, FailureReason, DisplayName, Nothing)
            ElseIf Settings.CompiledRegularExpression IsNot Nothing AndAlso Not MatchesRegularExpression(Settings.CompiledRegularExpression, TextValue) Then
                FailureReason = ValidationFailureReason.InvalidFormat
                ErrorMessage = ResolveErrorMessage(Settings, FailureReason, DisplayName, Nothing)
            End If
        End If
        If FailureReason = ValidationFailureReason.None AndAlso Settings.CompareWith IsNot Nothing AndAlso Not Settings.CompareWith.IsDisposed Then
            Dim ComparisonSettings As ControlValidationSettings = GetSettings(Settings.CompareWith)
            If ComparisonSettings Is Nothing Then ComparisonSettings = New ControlValidationSettings
            Dim ComparisonValue As Object = ResolveValidationValue(Settings.CompareWith, ComparisonSettings)
            If Not ValuesMatch(Value, ComparisonValue) Then
                FailureReason = ValidationFailureReason.Comparison
                ErrorMessage = ResolveErrorMessage(Settings, FailureReason, DisplayName, ResolveDisplayName(Settings.CompareWith, ComparisonSettings))
            End If
        End If
        Dim EventArgs As New ValidatingControlEventArgs(TargetControl, Value, Settings.ValidationGroup, FailureReason = ValidationFailureReason.None, ErrorMessage, FailureReason)
        OnValidatingControl(EventArgs)
        If EventArgs.IsValid Then
            EventArgs.ErrorMessage = String.Empty
            EventArgs.FailureReason = ValidationFailureReason.None
        Else
            If EventArgs.FailureReason = ValidationFailureReason.None Then EventArgs.FailureReason = ValidationFailureReason.Custom
            If String.IsNullOrWhiteSpace(EventArgs.ErrorMessage) Then EventArgs.ErrorMessage = ResolveErrorMessage(Settings, ValidationFailureReason.Custom, DisplayName, Nothing)
        End If
        SetError(TargetControl, EventArgs.ErrorMessage)
        Dim Result As New ValidationResult(TargetControl, Settings.ValidationGroup, EventArgs.IsValid, EventArgs.ErrorMessage, EventArgs.FailureReason)
        OnControlValidated(New ControlValidatedEventArgs(Result))
        Return Result
    End Function
    Private Function ResolveValidationValue(TargetControl As Control, Settings As ControlValidationSettings) As Object
        Dim ResolvedValue As Object = ResolveDefaultValue(TargetControl, Settings)
        Dim EventArgs As New ValidationValueRequestedEventArgs(TargetControl, ResolvedValue)
        OnValidationValueRequested(EventArgs)
        Return If(EventArgs.Handled, EventArgs.Value, ResolvedValue)
    End Function
    Private Shared Function ResolveDefaultValue(TargetControl As Control, Settings As ControlValidationSettings) As Object
        If Not String.IsNullOrEmpty(Settings.ValuePropertyName) Then Return ResolvePropertyPath(TargetControl, Settings.ValuePropertyName)
        Dim HasValueState As Boolean? = ResolveHasValueState(TargetControl)
        If HasValueState.HasValue AndAlso Not HasValueState.Value Then Return Nothing
        If TypeOf TargetControl Is CheckBox Then Return DirectCast(TargetControl, CheckBox).Checked
        If TypeOf TargetControl Is RadioButton Then Return DirectCast(TargetControl, RadioButton).Checked
        If TypeOf TargetControl Is CheckedListBox Then
            Dim CheckedControl As CheckedListBox = DirectCast(TargetControl, CheckedListBox)
            If CheckedControl.CheckedItems.Count = 0 Then Return Nothing
            Return New ArrayList(CheckedControl.CheckedItems)
        End If
        If TypeOf TargetControl Is ComboBox Then
            Dim ComboControl As ComboBox = DirectCast(TargetControl, ComboBox)
            If ComboControl.DropDownStyle = ComboBoxStyle.DropDownList AndAlso ComboControl.SelectedIndex < 0 Then Return Nothing
            Return ComboControl.Text
        End If
        If TypeOf TargetControl Is ListBox Then
            Dim ListControl As ListBox = DirectCast(TargetControl, ListBox)
            If ListControl.SelectedIndex < 0 Then Return Nothing
            Return If(ListControl.SelectedValue, ListControl.SelectedItem)
        End If
        If TypeOf TargetControl Is ListView Then
            Dim ListControl As ListView = DirectCast(TargetControl, ListView)
            If ListControl.SelectedItems.Count = 0 Then Return Nothing
            Return ListControl.SelectedItems
        End If
        If TypeOf TargetControl Is TreeView Then
            Dim TreeControl As TreeView = DirectCast(TargetControl, TreeView)
            Return TreeControl.SelectedNode?.Text
        End If
        If TypeOf TargetControl Is DateTimePicker Then
            Dim Picker As DateTimePicker = DirectCast(TargetControl, DateTimePicker)
            If Picker.ShowCheckBox AndAlso Not Picker.Checked Then Return Nothing
            Return Picker.Value
        End If
        If TypeOf TargetControl Is NumericUpDown Then Return DirectCast(TargetControl, NumericUpDown).Value
        Dim ValueProperty As PropertyDescriptor = TypeDescriptor.GetProperties(TargetControl)("Value")
        If ValueProperty IsNot Nothing AndAlso ValueProperty.IsBrowsable Then Return ValueProperty.GetValue(TargetControl)
        Dim SelectedValueProperty As PropertyDescriptor = TypeDescriptor.GetProperties(TargetControl)("SelectedValue")
        If SelectedValueProperty IsNot Nothing Then Return SelectedValueProperty.GetValue(TargetControl)
        Return TargetControl.Text
    End Function
    Private Shared Function ResolveHasValueState(TargetControl As Control) As Boolean?
        Dim PropertyNames As String() = {"HasDateTime", "HasDate", "HasTime", "HasValue"}
        Dim Properties As PropertyDescriptorCollection = TypeDescriptor.GetProperties(TargetControl)
        For Each PropertyName As String In PropertyNames
            Dim Descriptor As PropertyDescriptor = Properties(PropertyName)
            If Descriptor IsNot Nothing AndAlso Descriptor.PropertyType Is GetType(Boolean) Then Return CBool(Descriptor.GetValue(TargetControl))
        Next
        Return Nothing
    End Function
    Private Shared Function ResolvePropertyPath(Source As Object, PropertyPath As String) As Object
        Dim CurrentValue As Object = Source
        For Each PropertyName As String In PropertyPath.Split("."c)
            If CurrentValue Is Nothing Then Return Nothing
            Dim Descriptor As PropertyDescriptor = TypeDescriptor.GetProperties(CurrentValue)(PropertyName)
            If Descriptor Is Nothing Then Throw New InvalidOperationException($"Property '{PropertyName}' from path '{PropertyPath}' was not found on type '{CurrentValue.GetType().FullName}'.")
            CurrentValue = Descriptor.GetValue(CurrentValue)
        Next
        Return CurrentValue
    End Function
    Private Shared Function IsEmptyValue(Value As Object) As Boolean
        If Value Is Nothing OrElse Convert.IsDBNull(Value) Then Return True
        If TypeOf Value Is String Then Return String.IsNullOrWhiteSpace(DirectCast(Value, String))
        If TypeOf Value Is Boolean Then Return Not CBool(Value)
        If TypeOf Value Is DateTime Then Return CDate(Value) = DateTime.MinValue
        If TypeOf Value Is Array Then Return DirectCast(Value, Array).Length = 0
        If TypeOf Value Is ICollection Then Return DirectCast(Value, ICollection).Count = 0
        Return False
    End Function
    Private Function ConvertValueToText(Value As Object) As String
        Dim TextValue As String = Convert.ToString(Value, CultureInfo.CurrentCulture)
        If TextValue Is Nothing Then TextValue = String.Empty
        Return If(TrimTextValues, TextValue.Trim(), TextValue)
    End Function
    Private Function ValuesMatch(FirstValue As Object, SecondValue As Object) As Boolean
        If IsEmptyValue(FirstValue) AndAlso IsEmptyValue(SecondValue) Then Return True
        If IsEmptyValue(FirstValue) Xor IsEmptyValue(SecondValue) Then Return False
        Dim Comparison As StringComparison = If(CaseSensitiveComparison, StringComparison.CurrentCulture, StringComparison.CurrentCultureIgnoreCase)
        Return String.Equals(ConvertValueToText(FirstValue), ConvertValueToText(SecondValue), Comparison)
    End Function
    Private Shared Function MatchesRegularExpression(Expression As Regex, Value As String) As Boolean
        Try
            Return Expression.IsMatch(Value)
        Catch Failure As RegexMatchTimeoutException
            Return False
        End Try
    End Function
    Private Shared Function ResolveDisplayName(TargetControl As Control, Settings As ControlValidationSettings) As String
        If Not String.IsNullOrWhiteSpace(Settings.ValidationDisplayName) Then Return Settings.ValidationDisplayName
        If Not String.IsNullOrWhiteSpace(TargetControl.AccessibleName) Then Return TargetControl.AccessibleName
        If Not String.IsNullOrWhiteSpace(TargetControl.Name) Then Return TargetControl.Name
        Return TargetControl.GetType().Name
    End Function
    Private Function ResolveErrorMessage(Settings As ControlValidationSettings, FailureReason As ValidationFailureReason, DisplayName As String, ComparisonDisplayName As String) As String
        If Not String.IsNullOrWhiteSpace(Settings.ValidationMessage) Then Return Settings.ValidationMessage
        Select Case FailureReason
            Case ValidationFailureReason.Required
                Return FormatMessage(RequiredErrorMessage, DisplayName)
            Case ValidationFailureReason.MinimumLength
                Return FormatMessage(MinimumLengthErrorMessage, DisplayName, Settings.MinimumLength)
            Case ValidationFailureReason.MaximumLength
                Return FormatMessage(MaximumLengthErrorMessage, DisplayName, Settings.MaximumLength)
            Case ValidationFailureReason.InvalidFormat
                Return FormatMessage(InvalidFormatErrorMessage, DisplayName)
            Case ValidationFailureReason.Comparison
                Return FormatMessage(ComparisonErrorMessage, DisplayName, ComparisonDisplayName)
            Case Else
                Return FormatMessage(CustomErrorMessage, DisplayName)
        End Select
    End Function
    Private Shared Function FormatMessage(MessageFormat As String, ParamArray Arguments As Object()) As String
        Try
            Return String.Format(CultureInfo.CurrentCulture, MessageFormat, Arguments)
        Catch Failure As FormatException
            Return MessageFormat
        End Try
    End Function
    Private Shared Function NormalizeMessage(Value As String, Fallback As String) As String
        Return If(String.IsNullOrEmpty(Value), Fallback, Value)
    End Function
    Private Function ShouldValidateControl(TargetControl As Control) As Boolean
        If TargetControl.IsDisposed OrElse TargetControl.Disposing Then Return False
        If Not ValidateDisabledControls AndAlso Not TargetControl.Enabled Then Return False
        If Not ValidateHiddenControls AndAlso Not TargetControl.Visible Then Return False
        Return True
    End Function
    Private Shared Sub FocusFirstInvalidResult(Results As IEnumerable(Of ValidationResult))
        For Each Result As ValidationResult In Results
            If Not Result.IsValid AndAlso Result.TargetControl.CanSelect Then
                Result.TargetControl.Select()
                Return
            End If
        Next
    End Sub
    Private Function GetSettings(TargetControl As Control) As ControlValidationSettings
        EnsureTargetControl(TargetControl)
        Dim Settings As ControlValidationSettings = Nothing
        If _Settings.TryGetValue(TargetControl, Settings) Then Return Settings
        Return Nothing
    End Function
    Private Function GetOrCreateSettings(TargetControl As Control) As ControlValidationSettings
        EnsureTargetControl(TargetControl)
        Dim Settings As ControlValidationSettings = Nothing
        If _Settings.TryGetValue(TargetControl, Settings) Then Return Settings
        Settings = New ControlValidationSettings
        _Settings.Add(TargetControl, Settings)
        Return Settings
    End Function
    Private Sub UpdateControlRegistration(TargetControl As Control, Settings As ControlValidationSettings)
        If Settings.IsDefault Then
            RemoveControl(TargetControl)
            Return
        End If
        If Settings.IsRegistered Then Return
        AddHandler TargetControl.Validating, AddressOf TargetControl_Validating
        AddHandler TargetControl.TextChanged, AddressOf TargetControl_ValueChanged
        AddHandler TargetControl.Disposed, AddressOf TargetControl_Disposed
        If TypeOf TargetControl Is ListControl Then AddHandler DirectCast(TargetControl, ListControl).SelectedValueChanged, AddressOf TargetControl_ValueChanged
        If TypeOf TargetControl Is CheckBox Then AddHandler DirectCast(TargetControl, CheckBox).CheckedChanged, AddressOf TargetControl_ValueChanged
        If TypeOf TargetControl Is RadioButton Then AddHandler DirectCast(TargetControl, RadioButton).CheckedChanged, AddressOf TargetControl_ValueChanged
        If TypeOf TargetControl Is DateTimePicker Then AddHandler DirectCast(TargetControl, DateTimePicker).ValueChanged, AddressOf TargetControl_ValueChanged
        If TypeOf TargetControl Is NumericUpDown Then AddHandler DirectCast(TargetControl, NumericUpDown).ValueChanged, AddressOf TargetControl_ValueChanged
        If TypeOf TargetControl Is CheckedListBox Then AddHandler DirectCast(TargetControl, CheckedListBox).ItemCheck, AddressOf TargetControl_ItemCheck
        Settings.IsRegistered = True
    End Sub
    Private Sub RemoveControl(TargetControl As Control)
        Dim Settings As ControlValidationSettings = Nothing
        If Not _Settings.TryGetValue(TargetControl, Settings) Then Return
        If Settings.IsRegistered Then
            RemoveHandler TargetControl.Validating, AddressOf TargetControl_Validating
            RemoveHandler TargetControl.TextChanged, AddressOf TargetControl_ValueChanged
            RemoveHandler TargetControl.Disposed, AddressOf TargetControl_Disposed
            If TypeOf TargetControl Is ListControl Then RemoveHandler DirectCast(TargetControl, ListControl).SelectedValueChanged, AddressOf TargetControl_ValueChanged
            If TypeOf TargetControl Is CheckBox Then RemoveHandler DirectCast(TargetControl, CheckBox).CheckedChanged, AddressOf TargetControl_ValueChanged
            If TypeOf TargetControl Is RadioButton Then RemoveHandler DirectCast(TargetControl, RadioButton).CheckedChanged, AddressOf TargetControl_ValueChanged
            If TypeOf TargetControl Is DateTimePicker Then RemoveHandler DirectCast(TargetControl, DateTimePicker).ValueChanged, AddressOf TargetControl_ValueChanged
            If TypeOf TargetControl Is NumericUpDown Then RemoveHandler DirectCast(TargetControl, NumericUpDown).ValueChanged, AddressOf TargetControl_ValueChanged
            If TypeOf TargetControl Is CheckedListBox Then RemoveHandler DirectCast(TargetControl, CheckedListBox).ItemCheck, AddressOf TargetControl_ItemCheck
        End If
        If Not TargetControl.IsDisposed Then SetError(TargetControl, String.Empty)
        _Settings.Remove(TargetControl)
    End Sub
    Private Sub TargetControl_Validating(Sender As Object, E As CancelEventArgs)
        If Not AutomaticValidation Then Return
        Dim TargetControl As Control = TryCast(Sender, Control)
        If TargetControl Is Nothing Then Return
        Dim Settings As ControlValidationSettings = GetSettings(TargetControl)
        If Settings Is Nothing OrElse Not ShouldValidateControl(TargetControl) Then Return
        Dim Result As ValidationResult = ValidateSingleControl(TargetControl, Settings)
        If CancelValidationOnError AndAlso Not Result.IsValid Then E.Cancel = True
    End Sub
    Private Sub TargetControl_ValueChanged(Sender As Object, E As EventArgs)
        If Not ClearErrorOnValueChanged Then Return
        Dim TargetControl As Control = TryCast(Sender, Control)
        If TargetControl IsNot Nothing AndAlso Not TargetControl.IsDisposed Then SetError(TargetControl, String.Empty)
    End Sub
    Private Sub TargetControl_ItemCheck(Sender As Object, E As ItemCheckEventArgs)
        TargetControl_ValueChanged(Sender, EventArgs.Empty)
    End Sub
    Private Sub TargetControl_Disposed(Sender As Object, E As EventArgs)
        Dim TargetControl As Control = TryCast(Sender, Control)
        If TargetControl IsNot Nothing Then RemoveControl(TargetControl)
    End Sub
    Private Shared Sub EnsureTargetControl(TargetControl As Control)
        ArgumentNullException.ThrowIfNull(TargetControl)
    End Sub
End Class
