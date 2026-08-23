''' <summary>
''' Stores validation settings associated with a control managed by the <see cref="ValidationProvider"/>.
''' </summary>
''' <remarks>
''' This class contains the configured validation rules, visual feedback settings, cached regular expression,
''' and internal registration state for a single control.
''' </remarks>
Friend Class ControlValidationSettings
    ''' <summary>
    ''' Gets or sets a value indicating whether the control must contain a value.
    ''' </summary>
    Public Property Required As Boolean
    ''' <summary>
    ''' Gets or sets a value indicating whether the control participates in application-defined validation
    ''' even when no built-in validation rule is configured.
    ''' </summary>
    Public Property CustomValidationEnabled As Boolean
    ''' <summary>
    ''' Gets or sets the validation group associated with the control.
    ''' </summary>
    Public Property ValidationGroup As String = String.Empty
    ''' <summary>
    ''' Gets or sets the custom validation message displayed when validation fails.
    ''' </summary>
    Public Property ValidationMessage As String = String.Empty
    ''' <summary>
    ''' Gets or sets the friendly display name used when generating validation messages for the control.
    ''' </summary>
    Public Property ValidationDisplayName As String = String.Empty
    ''' <summary>
    ''' Gets or sets the minimum number of characters required for the represented value.
    ''' </summary>
    Public Property MinimumLength As Integer
    ''' <summary>
    ''' Gets or sets the maximum number of characters allowed for the represented value.
    ''' </summary>
    Public Property MaximumLength As Integer
    ''' <summary>
    ''' Gets or sets the regular expression pattern used to validate the represented text value.
    ''' </summary>
    Public Property RegularExpression As String = String.Empty
    ''' <summary>
    ''' Gets or sets the compiled regular expression used internally to validate the represented text value.
    ''' </summary>
    Public Property CompiledRegularExpression As Text.RegularExpressions.Regex
    ''' <summary>
    ''' Gets or sets the control whose represented value must match the value of the current control.
    ''' </summary>
    Public Property CompareWith As Control
    ''' <summary>
    ''' Gets or sets the property name or nested property path used to obtain the value represented by the control.
    ''' </summary>
    Public Property ValuePropertyName As String = String.Empty
    ''' <summary>
    ''' Gets or sets the control on which the validation error indicator is displayed.
    ''' </summary>
    ''' <remarks>
    ''' When this property is <see langword="Nothing"/>, the validated control itself is used to display the error indicator.
    ''' </remarks>
    Public Property ValidationIndicatorControl As Control
    ''' <summary>
    ''' Gets or sets the alignment of the validation error icon relative to the indicator control.
    ''' </summary>
    Public Property ValidationIndicatorAlignment As ErrorIconAlignment = ErrorIconAlignment.MiddleRight
    ''' <summary>
    ''' Gets or sets the amount of extra space, in pixels, between the validation error icon and the indicator control.
    ''' </summary>
    Public Property ValidationIndicatorPadding As Integer
    ''' <summary>
    ''' Gets or sets a value indicating whether the control's validation-related events are currently registered.
    ''' </summary>
    Public Property IsRegistered As Boolean
    ''' <summary>
    ''' Gets a value indicating whether all validation settings are currently using their default values.
    ''' </summary>
    Public ReadOnly Property IsDefault As Boolean
        Get
            Return Not Required AndAlso Not CustomValidationEnabled AndAlso String.IsNullOrEmpty(ValidationGroup) AndAlso String.IsNullOrEmpty(ValidationMessage) AndAlso String.IsNullOrEmpty(ValidationDisplayName) AndAlso MinimumLength = 0 AndAlso MaximumLength = 0 AndAlso String.IsNullOrEmpty(RegularExpression) AndAlso CompareWith Is Nothing AndAlso String.IsNullOrEmpty(ValuePropertyName) AndAlso ValidationIndicatorControl Is Nothing AndAlso ValidationIndicatorAlignment = ErrorIconAlignment.MiddleRight AndAlso ValidationIndicatorPadding = 0
        End Get
    End Property
End Class