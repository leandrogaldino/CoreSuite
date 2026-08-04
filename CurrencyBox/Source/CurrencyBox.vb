Imports System.ComponentModel
Imports System.Globalization

''' <summary>
''' Represents a text box specialized in culture-aware currency value input.
''' </summary>
<DefaultEvent("CurrencyValueChanged")>
<DefaultProperty("CurrencyValue")>
<DefaultBindingProperty("CurrencyValue")>
<ToolboxItem(True)>
<Designer(GetType(CurrencyBoxControlDesigner))>
Public Class CurrencyBox
    Inherits NumericBoxBase
    ''' <summary>
    ''' Occurs when the currency value changes.
    ''' </summary>
    <Category("CurrencyBox")>
    <Description("Occurs when the currency value changes.")>
    Public Event CurrencyValueChanged As EventHandler
    ''' <summary>
    ''' Initializes a new instance of the <see cref="CurrencyBox"/> class.
    ''' </summary>
    Public Sub New()
        InitializeNumericText()
    End Sub
    ''' <summary>
    ''' Gets or sets the currency value represented by the control.
    ''' </summary>
    <Category("CurrencyBox")>
    <Bindable(True)>
    <DefaultValue(GetType(Decimal), "0")>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the currency value represented by the control.")>
    Public Property CurrencyValue As Decimal
        Get
            Return NumericValue
        End Get
        Set(value As Decimal)
            SetNumericValue(value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the culture used to parse and format currency values.
    ''' </summary>
    <Category("CurrencyBox")>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the culture used to parse and format currency values.")>
    Public Property CurrencyCulture As CultureInfo
        Get
            Return MyBase.NumericCulture
        End Get
        Set(value As CultureInfo)
            MyBase.NumericCulture = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the number of decimal places displayed and accepted.
    ''' </summary>
    <Category("CurrencyBox")>
    <DefaultValue(2)>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the number of decimal places displayed and accepted.")>
    Public Overrides Property DecimalPlaces As Integer
        Get
            Return MyBase.DecimalPlaces
        End Get
        Set(value As Integer)
            MyBase.DecimalPlaces = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether group separators are displayed when the control
    ''' does not have input focus.
    ''' </summary>
    <Category("CurrencyBox")>
    <DefaultValue(True)>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets whether group separators are displayed when the control does not have input focus.")>
    Public Overrides Property IncludeThousandSeparator As Boolean
        Get
            Return MyBase.IncludeThousandSeparator
        End Get
        Set(value As Boolean)
            MyBase.IncludeThousandSeparator = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the rounding strategy used for currency values.
    ''' </summary>
    <Category("CurrencyBox")>
    <DefaultValue(MidpointRounding.ToEven)>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the rounding strategy used for currency values.")>
    Public Overrides Property RoundingMode As MidpointRounding
        Get
            Return MyBase.RoundingMode
        End Get
        Set(value As MidpointRounding)
            MyBase.RoundingMode = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the culture used internally to parse and format numeric values.
    ''' </summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows Property NumericCulture As CultureInfo
        Get
            Return MyBase.NumericCulture
        End Get
        Set(value As CultureInfo)
            MyBase.NumericCulture = value
        End Set
    End Property
    ''' <summary>
    ''' Formats the currency value for display while the control does not have
    ''' input focus.
    ''' </summary>
    ''' <param name="value">The currency value to format.</param>
    ''' <returns>The formatted currency text.</returns>
    Protected Overrides Function FormatNumericValueForDisplay(
        value As Decimal
    ) As String
        Dim FormatSpecifier As String =
            "C" &
            DecimalPlaces.ToString(
                CultureInfo.InvariantCulture)
        Dim NumberFormat As NumberFormatInfo =
            DirectCast(
                CurrencyCulture.NumberFormat.Clone(),
                NumberFormatInfo)
        If Not IncludeThousandSeparator Then
            NumberFormat.CurrencyGroupSizes =
                New Integer() {0}
        End If
        Return value.ToString(
            FormatSpecifier,
            NumberFormat)
    End Function
    ''' <summary>
    ''' Attempts to parse text using the configured currency culture.
    ''' </summary>
    ''' <param name="value">The text to parse.</param>
    ''' <param name="parsedValue">
    ''' When this method returns, contains the parsed currency value.
    ''' </param>
    ''' <returns>
    ''' <see langword="True"/> when the text can be converted; otherwise,
    ''' <see langword="False"/>.
    ''' </returns>
    Protected Overrides Function TryParseNumericValue(
        value As String,
        ByRef parsedValue As Decimal
    ) As Boolean
        If MyBase.TryParseNumericValue(
            value,
            parsedValue) Then
            Return True
        End If
        If String.IsNullOrWhiteSpace(value) Then
            parsedValue = 0D
            Return False
        End If
        Return Decimal.TryParse(
            value.Trim(),
            NumberStyles.Currency,
            CurrencyCulture,
            parsedValue)
    End Function
    ''' <summary>
    ''' Raises the <see cref="CurrencyValueChanged"/> event.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnNumericValueChanged(e As EventArgs)
        MyBase.OnNumericValueChanged(e)
        OnCurrencyValueChanged(e)
    End Sub
    ''' <summary>
    ''' Raises the <see cref="CurrencyValueChanged"/> event.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overridable Sub OnCurrencyValueChanged(e As EventArgs)
        RaiseEvent CurrencyValueChanged(Me, e)
    End Sub
End Class