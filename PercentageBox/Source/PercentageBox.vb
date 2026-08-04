Imports System.ComponentModel
Imports System.Globalization

''' <summary>
''' Represents a text box specialized in culture-aware percentage value input.
''' </summary>
''' <remarks>
''' The percentage value is represented in percentage points. For example,
''' a value of 25 represents 25%, while its fractional value is 0.25.
''' </remarks>
<DefaultEvent("PercentageValueChanged")>
<DefaultProperty("PercentageValue")>
<DefaultBindingProperty("PercentageValue")>
<ToolboxItem(True)>
<Designer(GetType(PercentageBoxControlDesigner))>
Public Class PercentageBox
    Inherits NumericBoxBase
    ''' <summary>
    ''' Occurs when the percentage value changes.
    ''' </summary>
    <Category("PercentageBox")>
    <Description("Occurs when the percentage value changes.")>
    Public Event PercentageValueChanged As EventHandler
    ''' <summary>
    ''' Initializes a new instance of the <see cref="PercentageBox"/> class.
    ''' </summary>
    Public Sub New()
        InitializeNumericText()
    End Sub
    ''' <summary>
    ''' Gets or sets the percentage value represented by the control.
    ''' </summary>
    ''' <remarks>
    ''' The value is expressed in percentage points. For example, assigning
    ''' 25 displays 25%.
    ''' </remarks>
    <Category("PercentageBox")>
    <Bindable(True)>
    <DefaultValue(GetType(Decimal), "0")>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the percentage value represented by the control.")>
    Public Property PercentageValue As Decimal
        Get
            Return NumericValue
        End Get
        Set(value As Decimal)
            SetNumericValue(value)
        End Set
    End Property
    ''' <summary>
    ''' Gets the percentage value represented as a fractional value.
    ''' </summary>
    ''' <remarks>
    ''' For example, a percentage value of 25 produces a fractional value
    ''' of 0.25.
    ''' </remarks>
    <Browsable(False)>
    <Description("Gets the percentage value represented as a fractional value.")>
    Public ReadOnly Property FractionalValue As Decimal
        Get
            Return NumericValue / 100D
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the culture used to parse and format percentage values.
    ''' </summary>
    <Category("PercentageBox")>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the culture used to parse and format percentage values.")>
    Public Property PercentageCulture As CultureInfo
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
    <Category("PercentageBox")>
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
    <Category("PercentageBox")>
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
    <Category("PercentageBox")>
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
    ''' Gets or sets the culture used to parse and format numeric values.
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
    ''' Formats the percentage value for display while the control does not
    ''' have input focus.
    ''' </summary>
    ''' <param name="value">The percentage value to format.</param>
    ''' <returns>The formatted percentage text.</returns>
    Protected Overrides Function FormatNumericValueForDisplay(
        value As Decimal
    ) As String
        Dim FormatSpecifier As String =
            "P" &
            DecimalPlaces.ToString(
                CultureInfo.InvariantCulture)
        Dim NumberFormat As NumberFormatInfo =
            CType(
                PercentageCulture.NumberFormat.Clone(),
                NumberFormatInfo)
        If Not IncludeThousandSeparator Then
            NumberFormat.PercentGroupSizes =
                New Integer() {0}
        End If
        Return (value / 100D).ToString(
            FormatSpecifier,
            NumberFormat)
    End Function
    ''' <summary>
    ''' Raises the <see cref="PercentageValueChanged"/> event.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnNumericValueChanged(e As EventArgs)
        MyBase.OnNumericValueChanged(e)
        OnPercentageValueChanged(e)
    End Sub
    ''' <summary>
    ''' Raises the <see cref="PercentageValueChanged"/> event.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overridable Sub OnPercentageValueChanged(e As EventArgs)
        RaiseEvent PercentageValueChanged(Me, e)
    End Sub
End Class