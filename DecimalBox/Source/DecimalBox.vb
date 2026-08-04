Imports System.ComponentModel
Imports System.Globalization

''' <summary>
''' Represents a text box specialized in culture-aware decimal value input.
''' </summary>
<DefaultEvent("DecimalValueChanged")>
<DefaultProperty("DecimalValue")>
<DefaultBindingProperty("DecimalValue")>
<ToolboxItem(True)>
<Designer(GetType(DecimalBoxControlDesigner))>
<ToolboxItemFilter("CoreSuite")>
Public Class DecimalBox
    Inherits NumericBoxBase
    ''' <summary>
    ''' Occurs when the decimal value changes.
    ''' </summary>
    <Category("DecimalBox")>
    <Description("Occurs when the decimal value changes.")>
    Public Event DecimalValueChanged As EventHandler
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DecimalBox"/> class.
    ''' </summary>
    Public Sub New()
        InitializeNumericText()
    End Sub
    ''' <summary>
    ''' Gets or sets the decimal value represented by the control.
    ''' </summary>
    <Category("DecimalBox")>
    <Bindable(True)>
    <DefaultValue(GetType(Decimal), "0")>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the decimal value represented by the control.")>
    Public Property DecimalValue As Decimal
        Get
            Return NumericValue
        End Get
        Set(value As Decimal)
            SetNumericValue(value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the culture used to parse and format decimal values.
    ''' </summary>
    <Category("DecimalBox")>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the culture used to parse and format decimal values.")>
    Public Property DecimalCulture As CultureInfo
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
    <Category("DecimalBox")>
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
    <Category("DecimalBox")>
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
    <Category("DecimalBox")>
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
    ''' Formats the decimal value for display while the control does not have
    ''' input focus.
    ''' </summary>
    ''' <param name="value">The decimal value to format.</param>
    ''' <returns>The formatted decimal text.</returns>
    Protected Overrides Function FormatNumericValueForDisplay(
        value As Decimal
    ) As String
        Dim FormatSpecifier As String =
            If(
                IncludeThousandSeparator,
                "N",
                "F") &
            DecimalPlaces.ToString(
                CultureInfo.InvariantCulture)
        Return value.ToString(
            FormatSpecifier,
            DecimalCulture)
    End Function
    ''' <summary>
    ''' Raises the <see cref="DecimalValueChanged"/> event.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overrides Sub OnNumericValueChanged(e As EventArgs)
        MyBase.OnNumericValueChanged(e)
        OnDecimalValueChanged(e)
    End Sub
    ''' <summary>
    ''' Raises the <see cref="DecimalValueChanged"/> event.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overridable Sub OnDecimalValueChanged(e As EventArgs)
        RaiseEvent DecimalValueChanged(Me, e)
    End Sub
End Class