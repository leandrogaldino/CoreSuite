Imports System.ComponentModel
Imports System.Globalization

''' <summary>
''' Provides a base implementation for culture-aware numeric input controls.
''' </summary>
<ToolboxItem(False)>
Public MustInherit Class NumericBoxBase
    Inherits TextBox
    Private Const NumericNumberStyles As NumberStyles = NumberStyles.AllowLeadingSign Or NumberStyles.AllowDecimalPoint Or NumberStyles.AllowThousands
    Private _DecimalPlaces As Integer = 2
    Private _NumericValue As Decimal
    Private _IncludeThousandSeparator As Boolean = True
    Private _NumericCulture As CultureInfo = CultureInfo.CurrentCulture
    Private _RoundingMode As MidpointRounding = MidpointRounding.ToEven
    Private _InternalTextChange As Boolean
    Private _TextInitialized As Boolean
    Private _LastAcceptedText As String = String.Empty
    ''' <summary>
    ''' Initializes a new instance of the <see cref="NumericBoxBase"/> class.
    ''' </summary>
    Protected Sub New()
        MyBase.Multiline = False
        TextAlign = HorizontalAlignment.Right
    End Sub
    ''' <summary>
    ''' Gets or sets the number of decimal places displayed and accepted.
    ''' </summary>
    ''' <exception cref="ArgumentOutOfRangeException">
    ''' The assigned value is less than zero or greater than 28.
    ''' </exception>
    <Category("NumericBox")>
    <DefaultValue(2)>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the number of decimal places displayed and accepted.")>
    Public Overridable Property DecimalPlaces As Integer
        Get
            Return _DecimalPlaces
        End Get
        Set(value As Integer)
            If value < 0 OrElse value > 28 Then
                Throw New ArgumentOutOfRangeException(
                    NameOf(value),
                    value,
                    "DecimalPlaces must be between 0 and 28.")
            End If
            If _DecimalPlaces = value Then Return
            _DecimalPlaces = value
            SetNumericValueInternal(
                _NumericValue,
                UpdateText:=True)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether group separators are displayed when the control
    ''' does not have input focus.
    ''' </summary>
    <Category("NumericBox")>
    <DefaultValue(True)>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets whether group separators are displayed when the control does not have input focus.")>
    Public Overridable Property IncludeThousandSeparator As Boolean
        Get
            Return _IncludeThousandSeparator
        End Get
        Set(value As Boolean)
            If _IncludeThousandSeparator = value Then Return
            _IncludeThousandSeparator = value
            RefreshNumericText()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the culture used to parse and format numeric values.
    ''' </summary>
    <Category("NumericBox")>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the culture used to parse and format numeric values.")>
    Public Property NumericCulture As CultureInfo
        Get
            Return _NumericCulture
        End Get
        Set(value As CultureInfo)
            Dim NewCulture As CultureInfo = If(value, CultureInfo.CurrentCulture)
            If _NumericCulture.Equals(NewCulture) Then Return
            _NumericCulture = NewCulture
            RefreshNumericText()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the rounding strategy used when a numeric value contains
    ''' more fractional digits than allowed by <see cref="DecimalPlaces"/>.
    ''' </summary>
    <Category("NumericBox")>
    <DefaultValue(MidpointRounding.ToEven)>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the rounding strategy used for numeric values.")>
    Public Overridable Property RoundingMode As MidpointRounding
        Get
            Return _RoundingMode
        End Get
        Set(value As MidpointRounding)
            If Not [Enum].IsDefined(
                GetType(MidpointRounding),
                value) Then
                Throw New InvalidEnumArgumentException(
                    NameOf(value),
                    CInt(value),
                    GetType(MidpointRounding))
            End If
            If _RoundingMode = value Then Return
            _RoundingMode = value
            SetNumericValueInternal(
                _NumericValue,
                UpdateText:=True)
        End Set
    End Property
    ''' <summary>
    ''' Gets the current numeric value stored by the control.
    ''' </summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    Protected ReadOnly Property NumericValue As Decimal
        Get
            Return _NumericValue
        End Get
    End Property
    ''' <summary>
    ''' Multiline input is not supported by numeric input controls.
    ''' </summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    <DefaultValue(False)>
    Public Overrides Property Multiline As Boolean
        Get
            Return False
        End Get
        Set(value As Boolean)
            MyBase.Multiline = False
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the text associated with the control.
    ''' </summary>
    ''' <exception cref="FormatException">
    ''' The assigned text cannot be converted to a numeric value using the
    ''' configured culture.
    ''' </exception>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overrides Property Text As String
        Get
            Return MyBase.Text
        End Get
        Set(value As String)
            If _NumericCulture Is Nothing Then
                MyBase.Text = If(value, String.Empty)
                Return
            End If
            If String.IsNullOrWhiteSpace(value) Then
                SetNumericValueInternal(0D, UpdateText:=False)
                If Focused Then
                    SetTextInternal(String.Empty)
                Else
                    SetTextInternal(FormatNumericValueForDisplay(_NumericValue))
                End If
                Return
            End If
            Dim ParsedValue As Decimal
            If Not TryParseNumericValue(value, ParsedValue) Then
                Throw New FormatException($"The value '{value}' is not a valid numeric value " & $"for culture '{_NumericCulture.Name}'.")
            End If
            SetNumericValueInternal(ParsedValue, UpdateText:=True)
        End Set
    End Property
    ''' <summary>
    ''' Sets the numeric value stored by the control.
    ''' </summary>
    ''' <param name="value">The value to assign.</param>
    Protected Sub SetNumericValue(value As Decimal)
        SetNumericValueInternal(value, UpdateText:=True)
    End Sub
    ''' <summary>
    ''' Initializes the displayed text after the derived control has been
    ''' constructed.
    ''' </summary>
    Protected Sub InitializeNumericText()
        RefreshNumericText()
    End Sub
    ''' <summary>
    ''' Refreshes the displayed text using the current focus state.
    ''' </summary>
    Protected Sub RefreshNumericText()
        If Not _TextInitialized AndAlso Not IsHandleCreated Then Return
        If Focused Then
            If _NumericValue = 0D Then
                SetTextInternal(String.Empty)
            Else
                SetTextInternal(FormatNumericValueForEditing(_NumericValue))
            End If
        Else
            SetTextInternal(FormatNumericValueForDisplay(_NumericValue))
        End If
    End Sub
    ''' <summary>
    ''' Formats a numeric value for editing while the control has input focus.
    ''' </summary>
    ''' <param name="Value">The value to format.</param>
    ''' <returns>The formatted numeric text.</returns>
    Protected Overridable Function FormatNumericValueForEditing(Value As Decimal) As String
        Dim FormatSpecifier As String = "F" & _DecimalPlaces.ToString(CultureInfo.InvariantCulture)
        Return Value.ToString(FormatSpecifier, _NumericCulture)
    End Function
    ''' <summary>
    ''' Formats a numeric value for display while the control does not have
    ''' input focus.
    ''' </summary>
    ''' <param name="value">The value to format.</param>
    ''' <returns>The formatted display text.</returns>
    Protected MustOverride Function FormatNumericValueForDisplay(value As Decimal) As String
    ''' <summary>
    ''' Raises the value-changed event implemented by the derived control.
    ''' </summary>
    ''' <param name="e">The event data.</param>
    Protected Overridable Sub OnNumericValueChanged(e As EventArgs)
    End Sub
    ''' <summary>
    ''' Rounds and normalizes a value before storing it.
    ''' </summary>
    ''' <param name="value">The value to normalize.</param>
    ''' <returns>The normalized numeric value.</returns>
    Protected Overridable Function NormalizeNumericValue(value As Decimal) As Decimal
        Return Decimal.Round(value, _DecimalPlaces, _RoundingMode)
    End Function
    ''' <summary>
    ''' Attempts to parse text using the configured numeric culture.
    ''' </summary>
    Protected Overridable Function TryParseNumericValue(Value As String, ByRef ParsedValue As Decimal) As Boolean
        If String.IsNullOrWhiteSpace(Value) Then
            ParsedValue = 0D
            Return False
        End If
        Dim NormalizedValue As String = Value.Trim()
        Dim NumberFormat As NumberFormatInfo = _NumericCulture.NumberFormat
        Dim DecimalSeparator As String = NumberFormat.NumberDecimalSeparator
        Dim NegativeSign As String = NumberFormat.NegativeSign
        Dim PositiveSign As String = NumberFormat.PositiveSign
        If NormalizedValue.StartsWith(NegativeSign & DecimalSeparator, StringComparison.Ordinal) Then
            NormalizedValue = String.Concat(NegativeSign, "0", NormalizedValue.AsSpan(NegativeSign.Length))
        ElseIf NormalizedValue.StartsWith(PositiveSign & DecimalSeparator, StringComparison.Ordinal) Then
            NormalizedValue = String.Concat(PositiveSign, "0", NormalizedValue.AsSpan(PositiveSign.Length))
        ElseIf NormalizedValue.StartsWith(DecimalSeparator, StringComparison.Ordinal) Then
            NormalizedValue = "0" & NormalizedValue
        End If
        If NormalizedValue.EndsWith(DecimalSeparator, StringComparison.Ordinal) Then
            NormalizedValue &= "0"
        End If
        Return Decimal.TryParse(NormalizedValue, NumericNumberStyles, _NumericCulture, ParsedValue)
    End Function
    ''' <summary>
    ''' Determines whether the supplied text represents a complete numeric
    ''' value or a valid intermediate numeric input.
    ''' </summary>
    Protected Overridable Function IsPotentialNumericText(Value As String) As Boolean
        If String.IsNullOrEmpty(Value) Then Return True
        Dim NumberFormat As NumberFormatInfo = _NumericCulture.NumberFormat
        Dim DecimalSeparator As String = NumberFormat.NumberDecimalSeparator
        Dim NegativeSign As String = NumberFormat.NegativeSign
        Dim PositiveSign As String = NumberFormat.PositiveSign
        If Value = NegativeSign OrElse Value = PositiveSign Then
            Return True
        End If
        If Not HasValidFractionalLength(Value, DecimalSeparator) Then
            Return False
        End If
        Dim ParsedValue As Decimal
        Return TryParseNumericValue(Value, ParsedValue)
    End Function
    ''' <summary>
    ''' Initializes the displayed numeric text when the control handle is
    ''' created.
    ''' </summary>
    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        If Not _TextInitialized Then
            SetTextInternal(FormatNumericValueForDisplay(_NumericValue))
        End If
    End Sub
    ''' <summary>
    ''' Processes changes to the displayed text.
    ''' </summary>
    Protected Overrides Sub OnTextChanged(e As EventArgs)
        If _NumericCulture Is Nothing OrElse
           _InternalTextChange Then
            MyBase.OnTextChanged(e)
            Return
        End If
        Dim CurrentText As String = MyBase.Text
        If Not IsPotentialNumericText(CurrentText) Then
            RestoreLastAcceptedText()
            Return
        End If
        _LastAcceptedText = CurrentText
        If String.IsNullOrEmpty(CurrentText) Then
            SetNumericValueInternal(0D, UpdateText:=False)
        Else
            Dim ParsedValue As Decimal
            If TryParseNumericValue(CurrentText, ParsedValue) Then
                SetNumericValueInternal(ParsedValue, UpdateText:=False)
            End If
        End If
        MyBase.OnTextChanged(e)
    End Sub
    ''' <summary>
    ''' Prepares the numeric value for editing when the control receives focus.
    ''' </summary>
    Protected Overrides Sub OnEnter(e As EventArgs)
        MyBase.OnEnter(e)
        If _NumericValue = 0D Then
            SetTextInternal(String.Empty)
        Else
            SetTextInternal(FormatNumericValueForEditing(_NumericValue))
        End If
        SelectionStart = TextLength
        SelectionLength = 0
    End Sub
    ''' <summary>
    ''' Commits and formats the numeric value when the control loses focus.
    ''' </summary>
    Protected Overrides Sub OnLostFocus(e As EventArgs)
        CommitCurrentText()
        SetTextInternal(FormatNumericValueForDisplay(_NumericValue))
        MyBase.OnLostFocus(e)
    End Sub
    ''' <summary>
    ''' Prevents characters that would produce an invalid numeric value.
    ''' </summary>
    Protected Overrides Sub OnKeyPress(e As KeyPressEventArgs)
        MyBase.OnKeyPress(e)
        If e.Handled OrElse
           Char.IsControl(e.KeyChar) Then
            Return
        End If
        Dim CandidateText As String = GetTextAfterKeyPress(e.KeyChar)
        If Not IsPotentialNumericText(CandidateText) Then
            e.Handled = True
        End If
    End Sub
    ''' <summary>
    ''' Determines whether the fractional portion does not exceed the
    ''' configured number of decimal places.
    ''' </summary>
    Private Function HasValidFractionalLength(Value As String, DecimalSeparator As String) As Boolean
        Dim SeparatorIndex As Integer = Value.IndexOf(DecimalSeparator, StringComparison.Ordinal)
        If SeparatorIndex < 0 Then Return True
        If _DecimalPlaces = 0 Then Return False
        Dim NextSeparatorIndex As Integer = Value.IndexOf(DecimalSeparator, SeparatorIndex + DecimalSeparator.Length, StringComparison.Ordinal)
        If NextSeparatorIndex >= 0 Then Return False
        Dim FractionalPart As String = Value.Substring(SeparatorIndex + DecimalSeparator.Length)
        If FractionalPart.Length > _DecimalPlaces Then
            Return False
        End If
        For Each Character As Char In FractionalPart
            If Not Char.IsDigit(Character) Then
                Return False
            End If
        Next
        Return True
    End Function
    ''' <summary>
    ''' Creates the text that would result from inserting a character at the
    ''' current selection.
    ''' </summary>
    Private Function GetTextAfterKeyPress(KeyChar As Char) As String
        Dim CurrentText As String = MyBase.Text
        Dim CurrentSelectionStart As Integer = SelectionStart
        Dim CurrentSelectionLength As Integer = SelectionLength
        Return CurrentText.Remove(CurrentSelectionStart, CurrentSelectionLength).Insert(CurrentSelectionStart, KeyChar.ToString())
    End Function
    ''' <summary>
    ''' Commits the current text to the stored numeric value.
    ''' </summary>
    Private Sub CommitCurrentText()
        Dim ParsedValue As Decimal
        If TryParseNumericValue(MyBase.Text, ParsedValue) Then
            SetNumericValueInternal(ParsedValue, UpdateText:=False)
        Else
            SetNumericValueInternal(0D, UpdateText:=False)
        End If
    End Sub
    ''' <summary>
    ''' Sets the internal numeric value and optionally updates the displayed
    ''' text.
    ''' </summary>
    Private Sub SetNumericValueInternal(Value As Decimal, UpdateText As Boolean)
        Dim NormalizedValue As Decimal = NormalizeNumericValue(Value)
        Dim ValueChanged As Boolean = _NumericValue <> NormalizedValue
        _NumericValue = NormalizedValue
        If UpdateText Then
            If Focused AndAlso
               _NumericValue = 0D Then
                SetTextInternal(String.Empty)
            ElseIf Focused Then
                SetTextInternal(FormatNumericValueForEditing(_NumericValue))
            Else
                SetTextInternal(FormatNumericValueForDisplay(_NumericValue))
            End If
        End If
        If ValueChanged Then
            OnNumericValueChanged(EventArgs.Empty)
        End If
    End Sub
    ''' <summary>
    ''' Changes the displayed text without attempting to parse it again.
    ''' </summary>
    Private Sub SetTextInternal(Value As String)
        Dim NormalizedText As String = If(Value, String.Empty)
        _TextInitialized = True
        _LastAcceptedText = NormalizedText
        _InternalTextChange = True
        Try
            MyBase.Text = NormalizedText
        Finally
            _InternalTextChange = False
        End Try
    End Sub
    ''' <summary>
    ''' Restores the last accepted text after an invalid edit or paste.
    ''' </summary>
    Private Sub RestoreLastAcceptedText()
        Dim CaretPosition As Integer = Math.Min(SelectionStart, _LastAcceptedText.Length)
        SetTextInternal(_LastAcceptedText)
        SelectionStart = CaretPosition
        SelectionLength = 0
    End Sub
End Class