Imports System.ComponentModel
Imports System.Globalization
Imports System.Runtime.InteropServices
''' <summary>
''' Provides a base implementation for culture-aware date and time input controls.
''' </summary>
<ToolboxItem(False)>
Public MustInherit Class DateTimeBoxBase
    Inherits MaskedTextBox
    Private Const EmSetMargins As Integer = &HD3
    Private Const EcRightMargin As Integer = &H2
    Private _CurrentValue As DateTime?
    Private _ValueCulture As CultureInfo = CultureInfo.CurrentCulture
    Private _InternalTextChange As Boolean
    Private _Initialized As Boolean
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DateTimeBoxBase"/> class.
    ''' </summary>
    Protected Sub New()
        MyBase.Multiline = False
        MyBase.InsertKeyMode = InsertKeyMode.Overwrite
        MyBase.TextMaskFormat = MaskFormat.IncludeLiterals
        MyBase.CutCopyMaskFormat = MaskFormat.IncludeLiterals
    End Sub
    ''' <summary>
    ''' Gets the current value stored by the control.
    ''' </summary>
    Protected ReadOnly Property CurrentValue As DateTime?
        Get
            Return _CurrentValue
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the culture used internally by the control.
    ''' </summary>
    Protected Property ValueCulture As CultureInfo
        Get
            Return _ValueCulture
        End Get
        Set(value As CultureInfo)
            Dim NewCulture As CultureInfo = If(value, CultureInfo.CurrentCulture)
            If _ValueCulture.Equals(NewCulture) Then Return
            _ValueCulture = NewCulture
            If _Initialized Then
                ApplyCultureAndMask()
                RefreshValueText()
                OnValueCultureChanged()
            End If
        End Set
    End Property
    ''' <summary>
    ''' Gets the format used to parse and display the current value.
    ''' </summary>
    Protected MustOverride ReadOnly Property ValueFormat As String
    ''' <summary>
    ''' Gets the mask used by the control.
    ''' </summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows ReadOnly Property Mask As String
        Get
            Return MyBase.Mask
        End Get
    End Property
    ''' <summary>
    ''' Gets the culture internally used by the masked text box.
    ''' </summary>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Shadows ReadOnly Property Culture As CultureInfo
        Get
            Return MyBase.Culture
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the text associated with the control.
    ''' </summary>
    ''' <exception cref="FormatException">
    ''' The assigned text cannot be converted using the configured culture
    ''' and format.
    ''' </exception>
    <Browsable(False)>
    <EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overrides Property Text As String
        Get
            Return MyBase.Text
        End Get
        Set(value As String)
            If Not _Initialized Then
                MyBase.Text = If(value, String.Empty)
                Return
            End If
            If String.IsNullOrWhiteSpace(value) Then
                SetCurrentValue(Nothing, UpdateText:=True)
                Return
            End If
            Dim ParsedValue As DateTime
            If Not TryParseValue(value, ParsedValue) Then
                Throw New FormatException($"The value '{value}' is not valid for culture " & $"'{ValueCulture.Name}' and format '{ValueFormat}'.")
            End If
            SetCurrentValue(ParsedValue, UpdateText:=True)
        End Set
    End Property
    ''' <summary>
    ''' Multiline input is not supported by date and time input controls.
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
    ''' Initializes the value text after the derived control has been constructed.
    ''' </summary>
    Protected Sub InitializeValueText()
        _Initialized = True
        ApplyCultureAndMask()
        RefreshValueText()
        OnValueCultureChanged()
    End Sub
    ''' <summary>
    ''' Recreates the mask and refreshes the displayed text.
    ''' </summary>
    Protected Sub RefreshValueFormat()
        If Not _Initialized Then Return
        ApplyCultureAndMask()
        RefreshValueText()
        OnValueCultureChanged()
    End Sub
    ''' <summary>
    ''' Sets the value stored by the control.
    ''' </summary>
    Protected Sub SetTemporalValue(Value As DateTime?)
        SetCurrentValue(Value, UpdateText:=True)
    End Sub
    ''' <summary>
    ''' Formats the supplied value using the configured culture and format.
    ''' </summary>
    Protected Overridable Function FormatValue(Value As DateTime) As String
        Return Value.ToString(ValueFormat, ValueCulture)
    End Function
    ''' <summary>
    ''' Attempts to parse a value using the configured culture and format.
    ''' </summary>
    Protected Overridable Function TryParseValue(Value As String, ByRef ParsedValue As DateTime) As Boolean
        If String.IsNullOrWhiteSpace(Value) Then
            ParsedValue = DateTime.MinValue
            Return False
        End If
        Return DateTime.TryParseExact(Value.Trim(), ValueFormat, ValueCulture, DateTimeStyles.AllowWhiteSpaces, ParsedValue)
    End Function
    ''' <summary>
    ''' Normalizes a value before it is stored.
    ''' </summary>
    Protected Overridable Function NormalizeValue(Value As DateTime) As DateTime
        Return Value
    End Function
    ''' <summary>
    ''' Creates a mask for the supplied format.
    ''' </summary>
    Protected MustOverride Function CreateMask(Format As String) As String
    ''' <summary>
    ''' Raises the value-changed event implemented by the derived control.
    ''' </summary>
    Protected Overridable Sub OnTemporalValueChanged(e As EventArgs)
    End Sub
    ''' <summary>
    ''' Called after the configured culture or format changes.
    ''' </summary>
    Protected Overridable Sub OnValueCultureChanged()
    End Sub
    ''' <summary>
    ''' Reserves space on the right side of the input area.
    ''' </summary>
    Protected Sub SetRightTextMargin(Width As Integer)
        If Not IsHandleCreated Then Return
        Dim PackedMargins As Integer = Math.Max(0, Width) << 16
        SendMessage(Handle, EmSetMargins, EcRightMargin, CType(PackedMargins, IntPtr))
    End Sub
    ''' <summary>
    ''' Processes changes to the displayed text.
    ''' </summary>
    Protected Overrides Sub OnTextChanged(e As EventArgs)
        If Not _Initialized OrElse _InternalTextChange Then
            MyBase.OnTextChanged(e)
            Return
        End If
        Dim ParsedValue As DateTime
        If MaskCompleted AndAlso
           TryParseValue(MyBase.Text, ParsedValue) Then
            SetCurrentValue(ParsedValue, UpdateText:=False)
        Else
            SetCurrentValue(Nothing, UpdateText:=False)
        End If
        MyBase.OnTextChanged(e)
    End Sub
    ''' <summary>
    ''' Selects the complete input when the control receives focus.
    ''' </summary>
    Protected Overrides Sub OnEnter(e As EventArgs)
        MyBase.OnEnter(e)
        If Not IsHandleCreated Then Return
        BeginInvoke(
            CType(
                Sub()
                    SelectAll()
                End Sub,
                Action))
    End Sub
    ''' <summary>
    ''' Commits and normalizes the current value after validation.
    ''' </summary>
    Protected Overrides Sub OnValidated(e As EventArgs)
        CommitCurrentText()
        MyBase.OnValidated(e)
    End Sub
    ''' <summary>
    ''' Refreshes the displayed text using the stored value.
    ''' </summary>
    Private Sub RefreshValueText()
        If _CurrentValue.HasValue Then
            SetTextInternal(FormatValue(_CurrentValue.Value))
        Else
            SetTextInternal(String.Empty)
        End If
    End Sub
    ''' <summary>
    ''' Applies the configured culture and recreates the input mask.
    ''' </summary>
    Private Sub ApplyCultureAndMask()
        MyBase.Culture = _ValueCulture
        MyBase.Mask = CreateMask(ValueFormat)
    End Sub
    ''' <summary>
    ''' Commits the current text to the stored value.
    ''' </summary>
    Private Sub CommitCurrentText()
        Dim ParsedValue As DateTime
        If MaskCompleted AndAlso
           TryParseValue(MyBase.Text, ParsedValue) Then
            SetCurrentValue(ParsedValue, UpdateText:=True)
        Else
            SetCurrentValue(Nothing, UpdateText:=True)
        End If
    End Sub
    ''' <summary>
    ''' Sets the current value and optionally updates the displayed text.
    ''' </summary>
    Private Sub SetCurrentValue(Value As DateTime?, UpdateText As Boolean)
        Dim NormalizedValue As DateTime?
        If Value.HasValue Then
            NormalizedValue = NormalizeValue(Value.Value)
        Else
            NormalizedValue = Nothing
        End If
        Dim ValueChanged As Boolean = Not Object.Equals(_CurrentValue, NormalizedValue)
        _CurrentValue = NormalizedValue
        If UpdateText Then
            RefreshValueText()
        End If
        If ValueChanged Then
            OnTemporalValueChanged(EventArgs.Empty)
        End If
    End Sub
    ''' <summary>
    ''' Changes the displayed text without parsing it again.
    ''' </summary>
    Private Sub SetTextInternal(Value As String)
        _InternalTextChange = True
        Try
            MyBase.Text = If(Value, String.Empty)
        Finally
            _InternalTextChange = False
        End Try
    End Sub
    ''' <summary>
    ''' Sends a message to a Windows control.
    ''' </summary>
    <DllImport("User32.dll", SetLastError:=True)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As Integer, lParam As IntPtr) As IntPtr
    End Function
End Class