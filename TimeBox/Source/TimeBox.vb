Imports System.ComponentModel
Imports System.Drawing.Drawing2D
Imports System.Globalization
Imports System.Text.RegularExpressions
''' <summary>
''' Represents a culture-aware time input control with an integrated time selector.
''' </summary>
<DefaultEvent("TimeValueChanged")>
<DefaultProperty("Time")>
<DefaultBindingProperty("Time")>
<ToolboxItem(True)>
<Designer(GetType(TimeBoxControlDesigner))>
Public Class TimeBox
    Inherits DateTimeBoxBase
    Private WithEvents ControlContainer As New ControlContainer
    Private WithEvents TimePicker As New DateTimePicker
    Private WithEvents DropDownButton As New PictureBox
    Private _TimeCulture As CultureInfo = CultureInfo.CurrentCulture
    Private _ButtonImage As Image
    Private _ShowSeconds As Boolean
    Private _InternalPickerChange As Boolean
    Private _PickerValueChanged As Boolean
    ''' <summary>
    ''' Occurs when the represented time changes.
    ''' </summary>
    <Category("TimeBox")>
    <Description("Occurs when the represented time changes.")>
    Public Event TimeValueChanged As EventHandler
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TimeBox"/> class.
    ''' </summary>
    Public Sub New()
        MinimumSize = New Size(60, 0)
        TimePicker.Format = DateTimePickerFormat.Custom
        TimePicker.ShowUpDown = True
        TimePicker.Width = 130
        TimePicker.Visible = False
        DropDownButton.Size = New Size(25, ClientSize.Height + 2)
        DropDownButton.Location = New Point(ClientSize.Width - DropDownButton.Width + 1, -1)
        DropDownButton.Cursor = Cursors.Default
        DropDownButton.TabStop = False
        DropDownButton.BackgroundImageLayout = ImageLayout.Center
        DropDownButton.BackColor = BackColor
        Controls.Add(DropDownButton)
        Controls.Add(TimePicker)
        ControlContainer.HostControl = DropDownButton
        ControlContainer.HostedControl = TimePicker
        InitializeValueText()
    End Sub
    ''' <summary>
    ''' Gets or sets the time represented by the control.
    ''' </summary>
    ''' <exception cref="ArgumentOutOfRangeException">
    ''' The assigned value is less than zero or equal to or greater than one day.
    ''' </exception>
    <Category("TimeBox")>
    <Bindable(True)>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the time represented by the control.")>
    Public Property Time As TimeSpan
        Get
            If CurrentValue.HasValue Then
                Return CurrentValue.Value.TimeOfDay
            End If
            Return TimeSpan.Zero
        End Get
        Set(value As TimeSpan)
            If value < TimeSpan.Zero OrElse
               value >= TimeSpan.FromDays(1) Then
                Throw New ArgumentOutOfRangeException(
                    NameOf(value),
                    value,
                    "Time must be greater than or equal to 00:00:00 " &
                    "and less than 24:00:00.")
            End If
            SetTemporalValue(
                DateTime.MinValue.Date.Add(value))
        End Set
    End Property
    ''' <summary>
    ''' Indicates whether the control contains a valid time.
    ''' </summary>
    <Browsable(False)>
    Public ReadOnly Property HasTime As Boolean
        Get
            Return CurrentValue.HasValue
        End Get
    End Property
    ''' <summary>
    ''' Clears the time represented by the control.
    ''' </summary>
    Public Sub ClearTime()
        SetTemporalValue(Nothing)
    End Sub
    ''' <summary>
    ''' Gets or sets the culture used to parse and format time values.
    ''' </summary>
    <Category("TimeBox")>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the culture used to parse and format time values.")>
    Public Property TimeCulture As CultureInfo
        Get
            Return _TimeCulture
        End Get
        Set(value As CultureInfo)
            Dim NewCulture As CultureInfo = If(value, CultureInfo.CurrentCulture)
            If _TimeCulture.Equals(NewCulture) Then Return
            _TimeCulture = NewCulture
            ValueCulture = _TimeCulture
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether seconds are displayed and accepted.
    ''' </summary>
    <Category("TimeBox")>
    <DefaultValue(False)>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets whether seconds are displayed and accepted.")>
    Public Property ShowSeconds As Boolean
        Get
            Return _ShowSeconds
        End Get
        Set(value As Boolean)
            If _ShowSeconds = value Then Return
            _ShowSeconds = value
            RefreshValueFormat()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the image displayed on the time selector button.
    ''' </summary>
    <Category("TimeBox")>
    <DefaultValue(GetType(Image), Nothing)>
    <Description("Gets or sets the image displayed on the time selector button.")>
    Public Property ButtonImage As Image
        Get
            Return _ButtonImage
        End Get
        Set(value As Image)
            If ReferenceEquals(_ButtonImage, value) Then Return
            _ButtonImage = value
            DropDownButton.BackgroundImage = value
            DropDownButton.Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets the normalized time format used by the control.
    ''' </summary>
    Protected Overrides ReadOnly Property ValueFormat As String
        Get
            Dim Format As String
            If ShowSeconds Then
                Format = TimeCulture.DateTimeFormat.LongTimePattern
            Else
                Format = TimeCulture.DateTimeFormat.ShortTimePattern
            End If
            Return NormalizeTimeFormat(Format)
        End Get
    End Property
    ''' <summary>
    ''' Converts a time format into a masked text box mask.
    ''' </summary>
    Protected Overrides Function CreateMask(Format As String) As String
        Return Format.
            Replace("HH", "00").
            Replace("hh", "00").
            Replace("mm", "00").
            Replace("ss", "00").
            Replace("tt", "LL")
    End Function
    ''' <summary>
    ''' Removes the date component from a stored time value.
    ''' </summary>
    Protected Overrides Function NormalizeValue(Value As DateTime) As DateTime
        Return DateTime.MinValue.Date.Add(Value.TimeOfDay)
    End Function
    ''' <summary>
    ''' Synchronizes the time selector format.
    ''' </summary>
    Protected Overrides Sub OnValueCultureChanged()
        MyBase.OnValueCultureChanged()
        If TimePicker IsNot Nothing Then
            TimePicker.CustomFormat = ValueFormat
        End If
    End Sub
    ''' <summary>
    ''' Raises the <see cref="TimeValueChanged"/> event.
    ''' </summary>
    Protected Overrides Sub OnTemporalValueChanged(e As EventArgs)
        MyBase.OnTemporalValueChanged(e)
        OnTimeValueChanged(e)
    End Sub
    ''' <summary>
    ''' Raises the <see cref="TimeValueChanged"/> event.
    ''' </summary>
    Protected Overridable Sub OnTimeValueChanged(e As EventArgs)
        RaiseEvent TimeValueChanged(Me, e)
    End Sub
    ''' <summary>
    ''' Configures the right text margin after the handle is created.
    ''' </summary>
    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        SetRightTextMargin(DropDownButton.Width)
    End Sub
    ''' <summary>
    ''' Repositions the selector button when the control size changes.
    ''' </summary>
    Protected Overrides Sub OnSizeChanged(e As EventArgs)
        MyBase.OnSizeChanged(e)
        If DropDownButton Is Nothing Then Return
        DropDownButton.Size = New Size(25, ClientSize.Height + 2)
        DropDownButton.Location = New Point(ClientSize.Width - DropDownButton.Width + 1, -1)
        SetRightTextMargin(DropDownButton.Width)
    End Sub
    ''' <summary>
    ''' Synchronizes the selector button background color.
    ''' </summary>
    Protected Overrides Sub OnBackColorChanged(e As EventArgs)
        MyBase.OnBackColorChanged(e)
        If DropDownButton IsNot Nothing Then
            DropDownButton.BackColor = BackColor
        End If
    End Sub
    ''' <summary>
    ''' Opens the time selector when ENTER is pressed.
    ''' </summary>
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        If e.KeyCode <> Keys.Enter Then Return
        ShowTimeSelector()
        e.SuppressKeyPress = True
    End Sub
    ''' <summary>
    ''' Normalizes single-character time format elements.
    ''' </summary>
    Private Shared Function NormalizeTimeFormat(Format As String) As String
        Dim Result As String = Format
        Result = Regex.Replace(Result, "(?<!H)H(?!H)", "HH")
        Result = Regex.Replace(Result, "(?<!h)h(?!h)", "hh")
        Result = Regex.Replace(Result, "(?<!m)m(?!m)", "mm")
        Result = Regex.Replace(Result, "(?<!s)s(?!s)", "ss")
        Result = Regex.Replace(Result, "(?<!t)t(?!t)", "tt")
        Return Result
    End Function
    ''' <summary>
    ''' Displays the time selector dropdown.
    ''' </summary>
    Private Sub ShowTimeSelector()
        TimePicker.Visible = True
        ControlContainer.ShowDropDown()
    End Sub
    ''' <summary>
    ''' Opens the time selector when the internal button is clicked.
    ''' </summary>
    Private Sub DropDownButton_Click(sender As Object, e As EventArgs) Handles DropDownButton.Click
        ShowTimeSelector()
    End Sub
    ''' <summary>
    ''' Draws a default clock image when no button image is assigned.
    ''' </summary>
    Private Sub DropDownButton_Paint(sender As Object, e As PaintEventArgs) Handles DropDownButton.Paint
        If ButtonImage IsNot Nothing Then Return
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Dim Size As Integer = Math.Min(DropDownButton.ClientSize.Width, DropDownButton.ClientSize.Height) - 10
        If Size <= 0 Then Return
        Dim Left As Integer = (DropDownButton.ClientSize.Width - Size) \ 2
        Dim Top As Integer = (DropDownButton.ClientSize.Height - Size) \ 2
        Dim CenterX As Single = Left + Size / 2.0F
        Dim CenterY As Single = Top + Size / 2.0F
        Using ClockPen As New Pen(SystemColors.ControlDarkDark, 1.4F)
            e.Graphics.DrawEllipse(ClockPen, Left, Top, Size, Size)
            e.Graphics.DrawLine(ClockPen, CenterX, CenterY, CenterX, Top + Size * 0.27F)
            e.Graphics.DrawLine(ClockPen, CenterX, CenterY, Left + Size * 0.72F, Top + Size * 0.62F)
        End Using
    End Sub
    ''' <summary>
    ''' Configures the initial value displayed by the time selector.
    ''' </summary>
    Private Sub ControlContainer_Dropped(sender As Object) Handles ControlContainer.Dropped
        _InternalPickerChange = True
        Try
            TimePicker.CustomFormat = ValueFormat
            If HasTime Then
                TimePicker.Value = Today.Add(Time)
            Else
                TimePicker.Value = Now
            End If
        Finally
            _InternalPickerChange = False
        End Try
        _PickerValueChanged = False
        TimePicker.Focus()
    End Sub
    ''' <summary>
    ''' Records that the time selector value was changed by the user.
    ''' </summary>
    Private Sub TimePicker_ValueChanged(sender As Object, e As EventArgs) Handles TimePicker.ValueChanged
        If _InternalPickerChange Then Return
        _PickerValueChanged = True
    End Sub
    ''' <summary>
    ''' Applies the selected time after the selector closes.
    ''' </summary>
    Private Sub ControlContainer_Closed(sender As Object) Handles ControlContainer.Closed
        If _PickerValueChanged Then
            Time = TimePicker.Value.TimeOfDay
        End If
        _PickerValueChanged = False
        Me.Select()
    End Sub
End Class