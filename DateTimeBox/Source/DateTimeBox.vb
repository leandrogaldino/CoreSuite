Imports System.ComponentModel
Imports System.Drawing.Drawing2D
Imports System.Globalization
Imports System.Text.RegularExpressions
''' <summary>
''' Represents a culture-aware date and time input control with an integrated dropdown editor.
''' </summary>
<DefaultEvent("DateTimeValueChanged")>
<DefaultProperty("DateTime")>
<DefaultBindingProperty("DateTime")>
<Designer(GetType(DateTimeBoxControlDesigner))>
Public Class DateTimeBox
    Inherits DateTimeBoxBase
    Private WithEvents ControlContainer As New ControlContainer
    Private WithEvents DropDownEditor As New DateTimeBoxDropDown
    Private WithEvents DropDownButton As New PictureBox
    Private _DateTimeCulture As CultureInfo = CultureInfo.CurrentCulture
    Private _ButtonImage As Image
    Private _ShowSeconds As Boolean
    Private _SelectionConfirmed As Boolean
    Private _TimeLabelText As String = "Time"
    Private _OKButtonText As String = "OK"
    Private _CancelButtonText As String = "Cancel"
    ''' <summary>
    ''' Occurs when the date and time represented by the control changes.
    ''' </summary>
    <Category("DateTimeBox")>
    <Description("Occurs when the date and time represented by the control changes.")>
    Public Event DateTimeValueChanged As EventHandler
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DateTimeBox"/> class.
    ''' </summary>
    Public Sub New()
        MinimumSize = New Size(140, 0)
        DropDownEditor.Visible = False
        DropDownButton.Size = New Size(25, ClientSize.Height + 2)
        DropDownButton.Location = New Point(ClientSize.Width - DropDownButton.Width + 1, -1)
        DropDownButton.Cursor = Cursors.Default
        DropDownButton.TabStop = False
        DropDownButton.BackgroundImageLayout = ImageLayout.Center
        DropDownButton.BackColor = BackColor
        Controls.Add(DropDownButton)
        Controls.Add(DropDownEditor)
        ControlContainer.HostControl = DropDownButton
        ControlContainer.HostedControl = DropDownEditor
        InitializeValueText()
    End Sub
    ''' <summary>
    ''' Gets or sets the date and time represented by the control.
    ''' </summary>
    ''' <remarks>Assigning <see cref="System.DateTime.MinValue"/> clears the control.</remarks>
    <Category("DateTimeBox")>
    <Bindable(True)>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the date and time represented by the control.")>
    Public Property [DateTime] As DateTime
        Get
            If CurrentValue.HasValue Then Return CurrentValue.Value
            Return System.DateTime.MinValue
        End Get
        Set(value As DateTime)
            If value = System.DateTime.MinValue Then
                SetTemporalValue(Nothing)
            Else
                SetTemporalValue(value)
            End If
        End Set
    End Property
    ''' <summary>
    ''' Gets a value indicating whether the control contains a valid date and time.
    ''' </summary>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property HasDateTime As Boolean
        Get
            Return CurrentValue.HasValue
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the culture used to parse, display, and edit date and time values.
    ''' </summary>
    <Category("DateTimeBox")>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the culture used to parse, display, and edit date and time values.")>
    Public Property DateTimeCulture As CultureInfo
        Get
            Return _DateTimeCulture
        End Get
        Set(value As CultureInfo)
            Dim NewCulture As CultureInfo = If(value, CultureInfo.CurrentCulture)
            If _DateTimeCulture.Equals(NewCulture) Then Return
            _DateTimeCulture = NewCulture
            ValueCulture = _DateTimeCulture
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether seconds are displayed and accepted.
    ''' </summary>
    <Category("DateTimeBox")>
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
    ''' Gets or sets the image displayed on the dropdown button.
    ''' </summary>
    <Category("DateTimeBox")>
    <DefaultValue(GetType(Image), Nothing)>
    <Description("Gets or sets the image displayed on the dropdown button.")>
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
    ''' Gets or sets the text displayed beside the time selector in the dropdown editor.
    ''' </summary>
    ''' <value>
    ''' The text displayed beside the time selector. The default value is
    ''' <c>"Time:"</c>.
    ''' </value>
    <Category("DateTimeBox")>
    <DefaultValue("Time:")>
    <Localizable(True)>
    <Description("Gets or sets the text displayed beside the time selector in the dropdown editor.")>
    Public Property TimeLabelText As String
        Get
            Return _TimeLabelText
        End Get
        Set(value As String)
            Dim NewValue As String = If(value, String.Empty)
            If _TimeLabelText = NewValue Then Return
            _TimeLabelText = NewValue
            DropDownEditor.TimeLabel.Text = NewValue
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the text displayed by the confirmation button in the dropdown editor.
    ''' </summary>
    ''' <value>
    ''' The confirmation button text. The default value is <c>"OK"</c>.
    ''' </value>
    <Category("DateTimeBox")>
    <DefaultValue("OK")>
    <Localizable(True)>
    <Description("Gets or sets the text displayed by the confirmation button in the dropdown editor.")>
    Public Property OKButtonText As String
        Get
            Return _OKButtonText
        End Get
        Set(value As String)
            Dim NewValue As String = If(value, String.Empty)
            If _OKButtonText = NewValue Then Return
            _OKButtonText = NewValue
            DropDownEditor.ConfirmButton.Text = NewValue
        End Set
    End Property

    ''' <summary>
    ''' Gets or sets the text displayed by the cancellation button in the dropdown editor.
    ''' </summary>
    ''' <value>
    ''' The cancellation button text. The default value is <c>"Cancel"</c>.
    ''' </value>
    <Category("DateTimeBox")>
    <DefaultValue("Cancel")>
    <Localizable(True)>
    <Description("Gets or sets the text displayed by the cancellation button in the dropdown editor.")>
    Public Property CancelButtonText As String
        Get
            Return _CancelButtonText
        End Get
        Set(value As String)
            Dim NewValue As String = If(value, String.Empty)
            If _CancelButtonText = NewValue Then Return
            _CancelButtonText = NewValue
            DropDownEditor.CancelButton.Text = NewValue
        End Set
    End Property
    ''' <summary>
    ''' Clears the date and time represented by the control.
    ''' </summary>
    Public Sub ClearDateTime()
        SetTemporalValue(Nothing)
    End Sub
    ''' <summary>
    ''' Gets the normalized date and time format used by the control.
    ''' </summary>
    Protected Overrides ReadOnly Property ValueFormat As String
        Get
            Dim DateFormat As String = NormalizeDateFormat(DateTimeCulture.DateTimeFormat.ShortDatePattern)
            Dim TimeFormat As String = If(ShowSeconds, DateTimeCulture.DateTimeFormat.LongTimePattern, DateTimeCulture.DateTimeFormat.ShortTimePattern)
            Return $"{DateFormat} {NormalizeTimeFormat(TimeFormat)}"
        End Get
    End Property
    ''' <summary>
    ''' Converts a date and time format into a masked text box mask.
    ''' </summary>
    Protected Overrides Function CreateMask(Format As String) As String
        Return Format.Replace("yyyy", "0000").Replace("yy", "00").Replace("MM", "00").Replace("dd", "00").Replace("HH", "00").Replace("hh", "00").Replace("mm", "00").Replace("ss", "00").Replace("tt", "LL")
    End Function
    ''' <summary>
    ''' Synchronizes the dropdown editor after the configured culture or format changes.
    ''' </summary>
    Protected Overrides Sub OnValueCultureChanged()
        MyBase.OnValueCultureChanged()
        If DropDownEditor Is Nothing Then Return
        DropDownEditor.DateTimeCulture = DateTimeCulture
        DropDownEditor.ShowSeconds = ShowSeconds
    End Sub
    ''' <summary>
    ''' Raises the <see cref="DateTimeValueChanged"/> event.
    ''' </summary>
    Protected Overrides Sub OnTemporalValueChanged(e As EventArgs)
        MyBase.OnTemporalValueChanged(e)
        OnDateTimeValueChanged(e)
    End Sub
    ''' <summary>
    ''' Raises the <see cref="DateTimeValueChanged"/> event.
    ''' </summary>
    ''' <param name="e">An <see cref="EventArgs"/> instance containing the event data.</param>
    Protected Overridable Sub OnDateTimeValueChanged(e As EventArgs)
        RaiseEvent DateTimeValueChanged(Me, e)
    End Sub
    ''' <summary>
    ''' Configures the right text margin after the control handle is created.
    ''' </summary>
    Protected Overrides Sub OnHandleCreated(e As EventArgs)
        MyBase.OnHandleCreated(e)
        SetRightTextMargin(DropDownButton.Width)
    End Sub
    ''' <summary>
    ''' Repositions the dropdown button when the control size changes.
    ''' </summary>
    Protected Overrides Sub OnSizeChanged(e As EventArgs)
        MyBase.OnSizeChanged(e)
        If DropDownButton Is Nothing Then Return
        DropDownButton.Size = New Size(25, ClientSize.Height + 2)
        DropDownButton.Location = New Point(ClientSize.Width - DropDownButton.Width + 1, -1)
        SetRightTextMargin(DropDownButton.Width)
    End Sub
    ''' <summary>
    ''' Synchronizes the dropdown button background color.
    ''' </summary>
    Protected Overrides Sub OnBackColorChanged(e As EventArgs)
        MyBase.OnBackColorChanged(e)
        If DropDownButton IsNot Nothing Then DropDownButton.BackColor = BackColor
    End Sub
    ''' <summary>
    ''' Opens the dropdown editor when ENTER is pressed.
    ''' </summary>
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        If e.KeyCode <> Keys.Enter Then Return
        ShowDateTimeSelector()
        e.SuppressKeyPress = True
    End Sub
    ''' <summary>
    ''' Normalizes single-character date format elements.
    ''' </summary>
    Private Shared Function NormalizeDateFormat(Format As String) As String
        Dim Result As String = Format
        Result = Regex.Replace(Result, "(?<!y)y(?!y)", "yy")
        Result = Regex.Replace(Result, "(?<!M)M(?!M)", "MM")
        Result = Regex.Replace(Result, "(?<!d)d(?!d)", "dd")
        Return Result
    End Function
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
    ''' Displays the date and time dropdown editor.
    ''' </summary>
    Private Sub ShowDateTimeSelector()
        DropDownEditor.Visible = True
        ControlContainer.ShowDropDown()
    End Sub
    ''' <summary>
    ''' Opens the dropdown editor when the internal button is clicked.
    ''' </summary>
    Private Sub DropDownButton_Click(sender As Object, e As EventArgs) Handles DropDownButton.Click
        ShowDateTimeSelector()
    End Sub
    ''' <summary>
    ''' Initializes the dropdown editor whenever it is displayed.
    ''' </summary>
    Private Sub ControlContainer_Dropped(sender As Object) Handles ControlContainer.Dropped
        _SelectionConfirmed = False
        DropDownEditor.DateTimeCulture = DateTimeCulture
        DropDownEditor.ShowSeconds = ShowSeconds
        DropDownEditor.SetValue(If(HasDateTime, Me.DateTime, System.DateTime.Now))
        DropDownEditor.FocusEditor()
    End Sub
    ''' <summary>
    ''' Marks the current selection as confirmed and closes the dropdown editor.
    ''' </summary>
    Private Sub DropDownEditor_ValueConfirmed(sender As Object, e As EventArgs) Handles DropDownEditor.ValueConfirmed
        _SelectionConfirmed = True
        ControlContainer.CloseDropDown()
    End Sub
    ''' <summary>
    ''' Cancels the current selection and closes the dropdown editor.
    ''' </summary>
    Private Sub DropDownEditor_SelectionCanceled(sender As Object, e As EventArgs) Handles DropDownEditor.SelectionCanceled
        _SelectionConfirmed = False
        ControlContainer.CloseDropDown()
    End Sub
    ''' <summary>
    ''' Applies a confirmed value after the dropdown editor closes.
    ''' </summary>
    Private Sub ControlContainer_Closed(sender As Object) Handles ControlContainer.Closed
        If _SelectionConfirmed Then Me.DateTime = DropDownEditor.SelectedValue
        _SelectionConfirmed = False
        [Select]()
    End Sub
    ''' <summary>
    ''' Draws a default combined calendar and clock image when no custom button image is assigned.
    ''' </summary>
    Private Sub DropDownButton_Paint(sender As Object, e As PaintEventArgs) Handles DropDownButton.Paint
        If ButtonImage IsNot Nothing Then Return
        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
        Using IconPen As New Pen(SystemColors.ControlDarkDark, 1.25F)
            e.Graphics.DrawRectangle(IconPen, 5, 6, 11, 10)
            e.Graphics.DrawLine(IconPen, 5, 9, 16, 9)
            e.Graphics.DrawLine(IconPen, 8, 4, 8, 8)
            e.Graphics.DrawLine(IconPen, 13, 4, 13, 8)
            e.Graphics.DrawEllipse(IconPen, 12, 11, 8, 8)
            e.Graphics.DrawLine(IconPen, 16, 15, 16, 12.5F)
            e.Graphics.DrawLine(IconPen, 16, 15, 18, 16)
        End Using
    End Sub
End Class