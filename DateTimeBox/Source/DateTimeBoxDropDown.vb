Imports System.ComponentModel
Imports System.Globalization

''' <summary>
''' Provides the dropdown editor used to select the date and time
''' represented by a <see cref="DateTimeBox"/>.
''' </summary>
<ToolboxItem(False)>
Partial Friend Class DateTimeBoxDropDown
    Inherits UserControl

    Private _DateTimeCulture As CultureInfo = CultureInfo.CurrentCulture
    Private _ShowSeconds As Boolean
    ''' <summary>
    ''' Occurs when the user confirms the selected date and time.
    ''' </summary>
    Public Event ValueConfirmed As EventHandler
    ''' <summary>
    ''' Occurs when the user cancels the current selection.
    ''' </summary>
    Public Event SelectionCanceled As EventHandler
    ''' <summary>
    ''' Initializes a new instance of the
    ''' <see cref="DateTimeBoxDropDown"/> class.
    ''' </summary>
    Public Sub New()
        InitializeComponent()
        ApplyCulture()
    End Sub
    ''' <summary>
    ''' Gets the date and time currently selected in the dropdown editor.
    ''' </summary>
    ''' <value>
    ''' A <see cref="System.DateTime"/> composed of the selected calendar
    ''' date and the selected time.
    ''' </value>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public ReadOnly Property SelectedValue As DateTime
        Get
            Return Calendar.SelectionStart.Date.Add(TimePicker.Value.TimeOfDay)
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the culture used to format the time selector and
    ''' localize the captions displayed by the editor.
    ''' </summary>
    ''' <value>
    ''' The culture used by the dropdown editor. If <see langword="Nothing"/>
    ''' is assigned, <see cref="CultureInfo.CurrentCulture"/> is used.
    ''' </value>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property DateTimeCulture As CultureInfo
        Get
            Return _DateTimeCulture
        End Get
        Set(value As CultureInfo)
            Dim NewCulture As CultureInfo = If(value, CultureInfo.CurrentCulture)
            If Equals(_DateTimeCulture, NewCulture) Then Return
            _DateTimeCulture = NewCulture
            ApplyCulture()
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the time selector displays and accepts seconds.
    ''' </summary>
    ''' <value>
    ''' <see langword="True"/> to display seconds; otherwise,
    ''' <see langword="False"/>.
    ''' </value>
    <Browsable(False)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Property ShowSeconds As Boolean
        Get
            Return _ShowSeconds
        End Get
        Set(value As Boolean)
            If _ShowSeconds = value Then Return
            _ShowSeconds = value
            ApplyCulture()
        End Set
    End Property
    ''' <summary>
    ''' Loads the specified date and time into the calendar and time selector.
    ''' </summary>
    ''' <param name="Value">
    ''' The date and time that should initially be displayed.
    ''' </param>
    Public Sub SetValue(Value As DateTime)
        Calendar.SetDate(Value.Date)
        TimePicker.Value = System.DateTime.Today.Add(Value.TimeOfDay)
    End Sub
    ''' <summary>
    ''' Moves the input focus to the calendar.
    ''' </summary>
    Public Sub FocusEditor()
        Calendar.Focus()
    End Sub
    ''' <summary>
    ''' Ensures that the dropdown editor and its child controls are ready
    ''' before the hosting container displays them.
    ''' </summary>
    Public Sub PrepareForDisplay()
        CreateControl()
        PerformLayout()
    End Sub
    ''' <summary>
    ''' Applies the configured culture, localized captions and time format
    ''' to the editor controls.
    ''' </summary>
    Private Sub ApplyCulture()
        If Calendar Is Nothing OrElse TimePicker Is Nothing Then Return
        Dim IsPortuguese As Boolean = String.Equals(DateTimeCulture.TwoLetterISOLanguageName, "pt", StringComparison.OrdinalIgnoreCase)
        TimeLabel.Text = If(IsPortuguese, "Hora:", "Time:")
        ConfirmButton.Text = "OK"
        CancelButton.Text = If(IsPortuguese, "Cancelar", "Cancel")
        TimePicker.CustomFormat = If(ShowSeconds, DateTimeCulture.DateTimeFormat.LongTimePattern, DateTimeCulture.DateTimeFormat.ShortTimePattern)
        PerformLayout()
    End Sub
    ''' <summary>
    ''' Raises the <see cref="ValueConfirmed"/> event when the confirmation
    ''' button is clicked.
    ''' </summary>
    ''' <param name="sender">The confirmation button.</param>
    ''' <param name="e">The event data.</param>
    Private Sub ConfirmButton_Click(sender As Object, e As EventArgs) Handles ConfirmButton.Click
        RaiseEvent ValueConfirmed(Me, EventArgs.Empty)
    End Sub
    ''' <summary>
    ''' Raises the <see cref="SelectionCanceled"/> event when the cancellation
    ''' button is clicked.
    ''' </summary>
    ''' <param name="sender">The cancellation button.</param>
    ''' <param name="e">The event data.</param>
    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
        RaiseEvent SelectionCanceled(Me, EventArgs.Empty)
    End Sub
End Class