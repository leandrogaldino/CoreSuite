Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
''' <summary>
''' Represents a culture-aware date input control with an integrated calendar.
''' </summary>
<DefaultEvent("DateValueChanged")>
<DefaultProperty("Date")>
<DefaultBindingProperty("Date")>
<Designer(GetType(DateBoxControlDesigner))>
<ToolboxItem(True)>
<ToolboxItemFilter("CoreSuite")>
Public Class DateBox
    Inherits DateTimeBoxBase
    Private WithEvents ControlContainer As New ControlContainer
    Private WithEvents Calendar As New MonthCalendar
    Private WithEvents DropDownButton As New PictureBox
    Private _DateCulture As CultureInfo = CultureInfo.CurrentCulture
    Private _ButtonImage As Image
    Private _CalendarSelectionCommitted As Boolean
    ''' <summary>
    ''' Occurs when the represented date changes.
    ''' </summary>
    <Category("DateBox")>
    <Description("Occurs when the represented date changes.")>
    Public Event DateValueChanged As EventHandler
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DateBox"/> class.
    ''' </summary>
    Public Sub New()
        MinimumSize = New Size(100, 0)
        Calendar.Visible = False
        Calendar.MaxSelectionCount = 1
        DropDownButton.Cursor = Cursors.Default
        DropDownButton.TabStop = False
        DropDownButton.BackColor = BackColor
        Controls.Add(DropDownButton)
        ConfigureDropDownButton(DropDownButton, AddressOf RenderDropDownButton, Sub() DropDownButton_Click(DropDownButton, EventArgs.Empty))
        Controls.Add(Calendar)
        ControlContainer.HostControl = DropDownButton
        ControlContainer.HostedControl = Calendar
        InitializeValueText()
    End Sub
    ''' <summary>
    ''' Gets or sets the date represented by the control.
    ''' </summary>
    ''' <remarks>
    ''' Assigning <see cref="DateTime.MinValue"/> clears the control.
    ''' </remarks>
    <Category("DateBox")>
    <Bindable(True)>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the date represented by the control.")>
    Public Property [Date] As DateTime
        Get
            If CurrentValue.HasValue Then
                Return CurrentValue.Value.Date
            End If
            Return DateTime.MinValue
        End Get
        Set(value As DateTime)
            If value = DateTime.MinValue Then
                SetTemporalValue(Nothing)
            Else
                SetTemporalValue(value.Date)
            End If
        End Set
    End Property
    ''' <summary>
    ''' Indicates whether the control contains a valid date.
    ''' </summary>
    <Browsable(False)>
    Public ReadOnly Property HasDate As Boolean
        Get
            Return CurrentValue.HasValue
        End Get
    End Property
    ''' <summary>
    ''' Gets or sets the culture used to parse and format dates.
    ''' </summary>
    <Category("DateBox")>
    <RefreshProperties(RefreshProperties.All)>
    <Description("Gets or sets the culture used to parse and format dates.")>
    Public Property DateCulture As CultureInfo
        Get
            Return _DateCulture
        End Get
        Set(value As CultureInfo)
            Dim NewCulture As CultureInfo = If(value, CultureInfo.CurrentCulture)
            If _DateCulture.Equals(NewCulture) Then Return
            _DateCulture = NewCulture
            ValueCulture = _DateCulture
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the image displayed on the calendar button.
    ''' </summary>
    <Category("DateBox")>
    <DefaultValue(GetType(Image), Nothing)>
    <Description("Gets or sets the image displayed on the calendar button.")>
    Public Property ButtonImage As Image
        Get
            Return _ButtonImage
        End Get
        Set(value As Image)
            If ReferenceEquals(_ButtonImage, value) Then Return
            _ButtonImage = value
            DropDownButton.Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Gets the normalized short-date format used by the control.
    ''' </summary>
    Protected Overrides ReadOnly Property ValueFormat As String
        Get
            Return NormalizeDateFormat(DateCulture.DateTimeFormat.ShortDatePattern)
        End Get
    End Property
    ''' <summary>
    ''' Converts a date format into a masked text box mask.
    ''' </summary>
    Protected Overrides Function CreateMask(Format As String) As String
        Return Format.
            Replace("yyyy", "0000").
            Replace("yy", "00").
            Replace("MM", "00").
            Replace("dd", "00")
    End Function
    ''' <summary>
    ''' Removes the time component from a stored date value.
    ''' </summary>
    Protected Overrides Function NormalizeValue(Value As DateTime) As DateTime
        Return Value.Date
    End Function
    ''' <summary>
    ''' Raises the <see cref="DateValueChanged"/> event.
    ''' </summary>
    Protected Overrides Sub OnTemporalValueChanged(e As EventArgs)
        MyBase.OnTemporalValueChanged(e)
        OnDateValueChanged(e)
    End Sub
    ''' <summary>
    ''' Raises the <see cref="DateValueChanged"/> event.
    ''' </summary>
    Protected Overridable Sub OnDateValueChanged(e As EventArgs)
        RaiseEvent DateValueChanged(Me, e)
    End Sub
    ''' <summary>
    ''' Synchronizes the button background color.
    ''' </summary>
    Protected Overrides Sub OnBackColorChanged(e As EventArgs)
        MyBase.OnBackColorChanged(e)
        If DropDownButton IsNot Nothing Then
            DropDownButton.BackColor = BackColor
        End If
    End Sub
    ''' <summary>
    ''' Opens the calendar when ENTER is pressed.
    ''' </summary>
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        If e.KeyCode <> Keys.Enter Then Return
        ShowCalendar()
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
    ''' Displays the calendar dropdown.
    ''' </summary>
    Private Sub ShowCalendar()
        Calendar.Visible = True
        ControlContainer.ShowDropDown()
    End Sub
    ''' <summary>
    ''' Opens the calendar when the internal button is clicked.
    ''' </summary>
    Private Sub DropDownButton_Click(sender As Object, e As EventArgs) Handles DropDownButton.Click
        ShowCalendar()
    End Sub
    ''' <summary>
    ''' Configures the initial date displayed by the calendar.
    ''' </summary>
    Private Sub ControlContainer_Dropped(sender As Object) Handles ControlContainer.Dropped
        _CalendarSelectionCommitted = False
        If HasDate Then
            Calendar.SetDate(Me.Date)
        Else
            Calendar.SetDate(Today)
        End If
    End Sub
    ''' <summary>
    ''' Marks the selected calendar date for confirmation.
    ''' </summary>
    Private Sub Calendar_DateSelected(sender As Object, e As DateRangeEventArgs) Handles Calendar.DateSelected
        _CalendarSelectionCommitted = True
        ControlContainer.CloseDropDown()
    End Sub
    ''' <summary>
    ''' Applies the selected date after the calendar closes.
    ''' </summary>
    Private Sub ControlContainer_Closed(sender As Object) Handles ControlContainer.Closed
        If _CalendarSelectionCommitted Then
            Me.Date = Calendar.SelectionStart.Date
        End If
        _CalendarSelectionCommitted = False
        Me.Select()
    End Sub
    ''' <summary>
    ''' Draws a default calendar image when no button image is assigned.
    ''' </summary>
    Private Sub RenderDropDownButton(G As Graphics, SurfaceSize As Size)
        If ButtonImage IsNot Nothing Then
            EmbeddedTextButtonLayout.DrawCenteredImage(G, ButtonImage, SurfaceSize, CInt(Math.Round(4 * DeviceDpi / 96.0R)), DropDownButton.Enabled, DropDownButton.BackColor)
            Return
        End If
        G.SmoothingMode = SmoothingMode.AntiAlias
        Dim IconWidth As Integer = Math.Min(SurfaceSize.Width - 10, 14)
        Dim IconHeight As Integer = Math.Min(SurfaceSize.Height - 10, 14)
        If IconWidth <= 0 OrElse IconHeight <= 0 Then Return
        Dim Left As Integer = (SurfaceSize.Width - IconWidth) \ 2
        Dim Top As Integer = (SurfaceSize.Height - IconHeight) \ 2
        Dim HeaderHeight As Integer = Math.Max(3, IconHeight \ 3)
        Using CalendarPen As New Pen(SystemColors.ControlDarkDark, 1.4F)
            Using HeaderBrush As New SolidBrush(SystemColors.ControlDarkDark)
                G.DrawRectangle(CalendarPen, Left, Top + 2, IconWidth, IconHeight - 2)
                G.FillRectangle(HeaderBrush, Left, Top + 2, IconWidth, HeaderHeight)
                G.DrawLine(CalendarPen, Left + 3, Top, Left + 3, Top + 4)
                G.DrawLine(CalendarPen, Left + IconWidth - 3, Top, Left + IconWidth - 3, Top + 4)
                Dim CellTop As Integer = Top + HeaderHeight + 4
                Dim CellBottom As Integer = Top + IconHeight - 2
                Dim CellHeight As Integer = CellBottom - CellTop
                If CellHeight > 2 Then
                    Dim FirstColumn As Integer = Left + IconWidth \ 3
                    Dim SecondColumn As Integer = Left + (IconWidth * 2) \ 3
                    Dim Row As Integer = CellTop + CellHeight \ 2
                    G.DrawLine(CalendarPen, FirstColumn, CellTop, FirstColumn, CellBottom)
                    G.DrawLine(CalendarPen, SecondColumn, CellTop, SecondColumn, CellBottom)
                    G.DrawLine(CalendarPen, Left + 1, Row, Left + IconWidth - 1, Row)
                End If
            End Using
        End Using
    End Sub
End Class
