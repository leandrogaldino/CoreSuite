Imports System.ComponentModel
Imports System.Drawing.Design
''' <summary>
''' Represents a button that renders multiple text parts with independent text, font, and color settings.
''' </summary>
<DefaultProperty("TextParts")>
Public Class RichTextButton
    Inherits NoFocusCueButton
    Private _TextParts As BindingList(Of RichTextPart)
    Private _HideBaseText As Boolean
    ''' <summary>
    ''' Gets the combined text of all configured text parts.
    ''' </summary>
    ''' <value>
    ''' A string containing the concatenated text of all items in <see cref="TextParts"/>.
    ''' </value>
    ''' <remarks>
    ''' Setting this property has no effect. Use <see cref="TextParts"/> to configure the displayed text.
    ''' </remarks>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    <Browsable(False), EditorBrowsable(EditorBrowsableState.Never)>
    Public Overrides Property Text As String
        Get
            If _HideBaseText Then Return String.Empty
            Return String.Concat(TextParts.Select(Function(Part) Part.Text))
        End Get
        Set(value As String)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the text displayed by the tooltip associated with the button.
    ''' </summary>
    ''' <value>
    ''' The tooltip text displayed for the button.
    ''' </value>
    <Category("RichTextButton")>
    Public Overrides Property TooltipText As String
        Get
            Return MyBase.TooltipText
        End Get
        Set(value As String)
            MyBase.TooltipText = value
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the collection of independently formatted text parts rendered by the button.
    ''' </summary>
    ''' <value>
    ''' A bindable collection containing the text parts displayed by the button.
    ''' </value>
    <Category("RichTextButton")>
    <Description("Gets or sets the collection of independently formatted text parts rendered by the button.")>
    <Editor(GetType(RichTextCollectionEditor), GetType(UITypeEditor))>
    <TypeConverter(GetType(RichTextConverter))>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property TextParts As BindingList(Of RichTextPart)
        Get
            Return _TextParts
        End Get
        Set(value As BindingList(Of RichTextPart))
            If ReferenceEquals(_TextParts, value) Then Return
            If _TextParts IsNot Nothing Then RemoveHandler _TextParts.ListChanged, AddressOf TextParts_ListChanged
            _TextParts = If(value, New BindingList(Of RichTextPart))
            AddHandler _TextParts.ListChanged, AddressOf TextParts_ListChanged
            Invalidate()
        End Set
    End Property
    ''' <summary>
    ''' Initializes a new instance of the <see cref="RichTextButton"/> class.
    ''' </summary>
    Public Sub New()
        TextParts = New BindingList(Of RichTextPart)
    End Sub
    ''' <summary>
    ''' Handles changes made to the text-parts collection and invalidates the control.
    ''' </summary>
    ''' <param name="sender">The object that raised the event.</param>
    ''' <param name="e">An object containing information about the collection change.</param>
    Private Sub TextParts_ListChanged(sender As Object, e As ListChangedEventArgs)
        Invalidate()
    End Sub
    ''' <summary>
    ''' Paints the button and renders each configured text part using its individual formatting.
    ''' </summary>
    ''' <param name="e">An object containing the graphics context and clipping information used to paint the control.</param>
    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        _HideBaseText = True
        Try
            MyBase.OnPaint(e)
        Finally
            _HideBaseText = False
        End Try
        If TextParts.Count = 0 Then Return
        Dim ContentBounds As Rectangle = ClientRectangle
        ContentBounds = Rectangle.FromLTRB(ContentBounds.Left + Padding.Left, ContentBounds.Top + Padding.Top, ContentBounds.Right - Padding.Right, ContentBounds.Bottom - Padding.Bottom)
        If ContentBounds.Width <= 0 OrElse ContentBounds.Height <= 0 Then Return
        Using Format As New StringFormat(StringFormat.GenericTypographic)
            Format.FormatFlags = Format.FormatFlags Or StringFormatFlags.MeasureTrailingSpaces
            Dim PartSizes As New List(Of SizeF)
            Dim TotalWidth As Single
            Dim TotalHeight As Single
            For Each Part As RichTextPart In TextParts
                If String.IsNullOrEmpty(Part.Text) Then
                    PartSizes.Add(SizeF.Empty)
                    Continue For
                End If
                Dim PartSize As SizeF = e.Graphics.MeasureString(Part.Text, Part.Font, Integer.MaxValue, Format)
                PartSizes.Add(PartSize)
                TotalWidth += PartSize.Width
                TotalHeight = Math.Max(TotalHeight, PartSize.Height)
            Next
            Dim StartX As Single = GetHorizontalPosition(ContentBounds, TotalWidth)
            Dim StartY As Single = GetVerticalPosition(ContentBounds, TotalHeight)
            For Index As Integer = 0 To TextParts.Count - 1
                Dim Part As RichTextPart = TextParts(Index)
                Dim PartSize As SizeF = PartSizes(Index)
                If PartSize.IsEmpty Then Continue For
                Dim PartY As Single = StartY + ((TotalHeight - PartSize.Height) / 2.0F)
                Dim DrawColor As Color = If(Enabled, Part.Color, SystemColors.GrayText)
                Using Brush As New SolidBrush(DrawColor)
                    e.Graphics.DrawString(Part.Text, Part.Font, Brush, New PointF(StartX, PartY), Format)
                End Using
                StartX += PartSize.Width
            Next
        End Using
    End Sub
    ''' <summary>
    ''' Calculates the horizontal position at which the combined text must be rendered.
    ''' </summary>
    ''' <param name="Bounds">The available content bounds inside the button.</param>
    ''' <param name="TotalWidth">The combined width of all text parts.</param>
    ''' <returns>The horizontal coordinate at which text rendering must begin.</returns>
    Private Function GetHorizontalPosition(Bounds As Rectangle, TotalWidth As Single) As Single
        Select Case TextAlign
            Case ContentAlignment.TopLeft, ContentAlignment.MiddleLeft, ContentAlignment.BottomLeft
                Return Bounds.Left
            Case ContentAlignment.TopCenter, ContentAlignment.MiddleCenter, ContentAlignment.BottomCenter
                Return Bounds.Left + ((Bounds.Width - TotalWidth) / 2.0F)
            Case Else
                Return Bounds.Right - TotalWidth
        End Select
    End Function
    ''' <summary>
    ''' Calculates the vertical position at which the combined text must be rendered.
    ''' </summary>
    ''' <param name="Bounds">The available content bounds inside the button.</param>
    ''' <param name="TotalHeight">The maximum height among all text parts.</param>
    ''' <returns>The vertical coordinate at which text rendering must begin.</returns>
    Private Function GetVerticalPosition(Bounds As Rectangle, TotalHeight As Single) As Single
        Select Case TextAlign
            Case ContentAlignment.TopLeft, ContentAlignment.TopCenter, ContentAlignment.TopRight
                Return Bounds.Top
            Case ContentAlignment.MiddleLeft, ContentAlignment.MiddleCenter, ContentAlignment.MiddleRight
                Return Bounds.Top + ((Bounds.Height - TotalHeight) / 2.0F)
            Case Else
                Return Bounds.Bottom - TotalHeight
        End Select
    End Function
End Class