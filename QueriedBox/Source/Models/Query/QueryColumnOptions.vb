Imports System.ComponentModel

Public Class QueryColumnOptions
    ''' <summary>
    ''' Gets or sets a value indicating whether the column should be displayed in the result view.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets whether the column should be displayed in the result view.")>
    Public Property Display As Boolean = True
    ''' <summary>
    ''' Gets or sets a value indicating whether the column should remain frozen in the result display.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets whether the column should remain frozen in the result display.")>
    Public Property Freeze As Boolean = True
    ''' <summary>
    ''' Gets or sets a value indicating whether the column participates in text searches.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets whether the column participates in text searches.")>
    Public Property Searchable As Boolean = True
    ''' <summary>
    ''' Gets or sets the text prefix applied to the displayed column value.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the text prefix applied to the displayed column value.")>
    Public Property Prefix As String
    ''' <summary>
    ''' Gets or sets the text suffix applied to the displayed column value.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the text suffix applied to the displayed column value.")>
    Public Property Suffix As String
    ''' <summary>
    ''' Gets or sets the automatic sizing mode used for the corresponding DataGridView column.
    ''' </summary>
    <Category("QueriedBox")>
    <Description("Gets or sets the automatic sizing mode used for the corresponding DataGridView column.")>
    Public Property SizeColumnMode As DataGridViewAutoSizeColumnMode
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryColumnOptions"/> class.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="QueryColumnOptions"/> class
    ''' with the specified display, freeze, search, formatting, and sizing options.
    ''' </summary>
    ''' <param name="Display">
    ''' Indicates whether the column should be displayed in the result view.
    ''' </param>
    ''' <param name="Freeze">
    ''' Indicates whether the column value should be stored when a result is frozen.
    ''' </param>
    ''' <param name="Searchable">
    ''' Indicates whether the column participates in text searches.
    ''' </param>
    ''' <param name="Prefix">
    ''' The text prefix applied to the displayed column value.
    ''' </param>
    ''' <param name="Suffix">
    ''' The text suffix applied to the displayed column value.
    ''' </param>
    ''' <param name="SizeColumnMode">
    ''' The automatic sizing mode used for the corresponding DataGridView column.
    ''' </param>
    Public Sub New(Display As Boolean, Freeze As Boolean, Searchable As Boolean, Prefix As String, Suffix As String, SizeColumnMode As DataGridViewAutoSizeColumnMode)
        Me.Display = Display
        Me.Freeze = Freeze
        Me.Searchable = Searchable
        Me.Prefix = Prefix
        Me.Suffix = Suffix
        Me.SizeColumnMode = SizeColumnMode
    End Sub
    ''' <summary>
    ''' Returns a string representation of the current <see cref="QueryColumnOptions"/> instance.
    ''' </summary>
    ''' <returns>
    ''' An empty string.
    ''' </returns>
    Public Overrides Function ToString() As String
        Return String.Empty
    End Function
End Class

