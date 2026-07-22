''' <summary>
''' Represents a single column configuration within a DataGridView layout system.
''' </summary>
''' <remarks>
''' This model defines how a column should be displayed, including:
''' - Visibility in grid and context menu
''' - Display order and name
''' - Width and auto-size behavior
''' - Cell and header alignment
''' - Formatting string for display values
''' </remarks>
Public Class DataGridViewLayoutColumn
    ''' <summary>
    ''' Indicates whether this column is allowed to appear in the context menu for visibility toggling.
    ''' </summary>
    Public ReadOnly Property VisibleInContext As Boolean = True
    ''' <summary>
    ''' Gets or sets the column width in pixels when AutoSize is disabled.
    ''' </summary>
    Public Property Width As Integer = 100
    ''' <summary>
    ''' Gets or sets the auto-size mode used to determine column width behavior.
    ''' </summary>
    Public Property WidthSizeMode As DataGridViewAutoSizeColumnMode = DataGridViewAutoSizeColumnMode.None
    ''' <summary>
    ''' Gets or sets whether the column is visible in the DataGridView.
    ''' </summary>
    Public Property VisibleInGrid As Boolean = True
    ''' <summary>
    ''' Gets or sets the display index (order) of the column in the DataGridView.
    ''' </summary>
    Public Property DisplayIndex As Integer
    ''' <summary>
    ''' Gets or sets the display name (header text) of the column.
    ''' </summary>
    Public Property DisplayName As String
    ''' <summary>
    ''' Gets or sets the alignment applied to cell values in this column.
    ''' </summary>
    Public Property CellAlignment As DataGridViewContentAlignment = DataGridViewContentAlignment.NotSet
    ''' <summary>
    ''' Gets or sets the alignment applied to the column header text.
    ''' </summary>
    Public Property HeaderAlignment As DataGridViewContentAlignment = DataGridViewContentAlignment.NotSet
    ''' <summary>
    ''' Gets or sets the format string used to display cell values.
    ''' </summary>
    Public Property Format As String
    ''' <summary>
    ''' Initializes a new instance of the DataGridViewLayoutColumn class.
    ''' </summary>
    ''' <param name="VisibleInContext">
    ''' Indicates whether the column should appear in the context menu for visibility control.
    ''' </param>
    Public Sub New(Optional VisibleInContext As Boolean = True)
        Me.VisibleInContext = VisibleInContext
    End Sub
End Class
