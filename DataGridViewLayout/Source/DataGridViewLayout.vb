''' <summary>
''' Represents a full DataGridView layout configuration for a specific routine/version.
''' </summary>
''' <remarks>
''' This class stores the complete state of a DataGridView layout, including:
''' - Column configurations
''' - Sorting state
''' - Version control for migration/updates
''' It supports deep cloning for safe manipulation without affecting the original instance.
''' </remarks>
Public Class DataGridViewLayout
    ''' <summary>
    ''' Gets the routine name associated with this layout configuration.
    ''' </summary>
    Public ReadOnly Property Routine As String
    ''' <summary>
    ''' Gets the version of the layout schema.
    ''' Used to determine whether a layout file is outdated.
    ''' </summary>
    Public ReadOnly Property Version As Integer
    ''' <summary>
    ''' Gets or sets the index of the column currently used for sorting.
    ''' A value of -1 indicates no sorting.
    ''' </summary>
    Public Property SortedColumn As Integer = -1
    ''' <summary>
    ''' Gets or sets the current sort direction applied to the grid.
    ''' </summary>
    Public Property SortDirection As SortOrder = SortOrder.None
    ''' <summary>
    ''' Gets the list of column configurations that define the layout structure.
    ''' </summary>
    Public Property Columns As New List(Of DataGridViewLayoutColumn)
    ''' <summary>
    ''' Initializes a new instance of the DataGridViewLayout class.
    ''' </summary>
    ''' <param name="Routine">The identifier name for the layout configuration.</param>
    ''' <param name="Version">The schema version of the layout.</param>
    Public Sub New(Routine As String, Version As Integer)
        Me.Routine = Routine
        Me.Version = Version
    End Sub
    ''' <summary>
    ''' Creates a deep copy of the current layout instance.
    ''' </summary>
    ''' <returns>
    ''' A new DataGridViewLayout instance with copied column configurations
    ''' and identical sorting state.
    ''' </returns>
    ''' <remarks>
    ''' All column objects are cloned individually to ensure full isolation
    ''' between the original and the cloned layout.
    ''' </remarks>
    Public Function Clone() As DataGridViewLayout
        Dim MyClone As New DataGridViewLayout(Routine, Version) With {
            .SortedColumn = Me.SortedColumn,
            .SortDirection = Me.SortDirection
        }
        For Each Column In Columns
            MyClone.Columns.Add(New DataGridViewLayoutColumn(Column.VisibleInContext) With {
                .Width = Column.Width,
                .WidthSizeMode = Column.WidthSizeMode,
                .VisibleInGrid = Column.VisibleInGrid,
                .DisplayIndex = Column.DisplayIndex,
                .DisplayName = Column.DisplayName,
                .CellAlignment = Column.CellAlignment,
                .HeaderAlignment = Column.HeaderAlignment,
                .Format = Column.Format
            })
        Next Column
        Return MyClone
    End Function
End Class