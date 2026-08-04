Imports System.ComponentModel
''' <summary>
''' Identifies a data-bound column used by a <see cref="DataGridViewFilterBox"/> configuration collection.
''' </summary>
<DefaultProperty("ColumnName")>
<TypeConverter(GetType(ExpandableObjectConverter))>
Public Class DataGridViewFilterColumn
    Implements INotifyPropertyChanged
    Private _ColumnName As String = String.Empty
    ''' <summary>
    ''' Occurs when the configured column name changes.
    ''' </summary>
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    ''' <summary>
    ''' Initializes a new empty instance of the <see cref="DataGridViewFilterColumn"/> class.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DataGridViewFilterColumn"/> class with the specified column name.
    ''' </summary>
    ''' <param name="ColumnName">The <see cref="DataColumn.ColumnName"/>, <see cref="DataGridViewColumn.Name"/>, or <see cref="DataGridViewColumn.DataPropertyName"/> to reference.</param>
    Public Sub New(ColumnName As String)
        Me.ColumnName = ColumnName
    End Sub
    ''' <summary>
    ''' Gets or sets the name of the referenced data-bound column.
    ''' </summary>
    ''' <value>The data column name, grid column name, or data property name used to locate the column.</value>
    <Category("DataGridViewFilterColumn")>
    <DefaultValue("")>
    <Description("Defines the DataColumn name, DataGridView column name, or data property name referenced by this item.")>
    <NotifyParentProperty(True)>
    <RefreshProperties(RefreshProperties.All)>
    Public Property ColumnName As String
        Get
            Return _ColumnName
        End Get
        Set(value As String)
            Dim NormalizedValue As String = If(value, String.Empty).Trim()
            If String.Equals(_ColumnName, NormalizedValue, StringComparison.Ordinal) Then Return
            _ColumnName = NormalizedValue
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(ColumnName)))
        End Set
    End Property
    ''' <summary>
    ''' Returns the configured column name for display in collection editors and property grids.
    ''' </summary>
    ''' <returns>The configured <see cref="ColumnName"/>, or <c>(Column)</c> when no name has been assigned.</returns>
    Public Overrides Function ToString() As String
        Return If(String.IsNullOrWhiteSpace(ColumnName), "(Column)", ColumnName)
    End Function
End Class
