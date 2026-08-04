Imports System.ComponentModel
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides smart tag actions and design-time property access for the <see cref="DataGridViewFilterBox"/> control.
''' </summary>
Public Class DataGridViewFilterBoxControlDesignerActionList
    Inherits DesignerActionList
    Private ReadOnly _Control As DataGridViewFilterBox
    ''' <summary>
    ''' Initializes a new instance of the <see cref="DataGridViewFilterBoxControlDesignerActionList"/> class.
    ''' </summary>
    ''' <param name="Designer">The designer associated with the filter box.</param>
    Public Sub New(Designer As DataGridViewFilterBoxControlDesigner)
        MyBase.New(Designer.Component)
        _Control = CType(Designer.Component, DataGridViewFilterBox)
    End Sub
    ''' <summary>
    ''' Gets the collection of smart tag items displayed in the Windows Forms designer.
    ''' </summary>
    ''' <returns>A collection containing the most frequently used filter configuration properties.</returns>
    Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
        Return New DesignerActionItemCollection From {
            New DesignerActionHeaderItem("Data source"),
            New DesignerActionPropertyItem(NameOf(DataGridView), "DataGridView", "Data source", "Defines the DataGridView whose source is filtered."),
            New DesignerActionHeaderItem("Filtering"),
            New DesignerActionPropertyItem(NameOf(FilterMode), "FilterMode", "Filtering", "Defines whether filtering is automatic, local-only, or custom."),
            New DesignerActionPropertyItem(NameOf(SearchMode), "SearchMode", "Filtering", "Defines how entered text is matched during local filtering."),
            New DesignerActionPropertyItem(NameOf(MinimumCharacters), "MinimumCharacters", "Filtering", "Defines the minimum number of characters required before filtering."),
            New DesignerActionPropertyItem(NameOf(FilterInterval), "FilterInterval", "Filtering", "Defines the debounce interval in milliseconds."),
            New DesignerActionPropertyItem(NameOf(CaseSensitive), "CaseSensitive", "Filtering", "Determines whether local text matching is case-sensitive."),
            New DesignerActionPropertyItem(NameOf(IncludeHiddenColumns), "IncludeHiddenColumns", "Filtering", "Determines whether hidden columns participate in automatic filtering."),
            New DesignerActionHeaderItem("Appearance"),
            New DesignerActionPropertyItem(NameOf(ShowClearButton), "ShowClearButton", "Appearance", "Determines whether the embedded clear button is displayed.")
        }
    End Function
    ''' <summary>
    ''' Gets or sets the target <see cref="DataGridView"/>.
    ''' </summary>
    Public Property DataGridView As DataGridView
        Get
            Return _Control.DataGridView
        End Get
        Set(value As DataGridView)
            SetProperty(NameOf(DataGridView), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets how filter requests are processed.
    ''' </summary>
    Public Property FilterMode As DataGridViewFilterMode
        Get
            Return _Control.FilterMode
        End Get
        Set(value As DataGridViewFilterMode)
            SetProperty(NameOf(FilterMode), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets how entered text is matched during local filtering.
    ''' </summary>
    Public Property SearchMode As DataGridViewFilterSearchMode
        Get
            Return _Control.SearchMode
        End Get
        Set(value As DataGridViewFilterSearchMode)
            SetProperty(NameOf(SearchMode), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the minimum character count required before filtering.
    ''' </summary>
    Public Property MinimumCharacters As Integer
        Get
            Return _Control.MinimumCharacters
        End Get
        Set(value As Integer)
            SetProperty(NameOf(MinimumCharacters), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the debounce interval in milliseconds.
    ''' </summary>
    Public Property FilterInterval As Integer
        Get
            Return _Control.FilterInterval
        End Get
        Set(value As Integer)
            SetProperty(NameOf(FilterInterval), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether local text matching is case-sensitive.
    ''' </summary>
    Public Property CaseSensitive As Boolean
        Get
            Return _Control.CaseSensitive
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(CaseSensitive), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether hidden columns participate in automatic filtering.
    ''' </summary>
    Public Property IncludeHiddenColumns As Boolean
        Get
            Return _Control.IncludeHiddenColumns
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(IncludeHiddenColumns), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the embedded clear button is displayed.
    ''' </summary>
    Public Property ShowClearButton As Boolean
        Get
            Return _Control.ShowClearButton
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(ShowClearButton), value)
        End Set
    End Property
    Private Sub SetProperty(PropertyName As String, Value As Object)
        Dim Descriptor As PropertyDescriptor = TypeDescriptor.GetProperties(_Control)(PropertyName)
        If Descriptor Is Nothing Then Throw New InvalidOperationException($"Property '{PropertyName}' was not found.")
        Descriptor.SetValue(_Control, Value)
    End Sub
End Class
