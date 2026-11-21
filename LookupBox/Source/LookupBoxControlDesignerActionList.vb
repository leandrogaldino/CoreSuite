Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Drawing
''' <summary>
''' Provides smart-tag actions and design-time property access for the <see cref="LookupBox"/> control.
''' </summary>
Public Class LookupBoxControlDesignerActionList
    Inherits DesignerActionList
    Private ReadOnly _Control As LookupBox
    ''' <summary>
    ''' Initializes a new instance of the <see cref="LookupBoxControlDesignerActionList"/> class.
    ''' </summary>
    ''' <param name="Designer">The designer associated with the lookup box.</param>
    Public Sub New(Designer As LookupBoxControlDesigner)
        MyBase.New(Designer.Component)
        _Control = CType(Designer.Component, LookupBox)
    End Sub
    ''' <summary>
    ''' Gets the smart-tag items displayed by the Windows Forms designer.
    ''' </summary>
    ''' <returns>A collection containing common lookup, result, and appearance settings.</returns>
    Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
        Return New DesignerActionItemCollection From {
            New DesignerActionHeaderItem("Selection"),
            New DesignerActionPropertyItem(NameOf(DisplayMember), "DisplayMember", "Selection", "Defines the property used as selected display text."),
            New DesignerActionPropertyItem(NameOf(ValueMember), "ValueMember", "Selection", "Defines the property used as the selected value."),
            New DesignerActionPropertyItem(NameOf(HighlightSelectedItem), "HighlightSelectedItem", "Selection", "Determines whether selected items use distinct text-box colors."),
            New DesignerActionPropertyItem(NameOf(SelectedItemBackColor), "SelectedItemBackColor", "Selection", "Defines the selected-state background color."),
            New DesignerActionPropertyItem(NameOf(SelectedItemForeColor), "SelectedItemForeColor", "Selection", "Defines the selected-state text color."),
            New DesignerActionPropertyItem(NameOf(ClearOnEdit), "ClearOnEdit", "Selection", "Clears the text when a selected record is edited."),
            New DesignerActionPropertyItem(NameOf(AutoSelectMatchedResultOnLeave), "AutoSelectMatchedResultOnLeave", "Selection", "Selects the unique matching result on leave."),
            New DesignerActionPropertyItem(NameOf(AllowSelectionLink), "AllowSelectionLink", "Selection", "Enables Ctrl+Click on the selected result."),
            New DesignerActionHeaderItem("Searching"),
            New DesignerActionPropertyItem(NameOf(SearchEnabled), "SearchEnabled", "Searching", "Determines whether entered text starts lookup requests."),
            New DesignerActionPropertyItem(NameOf(MinimumCharacters), "MinimumCharacters", "Searching", "Defines the minimum character count required before searching."),
            New DesignerActionPropertyItem(NameOf(SearchInterval), "SearchInterval", "Searching", "Defines the debounce interval in milliseconds."),
            New DesignerActionPropertyItem(NameOf(MaximumResults), "MaximumResults", "Searching", "Defines the maximum number of retained results."),
            New DesignerActionPropertyItem(NameOf(AutoSelectSingleResult), "AutoSelectSingleResult", "Searching", "Determines whether a single result is selected automatically."),
            New DesignerActionHeaderItem("Drop-down"),
            New DesignerActionPropertyItem(NameOf(ResultsViewMode), "ResultsViewMode", "Drop-down", "Grid or custom UserControl templates."),
            New DesignerActionPropertyItem(NameOf(TemplateItemHeight), "TemplateItemHeight", "Drop-down", "Height of each template row."),
            New DesignerActionPropertyItem(NameOf(DropDownWidth), "DropDownWidth", "Drop-down", "Defines the result drop-down width."),
            New DesignerActionPropertyItem(NameOf(DropDownHeight), "DropDownHeight", "Drop-down", "Defines the result drop-down height."),
            New DesignerActionPropertyItem(NameOf(ShowColumnHeaders), "ShowColumnHeaders", "Drop-down", "Determines whether result-column headers are displayed."),
            New DesignerActionPropertyItem(NameOf(ShowClearButton), "ShowClearButton", "Drop-down", "Determines whether the clear and cancel button is displayed."),
            New DesignerActionPropertyItem(NameOf(AutoStretchDropDown), "AutoStretchDropDown", "Drop-down", "Grows the drop-down width to fit columns."),
            New DesignerActionPropertyItem(NameOf(GridHeadersBold), "GridHeadersBold", "Drop-down", "Uses bold result-grid column headers.")
        }
    End Function
    ''' <summary>
    ''' Gets or sets the selected display-property path.
    ''' </summary>
    Public Property DisplayMember As String
        Get
            Return _Control.DisplayMember
        End Get
        Set(value As String)
            SetProperty(NameOf(DisplayMember), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the selected value-property path.
    ''' </summary>
    Public Property ValueMember As String
        Get
            Return _Control.ValueMember
        End Get
        Set(value As String)
            SetProperty(NameOf(ValueMember), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether selected items use distinct text-box colors.
    ''' </summary>
    Public Property HighlightSelectedItem As Boolean
        Get
            Return _Control.HighlightSelectedItem
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(HighlightSelectedItem), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the selected-state background color.
    ''' </summary>
    Public Property SelectedItemBackColor As Color
        Get
            Return _Control.SelectedItemBackColor
        End Get
        Set(value As Color)
            SetProperty(NameOf(SelectedItemBackColor), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the selected-state text color.
    ''' </summary>
    Public Property SelectedItemForeColor As Color
        Get
            Return _Control.SelectedItemForeColor
        End Get
        Set(value As Color)
            SetProperty(NameOf(SelectedItemForeColor), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether entered text starts searches.
    ''' </summary>
    Public Property SearchEnabled As Boolean
        Get
            Return _Control.SearchEnabled
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(SearchEnabled), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the minimum character count required before searching.
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
    ''' Gets or sets the search debounce interval.
    ''' </summary>
    Public Property SearchInterval As Integer
        Get
            Return _Control.SearchInterval
        End Get
        Set(value As Integer)
            SetProperty(NameOf(SearchInterval), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the maximum number of retained results.
    ''' </summary>
    Public Property MaximumResults As Integer
        Get
            Return _Control.MaximumResults
        End Get
        Set(value As Integer)
            SetProperty(NameOf(MaximumResults), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether a single result is selected automatically.
    ''' </summary>
    Public Property AutoSelectSingleResult As Boolean
        Get
            Return _Control.AutoSelectSingleResult
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(AutoSelectSingleResult), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the drop-down width.
    ''' </summary>
    Public Property ResultsViewMode As LookupResultsViewMode
        Get
            Return _Control.ResultsViewMode
        End Get
        Set(value As LookupResultsViewMode)
            SetProperty(NameOf(ResultsViewMode), value)
        End Set
    End Property
    Public Property TemplateItemHeight As Integer
        Get
            Return _Control.TemplateItemHeight
        End Get
        Set(value As Integer)
            SetProperty(NameOf(TemplateItemHeight), value)
        End Set
    End Property
    Public Property DropDownWidth As Integer
        Get
            Return _Control.DropDownWidth
        End Get
        Set(value As Integer)
            SetProperty(NameOf(DropDownWidth), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets the drop-down height.
    ''' </summary>
    Public Property DropDownHeight As Integer
        Get
            Return _Control.DropDownHeight
        End Get
        Set(value As Integer)
            SetProperty(NameOf(DropDownHeight), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether result-column headers are displayed.
    ''' </summary>
    Public Property ShowColumnHeaders As Boolean
        Get
            Return _Control.ShowColumnHeaders
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(ShowColumnHeaders), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether the clear and cancel button is displayed.
    ''' </summary>
    Public Property ShowClearButton As Boolean
        Get
            Return _Control.ShowClearButton
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(ShowClearButton), value)
        End Set
    End Property
    Public Property ClearOnEdit As Boolean
        Get
            Return _Control.ClearOnEdit
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(ClearOnEdit), value)
        End Set
    End Property
    Public Property AutoSelectMatchedResultOnLeave As Boolean
        Get
            Return _Control.AutoSelectMatchedResultOnLeave
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(AutoSelectMatchedResultOnLeave), value)
        End Set
    End Property
    Public Property AllowSelectionLink As Boolean
        Get
            Return _Control.AllowSelectionLink
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(AllowSelectionLink), value)
        End Set
    End Property
    Public Property AutoStretchDropDown As Boolean
        Get
            Return _Control.AutoStretchDropDown
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(AutoStretchDropDown), value)
        End Set
    End Property
    Public Property GridHeadersBold As Boolean
        Get
            Return _Control.GridHeadersBold
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(GridHeadersBold), value)
        End Set
    End Property
    Private Sub SetProperty(PropertyName As String, Value As Object)
        Dim Descriptor As PropertyDescriptor = TypeDescriptor.GetProperties(_Control)(PropertyName)
        If Descriptor Is Nothing Then Throw New InvalidOperationException($"Property '{PropertyName}' was not found.")
        Descriptor.SetValue(_Control, Value)
    End Sub
End Class
