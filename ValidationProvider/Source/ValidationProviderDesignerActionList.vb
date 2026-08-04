Imports System.ComponentModel
Imports Microsoft.DotNet.DesignTools.Designers.Actions
''' <summary>
''' Provides smart-tag actions and design-time property access for the <see cref="ValidationProvider"/> component.
''' </summary>
Public Class ValidationProviderDesignerActionList
    Inherits DesignerActionList
    Private ReadOnly _Provider As ValidationProvider
    ''' <summary>
    ''' Initializes a new instance of the <see cref="ValidationProviderDesignerActionList"/> class.
    ''' </summary>
    ''' <param name="Designer">The designer associated with the validation provider.</param>
    Public Sub New(Designer As ValidationProviderDesigner)
        MyBase.New(Designer.Component)
        _Provider = CType(Designer.Component, ValidationProvider)
    End Sub
    ''' <summary>
    ''' Gets the collection of smart-tag items displayed in the Windows Forms designer.
    ''' </summary>
    ''' <returns>A collection containing the most frequently used validation settings.</returns>
    Public Overrides Function GetSortedActionItems() As DesignerActionItemCollection
        Return New DesignerActionItemCollection From {
            New DesignerActionHeaderItem("Automatic validation"),
            New DesignerActionPropertyItem(NameOf(AutomaticValidation), "AutomaticValidation", "Automatic validation", "Validates configured controls during their Validating event."),
            New DesignerActionPropertyItem(NameOf(CancelValidationOnError), "CancelValidationOnError", "Automatic validation", "Cancels the Validating event when automatic validation fails."),
            New DesignerActionPropertyItem(NameOf(ClearErrorOnValueChanged), "ClearErrorOnValueChanged", "Automatic validation", "Clears stale feedback while a value is edited."),
            New DesignerActionHeaderItem("Batch validation"),
            New DesignerActionPropertyItem(NameOf(FocusFirstInvalidControl), "FocusFirstInvalidControl", "Batch validation", "Focuses the first invalid control after Validate or ValidateGroup."),
            New DesignerActionPropertyItem(NameOf(ValidateDisabledControls), "ValidateDisabledControls", "Batch validation", "Includes disabled controls in validation."),
            New DesignerActionPropertyItem(NameOf(ValidateHiddenControls), "ValidateHiddenControls", "Batch validation", "Includes hidden controls in validation."),
            New DesignerActionHeaderItem("Comparison"),
            New DesignerActionPropertyItem(NameOf(TrimTextValues), "TrimTextValues", "Comparison", "Ignores surrounding white space in length and comparison rules."),
            New DesignerActionPropertyItem(NameOf(CaseSensitiveComparison), "CaseSensitiveComparison", "Comparison", "Uses case-sensitive CompareWith rules.")
        }
    End Function
    ''' <summary>
    ''' Gets or sets whether controls validate during their Validating event.
    ''' </summary>
    Public Property AutomaticValidation As Boolean
        Get
            Return _Provider.AutomaticValidation
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(AutomaticValidation), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether invalid automatic validation cancels the Validating event.
    ''' </summary>
    Public Property CancelValidationOnError As Boolean
        Get
            Return _Provider.CancelValidationOnError
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(CancelValidationOnError), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether an existing error is removed as the value changes.
    ''' </summary>
    Public Property ClearErrorOnValueChanged As Boolean
        Get
            Return _Provider.ClearErrorOnValueChanged
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(ClearErrorOnValueChanged), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether batch validation focuses the first invalid control.
    ''' </summary>
    Public Property FocusFirstInvalidControl As Boolean
        Get
            Return _Provider.FocusFirstInvalidControl
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(FocusFirstInvalidControl), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether disabled controls participate in validation.
    ''' </summary>
    Public Property ValidateDisabledControls As Boolean
        Get
            Return _Provider.ValidateDisabledControls
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(ValidateDisabledControls), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether hidden controls participate in validation.
    ''' </summary>
    Public Property ValidateHiddenControls As Boolean
        Get
            Return _Provider.ValidateHiddenControls
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(ValidateHiddenControls), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether surrounding white space is removed before length and comparison rules.
    ''' </summary>
    Public Property TrimTextValues As Boolean
        Get
            Return _Provider.TrimTextValues
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(TrimTextValues), value)
        End Set
    End Property
    ''' <summary>
    ''' Gets or sets whether comparison rules are case-sensitive.
    ''' </summary>
    Public Property CaseSensitiveComparison As Boolean
        Get
            Return _Provider.CaseSensitiveComparison
        End Get
        Set(value As Boolean)
            SetProperty(NameOf(CaseSensitiveComparison), value)
        End Set
    End Property
    Private Sub SetProperty(PropertyName As String, Value As Object)
        Dim Descriptor As PropertyDescriptor = TypeDescriptor.GetProperties(_Provider)(PropertyName)
        If Descriptor Is Nothing Then Throw New InvalidOperationException($"Property '{PropertyName}' was not found.")
        Descriptor.SetValue(_Provider, Value)
    End Sub
End Class
