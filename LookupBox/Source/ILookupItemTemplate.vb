''' <summary>
''' Optional contract for UserControls displayed in LookupBox template mode.
''' Implement it on the designer-created UserControl to receive the current result
''' each time the view is recycled during scrolling.
''' </summary>
Public Interface ILookupItemTemplate
    Sub BindItem(Item As Object)
End Interface

''' <summary>
''' Optional contract allowing a template to customize its appearance when the
''' keyboard/mouse highlight changes. The LookupBox also paints a selection stripe.
''' </summary>
Public Interface ILookupItemSelectionAware
    Sub SetHighlighted(IsHighlighted As Boolean)
End Interface
