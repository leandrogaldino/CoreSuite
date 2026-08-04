Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.Windows.Forms.Design

''' <summary>
''' Base class for creating custom property editors for Windows Forms.
''' </summary>
''' <remarks>
''' Provides a reusable implementation of <see cref="UITypeEditor"/> that manages
''' the creation, display and retrieval of custom editing controls for property values.
''' </remarks>
Public MustInherit Class CUITypeEditor
    Inherits UITypeEditor

    Protected IEditorService As IWindowsFormsEditorService
    Protected WithEvents EditControl As Control
    Protected EscapePressed As Boolean
    ''' <summary>
    ''' Gets or sets the caption used when displaying editor messages.
    ''' </summary>
    Protected Friend Property Caption As String = "UITypeEditor"
    ''' <summary>
    ''' Creates the control used to edit the specified property value.
    ''' </summary>
    ''' <param name="PropertyName">
    ''' The name of the property being edited.
    ''' </param>
    ''' <param name="CurrentValue">
    ''' The current value of the property.
    ''' </param>
    ''' <returns>
    ''' A control instance used for editing the property value.
    ''' </returns>
    Protected MustOverride Function GetEditControl(PropertyName As String, CurrentValue As Object) As Control
    ''' <summary>
    ''' Retrieves the edited value from the editing control.
    ''' </summary>
    ''' <param name="EditControl">
    ''' The control containing the edited value.
    ''' </param>
    ''' <param name="PropertyName">
    ''' The name of the property being edited.
    ''' </param>
    ''' <param name="OldValue">
    ''' The original property value before editing.
    ''' </param>
    ''' <returns>
    ''' The updated property value.
    ''' </returns>
    Protected MustOverride Function GetEditedValue(EditControl As Control, PropertyName As String, OldValue As Object) As Object
    ''' <summary>
    ''' Loads the current property value into the editing control.
    ''' </summary>
    ''' <param name="Context">
    ''' Provides information about the property being edited.
    ''' </param>
    ''' <param name="Provider">
    ''' Provides access to editor-related services.
    ''' </param>
    ''' <param name="Value">
    ''' The current property value.
    ''' </param>
    Protected MustOverride Sub LoadValues(Context As ITypeDescriptorContext, Provider As IServiceProvider, Value As Object)
    ''' <summary>
    ''' Determines the editing style supported by the derived editor.
    ''' </summary>
    ''' <param name="context">
    ''' The context of the property being edited.
    ''' </param>
    ''' <returns>
    ''' The editing style supported by the editor.
    ''' </returns>
    Protected MustOverride Function SetEditStyle(context As ITypeDescriptorContext) As UITypeEditorEditStyle

    ''' <summary>
    ''' Gets the editing style used by this editor.
    ''' </summary>
    ''' <param name="context">
    ''' The context of the property being edited.
    ''' </param>
    ''' <returns>
    ''' The <see cref="UITypeEditorEditStyle"/> supported by the editor.
    ''' </returns>
    Public NotOverridable Overrides Function GetEditStyle(context As ITypeDescriptorContext) As UITypeEditorEditStyle
        Return SetEditStyle(context)
    End Function

    ''' <summary>
    ''' Edits the value of a property using the custom editor interface.
    ''' </summary>
    ''' <param name="Context">
    ''' Provides information about the property being edited.
    ''' </param>
    ''' <param name="Provider">
    ''' Provides access to editor services.
    ''' </param>
    ''' <param name="Value">
    ''' The current property value.
    ''' </param>
    ''' <returns>
    ''' The updated property value after editing.
    ''' </returns>
    Public Overrides Function EditValue(Context As ITypeDescriptorContext, Provider As IServiceProvider, Value As Object) As Object
        Try
            If Context IsNot Nothing AndAlso Provider IsNot Nothing Then
                IEditorService = DirectCast(Provider.GetService(GetType(IWindowsFormsEditorService)), IWindowsFormsEditorService)
                If IEditorService IsNot Nothing Then
                    Dim PropName As String = Context.PropertyDescriptor.Name
                    EditControl = Me.GetEditControl(PropName, Value)
                    Me.LoadValues(Context, Provider, Value)
                    If EditControl IsNot Nothing Then
                        EscapePressed = False
                        If TypeOf EditControl Is Form Then
                            IEditorService.ShowDialog(CType(EditControl, Form))
                        Else
                            IEditorService.DropDownControl(EditControl)
                        End If
                        If EscapePressed Then
                            Return Value
                        Else
                            Return GetEditedValue(EditControl, PropName, Value)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return MyBase.EditValue(Context, Provider, Value)
    End Function

    ''' <summary>
    ''' Gets the <see cref="IWindowsFormsEditorService"/> instance associated with this editor.
    ''' </summary>
    ''' <returns>
    ''' The editor service instance, or <see langword="Nothing"/> if unavailable.
    ''' </returns>
    Public Function GetIWindowsFormsEditorService() As IWindowsFormsEditorService
        Return IEditorService
    End Function

    ''' <summary>
    ''' Closes the currently opened drop-down editor window.
    ''' </summary>
    Public Sub CloseDropDownWindow()
        If IEditorService IsNot Nothing Then IEditorService.CloseDropDown()
    End Sub

    ''' <summary>
    ''' Displays an error message using a message box.
    ''' </summary>
    ''' <param name="msg">
    ''' The error message to display.
    ''' </param>
    Protected Friend Sub DisplayError(msg As String)
        MessageBox.Show(msg, Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub

    ''' <summary>
    ''' Represents an item containing a control reference and its display name.
    ''' </summary>
    Protected Friend Class ListItem
        ''' <summary>
        ''' Gets or sets the control associated with this item.
        ''' </summary>
        Public Property [Control] As Control
        ''' <summary>
        ''' Gets or sets the display name of this list item.
        ''' </summary>
        Public Property Name As String
        ''' <summary>
        ''' Initializes a new instance of the <see cref="ListItem"/> class.
        ''' </summary>
        ''' <param name="c">The control associated with this item.</param>
        Public Sub New(c As Control)
            [Control] = c
            Name = c.Name
        End Sub
        ''' <summary>
        ''' Returns the display name of the list item.
        ''' </summary>
        ''' <returns>The name of the item.</returns>
        Public Overrides Function ToString() As String
            Return Name
        End Function
    End Class
End Class