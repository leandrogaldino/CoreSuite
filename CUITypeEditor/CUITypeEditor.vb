Imports System.ComponentModel
Imports System.Drawing.Design
Imports System.Windows.Forms
Imports System.Windows.Forms.Design

''' <summary>
''' Base class for creating custom property editors for Windows Forms.
''' Inherits from <see cref="UITypeEditor"/> and provides methods to manage
''' editing of property values in a UI.
''' </summary>
Public MustInherit Class CUITypeEditor
    Inherits UITypeEditor

    Protected IEditorService As IWindowsFormsEditorService
    Protected WithEvents EditControl As Control
    Protected EscapePressed As Boolean
    Protected Friend Property Caption As String = "UITypeEditor"

    Protected MustOverride Function GetEditControl(PropertyName As String, CurrentValue As Object) As Control
    Protected MustOverride Function GetEditedValue(EditControl As Control, PropertyName As String, OldValue As Object) As Object
    Protected MustOverride Sub LoadValues(Context As ITypeDescriptorContext, Provider As IServiceProvider, Value As Object)
    Protected MustOverride Function SetEditStyle(context As ITypeDescriptorContext) As UITypeEditorEditStyle

    ''' <summary>
    ''' Returns the editing style used by the editor.
    ''' </summary>
    ''' <param name="context">The context of the property being edited.</param>
    ''' <returns>The <see cref="UITypeEditorEditStyle"/> for this editor.</returns>
    Public NotOverridable Overrides Function GetEditStyle(context As ITypeDescriptorContext) As UITypeEditorEditStyle
        Return SetEditStyle(context)
    End Function

    ''' <summary>
    ''' Edits the value of the property using the editor.
    ''' </summary>
    ''' <param name="Context">The type descriptor context.</param>
    ''' <param name="Provider">Service provider used to obtain editor services.</param>
    ''' <param name="Value">The current value of the property.</param>
    ''' <returns>The new value of the property.</returns>
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
    ''' Gets the <see cref="IWindowsFormsEditorService"/> used by this editor.
    ''' </summary>
    ''' <returns>The editor service instance.</returns>
    Public Function GetIWindowsFormsEditorService() As IWindowsFormsEditorService
        Return IEditorService
    End Function

    ''' <summary>
    ''' Closes the drop-down window if it is open.
    ''' </summary>
    Public Sub CloseDropDownWindow()
        If IEditorService IsNot Nothing Then IEditorService.CloseDropDown()
    End Sub

    ''' <summary>
    ''' Displays an error message in a message box.
    ''' </summary>
    ''' <param name="msg">The message to display.</param>
    Protected Friend Sub DisplayError(msg As String)
        MessageBox.Show(msg, Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub

    ''' <summary>
    ''' Represents an item in a list for use in the editor UI.
    ''' </summary>
    Protected Friend Class ListItem
        ''' <summary>
        ''' Gets or sets the control associated with this list item.
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
