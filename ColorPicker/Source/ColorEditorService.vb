Imports System.Windows.Forms.Design
''' <summary>
''' Provides communication between the <see cref="ColorEditor"/> and the hosting <see cref="ColorPicker"/>.
''' Implements <see cref="IWindowsFormsEditorService"/> to host the embedded color editor.
''' </summary>
Friend Class ColorEditorService
    Implements IServiceProvider, IWindowsFormsEditorService

    ''' <summary>
    ''' Indicates whether the embedded editor is currently being closed.
    ''' </summary>
    Private CloseEditor As Boolean
    ''' <summary>
    ''' Occurs when the embedded color editor becomes available or is released.
    ''' </summary>
    Public Event ColorUIAvailable As EventHandler(Of EditorServiceEventArgs)
    ''' <summary>
    ''' Occurs when the selected color changes.
    ''' </summary>
    Public Event ColorChanged As EventHandler
    ''' <summary>
    ''' Marks the embedded editor for closure.
    ''' </summary>
    Public Sub CloseDropDownInternal()
        CloseEditor = True
    End Sub
    ''' <summary>
    ''' Returns the requested service if it is supported.
    ''' </summary>
    ''' <param name="ServiceType">The type of service being requested.</param>
    ''' <returns>
    ''' The requested service instance, or <see langword="Nothing"/> if the service is not available.
    ''' </returns>
    Private Function IServiceProvider_GetService(ServiceType As Type) As Object Implements IServiceProvider.GetService
        If ServiceType = GetType(IWindowsFormsEditorService) Then
            Return Me
        End If
        Return Nothing
    End Function
    ''' <summary>
    ''' Closes the hosted drop-down editor.
    ''' </summary>
    Private Sub IWindowsFormsEditorService_CloseDropDown() Implements IWindowsFormsEditorService.CloseDropDown
        If Not CloseEditor Then
            RaiseEvent ColorChanged(Me, EventArgs.Empty)
        Else
            RaiseEvent ColorUIAvailable(Me, New EditorServiceEventArgs(Nothing))
        End If
    End Sub
    ''' <summary>
    ''' Displays the specified control as a drop-down editor.
    ''' </summary>
    ''' <param name="control">The control to display.</param>
    Private Sub IWindowsFormsEditorService_DropDownControl(control As Control) Implements IWindowsFormsEditorService.DropDownControl
        CloseEditor = (control Is Nothing)
        RaiseEvent ColorUIAvailable(Me, New EditorServiceEventArgs(control))
    End Sub
    ''' <summary>
    ''' Displays the specified dialog box.
    ''' </summary>
    ''' <param name="dialog">The dialog box to display.</param>
    ''' <returns>
    ''' The value returned by the dialog box.
    ''' </returns>
    ''' <exception cref="NotImplementedException">
    ''' This method is not implemented.
    ''' </exception>
    Private Function IWindowsFormsEditorService_ShowDialog(dialog As Form) As DialogResult Implements IWindowsFormsEditorService.ShowDialog
        Throw New Exception("The method or operation is not implemented.")
    End Function
End Class

