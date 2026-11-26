Imports System.Windows.Forms


''' <summary>
''' Provides data for an action executed by a <see cref="TextBoxActionPanel"/>.
''' </summary>
Public NotInheritable Class TextBoxActionClickEventArgs
    Inherits EventArgs
    Private ReadOnly _Source As TextBoxActionPanel
    Private ReadOnly _TargetControl As System.Windows.Forms.TextBoxBase
    Private ReadOnly _Action As TextBoxAction
    ''' <summary>
    ''' Initializes a new instance of the <see cref="TextBoxActionClickEventArgs"/> class.
    ''' </summary>
    ''' <param name="Source">The component that executed the action.</param>
    ''' <param name="TargetControl">The text box associated with the component.</param>
    ''' <param name="Action">The action that was executed.</param>
    Public Sub New(Source As TextBoxActionPanel, TargetControl As System.Windows.Forms.TextBoxBase, Action As TextBoxAction)
        If Source Is Nothing Then Throw New ArgumentNullException("Source")
        If TargetControl Is Nothing Then Throw New ArgumentNullException("TargetControl")
        If Action Is Nothing Then Throw New ArgumentNullException("Action")
        _Source = Source
        _TargetControl = TargetControl
        _Action = Action
    End Sub
    ''' <summary>
    ''' Gets the component that executed the action.
    ''' </summary>
    ''' <value>The originating <see cref="TextBoxActionPanel"/>.</value>
    Public ReadOnly Property Source As TextBoxActionPanel
        Get
            Return _Source
        End Get
    End Property
    ''' <summary>
    ''' Gets the text box associated with the action panel.
    ''' </summary>
    ''' <value>The current target control.</value>
    Public ReadOnly Property TargetControl As System.Windows.Forms.TextBoxBase
        Get
            Return _TargetControl
        End Get
    End Property
    ''' <summary>
    ''' Gets the action that was executed.
    ''' </summary>
    ''' <value>The selected <see cref="TextBoxAction"/>.</value>
    Public ReadOnly Property Action As TextBoxAction
        Get
            Return _Action
        End Get
    End Property
End Class
