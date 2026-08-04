''' <summary>
''' Represents the non-selectable button used internally by a <see cref="TextBoxActionPanel"/> popup.
''' </summary>
Friend NotInheritable Class TextBoxActionButton
    Inherits NoFocusCueButton
    Private ReadOnly _Action As TextBoxAction
    Friend Event ActionClick As EventHandler
    ''' <summary>
    ''' Initializes a new button for the specified action.
    ''' </summary>
    ''' <param name="Action">The action represented by the button.</param>
    Public Sub New(Action As TextBoxAction)
        ArgumentNullException.ThrowIfNull(Action)
        _Action = Action
        SetStyle(ControlStyles.Selectable, False)
        TabStop = False
        Text = String.Empty
        Cursor = Cursors.Hand
        FlatStyle = FlatStyle.Flat
        UseVisualStyleBackColor = False
        ImageAlign = ContentAlignment.MiddleCenter
        TextImageRelation = TextImageRelation.Overlay
    End Sub
    ''' <summary>
    ''' Gets the action represented by this button.
    ''' </summary>
    ''' <value>The associated action.</value>
    Friend ReadOnly Property Action As TextBoxAction
        Get
            Return _Action
        End Get
    End Property
    ''' <summary>
    ''' Raises the internal action-click notification without selecting the button.
    ''' </summary>
    ''' <param name="E">The event data.</param>
    Protected Overrides Sub OnClick(E As EventArgs)
        MyBase.OnClick(E)
        RaiseEvent ActionClick(Me, EventArgs.Empty)
    End Sub
End Class
