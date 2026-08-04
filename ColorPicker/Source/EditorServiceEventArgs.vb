''' <summary>
''' Provides event data for events related to the ColorPicker editor service.
''' </summary>
''' <remarks>
''' This class exposes the Windows Forms control used as the color selection
''' interface by the editor service.
''' </remarks>
Friend Class EditorServiceEventArgs
    Inherits EventArgs
    Private ReadOnly _ColorUI As Control
    ''' <summary>
    ''' Initializes a new instance of the <see cref="EditorServiceEventArgs"/> class
    ''' with the specified color selection interface.
    ''' </summary>
    ''' <param name="ColorUI">
    ''' The Windows Forms control used as the color selection interface.
    ''' </param>
    Public Sub New(ColorUI As Control)
        _ColorUI = ColorUI
    End Sub
    ''' <summary>
    ''' Gets the Windows Forms control used as the color selection interface.
    ''' </summary>
    ''' <value>
    ''' The <see cref="Control"/> that represents the color selection interface.
    ''' </value>
    Public ReadOnly Property ColorUI As Control
        Get
            Return _ColorUI
        End Get
    End Property
End Class