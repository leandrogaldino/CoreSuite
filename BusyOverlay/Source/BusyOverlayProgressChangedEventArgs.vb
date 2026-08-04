''' <summary>
''' Provides data for the <see cref="BusyOverlay.ProgressChanged"/> event.
''' </summary>
Public NotInheritable Class BusyOverlayProgressChangedEventArgs
    Inherits EventArgs
    Private ReadOnly _Value As Integer
    Private ReadOnly _Percentage As Double
    Private ReadOnly _DetailText As String
    ''' <summary>
    ''' Initializes a new instance of the <see cref="BusyOverlayProgressChangedEventArgs"/> class.
    ''' </summary>
    ''' <param name="Value">The current progress value.</param>
    ''' <param name="Percentage">The normalized progress percentage.</param>
    ''' <param name="DetailText">The current detail text.</param>
    Public Sub New(Value As Integer, Percentage As Double, DetailText As String)
        _Value = Value
        _Percentage = Percentage
        _DetailText = DetailText
    End Sub
    ''' <summary>
    ''' Gets the current progress value.
    ''' </summary>
    Public ReadOnly Property Value As Integer
        Get
            Return _Value
        End Get
    End Property
    ''' <summary>
    ''' Gets the current progress as a value from 0 through 100.
    ''' </summary>
    Public ReadOnly Property Percentage As Double
        Get
            Return _Percentage
        End Get
    End Property
    ''' <summary>
    ''' Gets the detail text supplied with the progress update.
    ''' </summary>
    Public ReadOnly Property DetailText As String
        Get
            Return _DetailText
        End Get
    End Property
End Class
