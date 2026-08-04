''' <summary>
''' Represents an error caused by invalid source code during runtime compilation.
''' </summary>
Public NotInheritable Class CodeCompilationException
    Inherits Exception
    Private ReadOnly _Diagnostics As IReadOnlyList(Of String)
    ''' <summary>
    ''' Initializes a new instance of the <see cref="CodeCompilationException"/> class.
    ''' </summary>
    ''' <param name="Diagnostics">The compiler diagnostics that describe the compilation errors.</param>
    Friend Sub New(Diagnostics As IEnumerable(Of String))
        Me.New(Diagnostics.ToArray())
    End Sub
    Private Sub New(Diagnostics As String())
        MyBase.New(CreateMessage(Diagnostics))
        _Diagnostics = Diagnostics
    End Sub
    ''' <summary>
    ''' Gets the compiler diagnostics that caused the compilation to fail.
    ''' </summary>
    ''' <value>
    ''' A read-only collection containing the compilation error messages.
    ''' </value>
    Public ReadOnly Property Diagnostics As IReadOnlyList(Of String)
        Get
            Return _Diagnostics
        End Get
    End Property
    Private Shared Function CreateMessage(Diagnostics As IEnumerable(Of String)) As String
        Dim DiagnosticList As String() = Diagnostics.ToArray()
        If DiagnosticList.Length = 0 Then Return "The source code could not be compiled."
        Return "The source code contains compilation errors:" & Environment.NewLine & String.Join(Environment.NewLine, DiagnosticList)
    End Function
End Class