''' <summary>
''' Provides progress information for a Firebase Storage transfer.
''' </summary>
Public NotInheritable Class FirebaseTransferProgressChangedEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Gets the number of bytes transferred so far.
    ''' </summary>
    Public ReadOnly Property BytesTransferred As Long
    ''' <summary>
    ''' Gets the total number of bytes when known.
    ''' </summary>
    Public ReadOnly Property TotalBytes As Long?
    ''' <summary>
    ''' Gets the completed percentage when the total size is known.
    ''' </summary>
    Public ReadOnly Property Percentage As Double?
    ''' <summary>
    ''' Gets a value indicating whether the transfer has completed successfully.
    ''' </summary>
    Public ReadOnly Property IsCompleted As Boolean
    Friend Sub New(BytesTransferred As Long, TotalBytes As Long?, IsCompleted As Boolean)
        Me.BytesTransferred = BytesTransferred
        Me.TotalBytes = TotalBytes
        Me.IsCompleted = IsCompleted
        If IsCompleted Then
            Percentage = 100
        ElseIf TotalBytes.HasValue AndAlso TotalBytes.Value > 0 Then
            Percentage = Math.Min(100, BytesTransferred * 100.0R / TotalBytes.Value)
        End If
    End Sub
End Class
