Friend NotInheritable Class FirebasePathHelper
    Private Sub New()
    End Sub
    Friend Shared Function NormalizeCollectionPath(Value As String, ParameterName As String) As String
        Return NormalizeFirestorePath(Value, ParameterName, True)
    End Function
    Friend Shared Function NormalizeDocumentPath(Value As String, ParameterName As String) As String
        Return NormalizeFirestorePath(Value, ParameterName, False)
    End Function
    Friend Shared Function EncodeFirestorePath(Value As String) As String
        Return String.Join("/", Value.Split("/"c).Select(Function(Segment) Uri.EscapeDataString(Segment)))
    End Function
    Friend Shared Function NormalizeDocumentId(Value As String, ParameterName As String) As String
        If String.IsNullOrWhiteSpace(Value) Then Throw New ArgumentException("The document identifier cannot be empty.", ParameterName)
        Dim DocumentId As String = Value.Trim()
        If DocumentId.Contains("/"c) OrElse DocumentId.Contains("\"c) Then Throw New ArgumentException("The document identifier cannot contain path separators.", ParameterName)
        Return DocumentId
    End Function
    Friend Shared Function NormalizeStoragePath(Value As String, ParameterName As String) As String
        If String.IsNullOrWhiteSpace(Value) Then Throw New ArgumentException("The storage path cannot be empty.", ParameterName)
        Dim StoragePath As String = Value.Trim().TrimStart("/"c)
        If String.IsNullOrWhiteSpace(StoragePath) Then Throw New ArgumentException("The storage path must identify an object and cannot point to the bucket root.", ParameterName)
        Return StoragePath
    End Function
    Private Shared Function NormalizeFirestorePath(Value As String, ParameterName As String, IsCollection As Boolean) As String
        If String.IsNullOrWhiteSpace(Value) Then Throw New ArgumentException("The Firestore path cannot be empty.", ParameterName)
        Dim FirestorePath As String = Value.Trim().Trim("/"c)
        If FirestorePath.Contains("\"c) Then Throw New ArgumentException("Firestore paths must use forward slashes.", ParameterName)
        Dim Segments As String() = FirestorePath.Split("/"c)
        If Segments.Any(Function(Segment) String.IsNullOrWhiteSpace(Segment)) Then Throw New ArgumentException("Firestore paths cannot contain empty segments.", ParameterName)
        Dim HasCollectionSegmentCount As Boolean = Segments.Length Mod 2 = 1
        If IsCollection AndAlso Not HasCollectionSegmentCount Then Throw New ArgumentException("A collection path must contain an odd number of segments.", ParameterName)
        If Not IsCollection AndAlso HasCollectionSegmentCount Then Throw New ArgumentException("A document path must contain an even number of segments.", ParameterName)
        Return String.Join("/", Segments)
    End Function
End Class
