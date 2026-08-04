Imports System.Threading
''' <summary>
''' Contains the configuration used by <see cref="FirebaseService"/>.
''' </summary>
Public NotInheritable Class FirebaseOptions
    ''' <summary>
    ''' Gets or sets the Firebase Web API key used by Firebase Authentication.
    ''' </summary>
    Public Property ApiKey As String
    ''' <summary>
    ''' Gets or sets the Firebase project identifier.
    ''' </summary>
    Public Property ProjectId As String
    ''' <summary>
    ''' Gets or sets the Cloud Storage bucket name, with or without the <c>gs://</c> prefix.
    ''' </summary>
    Public Property StorageBucket As String
    ''' <summary>
    ''' Gets or sets the Cloud Firestore database identifier.
    ''' </summary>
    ''' <value>The database identifier. The default value is <c>(default)</c>.</value>
    Public Property DatabaseId As String = "(default)"
    ''' <summary>
    ''' Gets or sets the timeout applied to Authentication and Firestore operations.
    ''' </summary>
    ''' <value>The operation timeout. The default value is 30 seconds.</value>
    Public Property RequestTimeout As TimeSpan = TimeSpan.FromSeconds(30)
    ''' <summary>
    ''' Gets or sets the timeout applied to Storage uploads and downloads.
    ''' </summary>
    ''' <value>The transfer timeout. The default value is <see cref="Timeout.InfiniteTimeSpan"/>.</value>
    Public Property TransferTimeout As TimeSpan = Timeout.InfiniteTimeSpan
    ''' <summary>
    ''' Initializes an empty configuration object for use with configuration binding or object initializers.
    ''' </summary>
    Public Sub New()
    End Sub
    ''' <summary>
    ''' Initializes a configuration object with the required Firebase project values.
    ''' </summary>
    ''' <param name="ApiKey">The Firebase Web API key.</param>
    ''' <param name="ProjectId">The Firebase project identifier.</param>
    ''' <param name="StorageBucket">The Cloud Storage bucket name.</param>
    Public Sub New(ApiKey As String, ProjectId As String, StorageBucket As String)
        Me.ApiKey = ApiKey
        Me.ProjectId = ProjectId
        Me.StorageBucket = StorageBucket
    End Sub
    Friend Function CreateValidatedCopy() As FirebaseOptions
        Dim ValidatedApiKey As String = RequireValue(ApiKey, NameOf(ApiKey))
        Dim ValidatedProjectId As String = RequireValue(ProjectId, NameOf(ProjectId))
        Dim ValidatedStorageBucket As String = NormalizeStorageBucket(StorageBucket)
        Dim ValidatedDatabaseId As String = RequireValue(DatabaseId, NameOf(DatabaseId))
        ValidateTimeout(RequestTimeout, NameOf(RequestTimeout))
        ValidateTimeout(TransferTimeout, NameOf(TransferTimeout))
        Return New FirebaseOptions(ValidatedApiKey, ValidatedProjectId, ValidatedStorageBucket) With {.DatabaseId = ValidatedDatabaseId, .RequestTimeout = RequestTimeout, .TransferTimeout = TransferTimeout}
    End Function
    Friend Function Copy() As FirebaseOptions
        Return New FirebaseOptions(ApiKey, ProjectId, StorageBucket) With {.DatabaseId = DatabaseId, .RequestTimeout = RequestTimeout, .TransferTimeout = TransferTimeout}
    End Function
    Private Shared Function RequireValue(Value As String, ParameterName As String) As String
        If String.IsNullOrWhiteSpace(Value) Then Throw New ArgumentException("The value cannot be empty.", ParameterName)
        Return Value.Trim()
    End Function
    Private Shared Function NormalizeStorageBucket(Value As String) As String
        Dim Bucket As String = RequireValue(Value, NameOf(StorageBucket))
        If Bucket.StartsWith("gs://", StringComparison.OrdinalIgnoreCase) Then Bucket = Bucket.Substring(5)
        Bucket = Bucket.TrimEnd("/"c)
        If String.IsNullOrWhiteSpace(Bucket) Then Throw New ArgumentException("The storage bucket cannot be empty.", NameOf(Value))
        Return Bucket
    End Function
    Private Shared Sub ValidateTimeout(Value As TimeSpan, ParameterName As String)
        If Value <> Timeout.InfiniteTimeSpan AndAlso Value <= TimeSpan.Zero Then Throw New ArgumentOutOfRangeException(ParameterName, "The timeout must be positive or Timeout.InfiniteTimeSpan.")
    End Sub
End Class
