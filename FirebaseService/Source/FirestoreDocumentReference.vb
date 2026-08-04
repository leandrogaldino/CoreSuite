''' <summary>
''' Represents a Cloud Firestore document reference value.
''' </summary>
Public NotInheritable Class FirestoreDocumentReference
    ''' <summary>
    ''' Gets the complete Firestore resource name of the referenced document.
    ''' </summary>
    Public ReadOnly Property ResourceName As String
    ''' <summary>
    ''' Initializes a reference from a complete Firestore document resource name.
    ''' </summary>
    ''' <param name="ResourceName">A resource name in the form <c>projects/{project}/databases/{database}/documents/{path}</c>.</param>
    Public Sub New(ResourceName As String)
        If String.IsNullOrWhiteSpace(ResourceName) Then Throw New ArgumentException("The resource name cannot be empty.", NameOf(ResourceName))
        If Not ResourceName.StartsWith("projects/", StringComparison.Ordinal) OrElse Not ResourceName.Contains("/documents/", StringComparison.Ordinal) Then Throw New ArgumentException("The value is not a complete Firestore document resource name.", NameOf(ResourceName))
        Me.ResourceName = ResourceName
    End Sub
    ''' <summary>
    ''' Returns the complete Firestore resource name.
    ''' </summary>
    Public Overrides Function ToString() As String
        Return ResourceName
    End Function
End Class
