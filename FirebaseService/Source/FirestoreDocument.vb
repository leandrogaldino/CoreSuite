Imports System.Collections.ObjectModel
''' <summary>
''' Represents a Cloud Firestore document and keeps its metadata separate from its stored fields.
''' </summary>
Public NotInheritable Class FirestoreDocument
    ''' <summary>
    ''' Gets the final document identifier.
    ''' </summary>
    Public ReadOnly Property Id As String
    ''' <summary>
    ''' Gets the document path relative to the Firestore document root.
    ''' </summary>
    Public ReadOnly Property Path As String
    ''' <summary>
    ''' Gets the complete Firestore resource name.
    ''' </summary>
    Public ReadOnly Property ResourceName As String
    ''' <summary>
    ''' Gets the document fields.
    ''' </summary>
    Public ReadOnly Property Fields As IReadOnlyDictionary(Of String, Object)
    ''' <summary>
    ''' Gets the time at which the document was created, when returned by Firestore.
    ''' </summary>
    Public ReadOnly Property CreateTimeUtc As DateTime?
    ''' <summary>
    ''' Gets the time at which the document was last updated, when returned by Firestore.
    ''' </summary>
    Public ReadOnly Property UpdateTimeUtc As DateTime?
    ''' <summary>
    ''' Gets a field by name.
    ''' </summary>
    ''' <param name="FieldName">The field name.</param>
    Default Public ReadOnly Property Item(FieldName As String) As Object
        Get
            Return Fields(FieldName)
        End Get
    End Property
    Friend Sub New(Id As String, Path As String, ResourceName As String, Fields As Dictionary(Of String, Object), CreateTimeUtc As DateTime?, UpdateTimeUtc As DateTime?)
        Me.Id = Id
        Me.Path = Path
        Me.ResourceName = ResourceName
        Me.Fields = New ReadOnlyDictionary(Of String, Object)(New Dictionary(Of String, Object)(Fields, StringComparer.Ordinal))
        Me.CreateTimeUtc = CreateTimeUtc
        Me.UpdateTimeUtc = UpdateTimeUtc
    End Sub
End Class
