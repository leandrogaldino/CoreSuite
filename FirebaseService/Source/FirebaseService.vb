''' <summary>
''' Central service for accessing Firebase, responsible for initializing and exposing
''' the Authentication, Firestore, and Storage modules from a single shared client.
''' </summary>
''' <remarks>
''' This class acts as a *facade* for the main Firebase services,
''' ensuring that all of them use the same <see cref="FirebaseClient"/> instance.
''' 
''' It can be initialized through the constructor or later via the <see cref="Initialize"/> method.
''' </remarks>
Public Class FirebaseService
    Private _Client As FirebaseClient
    Private _Auth As FirebaseAuth
    Private _Firestore As FirebaseFirestore
    Private _Storage As FirebaseStorage


    ''' <summary>
    ''' Firebase authentication service.
    ''' </summary>
    ''' <value>
    ''' An instance of <see cref="FirebaseAuth"/> associated with the current client.
    ''' </value>
    Public ReadOnly Property Auth As FirebaseAuth
        Get
            Return _Auth
        End Get
    End Property

    ''' <summary>
    ''' Firestore Database access service.
    ''' </summary>
    ''' <value>
    ''' An instance of <see cref="FirebaseFirestore"/> associated with the current client.
    ''' </value>
    Public ReadOnly Property Firestore As FirebaseFirestore
        Get
            Return _Firestore
        End Get
    End Property

    ''' <summary>
    ''' Firebase Storage access service.
    ''' </summary>
    ''' <value>
    ''' An instance of <see cref="FirebaseStorage"/> associated with the current client.
    ''' </value>
    Public ReadOnly Property Storage As FirebaseStorage
        Get
            Return _Storage
        End Get
    End Property


    Friend ReadOnly Property Client As FirebaseClient
        Get
            Return _Client
        End Get
    End Property
    ''' <summary>
    ''' Creates a FirebaseService instance without initialization.
    ''' Calling Initialize is required before using any service.
    ''' </summary>
    Public Sub New()
    End Sub

    ''' <summary>
    ''' Initializes a new instance of <see cref="FirebaseService"/> with the specified credentials.
    ''' </summary>
    ''' <param name="ApiKey">
    ''' Firebase project API key.
    ''' </param>
    ''' <param name="ProjectID">
    ''' Firebase project identifier.
    ''' </param>
    ''' <param name="StorageBucket">
    ''' Default Firebase Storage bucket name.
    ''' </param>
    ''' <remarks>
    ''' The constructor internally creates the <see cref="FirebaseClient"/> and initializes
    ''' all dependent services (Auth, Firestore, and Storage).
    ''' </remarks>
    Public Sub New(ApiKey As String, ProjectID As String, StorageBucket As String)
        _Client = New FirebaseClient(ApiKey, ProjectID, StorageBucket)
        _Auth = New FirebaseAuth(Client)
        _Firestore = New FirebaseFirestore(Client)
        _Storage = New FirebaseStorage(Client)
    End Sub
    ''' <summary>
    ''' Initializes or reinitializes the Firebase service with new credentials.
    ''' </summary>
    ''' <param name="ApiKey">
    ''' Firebase project API key.
    ''' </param>
    ''' <param name="ProjectID">
    ''' Firebase project identifier.
    ''' </param>
    ''' <param name="StorageBucket">
    ''' Default Firebase Storage bucket name.
    ''' </param>
    ''' <remarks>
    ''' This method is useful when the class instance is created without initial configuration
    ''' or when it is necessary to dynamically switch the Firebase project in use.
    ''' </remarks>
    Public Sub Initialize(ApiKey As String, ProjectID As String, StorageBucket As String)
        _Client = New FirebaseClient(ApiKey, ProjectID, StorageBucket)
        _Auth = New FirebaseAuth(Client)
        _Firestore = New FirebaseFirestore(Client)
        _Storage = New FirebaseStorage(Client)
    End Sub
End Class
