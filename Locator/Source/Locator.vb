''' <summary>
''' Simple service provider and dependency injection container.
''' Allows registering instances as singletons, lazy singletons or via factories,
''' and retrieving instances by type.
''' </summary>
Public Class Locator
    Private Enum RegistrationType
        Singleton
        LazySingleton
        Factory
    End Enum

    Private Class Registration
        Public Property Type As RegistrationType
        Public Property Instance As Object
        Public Property Factory As Func(Of Object)
        Public ReadOnly Property SyncRoot As New Object
    End Class

    Private Shared ReadOnly _Registrations As New Dictionary(Of (Type, String), Registration)
    Private Shared ReadOnly _RegistrationsSyncRoot As New Object

    ''' <summary>
    ''' Registers an instance as a singleton for the specified type.
    ''' </summary>
    ''' <typeparam name="T">The type of service to register.</typeparam>
    ''' <param name="Instance">The instance that will always be returned when the service is requested.</param>
    ''' <param name="Key">Optional key to differentiate multiple registrations of the same type.</param>
    Public Shared Sub RegisterSingleton(Of T)(Instance As T, Optional Key As String = "")
        Dim InstanceType = GetType(T)
        SyncLock _RegistrationsSyncRoot
            _Registrations((InstanceType, Key)) = New Registration With {
                .Type = RegistrationType.Singleton,
                .Instance = Instance
            }
        End SyncLock
    End Sub

    ''' <summary>
    ''' Registers a service as a lazy singleton.
    ''' The service is created only when it is requested for the first time
    ''' and the same instance is returned for all subsequent requests.
    ''' </summary>
    ''' <typeparam name="T">The type of service to register.</typeparam>
    ''' <param name="Factory">A function that creates the singleton instance when it is first requested.</param>
    ''' <param name="Key">Optional key to differentiate multiple registrations of the same type.</param>
    Public Shared Sub RegisterLazySingleton(Of T)(Factory As Func(Of T), Optional Key As String = "")
        ArgumentNullException.ThrowIfNull(Factory)
        Dim InstanceType = GetType(T)
        SyncLock _RegistrationsSyncRoot
            _Registrations((InstanceType, Key)) = New Registration With {
                .Type = RegistrationType.LazySingleton,
                .Factory = Function() Factory()
            }
        End SyncLock
    End Sub

    ''' <summary>
    ''' Registers a service via a factory, creating a new instance each time it is requested.
    ''' </summary>
    ''' <typeparam name="T">The type of service to register.</typeparam>
    ''' <param name="Factory">A function that creates a new instance of the service.</param>
    ''' <param name="Key">Optional key to differentiate multiple registrations of the same type.</param>
    Public Shared Sub RegisterFactory(Of T)(Factory As Func(Of T), Optional Key As String = "")
        ArgumentNullException.ThrowIfNull(Factory)
        Dim InstanceType = GetType(T)
        SyncLock _RegistrationsSyncRoot
            _Registrations((InstanceType, Key)) = New Registration With {
                .Type = RegistrationType.Factory,
                .Factory = Function() Factory()
            }
        End SyncLock
    End Sub

    ''' <summary>
    ''' Retrieves a registered instance of the specified type.
    ''' </summary>
    ''' <typeparam name="T">The type of service to retrieve.</typeparam>
    ''' <param name="Key">Optional key used during registration.</param>
    ''' <returns>An instance of the requested type.</returns>
    ''' <exception cref="ServiceNotRegisteredException">
    ''' Thrown if the requested service has not been registered.
    ''' </exception>
    Public Shared Function GetInstance(Of T)(Optional Key As String = "") As T
        Dim InstanceType = GetType(T)
        Dim RegistrationKey = (InstanceType, Key)
        Dim Registration As Registration = Nothing
        SyncLock _RegistrationsSyncRoot
            _Registrations.TryGetValue(RegistrationKey, Registration)
        End SyncLock
        If Registration Is Nothing Then
            Throw New ServiceNotRegisteredException($"Service '{InstanceType.FullName}'{If(String.IsNullOrWhiteSpace(Key), "", $" with key '{Key}'")} is not registered.")
        End If
        Select Case Registration.Type
            Case RegistrationType.Singleton
                Return CType(Registration.Instance, T)
            Case RegistrationType.LazySingleton
                SyncLock Registration.SyncRoot
                    If Registration.Instance Is Nothing Then
                        Registration.Instance = Registration.Factory.Invoke()
                    End If
                    Return CType(Registration.Instance, T)
                End SyncLock
            Case RegistrationType.Factory
                Return CType(Registration.Factory.Invoke(), T)
            Case Else
                Throw New InvalidOperationException($"Unknown registration type '{Registration.Type}'.")
        End Select
    End Function
End Class