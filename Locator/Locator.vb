''' <summary>
''' Simple service provider and dependency injection container.
''' Allows registering instances as singletons or via factories, 
''' and retrieving instances by type.
''' </summary>
Public Class Locator

    Private Enum RegistrationType
        Singleton
        Factory
    End Enum

    Private Class Registration
        Public Property Type As RegistrationType
        Public Property Instance As Object
        Public Property Factory As Func(Of Object)
    End Class

    Private Shared ReadOnly _Registrations As New Dictionary(Of (Type, String), Registration)

    ''' <summary>
    ''' Registers an instance as a singleton for the specified type.
    ''' </summary>
    ''' <typeparam name="T">The type of service to register.</typeparam>
    ''' <param name="Instance">The instance that will always be returned when the service is requested.</param>
    ''' <param name="Key">Optional key to differentiate multiple registrations of the same type.</param>
    Public Shared Sub RegisterSingleton(Of T)(Instance As T, Optional Key As String = "")
        Dim InstanceType = GetType(T)
        _Registrations((InstanceType, Key)) = New Registration With {
            .Type = RegistrationType.Singleton,
            .Instance = Instance
        }
    End Sub

    ''' <summary>
    ''' Registers a service via a factory, creating a new instance each time it is requested.
    ''' </summary>
    ''' <typeparam name="T">The type of service to register.</typeparam>
    ''' <param name="Factory">A function that creates a new instance of the service.</param>
    ''' <param name="Key">Optional key to differentiate multiple registrations of the same type.</param>
    Public Shared Sub RegisterFactory(Of T)(Factory As Func(Of T), Optional Key As String = "")
        Dim InstanceType = GetType(T)
        _Registrations((InstanceType, Key)) = New Registration With {
            .Type = RegistrationType.Factory,
            .Factory = Function() Factory()
        }
    End Sub

    ''' <summary>
    ''' Retrieves a registered instance of the specified type.
    ''' </summary>
    ''' <typeparam name="T">The type of service to retrieve.</typeparam>
    ''' <param name="Key">Optional key used during registration.</param>
    ''' <returns>An instance of the requested type.</returns>
    ''' <exception cref="Exception">Thrown if the service of the specified type is not registered.</exception>
    Public Shared Function GetInstance(Of T)(Optional Key As String = "") As T
        Dim InstanceType = GetType(T)
        Dim RegistrationKey = (InstanceType, Key)
        If _Registrations.ContainsKey(RegistrationKey) Then
            Dim registration = _Registrations(RegistrationKey)
            Select Case registration.Type
                Case RegistrationType.Singleton
                    Return CType(registration.Instance, T)
                Case RegistrationType.Factory
                    Return CType(registration.Factory.Invoke(), T)
            End Select
        Else
            Throw New Exception($"Service of type {InstanceType.Name} {If(Not String.IsNullOrEmpty(Key), $"with key {Key} ", String.Empty)} is not registered.")
        End If
    End Function

End Class
