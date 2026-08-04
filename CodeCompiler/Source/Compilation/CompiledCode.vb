Imports System.Globalization
Imports System.Reflection
Imports System.Runtime.ExceptionServices

''' <summary>
''' Represents a dynamically compiled assembly that can execute public methods.
''' </summary>
Public NotInheritable Class CompiledCode
    Implements IDisposable
    Private _LoadContext As CompilerAssemblyLoadContext
    Private _Assembly As System.Reflection.Assembly
    Private _DisposedValue As Boolean
    ''' <summary>
    ''' Initializes a new instance of the <see cref="CompiledCode"/> class.
    ''' </summary>
    ''' <param name="LoadContext">The load context containing the compiled assembly.</param>
    ''' <param name="RuntimeAssembly">The dynamically compiled assembly.</param>
    Friend Sub New(LoadContext As CompilerAssemblyLoadContext, RuntimeAssembly As System.Reflection.Assembly)
        _LoadContext = LoadContext
        _Assembly = RuntimeAssembly
    End Sub
    ''' <summary>
    ''' Gets the dynamically compiled assembly.
    ''' </summary>
    ''' <value>
    ''' The assembly generated from the supplied source code.
    ''' </value>
    ''' <exception cref="ObjectDisposedException">The compiled assembly has already been released.</exception>
    Public ReadOnly Property Assembly As System.Reflection.Assembly
        Get
            ThrowIfDisposed()
            Return _Assembly
        End Get
    End Property
    ''' <summary>
    ''' Invokes a public shared, static or instance method from the compiled assembly.
    ''' </summary>
    ''' <param name="TypeName">The full name of the type that contains the method.</param>
    ''' <param name="MethodName">The name of the method to invoke.</param>
    ''' <param name="Parameters">The arguments passed to the method.</param>
    ''' <returns>The value returned by the method, or <see langword="Nothing"/> when the method does not return a value.</returns>
    ''' <exception cref="TypeLoadException">The specified type was not found.</exception>
    ''' <exception cref="MissingMethodException">A compatible method or constructor was not found.</exception>
    ''' <exception cref="AmbiguousMatchException">More than one compatible method overload was found.</exception>
    Public Function Invoke(TypeName As String, MethodName As String, ParamArray Parameters As Object()) As Object
        ThrowIfDisposed()
        If String.IsNullOrWhiteSpace(TypeName) Then Throw New ArgumentException("The type name cannot be empty.", NameOf(TypeName))
        If String.IsNullOrWhiteSpace(MethodName) Then Throw New ArgumentException("The method name cannot be empty.", NameOf(MethodName))
        Dim TargetType As System.Type = _Assembly.GetType(TypeName, False, False)
        If TargetType Is Nothing Then Throw New TypeLoadException($"The type '{TypeName}' was not found in the compiled assembly.")
        Dim SuppliedParameters As Object() = If(Parameters, Array.Empty(Of Object)())
        Dim PreparedParameters As Object() = Nothing
        Dim TargetMethod As MethodInfo = FindMethod(TargetType, MethodName, SuppliedParameters, PreparedParameters)
        If TargetMethod Is Nothing Then Throw New MissingMethodException($"A compatible public method '{TypeName}.{MethodName}' with {SuppliedParameters.Length} parameter(s) was not found.")
        Dim TargetInstance As Object = Nothing
        If Not TargetMethod.IsStatic Then
            If TargetType.IsAbstract Then Throw New InvalidOperationException($"The type '{TypeName}' is abstract and cannot be instantiated.")
            Dim Constructor As ConstructorInfo = TargetType.GetConstructor(BindingFlags.Instance Or BindingFlags.Public, Nothing, System.Type.EmptyTypes, Nothing)
            If Constructor Is Nothing Then Throw New MissingMethodException($"The type '{TypeName}' must have a public parameterless constructor to invoke an instance method.")
            TargetInstance = Constructor.Invoke(Array.Empty(Of Object)())
        End If
        Try
            Return TargetMethod.Invoke(TargetInstance, PreparedParameters)
        Catch Ex As TargetInvocationException When Ex.InnerException IsNot Nothing
            ExceptionDispatchInfo.Capture(Ex.InnerException).Throw()
            Return Nothing
        End Try
    End Function
    Private Shared Function FindMethod(TargetType As System.Type, MethodName As String, SuppliedParameters As Object(), ByRef PreparedParameters As Object()) As MethodInfo
        Dim BestMethod As MethodInfo = Nothing
        Dim BestParameters As Object() = Nothing
        Dim BestScore As Integer = Integer.MaxValue
        Dim IsAmbiguous As Boolean
        Dim Flags As BindingFlags = BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.Static Or BindingFlags.FlattenHierarchy
        Dim CandidateMethods As IEnumerable(Of MethodInfo) = TargetType.GetMethods(Flags).Where(Function(Item) String.Equals(Item.Name, MethodName, StringComparison.Ordinal) AndAlso Not Item.ContainsGenericParameters)
        For Each CandidateMethod As MethodInfo In CandidateMethods
            Dim CandidateParameters As Object() = Nothing
            Dim CandidateScore As Integer
            If Not TryPrepareParameters(CandidateMethod, SuppliedParameters, CandidateParameters, CandidateScore) Then Continue For
            If CandidateScore < BestScore Then
                BestMethod = CandidateMethod
                BestParameters = CandidateParameters
                BestScore = CandidateScore
                IsAmbiguous = False
            ElseIf CandidateScore = BestScore Then
                IsAmbiguous = True
            End If
        Next CandidateMethod
        If IsAmbiguous Then Throw New AmbiguousMatchException($"More than one overload of '{TargetType.FullName}.{MethodName}' is compatible with the supplied parameters.")
        PreparedParameters = BestParameters
        Return BestMethod
    End Function
    Private Shared Function TryPrepareParameters(TargetMethod As MethodInfo, SuppliedParameters As Object(), ByRef PreparedParameters As Object(), ByRef Score As Integer) As Boolean
        Dim MethodParameters As ParameterInfo() = TargetMethod.GetParameters()
        If MethodParameters.Length <> SuppliedParameters.Length Then Return False
        If SuppliedParameters.Length = 0 Then
            PreparedParameters = Array.Empty(Of Object)()
        Else
            PreparedParameters = New Object(SuppliedParameters.Length - 1) {}
        End If
        Score = 0
        For Index As Integer = 0 To MethodParameters.Length - 1
            Dim ParameterType As System.Type = MethodParameters(Index).ParameterType
            If ParameterType.IsByRef Then Return False
            Dim SuppliedValue As Object = SuppliedParameters(Index)
            If SuppliedValue Is Nothing Then
                If ParameterType.IsValueType AndAlso Nullable.GetUnderlyingType(ParameterType) Is Nothing Then Return False
                PreparedParameters(Index) = Nothing
                Score += 2
                Continue For
            End If
            Dim SuppliedType As System.Type = SuppliedValue.GetType()
            If ParameterType Is SuppliedType Then
                PreparedParameters(Index) = SuppliedValue
                Continue For
            End If
            If ParameterType.IsAssignableFrom(SuppliedType) Then
                PreparedParameters(Index) = SuppliedValue
                Score += 1
                Continue For
            End If
            Dim NullableType As System.Type = Nullable.GetUnderlyingType(ParameterType)
            Dim ConversionType As System.Type = If(NullableType, ParameterType)
            Try
                If ConversionType.IsEnum Then
                    If TypeOf SuppliedValue Is String Then
                        PreparedParameters(Index) = [Enum].Parse(ConversionType, DirectCast(SuppliedValue, String), True)
                    Else
                        PreparedParameters(Index) = [Enum].ToObject(ConversionType, SuppliedValue)
                    End If
                    Score += 3
                    Continue For
                End If
                If ConversionType Is GetType(Guid) AndAlso TypeOf SuppliedValue Is String Then
                    PreparedParameters(Index) = Guid.Parse(DirectCast(SuppliedValue, String))
                    Score += 3
                    Continue For
                End If
                If TypeOf SuppliedValue Is IConvertible AndAlso GetType(IConvertible).IsAssignableFrom(ConversionType) Then
                    PreparedParameters(Index) = Convert.ChangeType(SuppliedValue, ConversionType, CultureInfo.InvariantCulture)
                    Score += 4
                    Continue For
                End If
            Catch Ex As Exception When TypeOf Ex Is InvalidCastException OrElse TypeOf Ex Is FormatException OrElse TypeOf Ex Is OverflowException OrElse TypeOf Ex Is ArgumentException
                Return False
            End Try
            Return False
        Next Index
        Return True
    End Function
    Private Sub ThrowIfDisposed()
        If Not _DisposedValue Then Return
        Throw New ObjectDisposedException(NameOf(CompiledCode))
    End Sub
    ''' <summary>
    ''' Releases the compiled assembly and requests unloading of its assembly load context.
    ''' </summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        If _DisposedValue Then Return
        _Assembly = Nothing
        Dim LoadContext As CompilerAssemblyLoadContext = _LoadContext
        _LoadContext = Nothing
        _DisposedValue = True
        LoadContext?.Unload()
        GC.SuppressFinalize(Me)
    End Sub
End Class