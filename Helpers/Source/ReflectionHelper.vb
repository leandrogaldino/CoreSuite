Imports System.Reflection
Imports System.Reflection.Emit

''' <summary>
''' Provides helper methods for working with reflection,
''' dynamic types, and runtime method invocation.
''' </summary>
Public Class ReflectionHelper

    ''' <summary>
    ''' Determines whether a property represents a collection type.
    ''' </summary>
    ''' <param name="PropertyInfo">
    ''' The property to be inspected.
    ''' </param>
    ''' <returns>
    ''' <c>True</c> if the property represents a collection type
    ''' (excluding <see cref="String"/>); otherwise, <c>False</c>.
    ''' </returns>
    Public Shared Function IsCollection(PropertyInfo As PropertyInfo) As Boolean
        If PropertyInfo.PropertyType.Name = "String" Then
            Return False
        Else
            If GetType(IEnumerable).IsAssignableFrom(PropertyInfo.PropertyType) Then
                Return True
            Else
                Return False
            End If
        End If
    End Function

    ''' <summary>
    ''' Analyzes a collection property and returns the type
    ''' of its elements.
    ''' </summary>
    ''' <param name="PropertyInfo">
    ''' The collection property to be analyzed.
    ''' </param>
    ''' <returns>
    ''' The <see cref="Type"/> of the elements contained in the collection.
    ''' </returns>
    Public Shared Function GetCollectionPropertyType(PropertyInfo As PropertyInfo) As Type
        Dim Type As Type = PropertyInfo.PropertyType
        Dim IntIndexer = Type.GetMethod("get_Item", {GetType(Integer)})
        Return IntIndexer.ReturnType
    End Function

    ''' <summary>
    ''' Analyzes a collection type and returns the type
    ''' of its elements.
    ''' </summary>
    ''' <param name="Type">
    ''' The collection type to be analyzed.
    ''' </param>
    ''' <returns>
    ''' The <see cref="Type"/> of the elements contained in the collection.
    ''' </returns>
    Public Shared Function GetCollectionPropertyType(Type As Type) As Type
        Dim IntIndexer = Type.GetMethod("get_Item", {GetType(Integer)})
        Return IntIndexer.ReturnType
    End Function

    ''' <summary>
    ''' Invokes a method on an object using reflection.
    ''' </summary>
    ''' <param name="Obj">
    ''' The object whose method will be invoked.
    ''' </param>
    ''' <param name="MethodName">
    ''' The name of the method to invoke.
    ''' </param>
    ''' <param name="Flags">
    ''' The <see cref="BindingFlags"/> used to locate the method.
    ''' </param>
    ''' <param name="MethodParams">
    ''' The parameters to be passed to the method.
    ''' </param>
    ''' <returns>
    ''' The return value of the invoked method,
    ''' or <c>Nothing</c> if the method does not return a value.
    ''' </returns>
    Public Shared Function InvokeMethod(Obj As Object, MethodName As String, Flags As BindingFlags, ParamArray MethodParams As Object()) As Object
        Dim MethodParamTypes = If(MethodParams?.[Select](Function(p) p.[GetType]()).ToArray(), New Type() {})
        'Dim BindingFlags As BindingFlags = Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.[Public] Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.[Static]
        Dim BindingFlags As BindingFlags = Flags
        Dim Method As MethodInfo = Nothing
        Dim Type = Obj.[GetType]()
        While Method Is Nothing AndAlso Type IsNot Nothing
            Method = Type.GetMethod(MethodName, BindingFlags, Type.DefaultBinder, MethodParamTypes, Nothing)
            Type = Type.BaseType
        End While
        Return Method?.Invoke(Obj, MethodParams)
    End Function

    ''' <summary>
    ''' Creates a runtime-generated <see cref="Type"/> with the specified
    ''' properties.
    ''' </summary>
    ''' <param name="PropertyNames">
    ''' A list containing the names of the properties.
    ''' </param>
    ''' <param name="PropertyTypes">
    ''' A list containing the <see cref="Type"/> of each property.
    ''' </param>
    ''' <returns>
    ''' A dynamically generated <see cref="Type"/> containing
    ''' the specified properties.
    ''' </returns>
    ''' <exception cref="Exception">
    ''' Thrown when the number of property names does not match
    ''' the number of property types.
    ''' </exception>
    Public Shared Function GetRunTimeType(PropertyNames As List(Of String), PropertyTypes As List(Of Type)) As Type
        If PropertyNames.Count <> PropertyTypes.Count Then
            Throw New Exception("PropertyNames and PropertyTypes do not have the same amount of elements.")
        End If
        Dim TypeSignature As String = "MyDynamicType"
        Dim AssemblyName As New AssemblyName(TypeSignature)
        Dim AssemblyBuilder As AssemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(AssemblyName, AssemblyBuilderAccess.Run)
        Dim ModuleBuilder As ModuleBuilder = AssemblyBuilder.DefineDynamicModule("MainModule")
        Dim TypeBuilder As TypeBuilder = ModuleBuilder.DefineType(TypeSignature, TypeAttributes.Public Or TypeAttributes.Class Or TypeAttributes.AutoClass Or TypeAttributes.AnsiClass Or TypeAttributes.BeforeFieldInit Or TypeAttributes.AutoLayout)
        TypeBuilder.DefineDefaultConstructor(MethodAttributes.Public Or MethodAttributes.SpecialName Or MethodAttributes.RTSpecialName)
        For i As Integer = 0 To PropertyNames.Count - 1
            Dim FieldBuilder As FieldBuilder = TypeBuilder.DefineField("_" & PropertyNames(i), PropertyTypes(i), FieldAttributes.Private)
            Dim PropertyBuilder As PropertyBuilder = TypeBuilder.DefineProperty(PropertyNames(i), PropertyAttributes.HasDefault, PropertyTypes(i), Nothing)
            Dim GetMethod As MethodBuilder = TypeBuilder.DefineMethod("get_" & PropertyNames(i), MethodAttributes.Public Or MethodAttributes.SpecialName Or MethodAttributes.HideBySig, PropertyTypes(i), Type.EmptyTypes)
            Dim GetIL As ILGenerator = GetMethod.GetILGenerator()
            GetIL.Emit(OpCodes.Ldarg_0)
            GetIL.Emit(OpCodes.Ldfld, FieldBuilder)
            GetIL.Emit(OpCodes.Ret)
            Dim SetMethod As MethodBuilder = TypeBuilder.DefineMethod("set_" & PropertyNames(i), MethodAttributes.Public Or MethodAttributes.SpecialName Or MethodAttributes.HideBySig, Nothing, {PropertyTypes(i)})
            Dim SetIL As ILGenerator = SetMethod.GetILGenerator()
            SetIL.Emit(OpCodes.Ldarg_0)
            SetIL.Emit(OpCodes.Ldarg_1)
            SetIL.Emit(OpCodes.Stfld, FieldBuilder)
            SetIL.Emit(OpCodes.Ret)
            PropertyBuilder.SetGetMethod(GetMethod)
            PropertyBuilder.SetSetMethod(SetMethod)
        Next i
        Dim ObjectType As Type = TypeBuilder.CreateTypeInfo().AsType()
        Return ObjectType
    End Function

End Class