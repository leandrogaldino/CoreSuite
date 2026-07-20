Namespace Extensions
    Public Module CollectionExtensions

        ''' <summary>
        ''' Attempts to return the single element of a sequence.
        ''' If no element exists, a new instance of the element's type is created and returned.
        ''' </summary>
        ''' <typeparam name="T">The type of element in the sequence.</typeparam>
        ''' <param name="Source">The sequence to search.</param>
        ''' <param name="Predicate">An optional function to filter the elements.</param>
        ''' <returns>
        ''' The single element of the sequence matching the predicate,
        ''' or a new instance of <typeparamref name="T"/> if none exists.
        ''' </returns>
        ''' <remarks>
        ''' <para>
        ''' Use this extension when you expect at most one item from a collection,
        ''' but want to ensure a valid object is returned even if the sequence is empty.
        ''' </para>
        ''' <para>
        ''' If the type <typeparamref name="T"/> does not have a parameterless constructor,
        ''' <c>Nothing</c> is returned.
        ''' </para>
        ''' </remarks>
        Public Function FirstOrNew(Of T)(Source As IEnumerable(Of T), Optional Predicate As Func(Of T, Boolean) = Nothing) As T
            Dim Element As T = Source.SingleOrDefault(Predicate)
            If Element Is Nothing Then
                Try
                    Return Activator.CreateInstance(GetType(T))
                Catch ex As Exception
                    Return Nothing
                End Try
            Else
                Return Element
            End If
        End Function

    End Module
End Namespace