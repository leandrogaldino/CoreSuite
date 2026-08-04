Imports System.ComponentModel.Design

Friend Class RichTextCollectionEditor
    ''' <summary>
    ''' Provides a design-time collection editor for managing <see cref="RichTextPart"/> objects.
    ''' </summary>
    Public Class RichTextCollectionEditor
        Inherits CollectionEditor
        ''' <summary>
        ''' Initializes a new instance of the <see cref="RichTextCollectionEditor"/> class.
        ''' </summary>
        ''' <param name="ItemType">The type of collection edited by this editor.</param>
        Public Sub New(ItemType As Type)
            MyBase.New(ItemType)
        End Sub
        ''' <summary>
        ''' Gets the types of items that can be created by the collection editor.
        ''' </summary>
        ''' <returns>An array containing the <see cref="RichTextPart"/> type.</returns>
        Protected Overrides Function CreateNewItemTypes() As Type()
            Return {GetType(RichTextPart)}
        End Function
    End Class
End Class
