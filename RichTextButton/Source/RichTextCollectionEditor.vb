Imports System
Imports System.ComponentModel.Design

''' <summary>
''' Provides a collection editor for <see cref="RichTextPart"/> objects.
''' </summary>
Public Class RichTextCollectionEditor
    Inherits CollectionEditor

    ''' <summary>
    ''' Initializes a new instance of the editor.
    ''' </summary>
    Public Sub New(CollectionType As Type)
        MyBase.New(CollectionType)
    End Sub

    ''' <summary>
    ''' Specifies the item types available in the collection editor.
    ''' </summary>
    Protected Overrides Function CreateNewItemTypes() As Type()
        Return {GetType(RichTextPart)}
    End Function
End Class
