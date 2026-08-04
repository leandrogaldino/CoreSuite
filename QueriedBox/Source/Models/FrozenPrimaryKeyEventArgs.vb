''' <summary>
''' Provides data for events that occur when the primary key value of a frozen record changes.
''' </summary>
Public Class FrozenPrimaryKeyEventArgs
    Inherits EventArgs
    ''' <summary>
    ''' Gets the previous primary key value before the change occurred.
    ''' </summary>
    Public ReadOnly Property OldPrimaryKey As Object
    ''' <summary>
    ''' Gets the new primary key value after the change occurred.
    ''' </summary>
    Public ReadOnly Property NewPrimaryKey As Object
    ''' <summary>
    ''' Initializes a new instance of the <see cref="FrozenPrimaryKeyEventArgs"/> class
    ''' with the new primary key value.
    ''' </summary>
    ''' <param name="NewPrimaryKey">
    ''' The new primary key value assigned to the record.
    ''' </param>
    Public Sub New(NewPrimaryKey As Object)
        Me.NewPrimaryKey = NewPrimaryKey
    End Sub
    ''' <summary>
    ''' Initializes a new instance of the <see cref="FrozenPrimaryKeyEventArgs"/> class
    ''' with both the previous and new primary key values.
    ''' </summary>
    ''' <param name="OldPrimaryKey">
    ''' The previous primary key value before the change.
    ''' </param>
    ''' <param name="NewPrimaryKey">
    ''' The new primary key value after the change.
    ''' </param>
    Public Sub New(OldPrimaryKey As Object, NewPrimaryKey As Object)
        Me.OldPrimaryKey = OldPrimaryKey
        Me.NewPrimaryKey = NewPrimaryKey
    End Sub
End Class