Imports System.Drawing
Imports System.Reflection
Imports System.Windows.Forms
''' <summary>
''' Provides utility methods for traversing, configuring, and inspecting
''' Windows Forms controls and forms.
''' </summary>
Public Class ControlHelper

    ''' <summary>
    ''' Returns a collection containing the root control and all of its child controls.
    ''' </summary>
    ''' <param name="Root">The root control.</param>
    ''' <param name="includeRoot">
    ''' Indicates whether the root control itself should be included in the result.
    ''' </param>
    ''' <returns>
    ''' An enumerable collection containing the root control (if specified) and all descendant controls.
    ''' </returns>
    Public Shared Iterator Function GetAllControls(Root As Control, includeRoot As Boolean) As IEnumerable(Of Control)
        Dim Stack = New Stack(Of Control)()
        If includeRoot Then
            Stack.Push(Root)
        End If
        For Each Child As Control In Root.Controls
            Stack.Push(Child)
        Next

        While Stack.Any()
            Dim [next] = Stack.Pop()
            Yield [next]
            For Each child As Control In [next].Controls
                Stack.Push(child)
            Next
        End While
    End Function

    ''' <summary>
    ''' Enables or disables double buffering for a control using reflection.
    ''' </summary>
    ''' <param name="Ctrl">The target control.</param>
    ''' <param name="Setting">
    ''' <c>True</c> to enable double buffering; <c>False</c> to disable it.
    ''' </param>
    Public Shared Sub EnableControlDoubleBuffer(Ctrl As Control, Setting As Boolean)
        Dim DgvType As Type = Ctrl.[GetType]()
        Dim pi As PropertyInfo = DgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
        pi.SetValue(Ctrl, Setting, Nothing)
    End Sub

    ''' <summary>
    ''' Enables or disables double buffering for a form using reflection.
    ''' </summary>
    ''' <param name="Form">The target form.</param>
    ''' <param name="Setting">
    ''' <c>True</c> to enable double buffering; <c>False</c> to disable it.
    ''' </param>
    Public Shared Sub EnableFormDoubleBuffer(Form As Form, Setting As Boolean)
        Dim DgvType As Type = Form.[GetType]()
        Dim pi As PropertyInfo = DgvType.GetProperty("DoubleBuffered", BindingFlags.Instance Or BindingFlags.NonPublic)
        pi.SetValue(Form, Setting, Nothing)
    End Sub

    ''' <summary>
    ''' Determines whether the specified child control is fully visible within its parent control.
    ''' </summary>
    ''' <param name="Parent">The parent control.</param>
    ''' <param name="Child">The child control to evaluate.</param>
    ''' <returns>
    ''' <c>True</c> if the child control is completely visible within the parent; otherwise, <c>False</c>.
    ''' </returns>
    Public Shared Function IsControlFullyVisible(Parent As Control, Child As Control) As Boolean
        Dim MyBounds As Rectangle = Parent.ClientRectangle
        If Not MyBounds.Contains(Child.Location) Then
            Return False
        End If
        If MyBounds.Right < Child.Right Then
            Return False
        End If
        If MyBounds.Height < Child.Bottom Then
            Return False
        End If
        Return True
    End Function

End Class
