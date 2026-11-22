''' <summary>
''' Provides helper methods for mathematical operations and evaluations.
''' </summary>
Public Class MathHelper

    ''' <summary>
    ''' Returns the closest value from a given sample set
    ''' relative to a specified target value.
    ''' </summary>
    ''' <param name="Samples">
    ''' The sample values to be evaluated.
    ''' </param>
    ''' <param name="Value">
    ''' The value to be approximated within the sample set.
    ''' </param>
    ''' <returns>
    ''' The value from <paramref name="Samples"/> that is closest
    ''' to the specified <paramref name="Value"/>.
    ''' </returns>
    Public Shared Function ApproximateValue(Samples() As Decimal, Value As Decimal) As Decimal
        Dim Dif As New List(Of Decimal)
        Dim Min As Decimal
        For Each s In Samples
            Dif.Add(Math.Abs(Value - s))
        Next s
        Min = Dif.Min
        Return Samples(Dif.IndexOf(Min))
    End Function

    ''' <summary>
    ''' Evaluates a mathematical expression provided as a string
    ''' and returns its numerical result.
    ''' </summary>
    ''' <param name="input">
    ''' The mathematical expression to be evaluated.
    ''' </param>
    ''' <returns>
    ''' The result of the evaluated mathematical expression.
    ''' </returns>
    Public Shared Function EvaluateExpression(input As String) As Double
        Dim Expr As String = "(" & input.Replace(" ", Nothing) & ")"
        Dim Ops As New Stack(Of String)()
        Dim Vals As New Stack(Of Double)()
        For i As Integer = 0 To Expr.Length - 1
            Dim s As String = Expr.Substring(i, 1)
            If s.Equals("(") Then
            ElseIf s.Equals("+") Then
                Ops.Push(s)
            ElseIf s.Equals("-") Then
                Ops.Push(s)
            ElseIf s.Equals("*") Then
                Ops.Push(s)
            ElseIf s.Equals("/") Then
                Ops.Push(s)
            ElseIf s.Equals("sqrt") Then
                Ops.Push(s)
            ElseIf s.Equals(")") Then
                Dim count As Integer = Ops.Count
                While count > 0
                    Dim op As String = Ops.Pop()
                    Dim v As Double = Vals.Pop()

                    If op.Equals("+") Then
                        v = Vals.Pop() + v
                    ElseIf op.Equals("-") Then
                        v = Vals.Pop() - v
                    ElseIf op.Equals("*") Then
                        v = Vals.Pop() * v
                    ElseIf op.Equals("/") Then
                        v = Vals.Pop() / v
                    ElseIf op.Equals("sqrt") Then
                        v = Math.Sqrt(v)
                    End If
                    Vals.Push(v)
                    count -= 1
                End While
            Else
                Vals.Push(Double.Parse(s))
            End If
        Next
        Return Vals.Pop()
    End Function
    ''' <summary>
    ''' Restricts a value to the specified minimum and maximum bounds.
    ''' </summary>
    ''' <param name="Value">The value to clamp.</param>
    ''' <param name="Min">The minimum allowed value.</param>
    ''' <param name="Max">The maximum allowed value.</param>
    ''' <returns>
    ''' The minimum value if <paramref name="Value"/> is less than
    ''' <paramref name="Min"/>; the maximum value if it is greater than
    ''' <paramref name="Max"/>; otherwise, the original value.
    ''' </returns>
    ''' <exception cref="ArgumentException">
    ''' Thrown when <paramref name="Min"/> is greater than
    ''' <paramref name="Max"/>.
    ''' </exception>
    Public Shared Function Clamp(Value As Integer, Min As Integer, Max As Integer) As Integer
        If Min > Max Then
            Throw New ArgumentException("The minimum value cannot be greater than the maximum value.", NameOf(Min))
        End If

        Return Math.Max(Min, Math.Min(Value, Max))
    End Function
End Class
