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
        Dim Ops As Stack(Of String) = New Stack(Of String)()
        Dim Vals As Stack(Of Double) = New Stack(Of Double)()
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
End Class
