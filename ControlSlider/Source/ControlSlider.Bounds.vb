Imports System.Drawing

Public Partial Class ControlSlider
    ''' <summary>
    ''' Converts coordinates between the child's immediate parent and the bounds
    ''' container, so nested child controls are constrained correctly.
    ''' </summary>
    Private Function ToBoundaryCoordinates(ChildLocation As Point) As Point
        Return _Parent.PointToClient(_Child.Parent.PointToScreen(ChildLocation))
    End Function

    Private Function FromBoundaryCoordinates(BoundaryLocation As Point) As Point
        Return _Child.Parent.PointToClient(_Parent.PointToScreen(BoundaryLocation))
    End Function

    Private Function ClampChildLocation(Candidate As Point) As Point
        If Not _ConstrainToParent OrElse _Parent Is Nothing OrElse _Child Is Nothing Then Return Candidate
        Return FromBoundaryCoordinates(ClampBoundaryCoordinates(ToBoundaryCoordinates(Candidate)))
    End Function

    Private Function ClampBoundaryCoordinates(Point As Point) As Point
        Dim MinX As Long = _BoundsPadding.Left
        Dim MinY As Long = _BoundsPadding.Top
        Dim MaxX As Long = Math.Max(MinX, CLng(_Parent.ClientSize.Width) - _BoundsPadding.Right - _Child.Width)
        Dim MaxY As Long = Math.Max(MinY, CLng(_Parent.ClientSize.Height) - _BoundsPadding.Bottom - _Child.Height)
        Dim X As Long = Math.Max(MinX, Math.Min(MaxX, CLng(Point.X)))
        Dim Y As Long = Math.Max(MinY, Math.Min(MaxY, CLng(Point.Y)))
        Return New Point(CInt(X), CInt(Y))
    End Function

    Private Function CalculateSnapLocation(Current As Point) As Point
        Dim Local As Point = ToBoundaryCoordinates(Current)
        Dim X As Integer = Local.X
        Dim Y As Integer = Local.Y

        If _MovementAxis <> SliderMovementAxis.Vertical Then X = SnapCoordinate(X, _BoundsPadding.Left, _GridSize)
        If _MovementAxis <> SliderMovementAxis.Horizontal Then Y = SnapCoordinate(Y, _BoundsPadding.Top, _GridSize)

        Dim Snapped As New Point(X, Y)
        If _ConstrainToParent Then Snapped = ClampBoundaryCoordinates(Snapped)
        Return FromBoundaryCoordinates(Snapped)
    End Function

    Private Shared Function SnapCoordinate(Value As Integer, Origin As Integer, GridSize As Integer) As Integer
        Dim Steps As Double = (CDbl(Value) - Origin) / GridSize
        Dim Snapped As Double = Origin + Math.Round(Steps, 0, MidpointRounding.AwayFromZero) * GridSize
        Return CInt(Math.Max(CDbl(Integer.MinValue), Math.Min(CDbl(Integer.MaxValue), Snapped)))
    End Function
End Class
