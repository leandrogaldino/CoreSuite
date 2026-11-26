Imports System.Windows.Forms

Public Partial Class TileView
    ''' <summary>Select a record by its current filtered/display position.</summary>
    Public Sub SelectIndex(DisplayIndex As Integer)
        If DisplayIndex < 0 OrElse DisplayIndex >= _VisibleSourceIndices.Count Then
            Throw New ArgumentOutOfRangeException(NameOf(DisplayIndex))
        End If
        Dim NewSourceIndex As Integer = _VisibleSourceIndices(DisplayIndex)
        If _SelectedSourceIndex = NewSourceIndex Then
            EnsureVisible(DisplayIndex)
            Return
        End If
        Dim OldSourceIndex As Integer = _SelectedSourceIndex
        Dim OldData As Object = GetSourceItemOrNothing(OldSourceIndex)
        Dim OldTile As TileViewItem = SelectedItem
        _SelectedSourceIndex = NewSourceIndex
        For Each Pair In _Realized
            Pair.Value.SetSelectedInternal(Pair.Key = DisplayIndex)
        Next
        EnsureVisible(DisplayIndex)
        RaiseSelectionChanged(OldSourceIndex, NewSourceIndex, OldData, _Data(NewSourceIndex), OldTile, SelectedItem)
    End Sub

    ''' <summary>Select by the original source index; returns False if filtered out.</summary>
    Public Function SelectSourceIndex(SourceIndex As Integer) As Boolean
        If SourceIndex < 0 OrElse SourceIndex >= _Data.Count Then Return False
        Dim DisplayIndex As Integer = GetDisplayIndex(SourceIndex)
        If DisplayIndex < 0 Then Return False
        SelectIndex(DisplayIndex)
        Return True
    End Function

    Public Sub SelectItem(Tile As TileViewItem)
        If Tile Is Nothing Then Throw New ArgumentNullException(NameOf(Tile))
        If Tile.OwnerView IsNot Me OrElse Tile.DisplayIndex < 0 OrElse
           Not _Realized.ContainsKey(Tile.DisplayIndex) OrElse
           Not Object.ReferenceEquals(_Realized(Tile.DisplayIndex), Tile) Then
            Throw New ArgumentException("The tile is not a realized item of this view.", NameOf(Tile))
        End If
        SelectIndex(Tile.DisplayIndex)
    End Sub

    Public Sub ClearSelection()
        If _SelectedSourceIndex = -1 Then Return
        Dim OldSourceIndex As Integer = _SelectedSourceIndex
        Dim OldData As Object = GetSourceItemOrNothing(OldSourceIndex)
        Dim OldTile As TileViewItem = SelectedItem
        _SelectedSourceIndex = -1
        For Each Tile In _Realized.Values
            Tile.SetSelectedInternal(False)
        Next
        RaiseSelectionChanged(OldSourceIndex, -1, OldData, Nothing, OldTile, Nothing)
    End Sub

    Public Function SelectFirst() As Boolean
        Return SelectAndFocusIndex(0)
    End Function

    Public Function SelectLast() As Boolean
        Return SelectAndFocusIndex(_VisibleSourceIndices.Count - 1)
    End Function

    Public Function SelectNext() As Boolean
        Return MoveSequential(1)
    End Function

    Public Function SelectPrevious() As Boolean
        Return MoveSequential(-1)
    End Function

    Public Function ActivateSelected() As Boolean
        If SelectedIndex < 0 Then Return False
        Dim Tile As TileViewItem = SelectedItem
        If Tile Is Nothing Then
            EnsureVisible(SelectedIndex)
            Tile = SelectedItem
        End If
        If Tile Is Nothing Then Return False
        Return ActivateItem(Tile)
    End Function

    ''' <summary>Raises activation for a currently realized tile.</summary>
    Public Function ActivateItem(Tile As TileViewItem) As Boolean
        If Tile Is Nothing OrElse Not Object.ReferenceEquals(Tile.OwnerView, Me) OrElse
           Tile.DisplayIndex < 0 OrElse Not Tile.Enabled Then Return False
        SelectItem(Tile)
        Tile.RaiseActivated()
        RaiseEvent ItemActivated(Me, New TileViewItemEventArgs(Tile))
        Return True
    End Function

    Private Function SelectAndFocusIndex(Index As Integer) As Boolean
        If Index < 0 OrElse Index >= _VisibleSourceIndices.Count Then Return False
        SelectIndex(Index)
        Dim Tile As TileViewItem = SelectedItem
        If Tile IsNot Nothing AndAlso Tile.CanFocus Then Tile.Focus()
        Return True
    End Function

    Private Function MoveSequential(Direction As Integer) As Boolean
        Dim Count As Integer = _VisibleSourceIndices.Count
        If Count = 0 Then Return False
        Dim Current As Integer = SelectedIndex
        If Current = -1 Then Return SelectAndFocusIndex(If(Direction >= 0, 0, Count - 1))
        Dim NextIndex As Integer = Current + Direction
        If NextIndex < 0 OrElse NextIndex >= Count Then
            If Not _WrapNavigation Then Return False
            NextIndex = If(NextIndex < 0, Count - 1, 0)
        End If
        Return SelectAndFocusIndex(NextIndex)
    End Function

    Private Function MoveDirectional(Direction As Keys) As Boolean
        Dim Count As Integer = _VisibleSourceIndices.Count
        If Count = 0 Then Return False
        Dim Current As Integer = SelectedIndex
        If Current = -1 Then Return SelectAndFocusIndex(0)
        If _layoutDirty Then UpdateViewport()
        Dim NextIndex As Integer = Current
        Select Case Direction
            Case Keys.Left
                NextIndex -= 1
            Case Keys.Right
                NextIndex += 1
            Case Keys.Up
                NextIndex -= _Columns
            Case Keys.Down
                NextIndex += _Columns
            Case Keys.PageUp
                If _RowOffsets Is Nothing Then
                    NextIndex -= Math.Max(1, ClientSize.Height \ (_ItemSize.Height + _ItemSpacing)) * _Columns
                Else
                    Dim CurrentRow As Integer = Current \ _Columns
                    Dim TargetRow As Integer = GetRowAtOffset(GetRowOffset(CurrentRow) + -1L * Math.Max(1, ClientSize.Height))
                    NextIndex += If(TargetRow = CurrentRow, -1, TargetRow - CurrentRow) * _Columns
                    If TargetRow <> CurrentRow AndAlso NextIndex >= Count Then NextIndex = Count - 1
                End If
            Case Keys.PageDown
                If _RowOffsets Is Nothing Then
                    NextIndex += Math.Max(1, ClientSize.Height \ (_ItemSize.Height + _ItemSpacing)) * _Columns
                Else
                    Dim CurrentRow As Integer = Current \ _Columns
                    Dim TargetRow As Integer = GetRowAtOffset(GetRowOffset(CurrentRow) + 1L * Math.Max(1, ClientSize.Height))
                    NextIndex += If(TargetRow = CurrentRow, 1, TargetRow - CurrentRow) * _Columns
                    If TargetRow <> CurrentRow AndAlso NextIndex >= Count Then NextIndex = Count - 1
                End If
        End Select
        If NextIndex < 0 OrElse NextIndex >= Count Then
            If Not _WrapNavigation Then Return False
            NextIndex = If(NextIndex < 0, Count - 1, 0)
        End If
        Return SelectAndFocusIndex(NextIndex)
    End Function

    Friend Function ProcessItemCommand(Tile As TileViewItem, KeyData As Keys) As Boolean
        If Tile Is Nothing OrElse Tile.OwnerView IsNot Me OrElse Tile.DisplayIndex < 0 Then Return False
        Return ProcessNavigationKey(KeyData)
    End Function

    Private Function ProcessNavigationKey(KeyData As Keys) As Boolean
        If (KeyData And (Keys.Control Or Keys.Alt Or Keys.Shift)) <> Keys.None Then Return False
        Select Case KeyData And Keys.KeyCode
            Case Keys.Left, Keys.Right, Keys.Up, Keys.Down, Keys.PageUp, Keys.PageDown
                Return MoveDirectional(KeyData And Keys.KeyCode)
            Case Keys.Home
                Return SelectFirst()
            Case Keys.End
                Return SelectLast()
            Case Keys.Enter
                Return ActivateSelected()
            Case Keys.Space
                If SelectedIndex = -1 Then Return SelectFirst()
                Return True
        End Select
        Return False
    End Function

    Protected Overrides Function ProcessCmdKey(ByRef Msg As Message, KeyData As Keys) As Boolean
        For Each Tile As TileViewItem In _Realized.Values
            If Tile.ContainsFocus AndAlso Not Tile.Focused Then Return MyBase.ProcessCmdKey(Msg, KeyData)
        Next
        If ProcessNavigationKey(KeyData) Then Return True
        Return MyBase.ProcessCmdKey(Msg, KeyData)
    End Function

    Friend Sub NotifyItemClick(Tile As TileViewItem, FocusTile As Boolean)
        If Not IsRealized(Tile) Then Return
        SelectItem(Tile)
        If FocusTile AndAlso Tile.FocusOnClick AndAlso Tile.CanFocus Then Tile.Focus()
        RaiseEvent ItemClick(Me, New TileViewItemEventArgs(Tile))
    End Sub

    Friend Sub NotifyItemDoubleClick(Tile As TileViewItem, FocusTile As Boolean)
        If Not IsRealized(Tile) Then Return
        SelectItem(Tile)
        If FocusTile AndAlso Tile.FocusOnClick AndAlso Tile.CanFocus Then Tile.Focus()
        RaiseEvent ItemDoubleClick(Me, New TileViewItemEventArgs(Tile))
        If _ActivateOnDoubleClick Then ActivateItem(Tile)
    End Sub

    Private Function IsRealized(Tile As TileViewItem) As Boolean
        If Tile Is Nothing OrElse Tile.OwnerView IsNot Me OrElse Tile.DisplayIndex < 0 Then Return False
        Dim Current As TileViewItem = Nothing
        Return _Realized.TryGetValue(Tile.DisplayIndex, Current) AndAlso Object.ReferenceEquals(Tile, Current)
    End Function
End Class
