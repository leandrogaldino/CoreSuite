Partial Public Class NavigationView
    Private Sub RebuildNavigationButtons()
        _NavigationFlow.SuspendLayout()
        Try
            While _NavigationFlow.Controls.Count > 0
                Dim ItemControl As Control = _NavigationFlow.Controls(0)
                _NavigationFlow.Controls.RemoveAt(0)
                RemoveHandler ItemControl.Click, AddressOf NavigationButtonClick
                _ToolTip.SetToolTip(ItemControl, Nothing)
                ItemControl.Dispose()
            End While
            Dim TabIndex As Integer
            For Each Page As NavigationPage In _Pages
                If Page.Visible Then
                    Dim Button As New NavigationButton With {.Page = Page, .TabIndex = TabIndex, .TabStop = True}
                    AddHandler Button.Click, AddressOf NavigationButtonClick
                    _NavigationFlow.Controls.Add(Button)
                    TabIndex += 1
                End If
            Next
            ApplyButtonAppearance()
            ApplyButtonToolTips()
            UpdateButtonLayout()
        Finally
            _NavigationFlow.ResumeLayout(True)
        End Try
    End Sub
    Private Sub ApplyButtonAppearance()
        If _NavigationFlow Is Nothing Then Return
        For Each ItemControl As Control In _NavigationFlow.Controls
            Dim Button As NavigationButton = TryCast(ItemControl, NavigationButton)
            If Button Is Nothing OrElse Button.Page Is Nothing Then Continue For
            Dim Page As NavigationPage = Button.Page
            Button.Text = If(String.IsNullOrWhiteSpace(Page.Text), Page.Key, Page.Text)
            Button.PageImage = If(_ShowImages, Page.Image, Nothing)
            Button.ImageSize = _ImageSize
            Button.Padding = _ButtonPadding
            Button.Font = Font
            Button.RightToLeft = RightToLeft
            Button.Enabled = Enabled AndAlso Page.Enabled
            Button.IsSelected = ReferenceEquals(Page, _SelectedPage)
            Button.NormalBackColor = _ButtonBackColor
            Button.HoverBackColor = _ButtonHoverBackColor
            Button.SelectedBackColor = _SelectedButtonBackColor
            Button.NormalForeColor = _ButtonForeColor
            Button.SelectedForeColor = _SelectedButtonForeColor
            Button.IndicatorColor = _SelectedIndicatorColor
            Button.IndicatorWidth = _SelectedIndicatorWidth
            Button.IndicatorOnRight = _NavigationPosition = NavigationPanePosition.Right
            Button.AccessibleName = ResolveAccessibleName(Page)
        Next
    End Sub
    Private Sub ApplyButtonToolTips()
        If _NavigationFlow Is Nothing OrElse _ToolTip Is Nothing Then Return
        For Each ItemControl As Control In _NavigationFlow.Controls
            Dim Button As NavigationButton = TryCast(ItemControl, NavigationButton)
            If Button Is Nothing OrElse Button.Page Is Nothing Then Continue For
            _ToolTip.SetToolTip(Button, If(_ShowToolTips, Button.Page.ToolTipText, String.Empty))
        Next
    End Sub
    Private Sub UpdateButtonLayout()
        If _NavigationFlow Is Nothing Then Return
        Dim AvailableWidth As Integer = Math.Max(0, _NavigationFlow.ClientSize.Width - _NavigationFlow.Padding.Horizontal)
        If _NavigationFlow.VerticalScroll.Visible Then AvailableWidth = Math.Max(0, AvailableWidth - SystemInformation.VerticalScrollBarWidth)
        For Each ItemControl As Control In _NavigationFlow.Controls
            Dim Button As NavigationButton = TryCast(ItemControl, NavigationButton)
            If Button Is Nothing Then Continue For
            Button.Margin = New Padding(0, 0, 0, _ButtonSpacing)
            Button.Size = New Size(AvailableWidth, _ButtonHeight)
        Next
    End Sub
    Private Sub UpdateSelectedButton()
        ApplyButtonAppearance()
    End Sub
    Private Sub NavigationFlowLayout(Sender As Object, E As LayoutEventArgs)
        UpdateButtonLayout()
    End Sub
    Private Sub NavigationButtonClick(Sender As Object, E As EventArgs)
        Dim Button As NavigationButton = TryCast(Sender, NavigationButton)
        If Button IsNot Nothing AndAlso Button.Page IsNot Nothing Then Navigate(Button.Page)
    End Sub
    Private Shared Function ResolveAccessibleName(Page As NavigationPage) As String
        If Not String.IsNullOrWhiteSpace(Page.AccessibleName) Then Return Page.AccessibleName
        If Not String.IsNullOrWhiteSpace(Page.Text) Then Return Page.Text
        Return Page.Key
    End Function
End Class
