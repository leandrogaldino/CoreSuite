Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Drawing
Imports System.Globalization
Imports System.Windows.Forms

Public Class RichTextButton
    Inherits NoFocusCueButton

    <Browsable(False), EditorBrowsable(EditorBrowsableState.Never)>
    <DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)>
    Public Overloads ReadOnly Property Text As String
        Get
            Return String.Join("", TextParts.Select(Function(part) part.Text))
        End Get
    End Property

    <Category("Appearance"),
    Description("Define partes do texto com diferentes estilos." & vbCrLf & "A estilização é aplicada apenas em tempo de execução"),
    Editor(GetType(RichTextCollectionEditor), GetType(System.Drawing.Design.UITypeEditor)),
    TypeConverter(GetType(RichTextConverter)),
    DesignerSerializationVisibility(DesignerSerializationVisibility.Content)>
    Public Property TextParts As New BindingList(Of RichTextPart)

    Public Sub New()
        AddHandler TextParts.ListChanged, Sub()
                                              Invalidate() ' Redesenha sempre que a lista mudar
                                              If DesignMode Then MyBase.Text = Text
                                          End Sub
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        If DesignMode Then Return
        Dim g As Graphics = e.Graphics
        Dim Format As New StringFormat()
        Select Case TextAlign
            Case ContentAlignment.TopLeft, ContentAlignment.MiddleLeft, ContentAlignment.BottomLeft
                Format.Alignment = StringAlignment.Near
            Case ContentAlignment.TopCenter, ContentAlignment.MiddleCenter, ContentAlignment.BottomCenter
                Format.Alignment = StringAlignment.Center
            Case ContentAlignment.TopRight, ContentAlignment.MiddleRight, ContentAlignment.BottomRight
                Format.Alignment = StringAlignment.Far
        End Select
        Select Case TextAlign
            Case ContentAlignment.TopLeft, ContentAlignment.TopCenter, ContentAlignment.TopRight
                Format.LineAlignment = StringAlignment.Near
            Case ContentAlignment.MiddleLeft, ContentAlignment.MiddleCenter, ContentAlignment.MiddleRight
                Format.LineAlignment = StringAlignment.Center
            Case ContentAlignment.BottomLeft, ContentAlignment.BottomCenter, ContentAlignment.BottomRight
                Format.LineAlignment = StringAlignment.Far
        End Select
        Dim TotalWidth As Integer = 0
        Dim TotalHeight As Integer = 0
        Dim PartSizes As New List(Of SizeF)
        For Each Part In TextParts
            Dim partSize = g.MeasureString(Part.Text, Part.Font)
            PartSizes.Add(partSize)
            TotalWidth += CInt(partSize.Width)
            TotalHeight = Math.Max(TotalHeight, CInt(partSize.Height))
        Next
        Dim StartX As Integer
        Dim StartY As Integer
        Select Case Format.Alignment
            Case StringAlignment.Near
                StartX = 5
            Case StringAlignment.Center
                StartX = (Width - TotalWidth) / 2
            Case StringAlignment.Far
                StartX = Width - TotalWidth - 5
        End Select
        Select Case Format.LineAlignment
            Case StringAlignment.Near
                StartY = 5
            Case StringAlignment.Center
                StartY = (Height - TotalHeight) / 2
            Case StringAlignment.Far
                StartY = Height - TotalHeight - 5
        End Select
        For i As Integer = 0 To TextParts.Count - 1
            Dim Part = TextParts(i)
            Dim partSize = PartSizes(i)
            Dim partY = StartY + (TotalHeight - partSize.Height) / 2
            g.DrawString(Part.Text, Part.Font, New SolidBrush(Part.Color), New RectangleF(StartX, partY, partSize.Width, TotalHeight), Format)
            StartX += CInt(partSize.Width)
        Next
    End Sub

    Public Class RichTextPart
        Public Property Text As String = ""
        Public Property Font As Font = New Font("Microsoft Sans Serif", 8.25, FontStyle.Regular)
        Public Property Color As Color = SystemColors.WindowText

        Public Overrides Function ToString() As String
            Return $"{Text} ({Font.Name}, {Font.Size}, {Font.Style})"
        End Function

    End Class

    Public Class RichTextConverter
        Inherits ExpandableObjectConverter

        Public Overrides Function CanConvertTo(context As ITypeDescriptorContext, destinationType As Type) As Boolean
            If destinationType Is GetType(String) Then Return True
            Return MyBase.CanConvertTo(context, destinationType)
        End Function

        Public Overrides Function ConvertTo(context As ITypeDescriptorContext, culture As CultureInfo, value As Object, destinationType As Type) As Object
            If destinationType Is GetType(String) Then
                Return $"({DirectCast(value, BindingList(Of RichTextPart)).Count} partes)"
            End If
            Return MyBase.ConvertTo(context, culture, value, destinationType)
        End Function

    End Class

    Public Class RichTextCollectionEditor
        Inherits CollectionEditor

        Public Sub New(type As Type)
            MyBase.New(type)
        End Sub

        Protected Overrides Function CreateNewItemTypes() As Type()
            Return New Type() {GetType(RichTextPart)}
        End Function

    End Class

End Class