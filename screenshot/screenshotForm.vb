Imports System.Drawing
Imports System.Windows.Forms

Friend Class screenshotForm
    Inherits Form

    Public bmp As Bitmap
    Public base64 As String
    Private _settings As misc.Settings

    Private lbl As New Label

    Private imgConv As New ImageConverter

#Region "move form"
    ' Make form moveable; Visual Studio snippet

    Dim mouseOffset As Point

    Private Sub Me_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MyBase.MouseDown
        mouseOffset = New Point(-e.X, -e.Y)
    End Sub

    Private Sub Me_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles MyBase.MouseMove

        If e.Button = MouseButtons.Left Then
            Dim mousePos = Control.MousePosition
            mousePos.Offset(mouseOffset.X, mouseOffset.Y)
            Location = mousePos
        End If
    End Sub
#End Region

    Private Sub centerText()
        lbl.Top = (Me.Height / 2) - (lbl.Height / 2)
        lbl.Left = (Me.Width / 2) - (lbl.Width / 2)
    End Sub

    Private Sub sForm_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        centerText()



    End Sub

    Private Sub sForm_paint(sender As Object, e As PaintEventArgs) Handles MyBase.Paint
        e.Graphics.DrawRectangle(_settings.Pen, New Rectangle(2, 2, Width - 4, Height - 4))
    End Sub

    Private Function takeAreaScreenshot(location As Point, size As Point) As Bitmap
        Dim bmp As New Bitmap(size.X, size.Y)
        Graphics.FromImage(bmp).CopyFromScreen(location, New Point(0, 0), size)
        Return bmp
    End Function

    Private Sub sForm_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        Select Case e.KeyCode
            ' Submit
            Case Keys.KeyCode.Enter
                Me.Hide()
                returnScreenshot(takeAreaScreenshot(Me.Location, Me.Size))
                Me.Close()
            ' Cancel
            Case Keys.KeyCode.Escape
                DialogResult = DialogResult.Cancel
                Me.Close()
        End Select
    End Sub

    Private Sub returnScreenshot(screen As Bitmap)
        Select Case _settings.Format
            Case misc.screenshotFormat.Bitmap
                bmp = screen
            Case misc.screenshotFormat.Base64
                base64 = Convert.ToBase64String(imgConv.ConvertTo(screen, GetType(Byte())))
        End Select
        DialogResult = DialogResult.OK
    End Sub

    Private Sub setup()
        Dim sFormResizer As New FormResizer(Me, _settings.ResizeWidth)

        Opacity = _settings.Opacity
        Size = _settings.Size
        FormBorderStyle = FormBorderStyle.None
        BackColor = _settings.BackColor
        ResizeRedraw = True
        _settings.Pen.DashStyle = _settings.DashStyle

        lbl.Text = _settings.Hint
        lbl.AutoSize = True
        centerText()
        Controls.Add(lbl)
    End Sub

    Public Sub New()
        Dim settings As misc.Settings
        _settings = settings
        setup()
    End Sub

    Public Sub New(settings As misc.Settings)
        _settings = settings
        setup()
    End Sub
End Class
