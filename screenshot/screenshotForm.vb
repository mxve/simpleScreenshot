Imports System.Drawing
Imports System.Windows.Forms

Friend Class screenshotForm
    Inherits Form

    Public bmp As Bitmap
    Public base64 As String

    Private _hint As String = "The green area marks the screenshot." + vbNewLine + vbNewLine + "Press <ENTER> to capture"

    Private _generateBase64 As Boolean = False

    Private lbl As New Label
    Private borderPen As Pen

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
        e.Graphics.DrawRectangle(borderPen, New Rectangle(2, 2, Width - 4, Height - 4))
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
                Me.Close()
        End Select
    End Sub

    Private Sub returnScreenshot(screen As Bitmap)
        Dim conv As New ImageConverter
        If _generateBase64 Then base64 = Convert.ToBase64String(conv.ConvertTo(screen, GetType(Byte())))
        bmp = screen
        DialogResult = DialogResult.OK
    End Sub

    Private Sub setup()
        Dim sFormResizer As New FormResizer(Me, 6)

        Opacity = 0.7
        Size = New Point(450, 300)
        FormBorderStyle = FormBorderStyle.None
        BackColor = Color.LightGreen
        ResizeRedraw = True
        borderPen = New Pen(Color.FromArgb(190, 100, 100, 100), 2)
        borderPen.DashStyle = Drawing2D.DashStyle.Dot

        lbl.Text = _hint
        lbl.AutoSize = True
        centerText()
        Controls.Add(lbl)
    End Sub

    Public Sub New()
        setup()
    End Sub

    Public Sub New(generateBase64 As Boolean)
        _generateBase64 = generateBase64

        setup()
    End Sub

    Public Sub New(generateBase64 As Boolean, hint As String)
        _generateBase64 = generateBase64
        _hint = hint

        setup()
    End Sub
End Class
